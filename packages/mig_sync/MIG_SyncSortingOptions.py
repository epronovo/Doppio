# MIG_SyncSortingOptions.py
"""
MIG_SyncSortingOptions.py
--------------------------
Syncs sort options from a SOURCE tenant to a DESTINATION tenant.

Steps
-----
1.  SOURCE tenant -> CRS021MI.LstSrtOpt -> store in memory.
2.  DEST   tenant -> CRS021MI.LstSrtOpt -> store in memory.
3.  Diff: find (FI01, SOPT) pairs present in SOURCE but missing from DEST.
4.  Export records to an EVS100-format Excel file
    (evs100/ToProcess/API_CRS021MI_<timestamp>.xlsx) with a Control sheet
    and data sheets per operation type.
5.  Optionally upload the file to DEST via the File Management REST API (PUT).
6.  Optionally trigger processing via EVS100MI.ImportFile.

Ported to a library used by MIG_App.py: input()/print() are gone and tenant
auth is an explicit MIG_Api.Tenant instead of the InforMI CONFIG global.
"""

from __future__ import annotations

import datetime
from collections import defaultdict
from pathlib import Path

import requests
import xlsxwriter

import MIG_Api

# =============================================================================
# Constants
# =============================================================================

VALID_SOPT_PREFIXES = frozenset("UVX")

ADD_SHEET = "API_CRS021MI_AddSrtOpt"
ACT_SHEET = "API_CRS021MI_ActSrtOpt"
STD_SHEET = "API_CRS021MI_CrtStdSrtOpt"


# =============================================================================
# Helpers
# =============================================================================

def is_valid_sopt(sopt: str) -> bool:
    """True if SOPT is in the ranges U1-U9, V1-V9, or X1-X9."""
    return len(sopt) == 2 and sopt[0] in VALID_SOPT_PREFIXES and sopt[1] in "123456789"


def normalise_record(record: dict) -> dict:
    """Ensure the file-name field is keyed as FI01, truncated to 6 chars."""
    rec = dict(record)
    if "FILE" in rec and "FI01" not in rec:
        rec["FI01"] = rec.pop("FILE")[:6]
    elif "FI01" in rec:
        rec["FI01"] = rec["FI01"][:6]
    return rec


# =============================================================================
# Step 1 / 2 - LstSrtOpt (full list, no filter)
# =============================================================================

def list_all_srt_opts(tenant: MIG_Api.Tenant, session: requests.Session,
                       label: str = "") -> list[dict]:
    """
    CRS021MI.LstSrtOpt with no key filter. Uses maxrecs=10000 to retrieve the
    full set in one request.
    """
    url = tenant.api_url.replace("maxrecs=0", "maxrecs=10000")
    payload = {"program": "CRS021MI", "transactions": [{"transaction": "LstSrtOpt", "record": {}}]}

    # post_to_m3() always targets tenant.api_url, so borrow the maxrecs=10000
    # variant for just this one call.
    headers = {"Authorization": f"Bearer {tenant.access_token}", "Content-Type": "application/json"}
    resp = session.post(url, json=payload, headers=headers, timeout=MIG_Api.DEFAULT_TIMEOUT)
    if resp.status_code == 401:
        MIG_Api.authenticate(tenant, session)
        headers["Authorization"] = f"Bearer {tenant.access_token}"
        resp = session.post(url, json=payload, headers=headers, timeout=MIG_Api.DEFAULT_TIMEOUT)
    resp.raise_for_status()
    data = resp.json()

    records: list[dict] = []
    for result in data.get("results", []):
        for record in result.get("records", []):
            records.append(normalise_record(record))
    return records


# =============================================================================
# Step 3 - Diff
# =============================================================================

def diff_sorting_options(source_all: list[dict], dest_all: list[dict]) -> dict:
    """
    Find (FI01, SOPT) pairs present in SOURCE but missing from DEST (SOPT 'JD'
    is always excluded - it is M3's own default, never migrated).

    Returns {"missing": [...], "by_file": {fi01: [...]}, "counts": {...}}.
    """
    dest_keys = {(r["FI01"], r["SOPT"]) for r in dest_all}
    missing = [r for r in source_all
               if (r["FI01"], r["SOPT"]) not in dest_keys
               and (r.get("SOPT", "") or "").strip() != "JD"]

    by_file: dict[str, list[dict]] = defaultdict(list)
    for r in missing:
        by_file[r["FI01"]].append(r)

    return {
        "missing": missing,
        "by_file": {k: v for k, v in by_file.items()},
        "counts": {"source": len(source_all), "dest": len(dest_all), "missing": len(missing),
                   "files": len(by_file)},
    }


# =============================================================================
# Excel export (EVS100 format)
# =============================================================================

def _ordered_keys(records: list[dict]) -> list[str]:
    """Unique field names in insertion order across all records."""
    seen: list[str] = []
    for rec in records:
        for k in rec:
            if k not in seen:
                seen.append(k)
    return seen


def _write_sheet(wb: xlsxwriter.Workbook, sheet_name: str, records: list[dict]) -> None:
    """Write a single EVS100 data sheet with header rows and data rows."""
    ws = wb.add_worksheet(sheet_name)
    fields = _ordered_keys(records)
    cols = ["MESSAGE"] + fields

    for col, name in enumerate(cols):
        ws.write(0, col, name)
    for col, name in enumerate(cols):
        ws.write(1, col, "" if name == "MESSAGE" else name)
    for col in range(len(cols)):
        ws.write(2, col, "no" if col == 0 else "yes")
    for row_idx, rec in enumerate(records, start=3):
        for col, field_name in enumerate(fields, start=1):
            val = rec.get(field_name, "")
            if val:
                ws.write(row_idx, col, str(val))


def build_record_sets(missing: list[dict]) -> tuple[list[dict], list[dict], list[dict]]:
    """
    Split the missing records into the three EVS100 sheets:
      add_records - custom sort options (U/V/X range) -> AddSrtOpt
      act_records - the same, for ActSrtOpt (FILE + SOPT only)
      std_records - one row per file needing a standard sort option (CrtStdSrtOpt)
    """
    add_records = [
        {**{k: v for k, v in r.items() if k != "FI01"}, "FILE": r["FI01"]}
        for r in missing if is_valid_sopt(r.get("SOPT", ""))
    ]
    act_records = [{"FILE": r["FILE"], "SOPT": r["SOPT"]} for r in add_records]

    seen_files: set[str] = set()
    std_records: list[dict] = []
    for r in missing:
        if not is_valid_sopt(r.get("SOPT", "")) and r["FI01"] not in seen_files:
            seen_files.add(r["FI01"])
            std_records.append({"FI01": r["FI01"]})

    return add_records, act_records, std_records


def export_evs100_xlsx(add_records: list[dict], act_records: list[dict],
                        std_records: list[dict], out_dir: Path) -> Path:
    """
    Writes sort-option records in the EVS100 import format. Sheets included
    (only when non-empty): AddSrtOpt, ActSrtOpt, CrtStdSrtOpt.
    """
    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = out_dir / f"API_CRS021MI_{ts}.xlsx"

    wb = xlsxwriter.Workbook(str(out_path))

    ws_ctrl = wb.add_worksheet("Control")
    for col, val in enumerate(["Worksheet", "Description", "Data"]):
        ws_ctrl.write(0, col, val)

    ctrl_row = 1
    sheets: list[tuple[str, str, list[dict]]] = [
        (ADD_SHEET, "Add Sort Option", add_records),
        (ACT_SHEET, "Activate Sort Option", act_records),
        (STD_SHEET, "Create Standard Sort Option", std_records),
    ]
    for sheet_name, description, records in sheets:
        if records:
            ws_ctrl.write(ctrl_row, 0, sheet_name)
            ws_ctrl.write(ctrl_row, 1, description)
            ws_ctrl.write(ctrl_row, 2, "x")
            ctrl_row += 1

    for sheet_name, _, records in sheets:
        if records:
            _write_sheet(wb, sheet_name, records)

    wb.close()
    return out_path
