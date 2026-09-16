# MIG_SyncSortingOrders.py
"""
MIG_SyncSortingOrders.py
-------------------------
Syncs sort orders from a SOURCE tenant to a DESTINATION tenant.

Steps
-----
1.  SOURCE tenant -> CRS022MI.LstSortOrder -> store in memory.
2.  DEST   tenant -> CRS022MI.LstSortOrder -> store in memory.
3.  Diff: find (PGNM, FILE, SOPT, QTTP) records missing from DEST or that
    differ from SOURCE.
4.  Export records to an EVS100-format Excel file
    (evs100/ToProcess/API_CRS022MI_<timestamp>.xlsx) with a Control sheet,
    an AddSortOrder sheet, and a ChgSortOrder sheet.
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

SKIP_PGNM = frozenset({"CMS100", "LISTMI"})

LST_SELECTED = [
    "PGNM", "QTTP", "TX40", "FILE", "SOPT", "PAV1", "TX15", "NFTR", "PAV2", "PAV3",
    "PAV4", "PAV5", "PAV6", "PSEQ", "TABL", "AGGR", "OBK1", "OBK2", "OBK3", "OBK4",
    "OBK5", "OBK6", "OBK7", "TXID", "SOZ1", "SOZ2", "SOZ3", "SOZ4", "SOZ5", "SOZ6",
]

# QTTP is required because multiple records share the same (PGNM, FILE, SOPT)
# - one per panel view / query type.
KEY_FIELDS = ("PGNM", "FILE", "SOPT", "QTTP")
DATA_FIELDS = [f for f in LST_SELECTED if f not in KEY_FIELDS]

ADD_SHEET = "API_CRS022MI_AddSortOrder"
CHG_SHEET = "API_CRS022MI_ChgSortOrder"


# =============================================================================
# Helpers
# =============================================================================

def norm(v) -> str:
    """Normalise a field value to a stripped string for comparison."""
    return "" if v is None else str(v).strip()


def record_key(rec: dict) -> tuple:
    return (norm(rec.get("PGNM")), norm(rec.get("FILE")), norm(rec.get("SOPT")), norm(rec.get("QTTP")))


def records_equal(a: dict, b: dict) -> bool:
    """Compare two records on all LST_SELECTED fields (normalised)."""
    return all(norm(a.get(f)) == norm(b.get(f)) for f in LST_SELECTED)


def diff_fields(source: dict, dest: dict) -> list[str]:
    """List of changed fields showing src vs dest values (normalised)."""
    diffs = []
    for f in LST_SELECTED:
        sv, dv = norm(source.get(f)), norm(dest.get(f))
        if sv != dv:
            diffs.append(f"{f}  src={sv!r}  dest={dv!r}")
    return diffs


def build_payload_record(rec: dict) -> dict:
    """Record dict with blank fields dropped (never send empty tags)."""
    return {k: v for k, v in rec.items() if norm(v)}


# =============================================================================
# Step 1 / 2 - LstSortOrder
# =============================================================================

def list_all_sort_orders(tenant: MIG_Api.Tenant, session: requests.Session,
                          label: str = "", pgnm_filter: str = "") -> list[dict]:
    """
    CRS022MI.LstSortOrder. Pass pgnm_filter to restrict to a single program;
    leave blank for all. Uses maxrecs=10000 to retrieve the full set in one
    request.
    """
    url = tenant.api_url.replace("maxrecs=0", "maxrecs=10000")
    record: dict = {}
    if pgnm_filter:
        record["PGNM"] = pgnm_filter

    payload = {
        "program": "CRS022MI",
        "transactions": [{"transaction": "LstSortOrder", "record": record,
                           "selectedColumns": LST_SELECTED}],
    }

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
            records.append(record)
    return [r for r in records if r.get("PGNM", "") not in SKIP_PGNM]


# =============================================================================
# Step 3 - Diff
# =============================================================================

def diff_sort_orders(source_all: list[dict], dest_all: list[dict]) -> dict:
    """
    Find records missing from DEST (to_add) or present but differing from
    SOURCE (to_chg). Returns a structured plan grouped by PGNM for display,
    each entry carrying the CLI's own diff_fields() message list.
    """
    dest_index: dict[tuple, dict] = {record_key(r): r for r in dest_all}

    to_add: list[dict] = []
    to_chg: list[tuple[dict, dict]] = []
    for rec in source_all:
        dest_rec = dest_index.get(record_key(rec))
        if dest_rec is None:
            to_add.append(rec)
        elif not records_equal(rec, dest_rec):
            to_chg.append((rec, dest_rec))

    by_pgnm: dict[str, list[dict]] = defaultdict(list)
    for rec in to_add:
        by_pgnm[rec.get("PGNM", "?")].append(
            {"OP": "Add", "FILE": rec.get("FILE", ""), "SOPT": rec.get("SOPT", ""),
             "VIEW": rec.get("PAV1", ""), "CHANGES": []})
    for src, dst in to_chg:
        by_pgnm[src.get("PGNM", "?")].append(
            {"OP": "Chg", "FILE": src.get("FILE", ""), "SOPT": src.get("SOPT", ""),
             "VIEW": src.get("PAV1", ""), "CHANGES": diff_fields(src, dst)})

    return {
        "to_add": to_add, "to_chg": to_chg,
        "by_pgnm": {k: v for k, v in by_pgnm.items()},
        "counts": {"source": len(source_all), "dest": len(dest_all),
                   "to_add": len(to_add), "to_chg": len(to_chg), "programs": len(by_pgnm)},
    }


# =============================================================================
# Excel export (EVS100 format)
# =============================================================================

def _write_sheet(wb: xlsxwriter.Workbook, sheet_name: str, fields: list[str],
                  records: list[dict]) -> None:
    """Write a single EVS100 data sheet with header rows and data rows."""
    ws = wb.add_worksheet(sheet_name)
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
            if norm(val):
                ws.write(row_idx, col, str(val))


def export_evs100_xlsx(add_records: list[dict], chg_records: list[dict],
                        out_dir: Path) -> Path:
    """
    Writes sort-order records in the EVS100 import format. Sheets included
    (only when non-empty): AddSortOrder (missing from DEST), ChgSortOrder
    (differs from SOURCE).
    """
    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = out_dir / f"API_CRS022MI_{ts}.xlsx"

    wb = xlsxwriter.Workbook(str(out_path))

    ws_ctrl = wb.add_worksheet("Control")
    for col, val in enumerate(["Worksheet", "Description", "Data"]):
        ws_ctrl.write(0, col, val)

    ctrl_row = 1
    sheets: list[tuple[str, str, list[dict]]] = [
        (ADD_SHEET, "Add Sort Order", add_records),
        (CHG_SHEET, "Change Sort Order", chg_records),
    ]
    for sheet_name, description, records in sheets:
        if records:
            ws_ctrl.write(ctrl_row, 0, sheet_name)
            ws_ctrl.write(ctrl_row, 1, description)
            ws_ctrl.write(ctrl_row, 2, "x")
            ctrl_row += 1

    for sheet_name, _, records in sheets:
        if records:
            _write_sheet(wb, sheet_name, LST_SELECTED, records)

    wb.close()
    return out_path
