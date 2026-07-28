# MIG_SyncSortingOptions.py
"""
MIG_SyncSortingOptions.py
-------------------
Syncs sort options from a SOURCE tenant to a DESTINATION tenant.

Steps
-----
1.  Ask for SOURCE tenant → CRS021MI.LstSrtOpt → store in memory.
2.  Ask for DEST   tenant → CRS021MI.LstSrtOpt → store in memory.
3.  Diff: find (FI01, SOPT) pairs present in SOURCE but missing from DEST.
4.  Display every FILE to be processed and prompt to export to Excel.
5.  Export records to an EVS100-format Excel file
    (evs100/ToProcess/API_CRS021MI_<timestamp>.xlsx) with a Control sheet
    and data sheets per operation type.
6.  Optionally upload the file to DEST via the File Management REST API (PUT).
7.  Optionally trigger processing via EVS100MI.ImportFile.

Usage:
    python MIG_SyncSortingOptions.py
"""

import datetime
import requests
from collections import defaultdict
from pathlib import Path
import xlsxwriter

from InforMI import (
    CONFIG,
    get_ion_token,
    post_to_m3,
    select_ionapi_file,
    prompt_for_company_division,
)
from UserDefaults import save_user_defaults

# =============================================================================
# Constants
# =============================================================================

VALID_SOPT_PREFIXES = frozenset("UVX")

ADD_SHEET = "API_CRS021MI_AddSrtOpt"
ACT_SHEET = "API_CRS021MI_ActSrtOpt"
STD_SHEET = "API_CRS021MI_CrtStdSrtOpt"

EVS100_TO_PROCESS = Path(__file__).parent / "evs100" / "ToProcess"


# =============================================================================
# Helpers
# =============================================================================

def is_valid_sopt(sopt: str) -> bool:
    """Return True if SOPT is in the ranges U1-U9, V1-V9, or X1-X9."""
    return (
        len(sopt) == 2
        and sopt[0] in VALID_SOPT_PREFIXES
        and sopt[1] in "123456789"
    )


def snapshot_config() -> dict:
    return dict(CONFIG)


def restore_config(snapshot: dict) -> None:
    CONFIG.clear()
    CONFIG.update(snapshot)


def setup_tenant(ionapi_dir: Path, label: str) -> dict:
    """Prompt for an ionapi file + company/division, authenticate, return snapshot."""
    print(f"\n{'─' * 60}")
    print(f"  📡  Configure {label} tenant")
    print(f"{'─' * 60}")
    input(f"  Press Enter to select the {label} ionapi file …")
    CONFIG["tenant"] = select_ionapi_file(ionapi_dir)
    prompt_for_company_division()
    get_ion_token()
    snap = snapshot_config()
    print(f"  ✅  {label} ready: {CONFIG.get('tenant', '')}  "
          f"CONO={CONFIG.get('company', '')}  DIVI={CONFIG.get('division', '')}")
    return snap


def normalise_record(record: dict) -> dict:
    """Ensure the file-name field is keyed as FI01, truncated to 6 chars."""
    rec = dict(record)
    if "FILE" in rec and "FI01" not in rec:
        rec["FI01"] = rec.pop("FILE")[:6]
    elif "FI01" in rec:
        rec["FI01"] = rec["FI01"][:6]
    return rec


# =============================================================================
# Step 1 / 2 – LstSrtOpt (full list, no filter)
# =============================================================================

def list_all_srt_opts(session: requests.Session, label: str = "") -> list[dict]:
    """
    Calls CRS021MI.LstSrtOpt with no key filter and returns every record.
    Uses maxrecs=10000 to retrieve the full set in one request.
    """
    list_url = CONFIG["api_url"].replace("maxrecs=0", "maxrecs=10000")

    payload = {
        "program": "CRS021MI",
        "transactions": [{"transaction": "LstSrtOpt", "record": {}}],
    }

    headers = {
        "Authorization": f"Bearer {CONFIG['access_token']}",
        "Content-Type": "application/json",
    }

    response = session.post(list_url, json=payload, headers=headers)
    if response.status_code == 401:
        get_ion_token()
        headers["Authorization"] = f"Bearer {CONFIG['access_token']}"
        response = session.post(list_url, json=payload, headers=headers)

    response.raise_for_status()
    data = response.json()

    records: list[dict] = []
    for result in data.get("results", []):
        for record in result.get("records", []):
            records.append(normalise_record(record))

    tag = f" ({label})" if label else ""
    print(f"  ✅  {len(records)} sort option records retrieved{tag}.")
    return records


# =============================================================================
# Upload and process (EVS100 pattern)
# =============================================================================

def upload_file_to_m3(
    file_path: Path,
    session: requests.Session,
) -> bool:
    """
    Uploads file_path to the M3 FileImport area via the File Management REST API.

    PUT {iu}/{ti}/M3/foundation-rest/file-management/v1/file/FileImport/{filename}

    Returns True on success, False on failure.
    """
    filename   = file_path.name
    upload_url = (
        f"{CONFIG['iu']}/{CONFIG['ti']}"
        f"/M3/foundation-rest/file-management/v1/file/FileImport/{filename}"
    )

    headers = {
        "Authorization": f"Bearer {CONFIG['access_token']}",
        "Content-Type": "application/octet-stream",
    }

    with open(file_path, "rb") as fh:
        file_bytes = fh.read()

    response = session.put(upload_url, data=file_bytes, headers=headers)

    if response.status_code == 401:
        get_ion_token()
        headers["Authorization"] = f"Bearer {CONFIG['access_token']}"
        response = session.put(upload_url, data=file_bytes, headers=headers)

    if response.status_code in (200, 201, 204):
        print(f"  ✅  Upload succeeded (HTTP {response.status_code}): {filename}")
        return True
    else:
        print(
            f"  ❌  Upload failed (HTTP {response.status_code}): "
            f"{response.text[:200]}"
        )
        return False


def process_file_in_m3(
    filename: str,
    session: requests.Session,
) -> bool:
    """
    Calls EVS100MI.ImportFile with FNAM=filename so M3 processes the uploaded
    spreadsheet through the EVS100 interface.

    Returns True on success, False on failure.
    """
    payload = {
        "program": "EVS100MI",
        "transactions": [{
            "transaction": "ImportFile",
            "record": {"FNAM": filename},
        }],
    }

    try:
        result = post_to_m3(payload, session)
        for res in result.get("results", []):
            err = res.get("errorMessage", "").strip() if isinstance(res, dict) else ""
            if err:
                print(f"  ❌  EVS100MI.ImportFile error: {err}")
                return False
        print(f"  ✅  EVS100MI.ImportFile processed successfully: {filename}")
        return True
    except Exception as exc:
        print(f"  ❌  EVS100MI.ImportFile failed: {exc}")
        return False


# =============================================================================
# Excel export (EVS100 format)
# =============================================================================

def _ordered_keys(records: list[dict]) -> list[str]:
    """Return unique field names in insertion order across all records."""
    seen: list[str] = []
    for rec in records:
        for k in rec:
            if k not in seen:
                seen.append(k)
    return seen


def _write_sheet(
    wb: xlsxwriter.Workbook,
    sheet_name: str,
    records: list[dict],
) -> None:
    """Write a single EVS100 data sheet with header rows and data rows."""
    ws = wb.add_worksheet(sheet_name)
    fields = _ordered_keys(records)
    cols   = ["MESSAGE"] + fields

    # Row 1 – field names
    for col, name in enumerate(cols):
        ws.write(0, col, name)
    # Row 2 – descriptions
    for col, name in enumerate(cols):
        ws.write(1, col, "" if name == "MESSAGE" else name)
    # Row 3 – required
    for col in range(len(cols)):
        ws.write(2, col, "no" if col == 0 else "yes")
    # Data rows
    for row_idx, rec in enumerate(records, start=3):
        for col, field in enumerate(fields, start=1):
            val = rec.get(field, "")
            if val:
                ws.write(row_idx, col, str(val))


def export_evs100_xlsx(
    add_records: list[dict],
    act_records: list[dict],
    std_records: list[dict],
    out_dir: Path,
) -> Path:
    """
    Writes sort-option records in the EVS100 import format.

    Sheets included (only when non-empty):
      API_CRS021MI_AddSrtOpt     – custom sort options (U/V/X range)
      API_CRS021MI_ActSrtOpt     – activate the above
      API_CRS021MI_CrtStdSrtOpt  – standard sort options

    Returns the path of the written file.
    """
    ts       = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    out_path = out_dir / f"API_CRS021MI_{ts}.xlsx"

    wb = xlsxwriter.Workbook(str(out_path))

    # ── Control sheet ─────────────────────────────────────────────────────
    ws_ctrl = wb.add_worksheet("Control")
    for col, val in enumerate(["Worksheet", "Description", "Data"]):
        ws_ctrl.write(0, col, val)

    ctrl_row = 1
    sheets: list[tuple[str, str, list[dict]]] = [
        (ADD_SHEET, "Add Sort Option",             add_records),
        (ACT_SHEET, "Activate Sort Option",        act_records),
        (STD_SHEET, "Create Standard Sort Option", std_records),
    ]
    for sheet_name, description, records in sheets:
        if records:
            ws_ctrl.write(ctrl_row, 0, sheet_name)
            ws_ctrl.write(ctrl_row, 1, description)
            ws_ctrl.write(ctrl_row, 2, "x")
            ctrl_row += 1

    # ── Data sheets ────────────────────────────────────────────────────────
    for sheet_name, _, records in sheets:
        if records:
            _write_sheet(wb, sheet_name, records)

    wb.close()
    return out_path


# =============================================================================
# Main
# =============================================================================

def sync_view_file_sort() -> None:
    ionapi_dir = Path(__file__).parent / "ionapi"

    # ------------------------------------------------------------------ #
    # Step 1 – SOURCE: LstSrtOpt                                          #
    # ------------------------------------------------------------------ #
    save_user_defaults({})

    source_snap = setup_tenant(ionapi_dir, "SOURCE")
    print("\n🔍  Step 1 – CRS021MI.LstSrtOpt (SOURCE) …")
    with requests.Session() as session:
        source_all = list_all_srt_opts(session, label="SOURCE")

    if not source_all:
        print("⚠️   No sort options found in SOURCE. Exiting.")
        return

    # ------------------------------------------------------------------ #
    # Step 2 – DEST: LstSrtOpt                                            #
    # ------------------------------------------------------------------ #
    save_user_defaults({})

    dest_snap = setup_tenant(ionapi_dir, "DEST")
    print("\n🔍  Step 2 – CRS021MI.LstSrtOpt (DEST) …")
    with requests.Session() as session:
        dest_all = list_all_srt_opts(session, label="DEST")

    # ------------------------------------------------------------------ #
    # Step 3 – Diff: missing in DEST                                      #
    # ------------------------------------------------------------------ #
    dest_keys = {(r["FI01"], r["SOPT"]) for r in dest_all}
    missing = [
        r for r in source_all
        if (r["FI01"], r["SOPT"]) not in dest_keys
        and r.get("SOPT", "").strip() != "JD"
    ]

    if not missing:
        print("\n✅  DEST already has all sort options from SOURCE. Nothing to do.")
        return

    # Group by FI01 for display
    by_file: dict[str, list[dict]] = defaultdict(list)
    for r in missing:
        by_file[r["FI01"]].append(r)

    # ------------------------------------------------------------------ #
    # Step 4a – Display                                                   #
    # ------------------------------------------------------------------ #
    print(f"\n{'═' * 60}")
    print(f"  📋  {len(missing)} sort option(s) missing from DEST across "
          f"{len(by_file)} file(s)")
    print(f"{'═' * 60}")

    for fi01 in sorted(by_file.keys()):
        records = by_file[fi01]
        valid   = [r for r in records if is_valid_sopt(r.get("SOPT", ""))]
        std     = [r for r in records if not is_valid_sopt(r.get("SOPT", ""))]
        print(f"\n  FILE: {fi01}")
        for r in sorted(valid, key=lambda x: x.get("SOPT", "")):
            print(f"    SOPT={r.get('SOPT',''):<4}  →  AddSrtOpt + ActSrtOpt")
        if std:
            sopts = ", ".join(r.get("SOPT", "") for r in sorted(std, key=lambda x: x.get("SOPT", "")))
            print(f"    SOPT={sopts}  →  CrtStdSrtOpt  (runs once for this file)")

    print(f"\n{'═' * 60}")
    confirm = input(f"  Export these {len(missing)} record(s) to Excel? [y/N]: ").strip().lower()
    if confirm != "y":
        print("⛔  Aborted — no file created.")
        return
    print()

    # ------------------------------------------------------------------ #
    # Step 4b – Build record sets                                         #
    # ------------------------------------------------------------------ #
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

    # ------------------------------------------------------------------ #
    # Step 5 – Export to EVS100 Excel file                               #
    # ------------------------------------------------------------------ #
    print(f"▶   Step 5 – Exporting {len(missing)} record(s) to EVS100 Excel file …")
    xlsx_path = export_evs100_xlsx(add_records, act_records, std_records, EVS100_TO_PROCESS)
    print(f"  📄  Saved to: evs100/ToProcess/{xlsx_path.name}")

    # ------------------------------------------------------------------ #
    # Summary                                                              #
    # ------------------------------------------------------------------ #
    print(f"\n{'═' * 60}")
    print(f"  MIG_SyncSortingOptions – Run Summary")
    print(f"{'═' * 60}")
    print(f"  📋  Records in SOURCE      : {len(source_all):,}")
    print(f"  📋  Missing in DEST        : {len(missing):,}")
    print(f"  📄  AddSrtOpt records      : {len(add_records):,}")
    print(f"  📄  ActSrtOpt records      : {len(act_records):,}")
    print(f"  📄  CrtStdSrtOpt records   : {len(std_records):,}")
    print(f"  📄  Excel file             : {xlsx_path.name}")
    print(f"{'═' * 60}")

    # ------------------------------------------------------------------ #
    # Step 6 – Optional: upload and process in DEST                      #
    # ------------------------------------------------------------------ #
    restore_config(dest_snap)
    get_ion_token()

    send = input("\n  Send this file to M3? [y/N]: ").strip().lower()
    if send != "y":
        print("  ℹ️   File not sent. Done.")
        return

    with requests.Session() as session:
        print(f"\n▶   Step 6a – Uploading {xlsx_path.name} to M3 …")
        uploaded = upload_file_to_m3(xlsx_path, session)

        if uploaded:
            trigger = input(
                "\n  Process file in M3 (EVS100MI.ImportFile)? [y/N]: "
            ).strip().lower()
            if trigger == "y":
                print(f"\n▶   Step 6b – Triggering EVS100MI.ImportFile …")
                process_file_in_m3(xlsx_path.name, session)
            else:
                print("  ℹ️   File uploaded but not processed. Done.")


# =============================================================================
# Entry Point
# =============================================================================

if __name__ == "__main__":
    sync_view_file_sort()
