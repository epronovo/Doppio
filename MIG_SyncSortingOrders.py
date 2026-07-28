# MIG_SyncSortingOrders.py
"""
MIG_SyncSortingOrders.py
------------------------
Syncs sort orders from a SOURCE tenant to a DESTINATION tenant.

Steps
-----
1.  Ask for SOURCE tenant → CRS022MI.LstSortOrder → store in memory.
2.  Ask for DEST   tenant → CRS022MI.LstSortOrder → store in memory.
3.  Diff: find (PGNM, FILE, SOPT) records that are missing from DEST or differ
    from SOURCE.  Display each affected PGNM with Add / Chg label and prompt
    to export.
4.  Export records to an EVS100-format Excel file
    (evs100/ToProcess/API_CRS022MI_<timestamp>.xlsx) with a Control sheet,
    an AddSortOrder sheet, and a ChgSortOrder sheet.
5.  Optionally upload the file to DEST via the File Management REST API (PUT).
6.  Optionally trigger processing via EVS100MI.ImportFile.

Usage:
    python MIG_SyncSortingOrders.py
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

SKIP_PGNM = frozenset({"CMS100", "LISTMI"})

LST_SELECTED = [
    "PGNM","QTTP","TX40","FILE","SOPT","PAV1","TX15","NFTR","PAV2","PAV3",
    "PAV4","PAV5","PAV6","PSEQ","TABL","AGGR","OBK1","OBK2","OBK3","OBK4",
    "OBK5","OBK6","OBK7","TXID","SOZ1","SOZ2","SOZ3","SOZ4","SOZ5","SOZ6",
    # "RESP","USEQ","UPV1","UPV2","UPV3","UPV4","UPV5","UPV6","IBCA","SNDI",
    # "CHNF","CHAG","AGRG","MXRE","JNSO","SLF1","SLF2","SLF3","SUB1","SUB2",
    # "SUB3","HLAL","SBFD","SDSN","CON1","LUFN","LUVE",
]

# Key fields that identify a unique sort order record.
# QTTP is required because multiple records share the same (PGNM, FILE, SOPT)
# — one per panel view / query type.
KEY_FIELDS = ("PGNM", "FILE", "SOPT", "QTTP")

# Data fields sent on Add / Chg (all non-key fields from LstSortOrder)
DATA_FIELDS = [f for f in LST_SELECTED if f not in KEY_FIELDS]

ADD_SHEET = "API_CRS022MI_AddSortOrder"
CHG_SHEET = "API_CRS022MI_ChgSortOrder"

EVS100_TO_PROCESS = Path(__file__).parent / "evs100" / "ToProcess"


# =============================================================================
# Helpers
# =============================================================================

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


def norm(v) -> str:
    """Normalise a field value to a stripped string for comparison."""
    if v is None:
        return ""
    return str(v).strip()


def record_key(rec: dict) -> tuple:
    return (norm(rec.get("PGNM")), norm(rec.get("FILE")), norm(rec.get("SOPT")), norm(rec.get("QTTP")))


def records_equal(a: dict, b: dict) -> bool:
    """Compare two records on all LST_SELECTED fields (normalised)."""
    return all(norm(a.get(f)) == norm(b.get(f)) for f in LST_SELECTED)


def diff_fields(source: dict, dest: dict) -> list[str]:
    """Return a list of changed fields showing src vs dest values (normalised)."""
    diffs = []
    for f in LST_SELECTED:
        sv = norm(source.get(f))
        dv = norm(dest.get(f))
        if sv != dv:
            diffs.append(f"{f}  src={sv!r}  dest={dv!r}")
    return diffs


def build_payload_record(rec: dict) -> dict:
    """Return a record dict with blank fields dropped (never send empty tags)."""
    return {k: v for k, v in rec.items() if norm(v)}


# =============================================================================
# Step 1 / 2 – LstSortOrder
# =============================================================================

def list_all_sort_orders(session: requests.Session, label: str = "") -> list[dict]:
    """
    Calls CRS022MI.LstSortOrder with no filter and returns every record.
    Uses maxrecs=10000 to retrieve the full set in one request.
    """
    list_url = CONFIG["api_url"].replace("maxrecs=0", "maxrecs=10000")

    payload = {
        "program": "CRS022MI",
        "transactions": [{
            "transaction": "LstSortOrder",
            "record": {},
            "selectedColumns": LST_SELECTED,
        }],
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
            records.append(record)

    tag = f" ({label})" if label else ""
    print(f"  ✅  {len(records)} sort order records retrieved{tag}.")
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

def _write_sheet(
    wb: xlsxwriter.Workbook,
    sheet_name: str,
    fields: list[str],
    records: list[dict],
) -> None:
    """Write a single EVS100 data sheet with header rows and data rows."""
    ws   = wb.add_worksheet(sheet_name)
    cols = ["MESSAGE"] + fields

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
            if norm(val):
                ws.write(row_idx, col, str(val))


def export_evs100_xlsx(
    add_records: list[dict],
    chg_records: list[dict],
    out_dir: Path,
) -> Path:
    """
    Writes sort-order records in the EVS100 import format.

    Sheets included (only when non-empty):
      API_CRS022MI_AddSortOrder  – records missing from DEST
      API_CRS022MI_ChgSortOrder  – records that differ from SOURCE

    Returns the path of the written file.
    """
    ts       = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    out_path = out_dir / f"API_CRS022MI_{ts}.xlsx"

    wb = xlsxwriter.Workbook(str(out_path))

    # ── Control sheet ─────────────────────────────────────────────────────
    ws_ctrl = wb.add_worksheet("Control")
    for col, val in enumerate(["Worksheet", "Description", "Data"]):
        ws_ctrl.write(0, col, val)

    ctrl_row = 1
    sheets: list[tuple[str, str, list[dict]]] = [
        (ADD_SHEET, "Add Sort Order",    add_records),
        (CHG_SHEET, "Change Sort Order", chg_records),
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
            _write_sheet(wb, sheet_name, LST_SELECTED, records)

    wb.close()
    return out_path


# =============================================================================
# Main
# =============================================================================

def sync_sort_orders() -> None:
    ionapi_dir = Path(__file__).parent / "ionapi"

    # ------------------------------------------------------------------ #
    # Step 1 – SOURCE: LstSortOrder                                       #
    # ------------------------------------------------------------------ #
    save_user_defaults({})

    source_snap = setup_tenant(ionapi_dir, "SOURCE")
    print("\n🔍  Step 1 – CRS022MI.LstSortOrder (SOURCE) …")
    with requests.Session() as session:
        source_all = list_all_sort_orders(session, label="SOURCE")
    source_all = [r for r in source_all if r.get("PGNM", "") not in SKIP_PGNM]

    if not source_all:
        print("⚠️   No sort orders found in SOURCE. Exiting.")
        return

    # ------------------------------------------------------------------ #
    # Step 2 – DEST: LstSortOrder                                         #
    # ------------------------------------------------------------------ #
    save_user_defaults({})

    dest_snap = setup_tenant(ionapi_dir, "DEST")
    print("\n🔍  Step 2 – CRS022MI.LstSortOrder (DEST) …")
    with requests.Session() as session:
        dest_all = list_all_sort_orders(session, label="DEST")
    dest_all = [r for r in dest_all if r.get("PGNM", "") not in SKIP_PGNM]

    # ------------------------------------------------------------------ #
    # Step 3 – Diff: missing or changed                                   #
    # ------------------------------------------------------------------ #
    dest_index: dict[tuple, dict] = {record_key(r): r for r in dest_all}

    to_add: list[dict] = []
    to_chg: list[tuple[dict, dict]] = []   # (source_rec, dest_rec)

    for rec in source_all:
        key = record_key(rec)
        dest_rec = dest_index.get(key)
        if dest_rec is None:
            to_add.append(rec)
        elif not records_equal(rec, dest_rec):
            to_chg.append((rec, dest_rec))

    if not to_add and not to_chg:
        print("\n✅  DEST sort orders already match SOURCE. Nothing to do.")
        return

    # Group by PGNM for display
    by_pgnm: dict[str, list] = defaultdict(list)
    for rec in to_add:
        by_pgnm[rec.get("PGNM", "?")].append(("Add", rec, None))
    for src, dst in to_chg:
        by_pgnm[src.get("PGNM", "?")].append(("Chg", src, dst))

    print(f"\n{'═' * 60}")
    print(f"  📋  {len(to_add)} to Add, {len(to_chg)} to Chg  "
          f"across {len(by_pgnm)} program(s)")
    print(f"{'═' * 60}")

    for pgnm in sorted(by_pgnm.keys()):
        entries = by_pgnm[pgnm]
        adds = [(op, r, d) for op, r, d in entries if op == "Add"]
        chgs = [(op, r, d) for op, r, d in entries if op == "Chg"]
        print(f"\n  PGNM: {pgnm}")
        for _, r, _ in sorted(adds, key=lambda x: (x[1].get("FILE",""), x[1].get("SOPT",""))):
            print(f"    Add  FILE={r.get('FILE',''):<8}  SOPT={r.get('SOPT',''):<4}  VIEW={r.get('PAV1','')}")
        for _, r, d in sorted(chgs, key=lambda x: (x[1].get("FILE",""), x[1].get("SOPT",""))):
            changes = diff_fields(r, d)
            print(f"    Chg  FILE={r.get('FILE',''):<8}  SOPT={r.get('SOPT',''):<4}  VIEW={r.get('PAV1','')}")
            for change in changes:
                print(f"           {change}")

    print(f"\n{'═' * 60}")
    confirm = input(
        f"  Export {len(to_add)} Add + {len(to_chg)} Chg record(s) to Excel? [y/N]: "
    ).strip().lower()
    if confirm != "y":
        print("⛔  Aborted — no file created.")
        return
    print()

    # ------------------------------------------------------------------ #
    # Step 4 – Export to EVS100 Excel file                               #
    # ------------------------------------------------------------------ #
    add_records = [build_payload_record(r) for r in to_add]
    chg_records = [build_payload_record(src) for src, _ in to_chg]

    print(f"▶   Step 4 – Exporting to EVS100 Excel file …")
    xlsx_path = export_evs100_xlsx(add_records, chg_records, EVS100_TO_PROCESS)
    print(f"  📄  Saved to: evs100/ToProcess/{xlsx_path.name}")

    # ------------------------------------------------------------------ #
    # Summary                                                              #
    # ------------------------------------------------------------------ #
    print(f"\n{'═' * 60}")
    print(f"  MIG_SyncSortingOrders – Run Summary")
    print(f"{'═' * 60}")
    print(f"  📋  Records in SOURCE      : {len(source_all):,}")
    print(f"  📄  AddSortOrder records   : {len(add_records):,}")
    print(f"  📄  ChgSortOrder records   : {len(chg_records):,}")
    print(f"  📄  Excel file             : {xlsx_path.name}")
    print(f"{'═' * 60}")

    # ------------------------------------------------------------------ #
    # Step 5 – Optional: upload and process in DEST                      #
    # ------------------------------------------------------------------ #
    restore_config(dest_snap)
    get_ion_token()

    send = input("\n  Send this file to M3? [y/N]: ").strip().lower()
    if send != "y":
        print("  ℹ️   File not sent. Done.")
        return

    with requests.Session() as session:
        print(f"\n▶   Step 5a – Uploading {xlsx_path.name} to M3 …")
        uploaded = upload_file_to_m3(xlsx_path, session)

        if uploaded:
            trigger = input(
                "\n  Process file in M3 (EVS100MI.ImportFile)? [y/N]: "
            ).strip().lower()
            if trigger == "y":
                print(f"\n▶   Step 5b – Triggering EVS100MI.ImportFile …")
                process_file_in_m3(xlsx_path.name, session)
            else:
                print("  ℹ️   File uploaded but not processed. Done.")


# =============================================================================
# Entry Point
# =============================================================================

if __name__ == "__main__":
    sync_sort_orders()
