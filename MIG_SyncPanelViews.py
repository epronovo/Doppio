# MIG_SyncPanelViews.py
"""
MIG_SyncPanelViews.py
---------------------
Syncs panel view column configurations from a SOURCE tenant to a DESTINATION tenant.

Steps
-----
1.  Ask for SOURCE tenant → EXPORTMI.Select (CSYSPV) → parse REPL rows into memory.
2.  Ask for DEST   tenant → EXPORTMI.Select (CSYSPV) → parse REPL rows into memory.
3.  Diff: find views where C9PARA differs between SOURCE and DEST
    (matched on C9PGNM + C9PIC1 + C9PAVR).  Display and prompt to export.
4.  Export differing views to an EVS100-format Excel file
    (evs100/ToProcess/API_CRS020MI_<timestamp>.xlsx) with a Control sheet,
    a DelPanelVersion sheet, and an ImportView sheet.
5.  Optionally upload the file to DEST via the File Management REST API (PUT).
6.  Optionally trigger processing via EVS100MI.ImportFile.

Usage:
    python MIG_SyncPanelViews.py
"""

import datetime
import os
import requests
from pathlib import Path
import xlsxwriter

from InforMI import (
    CONFIG,
    get_ion_token,
    post_to_m3,
)
from UserDefaults import load_user_defaults, save_user_defaults

# =============================================================================
# Constants
# =============================================================================

EXPORTMI_QUERY = (
    "C9PAVR,C9PGNM,C9TX40,C9PARA,C9IBCA,C9RESP,C9PIC1,C9TX15,C9PVTP,C9UPDC,C9LUFN,C9LUVE,C9MCRT "
    # "from CSYSPV where C9PGNM <> CMS100 and C9PGNM <> LISTMI and C9PAVR <> ''"
    "from CSYSPV where C9PAVR <> ''"
)
EXPORTMI_SEP   = "^"

# Mapping from EXPORTMI C9xxx column names → CRS020MI.ImportView field names.
# C9PARA is renamed to PAR1 in the API.
IMPORT_VIEW_MAP: dict[str, str] = {
    "C9PAVR": "PAVR",
    "C9PGNM": "PGNM",
    "C9TX40": "TX40",
    "C9PARA": "PAR1",
    "C9IBCA": "IBCA",
    "C9RESP": "RESP",
    "C9PIC1": "PIC1",
    "C9TX15": "TX15",
    "C9PVTP": "PVTP",
    "C9UPDC": "UPDC",
    "C9LUFN": "LUFN",
    "C9LUVE": "LUVE",
    "C9MCRT": "MCRT",
}

IMPORT_VIEW_SELECTED = [
    "PAVR","PGNM","TX40","PAR1","IBCA","RESP","PIC1","TX15","PVTP","UPDC","LUFN","LUVE","MCRT",
]

DEL_SHEET    = "API_CRS020MI_DelPanelVersion"
IMPORT_SHEET = "API_CRS020MI_ImportView"
DEL_FIELDS   = ["PGNM", "PIC1", "PAVR"]

EVS100_TO_PROCESS = Path(__file__).parent / "evs100" / "ToProcess"


# =============================================================================
# Tenant helpers
# =============================================================================

def snapshot_config() -> dict:
    return dict(CONFIG)


def restore_config(snapshot: dict) -> None:
    CONFIG.clear()
    CONFIG.update(snapshot)


def _select_ionapi_forced(ionapi_dir: Path, label: str) -> Path:
    files = sorted(
        [f for f in os.listdir(ionapi_dir) if f.endswith(".ionapi")],
        key=str.lower,
    )
    if not files:
        raise FileNotFoundError(f"No .ionapi files found in: {ionapi_dir}")

    print(f"\n  🔹  Select the {label} ION API file:")
    for i, f in enumerate(files, start=1):
        print(f"    {i}. {f}")

    while True:
        try:
            choice = int(input(f"\n  Enter choice (1-{len(files)}): "))
            if 1 <= choice <= len(files):
                return ionapi_dir / files[choice - 1]
            print("  ❌  Invalid choice. Try again.")
        except ValueError:
            print("  ❌  Please enter a number.")


def _prompt_company_division_forced(label: str) -> None:
    defaults = load_user_defaults()
    default_company  = defaults.get("company", "100")
    default_division = defaults.get("division", "")

    print(f"\n  Configure {label} company / division:")
    company  = input(f"    Company  (default: {default_company}):  ").strip()
    division = input(f"    Division (default: {default_division}):  ").strip()

    CONFIG["company"]  = company  if company  else default_company
    CONFIG["division"] = division if division else default_division

    defaults["company"]  = CONFIG["company"]
    defaults["division"] = CONFIG["division"]
    save_user_defaults(defaults)


def setup_tenant(ionapi_dir: Path, label: str) -> dict:
    print(f"\n{'─' * 60}")
    print(f"  📡  Configure {label} tenant")
    print(f"{'─' * 60}")
    CONFIG["tenant"] = _select_ionapi_forced(ionapi_dir, label)
    _prompt_company_division_forced(label)
    get_ion_token()
    snap = snapshot_config()
    print(f"  ✅  {label} ready: {Path(CONFIG['tenant']).name}  "
          f"CONO={CONFIG.get('company', '')}  DIVI={CONFIG.get('division', '')}")
    return snap


# =============================================================================
# Step 1 / 2 – EXPORTMI.Select → parse REPL rows
# =============================================================================

def fetch_panel_views(session: requests.Session, label: str) -> list[dict]:
    """
    Calls EXPORTMI.Select against CSYSPV and returns a list of dicts,
    one per panel-view row, keyed by the column names in the header row.
    """
    list_url = CONFIG["api_url"]

    payload = {
        "program": "EXPORTMI",
        "transactions": [{
            "transaction": "Select",
            "record": {
                "QERY": EXPORTMI_QUERY,
                "SEPC": EXPORTMI_SEP,
                "HDRS": "1",
            },
            "selectedColumns": ["QERY", "SEPC", "HDRS", "REPL"],
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

    repl_rows: list[str] = []
    for result in data.get("results", []):
        for record in result.get("records", []):
            repl = record.get("REPL", "")
            if repl:
                repl_rows.append(repl)

    if not repl_rows:
        print(f"  ⚠️   No rows returned from EXPORTMI ({label}).")
        return []

    # First row is the header (HDRS:1)
    col_names = [c.strip() for c in repl_rows[0].rstrip(EXPORTMI_SEP).split(EXPORTMI_SEP)]

    parsed: list[dict] = []
    for raw in repl_rows[1:]:
        values = raw.rstrip(EXPORTMI_SEP).split(EXPORTMI_SEP)
        # Pad / trim to match header width
        values += [""] * (len(col_names) - len(values))
        # C9PARA is a fixed-width positional string — preserve it as-is so
        # character positions remain intact.  All other fields are stripped.
        row = {
            col_names[i]: (values[i] if col_names[i] == "C9PARA" else values[i].strip())
            for i in range(len(col_names))
        }
        parsed.append(row)

    print(f"  ✅  {len(parsed)} panel-view records retrieved ({label}).")
    return parsed


# =============================================================================
# C9PARA helpers
# =============================================================================

PARA_CMP_LEN = 1262   # only the first 1262 characters are meaningful for comparison


def cmp_para(c9para: str) -> str:
    """Return the comparison slice of C9PARA (first 1262 characters)."""
    return c9para[:PARA_CMP_LEN]


def para_diff_summary(src: str, dst: str) -> str:
    """
    Return a short human-readable summary of where two C9PARA strings differ.
    E.g. 'pos 40-79 (40 chars)' or 'multiple ranges (120 chars total)'
    """
    s, d = src[:PARA_CMP_LEN], dst[:PARA_CMP_LEN]
    max_len = max(len(s), len(d))
    s = s.ljust(max_len)
    d = d.ljust(max_len)

    # Collect contiguous differing ranges
    ranges: list[tuple[int, int]] = []
    in_range = False
    start = 0
    for i in range(max_len):
        if s[i] != d[i]:
            if not in_range:
                in_range = True
                start = i
        else:
            if in_range:
                ranges.append((start, i - 1))
                in_range = False
    if in_range:
        ranges.append((start, max_len - 1))

    if not ranges:
        return "identical"

    total = sum(e - b + 1 for b, e in ranges)
    if len(ranges) == 1:
        b, e = ranges[0]
        span = f"pos {b}-{e}" if b != e else f"pos {b}"
        return f"{span} ({total} char{'s' if total != 1 else ''})"
    return f"{len(ranges)} ranges ({total} chars total)"


# =============================================================================
# Record builders
# =============================================================================

def build_del_record(source_row: dict) -> dict:
    """Build the key record for CRS020MI.DelPanelVersion from an EXPORTMI source row."""
    return {
        "PGNM": source_row.get("C9PGNM", ""),
        "PIC1": source_row.get("C9PIC1", ""),
        "PAVR": source_row.get("C9PAVR", ""),
    }


def build_import_record(source_row: dict) -> dict:
    """
    Build the record for CRS020MI.ImportView from an EXPORTMI source row.
    Maps C9xxx keys to API field names (C9PARA → PAR1).
    Skips any field whose value is blank.
    """
    rec: dict = {}
    for c9_key, api_key in IMPORT_VIEW_MAP.items():
        val = source_row.get(c9_key, "")
        if val != "":
            rec[api_key] = val
    return rec


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

def export_evs100_xlsx(differing: list[dict], out_dir: Path) -> Path:
    """
    Writes differing panel-view records in the EVS100 import format.

    Layout
    ------
    Control sheet:
        Row 1 header : Worksheet | Description | Data
        Row 2 entry  : API_CRS020MI_DelPanelVersion | Delete Panel Version | x
        Row 3 entry  : API_CRS020MI_ImportView      | Import Panel View    | x

    Data sheets (one per operation):
        Row 1  – field names
        Row 2  – field descriptions
        Row 3  – required flags
        Row 4+ – data rows

    Returns the path of the written file.
    """
    del_records = [build_del_record(row) for row in differing]
    imp_records = [build_import_record(row) for row in differing]

    ts       = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    out_path = out_dir / f"API_CRS020MI_{ts}.xlsx"

    wb = xlsxwriter.Workbook(str(out_path))

    # ── Control sheet ─────────────────────────────────────────────────────
    ws_ctrl = wb.add_worksheet("Control")
    for col, val in enumerate(["Worksheet", "Description", "Data"]):
        ws_ctrl.write(0, col, val)
    ws_ctrl.write(1, 0, DEL_SHEET);    ws_ctrl.write(1, 1, "Delete Panel Version"); ws_ctrl.write(1, 2, "x")
    ws_ctrl.write(2, 0, IMPORT_SHEET); ws_ctrl.write(2, 1, "Import Panel View");    ws_ctrl.write(2, 2, "x")

    # ── Sheet 1: DelPanelVersion ───────────────────────────────────────────
    ws_del  = wb.add_worksheet(DEL_SHEET)
    del_cols = ["MESSAGE"] + DEL_FIELDS

    for col, name in enumerate(del_cols):
        ws_del.write(0, col, name)
    for col, name in enumerate(del_cols):
        ws_del.write(1, col, "" if name == "MESSAGE" else name)
    for col in range(len(del_cols)):
        ws_del.write(2, col, "no" if col == 0 else "yes")

    for row_idx, rec in enumerate(del_records, start=3):
        for col, field in enumerate(DEL_FIELDS, start=1):
            val = rec.get(field, "")
            if val:
                ws_del.write(row_idx, col, val)

    # ── Sheet 2: ImportView ────────────────────────────────────────────────
    ws_imp  = wb.add_worksheet(IMPORT_SHEET)
    imp_cols = ["MESSAGE"] + IMPORT_VIEW_SELECTED

    for col, name in enumerate(imp_cols):
        ws_imp.write(0, col, name)
    for col, name in enumerate(imp_cols):
        ws_imp.write(1, col, "" if name == "MESSAGE" else name)
    for col in range(len(imp_cols)):
        ws_imp.write(2, col, "no" if col == 0 else "yes")

    for row_idx, rec in enumerate(imp_records, start=3):
        for col, field in enumerate(IMPORT_VIEW_SELECTED, start=1):
            val = rec.get(field, "")
            if val:
                ws_imp.write(row_idx, col, val)

    wb.close()
    return out_path


# =============================================================================
# Main
# =============================================================================

def sync_panel_views() -> None:
    ionapi_dir = Path(__file__).parent / "ionapi"

    # ------------------------------------------------------------------ #
    # Step 1 – SOURCE: EXPORTMI.Select                                    #
    # ------------------------------------------------------------------ #
    source_snap = setup_tenant(ionapi_dir, "SOURCE")
    print("\n🔍  Step 1 – EXPORTMI.Select CSYSPV (SOURCE) …")
    with requests.Session() as session:
        source_views = fetch_panel_views(session, label="SOURCE")

    if not source_views:
        print("⚠️   No panel views found in SOURCE. Exiting.")
        return

    # ------------------------------------------------------------------ #
    # Step 2 – DEST: EXPORTMI.Select                                      #
    # ------------------------------------------------------------------ #
    dest_snap = setup_tenant(ionapi_dir, "DEST")
    print("\n🔍  Step 2 – EXPORTMI.Select CSYSPV (DEST) …")
    with requests.Session() as session:
        dest_views = fetch_panel_views(session, label="DEST")

    # ------------------------------------------------------------------ #
    # Step 3 – Diff: find views where C9PARA differs                      #
    # ------------------------------------------------------------------ #
    dest_index: dict[tuple, dict] = {
        (r["C9PGNM"], r["C9PIC1"], r["C9PAVR"]): r
        for r in dest_views
    }

    differing: list[dict] = []
    for row in source_views:
        key = (row["C9PGNM"], row["C9PIC1"], row["C9PAVR"])
        dest_row = dest_index.get(key)
        dest_para = cmp_para(dest_row.get("C9PARA", "")) if dest_row else None
        if dest_para is None or dest_para != cmp_para(row.get("C9PARA", "")):
            differing.append(row)

    if not differing:
        print("\n✅  DEST panel views all match SOURCE. Nothing to do.")
        return

    # Display plan
    print(f"\n{'═' * 60}")
    print(f"  📋  {len(differing)} view(s) differ between SOURCE and DEST")
    print(f"{'═' * 60}")
    for row in differing[:10]:
        key = (row["C9PGNM"], row["C9PIC1"], row["C9PAVR"])
        dest_row = dest_index.get(key)
        if dest_row is None:
            status = "NEW"
            diff_info = "not in DEST"
        else:
            status = "CHG"
            diff_info = para_diff_summary(
                cmp_para(row.get("C9PARA", "")),
                cmp_para(dest_row.get("C9PARA", "")),
            )
        print(f"  [{status}]  PGNM={row['C9PGNM']:<8}  PIC1={row['C9PIC1']}  PAVR={row['C9PAVR']}  diff={diff_info}")
    if len(differing) > 10:
        print(f"  … and {len(differing) - 10} more")
    print(f"\n{'═' * 60}")

    confirm = input(f"  Export these {len(differing)} view(s) to Excel? [y/N]: ").strip().lower()
    if confirm != "y":
        print("⛔  Aborted — no file created.")
        return
    print()

    # ------------------------------------------------------------------ #
    # Step 4 – Export to EVS100 Excel file                               #
    # ------------------------------------------------------------------ #
    print(f"▶   Step 4 – Exporting {len(differing)} view(s) to EVS100 Excel file …")
    xlsx_path = export_evs100_xlsx(differing, EVS100_TO_PROCESS)
    print(f"  📄  Saved to: evs100/ToProcess/{xlsx_path.name}")

    # ------------------------------------------------------------------ #
    # Summary                                                              #
    # ------------------------------------------------------------------ #
    print(f"\n{'═' * 60}")
    print(f"  MIG_SyncPanelViews – Run Summary")
    print(f"{'═' * 60}")
    print(f"  📋  Views in SOURCE        : {len(source_views):,}")
    print(f"  📋  Views differing        : {len(differing):,}")
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
    sync_panel_views()
