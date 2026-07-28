# ExtractTableCols.py
import sqlite3
import sys
from pathlib import Path
import requests
from tqdm import tqdm

from MpGd_Extract_FieldHelp import fetch_fieldhelp_xml, parse_fieldhelp_xml, save_to_sqlite
from InforMI import CONFIG, select_ionapi_file, prompt_for_company_division, get_ion_token

DB_PATH = "/Users/ericpronovost/sqlite/doppio.db"

# -------------------------------------------------------------------
# Create authenticated requests session
# -------------------------------------------------------------------
def create_authenticated_session() -> requests.Session:
    session = requests.Session()
    session.headers.update({
        "Authorization": f"Bearer {CONFIG['access_token']}",
        "Accept": "application/json"
    })
    return session

# -------------------------------------------------------------------
# Fetch table columns using LES endpoint
# -------------------------------------------------------------------
def fetch_table_columns(table_name: str, session: requests.Session) -> dict:
    tenant = CONFIG["tenant_id"]
    url = (
        f"https://mingle-ionapi.inforcloudsuite.com/{tenant}/M3/mdprest/les/"
        f"getColumnsUsedByTable/{table_name}/MVX?langId=GB"
    )

    response = session.get(url)
    if response.status_code == 401:
        print("🔄 Token expired. Refreshing token...")
        get_ion_token()
        session.headers.update({"Authorization": f"Bearer {CONFIG['access_token']}"})
        response = session.get(url)

    response.raise_for_status()
    return response.json()

# -------------------------------------------------------------------
# Save table metadata to m3TableCols with no duplicates
# -------------------------------------------------------------------
def save_table_columns_to_sqlite(table_name: str, api_json: dict):
    dataset = api_json.get("list", [])

    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()

    cur.execute("""
        CREATE TABLE IF NOT EXISTS m3TableCols (
            Sequence INTEGER PRIMARY KEY AUTOINCREMENT,
            TableName TEXT,
            ColumnName TEXT,
            Description TEXT,
            DataType TEXT,
            Length TEXT,
            Decimals TEXT,
            EditCode TEXT,
            Indexes TEXT,
            UNIQUE (TableName, ColumnName) ON CONFLICT IGNORE
        )
    """)

    for row in dataset:
        cur.execute("""
            INSERT INTO m3TableCols
            (TableName, ColumnName, Description, DataType, Length, Decimals, EditCode, Indexes)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            table_name,
            row.get("columnName"),
            row.get("description"),
            row.get("dataType"),
            row.get("length"),
            row.get("decimals"),
            row.get("editCode"),
            row.get("indexes")
        ))

    conn.commit()
    conn.close()

    print(f"✔ Saved {len(dataset)} columns for table {table_name}")

# -------------------------------------------------------------------
# Check if FieldHelp exists BEFORE calling API (performance boost)
# -------------------------------------------------------------------
def fieldhelp_exists(field_id: str) -> bool:
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute("SELECT 1 FROM m3FieldHelp WHERE FieldHelpID = ?", (field_id,))
    exists = cur.fetchone() is not None
    conn.close()
    return exists

# -------------------------------------------------------------------
# Extract table + FieldHelp with progress bars & summary
# -------------------------------------------------------------------
def extract_table_and_fieldhelp(table_name: str):
    print(f"\n📘 Extracting table metadata for {table_name}...")

    # -------------------------------------------------------
    # Ensure m3FieldHelp table exists (before any API calls)
    # -------------------------------------------------------
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute("""
        CREATE TABLE IF NOT EXISTS m3FieldHelp (
            FieldHelpID TEXT PRIMARY KEY,
            HeadingID TEXT,
            HeadingText TEXT,
            Definition TEXT,
            Alternatives TEXT,
            ProgramReferences TEXT,
            RawXML TEXT
        )
    """)
    conn.commit()
    conn.close()
    
    # ---------------------------------------------
    # Fetch & Save table column metadata
    # ---------------------------------------------
    session = create_authenticated_session()
    api_json = fetch_table_columns(table_name, session)
    save_table_columns_to_sqlite(table_name, api_json)

    columns = [c["columnName"] for c in api_json.get("list", [])]

    # Summary counters
    success_full = 0
    success_fallback = 0
    skipped = 0
    failed = 0

    print(f"🔍 Fetching FieldHelp for {len(columns)} fields...")

    # ---------------------------------------------
    # Progress bar for FieldHelp
    # ---------------------------------------------
    for col in tqdm(columns, desc="FieldHelp", unit="field"):
        # Skip if already exists
        if fieldhelp_exists(col) or fieldhelp_exists(col[-4:]):
            skipped += 1
            continue

        success = False

        # Try full column name
        try:
            xml = fetch_fieldhelp_xml(col)
            parsed = parse_fieldhelp_xml(xml)
            save_to_sqlite(parsed)
            success_full += 1
            success = True
        except Exception:
            success = False

        # Try fallback (last 4 chars)
        if not success:
            short_col = col[-4:]
            try:
                xml = fetch_fieldhelp_xml(short_col)
                parsed = parse_fieldhelp_xml(xml)
                save_to_sqlite(parsed)
                success_fallback += 1
                success = True
            except Exception:
                # FINAL FAIL: create blank fallback entry
                fallback = {
                    "FieldHelpID": short_col,
                    "HeadingID": "",
                    "HeadingText": "",
                    "Definition": "",
                    "Alternatives": "",
                    "ProgramReferences": "",
                    "RawXML": ""
                }
                save_to_sqlite(fallback)
                failed += 1

    # ---------------------------------------------
    # Summary
    # ---------------------------------------------
    print("\n📊 Summary")
    print(f"   ✔ Success (full):       {success_full}")
    print(f"   ✔ Success (fallback):   {success_fallback}")
    print(f"   ➖ Skipped (duplicate):  {skipped}")
    print(f"   ❌ Failed:               {failed}")

# -------------------------------------------------------------------
# Initialize authentication
# -------------------------------------------------------------------
def initialize_auth():
    ionapi_dir = Path(__file__).parent / "ionapi"
    CONFIG["tenant"] = select_ionapi_file(ionapi_dir)
    prompt_for_company_division()
    print("🔐 Getting OAuth token...")
    get_ion_token()
    print("✔ OAuth token acquired.")

# -------------------------------------------------------------------
# Run script
# -------------------------------------------------------------------
if __name__ == "__main__":
    initialize_auth()

    if len(sys.argv) > 1:
        tables = [sys.argv[1]]
    else:
        tables = [
            "CBANAC", "CIDADR", "CIDMAS", "CIDVEN", "CSYTAB",
            "FCHACC", "MITBAL", "MITFAC", "MITMAS", "MITPCE",
            "MPDHED", "MPDMAT", "MPDOPE", "OCUSAD", "OCUSMA",
            "OPRICH", "OPRICL", "MITVEN", "MITLOC", "MITAUN",
            "CSYCSN", "MPDWCT", "MGHEAD"
        ]

    for tbl in tables:
        try:
            extract_table_and_fieldhelp(tbl)
        except Exception:
            pass
