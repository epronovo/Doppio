# MpGd_Extract_M3Info.py
# -----------------------------------------------------------------------
# PURPOSE
#   Main orchestrator for the M3 metadata extraction pipeline.
#   Authenticates via IonAPI OAuth, then iterates over every distinct
#   table in ExtractRules to: (1) fetch column metadata from the M3 MDP
#   getColumnsUsedByTable endpoint, (2) extract FieldHelp for each column
#   concurrently (12 threads, with a short-ID fallback and empty-placeholder
#   on failure), and (3) handle transparent token refresh on 401 responses.
#   After extraction, calls MpGd_Build_m3Api2Table to build the API-to-column
#   mapping, then seeds DoppioGuide and HerffJonesMappingGuide for any
#   newly discovered tables or API transactions.
#
# INPUTS
#   - ExtractRules  (SQLite) — Tbl column drives the table list
#   - M3 IonAPI /M3/mdprest/les/getColumnsUsedByTable endpoint (HTTP)
#   - M3 FieldHelp XML endpoint (HTTP, via MpGd_Extract_FieldHelp)
#   - ionapi/ folder — .ionapi tenant config files
#
# OUTPUTS
#   - m3TableCols           (SQLite) — table column metadata
#   - m3FieldHelp           (SQLite) — field help content
#   - m3Api2Table           (SQLite) — API ↔ table/column mapping
#   - DoppioGuide           (SQLite) — seeded for new tables
#   - HerffJonesMappingGuide (SQLite) — seeded for new API transactions
#
# DEPENDENCIES
#   - config.py, MpGd_Build_m3Api2Table.py, MpGd_Extract_FieldHelp.py,
#     InforMI.py
# -----------------------------------------------------------------------

import logging
import sqlite3
import requests
import threading
from pathlib import Path
from typing import List, Optional
from tqdm import tqdm
from concurrent.futures import ThreadPoolExecutor, as_completed

logger = logging.getLogger(__name__)

# -------------------------------------------------------------------
# Local project imports
# -------------------------------------------------------------------
# build_m3Api2Table:
#   Builds the API ↔ table/column mapping table from extracted metadata
#
# fetch_fieldhelp_xml / parse_fieldhelp_xml / save_to_sqlite:
#   End-to-end FieldHelp extraction and persistence helpers
#
# InforMI helpers:
#   Handle IonAPI configuration, tenant selection, and OAuth tokens
# -------------------------------------------------------------------
from MpGd_Build_m3Api2Table import build_m3Api2Table
from MpGd_Extract_FieldHelp import fetch_fieldhelp_xml, parse_fieldhelp_xml, save_to_sqlite, process_all_missing_fieldhelp
from MpGd_Extract_Metadata import main as extract_metadata
from MpGd_Extract_Prompts import run as extract_prompts
from MpGd_Extract_DataStructures import run as extract_data_structures
from InforMI import CONFIG, select_ionapi_file, prompt_for_company_division, get_ion_token
from config import get_sqlite_db_path

# SQLite database location
DB_PATH = get_sqlite_db_path()


def create_authenticated_session() -> requests.Session:
    """
    Create a reusable requests.Session preconfigured with
    the current IonAPI OAuth access token.
    """
    session = requests.Session()
    session.headers.update({
        "Authorization": f"Bearer {CONFIG['access_token']}",
        "Accept": "application/json"
    })
    return session


def fetch_table_columns(table_name: str, session: requests.Session) -> dict:
    """
    Call the M3 metadata REST service to retrieve column definitions
    for a given table.

    Automatically refreshes the OAuth token if a 401 is encountered.
    """
    tenant = CONFIG["ti"]

    # M3 metadata endpoint (table name limited to 6 chars)
    url = (
        f"https://mingle-ionapi.inforcloudsuite.com/{tenant}/M3/mdprest/les/"
        f"getColumnsUsedByTable/{table_name[:6]}/MVX?langId=GB"
    )

    response = session.get(url)

    # Handle expired token transparently
    if response.status_code == 401:
        logger.warning("🔄 Token expired. Refreshing token...")
        get_ion_token()
        session.headers.update({"Authorization": f"Bearer {CONFIG['access_token']}"})
        response = session.get(url)

    response.raise_for_status()
    return response.json()


def save_table_columns_to_sqlite(table_name: str, api_json: dict):
    """
    Persist table column metadata returned by the API into SQLite
    (table: m3TableCols).

    Uses a UNIQUE constraint to safely ignore duplicates.
    """
    table_name = table_name[:6] if table_name else None
    dataset = api_json.get("list") or []

    # Guard against unexpected payloads
    if not isinstance(dataset, list):
        logger.warning(f"⚠  Unexpected payload for {table_name}")
        return

    # Normalize API payload into insertable rows
    rows = [
        (
            table_name,
            row.get("columnName"),
            row.get("description"),
            row.get("dataType"),
            row.get("length"),
            row.get("decimals"),
            row.get("editCode"),
            row.get("indexes"),
        )
        for row in dataset
    ]

    with sqlite3.connect(DB_PATH) as conn:
        cur = conn.cursor()

        # Create metadata table if needed
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

        # Bulk insert
        cur.executemany("""
            INSERT INTO m3TableCols
            (TableName, ColumnName, Description, DataType, Length, Decimals, EditCode, Indexes)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """, rows)

    # Optional debug output
    # print(f"✔ Saved {len(rows)} columns for {table_name}")


def fieldhelp_exists(field_id: str) -> bool:
    """
    Check whether FieldHelp already exists in SQLite for a given ID.
    Used to avoid redundant API calls.
    """
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute(
        "SELECT 1 FROM m3FieldHelp WHERE FieldHelpID = ?",
        (field_id,)
    )
    exists = cur.fetchone() is not None
    conn.close()
    return exists


def extract_table_and_fieldhelp(table_name: str):
    """
    End-to-end extraction for a single table:
      1. Ensure FieldHelp table exists
      2. Fetch and store table column metadata
      3. Extract FieldHelp for each column (threaded)
    """

    # Ensure FieldHelp target table exists
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

    # Fetch and persist table column metadata
    session = create_authenticated_session()
    api_json = fetch_table_columns(table_name, session)
    save_table_columns_to_sqlite(table_name, api_json)

    # Extract column names from API payload
    columns = [c["columnName"] for c in api_json.get("list", [])]

    # Track outcomes for diagnostics
    result_counters = {
        "success_full": 0,
        "success_fallback": 0,
        "skipped": 0,
        "failed": 0
    }

    # SQLite writes must be serialized across threads
    sqlite_lock = threading.Lock()

    def process_column(col):
        """
        Attempt FieldHelp extraction for a single column.

        Strategy:
          1. Skip if FieldHelp already exists
          2. Try full column name
          3. Fallback to last 4 characters
          4. Insert empty placeholder if all attempts fail
        """
        # Skip if already processed (full or short ID)
        if fieldhelp_exists(col) or fieldhelp_exists(col[-4:]):
            return ("skipped", col)

        # Attempt full column name
        try:
            xml = fetch_fieldhelp_xml(col)
            parsed = parse_fieldhelp_xml(xml)
            with sqlite_lock:
                save_to_sqlite(parsed)
            return ("success_full", col)
        except Exception:
            pass

        # Attempt fallback (last 4 characters)
        short_col = col[-4:]
        try:
            xml = fetch_fieldhelp_xml(short_col)
            parsed = parse_fieldhelp_xml(xml)
            with sqlite_lock:
                save_to_sqlite(parsed)
            return ("success_fallback", col)
        except Exception:
            # Persist empty placeholder to prevent reprocessing
            fallback = {
                "FieldHelpID": short_col,
                "HeadingID": "",
                "HeadingText": "",
                "Definition": "",
                "Alternatives": "",
                "ProgramReferences": "",
                "RawXML": ""
            }
            with sqlite_lock:
                save_to_sqlite(fallback)
            return ("failed", col)

    # Run FieldHelp extraction concurrently (no progress bar)
    with ThreadPoolExecutor(max_workers=12) as executor:
        futures = [executor.submit(process_column, col) for col in columns]

        for future in as_completed(futures):
            status, col = future.result()
            result_counters[status] += 1

    # Optional diagnostics
    # print(f"FieldHelp results: {result_counters}")


def initialize_auth():
    """
    Initialize IonAPI authentication:
      - Select tenant (.ionapi file)
      - Prompt for company/division
      - Acquire OAuth access token
    """
    ionapi_dir = Path(__file__).parent / "ionapi"
    CONFIG["tenant"] = select_ionapi_file(ionapi_dir)
    prompt_for_company_division()
    logger.info("🔐 Getting OAuth token...")
    get_ion_token()
    logger.info("✔ OAuth token acquired.")


def extract_tables_from_rules():
    """
    Drive table metadata and FieldHelp extraction using ExtractRules.
    Each distinct table becomes a unit of work.
    """
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()

    tables = cur.execute("""
        SELECT DISTINCT substr(Tbl, 1, 6) AS TableName
        FROM ExtractRules
        WHERE Tbl IS NOT NULL AND Tbl <> ''
    """).fetchall()

    conn.close()

    for (tbl,) in tqdm(tables, desc="Extracting table metadata"):
        extract_table_and_fieldhelp(tbl)


def finalize_mapping_guides():
    """
    Populate downstream mapping guide tables based on newly
    extracted API and metadata information.
    """
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()

    # -------------------------------------------------------
    # Populate DoppioGuide for newly discovered tables
    # -------------------------------------------------------
    cur.executescript("""
    CREATE TABLE IF NOT EXISTS DoppioGuide (
        Sequence    INTEGER PRIMARY KEY AUTOINCREMENT,
        TableName   TEXT,
        ColumnName  TEXT,
        Program     TEXT,
        Panel       TEXT,
        API         TEXT,
        TransactionName TEXT,
        FieldName   TEXT,
        MappingNotes TEXT,
        Responsible TEXT,
        Required    TEXT,
        DefaultValue TEXT,
        UNIQUE (TableName, ColumnName) ON CONFLICT IGNORE
    );

    INSERT INTO DoppioGuide (TableName,ColumnName,Program,Panel,API,TransactionName,FieldName,MappingNotes,Responsible,Required,DefaultValue)
    SELECT TableName,ColumnName,Program,Panel,API,TransactionName,FieldName,
           '' AS MappingNotes,
           '' AS Responsible,
           CASE WHEN API <> '' THEN 'yes' ELSE 'no' END AS Required,
           '' AS DefaultValue
    FROM m3Api2Table
    WHERE TableName NOT IN (
        SELECT DISTINCT TableName FROM DoppioGuide
    );
    """)

    # -------------------------------------------------------
    # Populate HerffJonesMappingGuide for new API/transactions
    # -------------------------------------------------------
    cur.executescript("""
    CREATE TABLE IF NOT EXISTS HerffJonesMappingGuide (
        Sequence        INTEGER PRIMARY KEY AUTOINCREMENT,
        API             TEXT,
        TransactionName TEXT,
        FieldName       TEXT,
        MappingNotes    TEXT,
        Responsible     TEXT,
        Required        TEXT,
        DefaultValue    TEXT,
        UNIQUE (API, TransactionName, FieldName) ON CONFLICT IGNORE
    );

    INSERT INTO HerffJonesMappingGuide (API,TransactionName,FieldName,MappingNotes,Responsible,Required,DefaultValue)
    SELECT
        MINM AS API,
        TRNM AS TransactionName,
        FLNM AS FieldName,
        FLDS AS MappingNotes,
        '' AS Responsible,
        CASE WHEN MAND = 1 THEN 'yes' ELSE 'no' END AS Required,
        '' AS DefaultValue
    FROM cmifld
    WHERE TRTP = 'I'
      AND MINM || TRNM IN (
        SELECT DISTINCT t2a.API || t2a.TransactionName
        FROM m3Api2Table t2a
        LEFT JOIN HerffJonesMappingGuide hj
            ON hj.API = t2a.API
           AND hj.TransactionName = t2a.TransactionName
        WHERE t2a.API <> ''
          AND hj.API IS NULL
      );
    """)

    conn.commit()
    conn.close()

    logger.info("✔ Mapping guides finalized.")


# ===================================================================
# MAIN ENTRY POINT
# ===================================================================
if __name__ == "__main__":
    import argparse

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s  %(levelname)-8s  %(message)s",
        datefmt="%H:%M:%S"
    )

    parser = argparse.ArgumentParser(description="M3 metadata extraction pipeline")
    parser.add_argument(
        "--table", "-t",
        metavar="TABLE",
        help="Extract a single table by name instead of processing ExtractRules"
    )
    args = parser.parse_args()

    initialize_auth()

    if args.table:
        logger.info(f"Extracting single table: {args.table}")
        extract_table_and_fieldhelp(args.table)
    else:
        extract_tables_from_rules()

    build_m3Api2Table()
    finalize_mapping_guides()
    extract_metadata()
    process_all_missing_fieldhelp()
    extract_prompts()
    extract_data_structures()
    logger.info("🎉 All done!")
