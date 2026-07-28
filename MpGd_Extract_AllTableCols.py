# MpGd_ExtractAllCols.py
# -----------------------------------------------------------------------
# PURPOSE
#   Fetch and persist M3 column metadata for every table listed in the
#   m3Tables SQLite table.  Unlike MpGd__M3Info.py (which is driven by
#   ExtractRules), this routine processes the full set of known tables so
#   that m3TableCols is complete regardless of active extraction rules.
#
# INPUTS
#   - m3Tables  (SQLite) — SELECT DISTINCT TableName drives the table list
#   - M3 IonAPI /M3/mdprest/les/getColumnsUsedByTable endpoint (HTTP)
#   - ionapi/ folder — .ionapi tenant config files
#
# OUTPUTS
#   - m3TableCols  (SQLite) — table column metadata (upserted, no duplicates)
#
# DEPENDENCIES
#   - config.py, InforMI.py
# -----------------------------------------------------------------------

import logging
import sqlite3
import requests
from pathlib import Path
from tqdm import tqdm

from InforMI import CONFIG, select_ionapi_file, prompt_for_company_division, get_ion_token
from config import get_sqlite_db_path

logger = logging.getLogger(__name__)

DB_PATH = get_sqlite_db_path()


def create_authenticated_session() -> requests.Session:
    session = requests.Session()
    session.headers.update({
        "Authorization": f"Bearer {CONFIG['access_token']}",
        "Accept": "application/json"
    })
    return session


def fetch_table_columns(table_name: str, session: requests.Session) -> dict:
    tenant = CONFIG["ti"]
    url = (
        f"https://mingle-ionapi.inforcloudsuite.com/{tenant}/M3/mdprest/les/"
        f"getColumnsUsedByTable/{table_name[:6]}/MVX?langId=GB"
    )

    response = session.get(url)

    if response.status_code == 401:
        logger.warning("Token expired. Refreshing...")
        get_ion_token()
        session.headers.update({"Authorization": f"Bearer {CONFIG['access_token']}"})
        response = session.get(url)

    response.raise_for_status()
    return response.json()


def save_table_columns_to_sqlite(table_name: str, api_json: dict):
    table_name = table_name[:6] if table_name else None
    dataset = api_json.get("list") or []

    if not isinstance(dataset, list):
        logger.warning(f"Unexpected payload for {table_name}")
        return

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
        cur.executemany("""
            INSERT INTO m3TableCols
            (TableName, ColumnName, Description, DataType, Length, Decimals, EditCode, Indexes)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """, rows)


def load_table_list() -> list[str]:
    with sqlite3.connect(DB_PATH) as conn:
        rows = conn.execute("""
            SELECT DISTINCT TableName FROM m3Tables
            WHERE TableName NOT IN (
                SELECT DISTINCT TableName FROM m3TableCols
            )
            AND tableComponent='MVX' 
            AND tableName not like '______J_'
        """).fetchall()
    return [r[0] for r in rows if r[0]]


def run():
    tables = load_table_list()
    logger.info(f"Found {len(tables)} tables in m3Tables.")

    session = create_authenticated_session()

    for table_name in tqdm(tables, desc="Extracting columns"):
        try:
            api_json = fetch_table_columns(table_name, session)
            save_table_columns_to_sqlite(table_name, api_json)
        except Exception as exc:
            logger.warning(f"Skipping {table_name}: {exc}")


if __name__ == "__main__":
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s  %(levelname)-8s  %(message)s",
        datefmt="%H:%M:%S"
    )

    ionapi_dir = Path(__file__).parent / "ionapi"
    CONFIG["tenant"] = select_ionapi_file(ionapi_dir)
    prompt_for_company_division()
    logger.info("Getting OAuth token...")
    get_ion_token()
    logger.info("OAuth token acquired.")

    run()
    logger.info("Done.")
