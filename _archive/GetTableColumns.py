# GetTableColumns.py
import sys
import sqlite3
import requests
from pathlib import Path
from tqdm import tqdm

from InforMI import CONFIG, select_ionapi_file, prompt_for_company_division, get_ion_token
from config import get_sqlite_db_path
from tqdm import tqdm

# =============================================================================
# Table Metadata Extract
# =============================================================================

def get_table_columns(table_name, lang="GB"):
    """
    Calls the M3 metadata REST API to get column information for a given table.
    """
    base_url = CONFIG['api_url'].split("/M3/m3api-rest")[0]
    endpoint = f"/M3/mdprest/les/getColumnsUsedByTable/{table_name}/MVX?langId={lang}"
    full_url = base_url + endpoint

    headers = {
        "Authorization": f"Bearer {CONFIG['access_token']}",
        "accept": "application/json"
    }

    with requests.Session() as session:
        response = session.get(full_url, headers=headers)

        # Refresh token once if unauthorized
        if response.status_code == 401:
            get_ion_token()
            headers["Authorization"] = f"Bearer {CONFIG['access_token']}"
            response = session.get(full_url, headers=headers)

        response.raise_for_status()
        return response.json()


def append_to_M3TableColumns(cur, table_name, columns_data):
    """
    Replaces metadata for a specific table in M3TableColumns.
    Adds columnSeq to preserve order.
    """
    cur.execute("""
        CREATE TABLE IF NOT EXISTS M3TableColumns (
            tableName   TEXT,
            columnSeq   INTEGER,
            columnName  TEXT,
            description TEXT,
            dataType    TEXT,
            length      TEXT,
            decimals    TEXT,
            editCode    TEXT,
            indexes     TEXT
        )
    """)

    # Delete existing records for this table to avoid duplicates
    cur.execute("DELETE FROM M3TableColumns WHERE tableName = ?", (table_name,))

    insert_sql = """
        INSERT INTO M3TableColumns
        (tableName, columnSeq, columnName, description, dataType, length, decimals, editCode, indexes)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
    """

    columns_list = columns_data.get("list", [])
    if not columns_list:
        return  # nothing to insert

    for seq, col in enumerate(columns_list, start=1):
        cur.execute(insert_sql, (
            table_name,
            seq,
            col.get("columnName"),
            col.get("description"),
            col.get("dataType"),
            col.get("length"),
            col.get("decimals"),
            col.get("editCode"),
            col.get("indexes")
        ))

# =============================================================================
# Main
# =============================================================================
if __name__ == "__main__":
    # Initialize authentication & config
    ionapi_dir = Path(__file__).parent / "ionapi"
    CONFIG["tenant"] = select_ionapi_file(ionapi_dir)
    prompt_for_company_division()
    get_ion_token()

    SQLITE_DB_PATH = get_sqlite_db_path()

    try:
        with sqlite3.connect(SQLITE_DB_PATH) as conn:
            cur = conn.cursor()

            # Get list of tables from M3Tables
            cur.execute("SELECT tableName FROM M3Tables")
            tables = [row[0] for row in cur.fetchall()]

            if not tables:
                print("⚠️ No tables found in M3Tables.")
                sys.exit(1)

            for table_name in tqdm(tables, desc="Processing tables"):
                try:
                    columns_data = get_table_columns(table_name)
                    append_to_M3TableColumns(cur, table_name, columns_data)
                    conn.commit()
                except Exception as e:
                    # Show failures without breaking progress bar
                    tqdm.write(f"❌ {table_name} failed: {e}")

        print(f"🏁 Finished processing {len(tables)} tables.")

    except Exception as e:
        print(f"❌ Fatal error: {e}")