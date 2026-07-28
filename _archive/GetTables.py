# GetTables.py
import json
import sqlite3
import http.client
import re
from pathlib import Path
from tqdm import tqdm

from InforMI import CONFIG, select_ionapi_file, prompt_for_company_division, get_ion_token
from config import get_sqlite_db_path

# =============================================================================
# Clean description
# =============================================================================
def clean_description(desc):
    """
    Cleans M3 table description:
    - Removes leading category + colon (e.g., "TF: ")
    - Removes all parentheses and everything inside
    - Removes trailing 0/, 1/, or any remaining /
    - Strips extra whitespace
    """
    if not desc:
        return ""
    
    # Remove leading category + colon
    desc = re.sub(r'^[A-Z]{2}:\s*', '', desc)
    
    # Remove parentheses and everything inside
    desc = re.sub(r'\(.*?\)', '', desc)

    # Remove trailing 0/, 1/, 2/, or any single /
    desc = re.sub(r'\s*(0/|1/|2/|/)\s*$', '', desc)
    
    # Strip leading/trailing whitespace
    return desc.strip()

# =============================================================================
# Fetch Tables from M3
# =============================================================================

def get_tables():
    """
    Calls the M3 analytics REST API to get the list of tables.
    Returns the JSON response.
    """
    conn = http.client.HTTPSConnection(CONFIG["iu"].replace("https://", ""))

    endpoint = f"/{CONFIG['ti']}/M3/mdprest/analytics/getTables"

    headers = {
        'accept': 'application/json',
        'Authorization': f'Bearer {CONFIG["access_token"]}'
    }

    conn.request("GET", endpoint, "", headers)
    res = conn.getresponse()
    data = res.read()
    return json.loads(data.decode("utf-8"))


# =============================================================================
# Save to SQLite
# =============================================================================

def append_to_m3tables(cur, tables_data):
    """
    Appends table metadata into M3Tables table.
    Clears previous table records first to avoid duplicates.
    Stores tableName, description, and category.
    """
    cur.execute("""
        CREATE TABLE IF NOT EXISTS M3Tables (
            tableName   TEXT,
            description TEXT,
            category    TEXT
        )
    """)

    insert_sql = "INSERT INTO M3Tables (tableName, description, category) VALUES (?, ?, ?)"

    tables_list = tables_data.get("list", [])  # Adjust key if JSON structure differs
    if not tables_list:
        print("⚠️ No tables found in API response.")
        return

    # Delete existing records for all table names we are about to insert
    table_names = [t.get("tableName") for t in tables_list if t.get("tableName")]
    if table_names:
        cur.execute(
            f"DELETE FROM M3Tables WHERE tableName IN ({','.join(['?']*len(table_names))})",
            table_names
        )

    for table in tqdm(tables_list, desc="Appending tables"):
        description = clean_description(table.get("description"))
        cur.execute(insert_sql, (
            table.get("tableName"),
            description,
            table.get("category")
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
        tables_data = get_tables()

        with sqlite3.connect(SQLITE_DB_PATH) as conn:
            cur = conn.cursor()
            append_to_m3tables(cur, tables_data)
            conn.commit()

        print(f"✅ Table metadata appended to M3Tables in database '{SQLITE_DB_PATH}'")

    except Exception as e:
        print(f"❌ Failed to fetch/save table metadata: {e}")