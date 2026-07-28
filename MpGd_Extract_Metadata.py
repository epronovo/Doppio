# MpGd_Extract_Metadata.py
# -----------------------------------------------------------------------
# PURPOSE
#   Authenticates via IonAPI OAuth and fetches three M3 reference metadata
#   sets from the M3 MDP REST endpoint: programs (m3Programs), reference
#   fields (m3ReferenceFields), and tables (m3Tables).  Only records whose
#   primary key does not already exist are inserted (INSERT OR IGNORE).
#   After loading, syncs corrected descriptions from m3ReferenceDescription
#   into m3ReferenceFields to overwrite any stale values.
#
# INPUTS
#   - M3 IonAPI MDP REST endpoints: getPrograms, getReferenceFields,
#     getTables (authenticated HTTP GET)
#   - m3ReferenceDescription (SQLite) — authoritative descriptions used
#     in the post-load sync
#   - ionapi/ folder — .ionapi tenant config files
#
# OUTPUTS
#   - m3Programs        (SQLite) — M3 program list
#   - m3ReferenceFields (SQLite) — reference field definitions (descriptions
#                                   updated from m3ReferenceDescription)
#   - m3Tables          (SQLite) — M3 table list
#
# DEPENDENCIES
#   - config.py, InforMI.py, UserDefaults.py
# -----------------------------------------------------------------------
import logging
import time
import sqlite3
import requests
from pathlib import Path

logger = logging.getLogger(__name__)

from InforMI import (
    select_ionapi_file,
    CONFIG,
    get_ion_token,
    prompt_for_company_division,
    get_m3_mdprest_base_url,
    post_to_m3
)
from UserDefaults import load_user_defaults, save_user_defaults
from config import BASE_DIR, get_sqlite_db_path
from tqdm import tqdm

# Added getReferenceDescription to the loop so we have the source data for the fix
EXTRACTS = [
    {"name": "getPrograms", "endpoint": "getPrograms", "table": "m3Programs"},
    {"name": "getReferenceFields", "endpoint": "getReferenceFields", "table": "m3ReferenceFields"},
    {"name": "getTables", "endpoint": "getTables", "table": "m3Tables"},
]

# ============================================================
# Helpers
# ============================================================

def fetch_metadata(endpoint, base_url):
    url = f"{base_url}/{endpoint}"
    headers = {
        "Authorization": f"Bearer {CONFIG['access_token']}",
        "Accept": "application/json",
    }
    resp = requests.get(url, headers=headers)
    if resp.status_code != 200:
        logger.error(f"❌ Error fetching {endpoint}")
        resp.raise_for_status()
    return resp.json()


def ensure_table_exists(cur, table_name, sample_row):
    """Creates table if missing, defining the first column as PRIMARY KEY."""
    columns = list(sample_row.keys())
    pk_col = columns[0]
    
    col_defs = []
    for col in columns:
        if col == pk_col:
            col_defs.append(f"{col} TEXT PRIMARY KEY")
        else:
            col_defs.append(f"{col} TEXT")
            
    create_sql = f"CREATE TABLE IF NOT EXISTS {table_name} ({', '.join(col_defs)})"
    cur.execute(create_sql)


def insert_missing_rows(cur, table_name, rows):
    """Inserts only records where the Primary Key doesn't already exist."""
    if not rows: return 0
    
    columns = list(rows[0].keys())
    placeholders = ",".join(["?"] * len(columns))
    
    # "INSERT OR IGNORE" is the magic that skips existing keys
    insert_sql = f"INSERT OR IGNORE INTO {table_name} ({', '.join(columns)}) VALUES ({placeholders})"
    
    cur.execute(f"SELECT COUNT(*) FROM {table_name}")
    before = cur.fetchone()[0]
    
    for row in rows:
        cur.execute(insert_sql, [row.get(col) for col in columns])
        
    cur.execute(f"SELECT COUNT(*) FROM {table_name}")
    return cur.fetchone()[0] - before


def sync_reference_descriptions(cur):
    """Fixes descriptions in m3ReferenceFields using m3ReferenceDescription."""
    logger.info("🧹 Syncing reference descriptions...")
    update_sql = """
        UPDATE m3ReferenceFields
        SET referenceFieldDescription = (
            SELECT src.referenceFieldDescription
            FROM m3ReferenceDescription AS src
            WHERE src.referenceFieldName = m3ReferenceFields.referenceFieldName
        )
        WHERE EXISTS (
            SELECT 1
            FROM m3ReferenceDescription AS src
            WHERE src.referenceFieldName = m3ReferenceFields.referenceFieldName
            AND src.referenceFieldDescription IS NOT NULL
        );
    """
    cur.execute(update_sql)
    logger.info(f"✅ Updated {cur.rowcount} descriptions.")

# ============================================================
# Main
# ============================================================

def main():
    # ============================================================
    # Setup / Config  (runs only when executed directly)
    # ============================================================
    ionapi_dir = Path(__file__).parent / "ionapi"

    CONFIG["tenant"] = select_ionapi_file(ionapi_dir)
    prompt_for_company_division()

    defaults = load_user_defaults()
    save_user_defaults(defaults)

    get_ion_token()

    sqlite_db_path = get_sqlite_db_path()
    base_url = get_m3_mdprest_base_url()

    start_time = time.time()

    with sqlite3.connect(sqlite_db_path) as conn:
        cur = conn.cursor()

        for extract in EXTRACTS:
            name = extract["name"]
            endpoint = extract["endpoint"]
            table = extract["table"]

            logger.info(f"🛠️  Extracting {name} → {table}")

            try:
                data = fetch_metadata(endpoint, base_url)
                rows = data.get("list", [])

                if not rows:
                    logger.warning(f"⚠️  {name}: empty response")
                    continue

                ensure_table_exists(cur, table, rows[0])
                added = insert_missing_rows(cur, table, rows)

                conn.commit()
                logger.info(f"✅ {table}: {added} new records added.")

            except Exception as e:
                logger.error(f"❌ Error processing {name}: {e}")

        # Post-process the description fix
        try:
            sync_reference_descriptions(cur)
            conn.commit()
        except Exception as e:
            logger.error(f"❌ Error syncing descriptions: {e}")

    elapsed = time.time() - start_time
    m, s = divmod(elapsed, 60)
    logger.info(f"✅ Total run time: {int(m):02d}:{int(s):02d}")


if __name__ == "__main__":
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s  %(levelname)-8s  %(message)s",
        datefmt="%H:%M:%S"
    )
    main()