# ExtractAPI.py
import os
import time
import sqlite3
import requests

from tqdm import tqdm
from pathlib import Path
from InforMI import select_ionapi_file, CONFIG, get_ion_token, post_to_m3, prompt_for_company_division
from UserDefaults import load_user_defaults, save_user_defaults
from config import BASE_DIR, get_sqlite_db_path

# initialize field
print(f"")
QUERIES_DIR = BASE_DIR / "queries/api"
ionapi_dir = Path(__file__).parent / "ionapi"
CONFIG["tenant"] = select_ionapi_file(ionapi_dir)
prompt_for_company_division()
SQLITE_DB_PATH = get_sqlite_db_path()


def extract_table(cur, metadata_file, session):
    # Step 1: Read and parse the metadata file
    with open(metadata_file, "r") as f:
        line = f.read().strip()

    if " from " not in line.lower():
        raise ValueError(f"Invalid format in file {metadata_file}: expected 'column1,column2 from TABLENAME'")

    # Normalize casing for splitting, preserve original
    lower_line = line.lower()

    # Split to get columns and rest of query
    pre_from, post_from = line[:lower_line.index(" from ")], line[lower_line.index(" from ") + 6:]

    # Determine if there's a WHERE clause
    if " where " in post_from.lower():
        table_part, where_part = post_from[:post_from.lower().index(" where ")], post_from[post_from.lower().index(" where "):]
    else:
        table_part, where_part = post_from, ""

    columns = [col.strip() for col in pre_from.split(",")]

    # Use filename (without extension) as table name
    table_name = os.path.splitext(os.path.basename(metadata_file))[0]

    # print(f"Processing file: {metadata_file}")
    # print(f"Creating table: {table_name}")

    # Step 2: Drop table if exists
    cur.execute(f"DROP TABLE IF EXISTS {table_name}")

    # Step 3: Create table
    create_sql = f"CREATE TABLE {table_name} ({', '.join(col + ' TEXT' for col in columns)})"
    cur.execute(create_sql)

    # Step 4: Prepare API payload (keep original table from SQL line for QERY)
    source_table = table_part.strip()
    query_string = f"{','.join(columns)} from {source_table}{where_part}"
    payload = {
        "program": "EXPORTMI",
        "transactions": [{
            "transaction": "Select",
            "record": {
                "SEPC": "^",
                "HDRS": "0",
                "QERY": query_string
            },
            "selectedColumns": ["REPL"]
        }]
    }

    # Step 5: Insert data
    placeholders = ", ".join(["?"] * len(columns))
    insert_sql = f"INSERT INTO {table_name} ({', '.join(columns)}) VALUES ({placeholders})"

    print(f"🛠️  Extracting and inserting data for table: {table_name} ...")
    data = post_to_m3(payload, session)
    records = data['results'][0]['records']
    total = len(records)
    if total == 0:
        print(f"No records found for {table_name}.")
        return

    with tqdm(total=total, desc=f"🛠️  Processing {table_name}") as pbar:
        for record in records:
            cur.execute(insert_sql, record["REPL"].split("^")[:len(columns)])
            pbar.update(1)
    cur.connection.commit()
    # print(f"✅ Finished processing {table_name} ({total} records).")


def process_api_files():
    return [
        os.path.join(QUERIES_DIR, f)
        for f in os.listdir(QUERIES_DIR)
        if f.endswith(".api")
    ]

if __name__ == "__main__":

    # Load user defaults once
    defaults = load_user_defaults()
    save_user_defaults(defaults)

    # === Get token and prepare session ===
    start_time = time.time()
    token = get_ion_token()
    session = requests.Session()
    session.headers.update({
        "Authorization": f"Bearer {token}",
        "Accept": "application/json",
        "Content-Type": "application/json"
    })

    api_files = process_api_files()

    with sqlite3.connect(SQLITE_DB_PATH) as conn:
        cur = conn.cursor()
        for filepath in api_files:
            try:
                extract_table(cur, filepath, session)
            except Exception as e:
                print(f"❌ Error processing {filepath}: {e}")

    elapsed = time.time() - start_time
    hours, rem = divmod(elapsed, 3600)
    minutes, seconds = divmod(rem, 60)
    print(f"✅ Total run time: {int(hours):02d}:{int(minutes):02d}:{int(seconds):02d} (hh:mm:ss)")