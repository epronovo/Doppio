# ExtractAPIs_MRS001.py
#
# this is a standalone script that extracts API information from MRS001MI
# and populates three tables: cmipgm, cmitrn, cmifld

import sqlite3
import time
import requests

from tqdm import tqdm
from pathlib import Path
from InforMI import select_ionapi_file, CONFIG, get_ion_token, post_to_m3, prompt_for_company_division
from UserDefaults import load_user_defaults, save_user_defaults
from config import get_sqlite_db_path

ionapi_dir = Path(__file__).parent / "ionapi"
CONFIG["tenant"] = select_ionapi_file(ionapi_dir)
prompt_for_company_division()
SQLITE_DB_PATH = get_sqlite_db_path()

def insert_records(cur, table, columns, records):
    """Generic helper to insert records into SQLite."""
    placeholders = ", ".join(["?"] * len(columns))
    insert_sql = f"INSERT INTO {table} ({', '.join(columns)}) VALUES ({placeholders})"
    for record in records:
        cur.execute(insert_sql, [record.get(col, None) for col in columns])
    cur.connection.commit()


def ensure_table_exists(cur, table_name, columns):
    """Create table if it doesn't exist yet."""
    create_sql = f"""
        CREATE TABLE IF NOT EXISTS {table_name} (
            {', '.join(f'{c} TEXT' for c in columns)}
        )
    """
    cur.execute(create_sql)


def load_cmipgm(cur, data):
    """Insert cmipgm data (append only)"""
    records = data.get("records", [])
    if not records:
        return

    columns = ["MINM", "OBNM", "MIDS", "TXT1", "TXT2", "RGDT", "RGTM",
               "VERS", "LMDT", "CHNO", "CHID", "MNID"]
    ensure_table_exists(cur, "cmipgm", columns)

    insert_records(cur, "cmipgm", columns, records)


def load_cmitrn(cur, data):
    """Insert cmitrn data (append only)"""
    records = data.get("records", [])
    if not records:
        return

    columns = ["MINM", "TRNM", "TRDS", "TXT1", "TXT2", "VERS", "STAT",
               "PRFI", "PRFO", "SIMU", "RGDT", "RGTM", "LMDT", "CHNO", "CHID"]
    ensure_table_exists(cur, "cmitrn", columns)

    insert_records(cur, "cmitrn", columns, records)


def load_cmifld(cur, data):
    """Insert cmifld data (append only)"""
    records = data.get("records", [])
    if not records:
        return

    columns = ["MINM", "TRNM", "TRTP", "FLNM", "FLDS", "TXT1",
               "FRPO", "TOPO", "LENG", "TYPE", "MAND", "RGDT",
               "RGTM", "LMDT", "CHNO", "CHID"]
    ensure_table_exists(cur, "cmifld", columns)

    insert_records(cur, "cmifld", columns, records)


def call_m3(session, endpoint):
    """Generic API call to M3 using your post_to_m3 helper."""
    program = "MRS001MI"
    transaction = endpoint.split(";")[0]  # e.g. "LstPrograms"
    params_part = endpoint.split("?")[-1] if "?" in endpoint else ""
    params = {}

    if params_part:
        for pair in params_part.split("&"):
            if "=" in pair:
                k, v = pair.split("=", 1)
                params[k] = v

    payload = {
        "program": program,
        "transactions": [{
            "transaction": transaction,
            "record": params,
        }]
    }

    response = post_to_m3(payload, session)
    # M3 MI returns results in this nested format
    if "results" in response and len(response["results"]) > 0:
        return response["results"][0]
    else:
        return {"records": []}


def extract_apis(session):
    """Equivalent of Java extractApis()"""
    start_time = time.time()
    with sqlite3.connect(SQLITE_DB_PATH) as conn:
        cur = conn.cursor()

        # === Step 1: LstPrograms ===
        print("🔹 Extracting MRS001MI/LstPrograms ...")
        url = "LstPrograms;maxrecs=0?"
        api_response = call_m3(session, url)
        load_cmipgm(cur, api_response)

        records = api_response.get("records", [])
        if not records:
            print("⚠️ No programs found.")
            return

        # === Step 2: For each MINM, get transactions ===
        for record in tqdm(records, desc="📦 Processing programs"):
            minm = record.get("MINM")
            if not minm:
                continue
            extract_transactions(session, cur, minm)

    elapsed = time.time() - start_time
    print(f"✅ Completed in {elapsed:.2f} seconds.")


def extract_transactions(session, cur, minm):
    """Equivalent of apiTransactions(minm)"""
    # print(f"🔹 Extracting MRS001MI/LstTransactions for {minm} ...")
    url = f"LstTransactions;maxrecs=0?MINM={minm}"
    api_response = call_m3(session, url)
    load_cmitrn(cur, api_response)

    transactions = api_response.get("records", [])
    if not transactions:
        return

    for trn in tqdm(transactions, desc=f"📦 Transactions for {minm}", leave=False):
        trnm = trn.get("TRNM")
        if not trnm:
            continue
        extract_fields(session, cur, minm, trnm, "I")
        extract_fields(session, cur, minm, trnm, "O")


def extract_fields(session, cur, minm, trnm, trtp):
    """Equivalent of apiFields(minm, trnm, trtp)"""
    url = f"LstFields;maxrecs=0?MINM={minm}&TRNM={trnm}&TRTP={trtp}"
    # print(f"   ↳ Extracting Fields {trtp} for {minm}/{trnm}")
    api_response = call_m3(session, url)
    load_cmifld(cur, api_response)


if __name__ == "__main__":
    defaults = load_user_defaults()
    save_user_defaults(defaults)

    # === Prepare M3 session ===
    token = get_ion_token()
    session = requests.Session()
    session.headers.update({
        "Authorization": f"Bearer {token}",
        "Accept": "application/json",
        "Content-Type": "application/json"
    })

    extract_apis(session)