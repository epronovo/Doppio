# ViewFixes.py
import json
import time
import sqlite3
import requests
from tqdm import tqdm
from pathlib import Path

from InforMI import select_ionapi_file, CONFIG, get_ion_token, post_to_m3, prompt_for_company_division
from UserDefaults import load_user_defaults, save_user_defaults
from config import get_sqlite_db_path

# =============================================================================
# Setup
# =============================================================================
ionapi_dir = Path(__file__).parent / "ionapi"
CONFIG["tenant"] = select_ionapi_file(ionapi_dir)
prompt_for_company_division()
SQLITE_DB_PATH = get_sqlite_db_path()


def create_output_table(cur):
    cur.execute("DROP TABLE IF EXISTS LstAllColumns")
    cur.execute("""
        CREATE TABLE LstAllColumns (
            PGNM TEXT, PAVR TEXT, PIC1 TEXT,
            RESP TEXT, IBCA TEXT, COID TEXT, FLDI TEXT, FLTY TEXT, MSID TEXT,
            FLDL TEXT, FLLU TEXT, FNDC TEXT, UNDC TEXT, ECDE TEXT, ECDU TEXT,
            SMFL TEXT, MXAV TEXT, SUMA TEXT, OMAX TEXT, SOSQ TEXT, FDCA TEXT,
            DCNY TEXT, UDCN TEXT, EDUS TEXT, EDFL TEXT, AGDR TEXT, AGUR TEXT,
            SUBR TEXT, DTFR TEXT, DCFM TEXT, FLDH TEXT, DDBX TEXT
        )
    """)
    cur.connection.commit()


def extract_lstallcolumns(cur, session):
    cur.execute("SELECT C9PGNM, C9PAVR, C9PIC1 FROM CSYSPV")
    driving_rows = cur.fetchall()
    if not driving_rows:
        print("❌ No rows found in CSYSPV")
        return

    insert_cols = [d[1] for d in cur.execute("PRAGMA table_info(LstAllColumns)")]
    placeholders = ", ".join(["?"] * len(insert_cols))
    insert_sql = f"INSERT INTO LstAllColumns ({', '.join(insert_cols)}) VALUES ({placeholders})"

    with tqdm(total=len(driving_rows), desc="Calling API") as pbar:
        for pgmn, pavr, pic1 in driving_rows:
            payload = {
                "program": "CRS020MI",
                "transactions": [{
                    "transaction": "LstAllColumns",
                    "record": {
                        "PGNM": pgmn,
                        "PAVR": pavr,
                        "PIC1": pic1
                    },
                    "selectedColumns": insert_cols
                }]
            }

            data = post_to_m3(payload, session)
            results = data.get("results", [])
            if results:
                records = results[0].get("records", [])
                for record in records:
                    row = tuple(record.get(col, "") for col in insert_cols)
                    # print(insert_sql)
                    # print(row)
                    cur.execute(insert_sql, row)

            pbar.update(1)

    cur.connection.commit()
    print("✅ Finished writing to LstAllColumns")


if __name__ == "__main__":
    # Load/save defaults
    defaults = load_user_defaults()
    save_user_defaults(defaults)

    # Token + session
    start_time = time.time()
    token = get_ion_token()
    session = requests.Session()
    session.headers.update({
        "Authorization": f"Bearer {token}",
        "Accept": "application/json",
        "Content-Type": "application/json"
    })

    with sqlite3.connect(SQLITE_DB_PATH) as conn:
        cur = conn.cursor()
        create_output_table(cur)
        extract_lstallcolumns(cur, session)

    elapsed = time.time() - start_time
    hours, rem = divmod(elapsed, 3600)
    minutes, seconds = divmod(rem, 60)
    print(f"✅ Total run time: {int(hours):02d}:{int(minutes):02d}:{int(seconds):02d} (hh:mm:ss)")