# ViewFixes_Basic.py
import json
import time
import sqlite3
import requests
import pprint
import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

from tqdm import tqdm
from pathlib import Path
from InforMI import select_ionapi_file, CONFIG, prompt_for_company_division
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
    cur.execute("DROP TABLE IF EXISTS LstAllColumns_ST")
    cur.execute("""
        CREATE TABLE LstAllColumns_ST (
            PGNM TEXT, RESP TEXT, PAVR TEXT, PIC1 TEXT,
            COID TEXT, FLDI TEXT, FLTY TEXT, MSID TEXT,
            FLDL TEXT, FLLU TEXT, FNDC TEXT, UNDC TEXT,
            ECDE TEXT, ECDU TEXT, SMFL TEXT, MXAV TEXT,
            SUMA TEXT, OMAX TEXT, SOSQ TEXT, IBCA TEXT,
            FDCA TEXT, DCNY TEXT, UDCN TEXT, EDUS TEXT,
            EDFL TEXT, AGDR TEXT, AGUR TEXT, SUBR TEXT,
            DTFR TEXT, DCFM TEXT, FLDH TEXT, DDBX TEXT
        )
    """)
    cur.connection.commit()


def extract_lstallcolumns(cur, session, ionapi):
    cur.execute("SELECT C9PGNM, C9PAVR, C9PIC1 FROM CSYSPV")
    driving_rows = cur.fetchall()
    if not driving_rows:
        print("❌ No rows found in CSYSPV")
        return

    insert_cols = [d[1] for d in cur.execute("PRAGMA table_info(LstAllColumns_ST)")]
    placeholders = ", ".join(["?"] * len(insert_cols))
    insert_sql = f"INSERT INTO LstAllColumns_ST ({', '.join(insert_cols)}) VALUES ({placeholders})"

    with tqdm(total=len(driving_rows), desc="Calling CRS020MI/LstAllColumns") as pbar:
        for pgmn, pavr, pic1 in driving_rows:
            url = f"{ionapi['api_url']}/CRS020MI/LstAllColumns"
            params = {"PGNM": pgmn, "PAVR": pavr, "PIC1": pic1}

            headers = {
                "Authorization": f"Basic {ionapi['basic_auth']}",
                "Accept": "application/json; charset=UTF-8",
                "Content-Type": "application/json; charset=UTF-8"
            }
            
            # Build the full request to see the final URL
            req = requests.Request("GET", url, headers=headers, params=params)
            prepared = session.prepare_request(req)
            # print(f"➡️ Calling URL: {prepared.url}")
            
            resp = session.get(url, headers=headers, params=params, verify=False)
            resp.raise_for_status()
            data = resp.json()
            
            # Pretty print the response for visibility
            # print("📦 Response:")
            # pprint.pprint(data)
            
            rows = []
            for record in data.get("MIRecord", []):
                values = {nv["Name"]: nv.get("Value", "").strip()
                          for nv in record.get("NameValue", [])}
                row = tuple(values.get(col, "") for col in insert_cols)
                rows.append(row)

            if rows:
                cur.executemany(insert_sql, rows)

            pbar.update(1)

    cur.connection.commit()
    print("✅ Finished writing to LstAllColumns_ST")


if __name__ == "__main__":
    defaults = load_user_defaults()
    save_user_defaults(defaults)

    start_time = time.time()
    ionapi_path = CONFIG["tenant"]
    with open(ionapi_path, "r") as f:
        ionapi = json.load(f)

    session = requests.Session()

    with sqlite3.connect(SQLITE_DB_PATH) as conn:
        cur = conn.cursor()
        create_output_table(cur)
        if ionapi.get("atype") == "basic":
            extract_lstallcolumns(cur, session, ionapi)
        else:
            print("⚠️ This script is for Basic Auth JSON calls. Use ViewFixes.py for OAuth2.")

    elapsed = time.time() - start_time
    hours, rem = divmod(elapsed, 3600)
    minutes, seconds = divmod(rem, 60)
    print(f"✅ Total run time: {int(hours):02d}:{int(minutes):02d}:{int(seconds):02d} (hh:mm:ss)")