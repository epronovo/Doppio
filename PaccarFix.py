# PaccarFix.py
import sqlite3

from datetime import datetime
from pathlib import Path
from tqdm import tqdm
from InforMI import CONFIG, get_ion_token, post_to_m3, prompt_for_company_division, select_ionapi_file
from config import BASE_DIR, get_sqlite_db_path  # adjust import if needed

# initialize field
print(f"")
QUERIES_DIR = BASE_DIR / "queries/api"
ionapi_dir = Path(__file__).parent / "ionapi"
CONFIG["tenant"] = select_ionapi_file(ionapi_dir)
prompt_for_company_division()
SQLITE_DB_PATH = get_sqlite_db_path()

# =============================================================================
# PACCAR Routine
# =============================================================================

def fetch_orders(db_path):
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    cursor.execute("""
        SELECT DISTINCT ORNO FROM fixes ORDER BY ORNO
    """)
    orders = cursor.fetchall()
    conn.close()
    return orders


def ensure_text_id(order_no, session):
    """Return TXID for an order, creating if necessary."""
    # Step 1: GetTextID
    payload = {
        "program": "CRS980MI",
        "transactions": [
            {
                "transaction": "GetTextID",
                "record": {
                    "FILE": "OOHEAD00",
                    "KV01": CONFIG["company"],   # ✅ company from config
                    "KV02": str(order_no),
                    "FLDI": "POTX"
                },
                "selectedColumns": ["TXID"]
            }
        ]
    }
    result = post_to_m3(payload, session)
    records = result.get("results", [])[0].get("records", [])
    txid = int(records[0]["TXID"]) if records else 0

    if txid != 0:
        return txid

    # Step 2: RtvNewTextID
    payload = {
        "program": "CRS980MI",
        "transactions": [
            {
                "transaction": "RtvNewTextID",
                "record": {"FILE": "OSYTXH"},
                "selectedColumns": ["TXID"]
            }
        ]
    }
    result = post_to_m3(payload, session)
    records = result.get("results", [])[0].get("records", [])
    txid = int(records[0]["TXID"]) if records else 0

    # Step 3: SetTextID
    payload = {
        "program": "CRS980MI",
        "transactions": [
            {
                "transaction": "SetTextID",
                "record": {
                    "FILE": "OOHEAD00",
                    "TXID": str(txid),
                    "KV01": CONFIG["company"],
                    "KV02": str(order_no),
                    "FLDI": "POTX"
                }
            }
        ]
    }
    post_to_m3(payload, session)

    return txid

def add_text_blocks(txid, session):
    """Add header + line text blocks for a given TXID."""

    # AddTxtBlockHead
    payload = {
        "program": "CRS980MI",
        "transactions": [
            {
                "transaction": "AddTxtBlockHead",
                "record": {
                    "TXID": str(txid),
                    "FILE": "OOHEAD00",
                    "KFLD": "POTX",
                    "USID": "USEPRON",   # 👈 adjust user ID if needed
                    "TFIL": "OSYTXH",
                    "TXVR": "CO01",
                    "LNCD": "GB",
                    "TX40": "Go Live Fix",
                    "TXEI": "2"
                }
            }
        ]
    }
    post_to_m3(payload, session)

    # AddTxtBlockLine 1
    today_str = datetime.today().strftime("%m/%d/%y")
    payload = {
        "program": "CRS980MI",
        "transactions": [
            {
                "transaction": "AddTxtBlockLine",
                "record": {
                    "TXID": str(txid),
                    "TFIL": "OSYTXH",
                    "FILE": "OOHEAD00",
                    "TXVR": "CO01",
                    "LNCD": "GB",
                    "LINO": "1",
                    "TX60": f"One or more lines of this order was cancelled on"
                }
            }
        ]
    }
    post_to_m3(payload, session)

    # AddTxtBlockLine 2
    payload = {
        "program": "CRS980MI",
        "transactions": [
            {
                "transaction": "AddTxtBlockLine",
                "record": {
                    "TXID": str(txid),
                    "TFIL": "OSYTXH",
                    "FILE": "OOHEAD",
                    "TXVR": "CO01",
                    "LNCD": "GB",
                    "LINO": "2",
                    "TX60": "October 19/20, 2025 (2025-10-19/20) for the"
                }
            }
        ]
    }
    post_to_m3(payload, session)
    # AddTxtBlockLine 3
    payload = {
        "program": "CRS980MI",
        "transactions": [
            {
                "transaction": "AddTxtBlockLine",
                "record": {
                    "TXID": str(txid),
                    "TFIL": "OSYTXH",
                    "FILE": "OOHEAD",
                    "TXVR": "CO01",
                    "LNCD": "GB",
                    "LINO": "3",
                    "TX60": "Heavy Truck EDI Order Duplication cleanup."
                }
            }
        ]
    }
    post_to_m3(payload, session)

def run_paccar():
    import requests
    from tqdm import tqdm

    orders = fetch_orders(SQLITE_DB_PATH)
    tqdm.write(f"Found {len(orders)} orders to process.")

    # Acquire token
    get_ion_token()  # sets CONFIG['access_token']

    # Create session with auth headers
    with requests.Session() as session:
        session.headers.update({
            "Authorization": f"Bearer {CONFIG['access_token']}",
            "Accept": "application/json",
            "Content-Type": "application/json"
        })

        # ✅ Unpack (order_no, order_line) from each row
        for (order_no,) in tqdm(orders, desc="Processing orders"):
            try:
                txid = ensure_text_id(order_no, session)
                add_text_blocks(txid, session)
            except Exception as e:
                tqdm.write(f"❌ Failed on order {order_no}: {e}")

# =============================================================================
# Entry Point
# =============================================================================

if __name__ == "__main__":
    # ionapi_dir = Path(__file__).parent / "ionapi"
    # CONFIG["tenant"] = select_ionapi_file(ionapi_dir)
    # prompt_for_company_division()
    # get_ion_token()
    run_paccar()