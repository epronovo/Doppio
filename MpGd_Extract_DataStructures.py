# MpGd_Extract_DataStructures.py
# -----------------------------------------------------------------------
# PURPOSE
#   For each M3 Data Structure referenced in ExtractRules that has not yet
#   been extracted, fetches the field layout from the M3 MDP
#   ViewDSRelation/getDSFields JSON endpoint via Selenium and stores the
#   field definitions (name, type, length, from/to offsets) in
#   m3DataStructures.  Requires an active SSO browser session.
#
# INPUTS
#   - ExtractRules      (SQLite) — DataStructure column identifies which DS
#                                  names to fetch
#   - m3DataStructures  (SQLite) — checked to skip already-extracted DS names
#   - M3 MDP /ViewDSRelation?operation=getDSFields endpoint (via Selenium)
#
# OUTPUTS
#   - m3DataStructures  (SQLite) — DataStructure, FieldName, Type, Length,
#                                  dsFrom, dsTo
#
# DEPENDENCIES
#   - config.py, Selenium / ChromeDriver, active M3 SSO browser session
# -----------------------------------------------------------------------

import json
import logging
import sqlite3
import tempfile
import time
from typing import List

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By

from config import get_sqlite_db_path

logger = logging.getLogger(__name__)

# -------------------------------------------------------------------
# Config
# -------------------------------------------------------------------
DB_PATH = get_sqlite_db_path()

BASE_URL = (
    "https://m3-cm3xprduse1b.m32.m3.us1.mprd.inforcloudsuite.com"
    "/mdp/ViewDSRelation"
    "?operation=getDSFields"
    "&beSystem=M3BE_16.0_ALL"
    "&dsName={ds}"
    "&component=MVX"
)

BATCH_SIZE = 50

# -------------------------------------------------------------------
# Selenium setup
# -------------------------------------------------------------------
def get_driver() -> webdriver.Chrome:
    options = Options()
    options.add_argument(f"--user-data-dir={tempfile.mkdtemp()}")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    return webdriver.Chrome(options=options)

# -------------------------------------------------------------------
# DB utilities
# -------------------------------------------------------------------
def ensure_table():
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute("""
        CREATE TABLE IF NOT EXISTS m3DataStructures (
            DataStructure TEXT,
            FieldName     TEXT,
            Type          TEXT,
            Length        INTEGER,
            dsFrom        INTEGER,
            dsTo          INTEGER,
            PRIMARY KEY (DataStructure, FieldName)
        );
    """)
    conn.commit()
    conn.close()

def get_missing_datastructures() -> List[str]:
    """
    Return DataStructures from ExtractRules that do not yet exist
    in m3DataStructures.
    """
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute("""
        SELECT DISTINCT r.DataStructure
        FROM ExtractRules r
        LEFT JOIN m3DataStructures ds
            ON ds.DataStructure = r.DataStructure
        WHERE r.DataStructure <> ''
          AND ds.DataStructure IS NULL
    """)
    rows = [row[0] for row in cur.fetchall()]
    conn.close()
    return rows

def save_batch(rows: List[tuple]):
    if not rows:
        return

    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.executemany("""
        INSERT INTO m3DataStructures
            (DataStructure, FieldName, Type, Length, dsFrom, dsTo)
        VALUES (?, ?, ?, ?, ?, ?)
        ON CONFLICT(DataStructure, FieldName) DO UPDATE SET
            Type   = excluded.Type,
            Length = excluded.Length,
            dsFrom = excluded.dsFrom,
            dsTo   = excluded.dsTo
    """, rows)
    conn.commit()
    conn.close()

# -------------------------------------------------------------------
# Extraction helpers
# -------------------------------------------------------------------
def fetch_ds_json(driver: webdriver.Chrome, ds_name: str) -> str:
    url = BASE_URL.format(ds=ds_name)
    driver.get(url)
    time.sleep(0.5)

    body = driver.find_element(By.TAG_NAME, "body")
    return body.get_attribute("textContent")

def extract_fields_from_json(source: str) -> List[tuple]:
    data = json.loads(source)

    ds_name = data.get("resultElements", {}).get("entityName")
    dataset = data.get("dataset", [])

    rows = []
    for f in dataset:
        rows.append((
            ds_name,
            f.get("fieldName"),
            f.get("type"),
            int(f["length"]) if f.get("length") else None,
            int(f["from"]) if f.get("from") else None,
            int(f["to"]) if f.get("to") else None
        ))

    return rows

# -------------------------------------------------------------------
# Runner
# -------------------------------------------------------------------
def run():
    ensure_table()

    datastructures = get_missing_datastructures()
    if not datastructures:
        logger.info("✅ No missing data structures found — nothing to do.")
        return

    logger.info(f"🔍 Found {len(datastructures)} data structures to process")

    driver = get_driver()

    logger.info("🔐 Opening browser for SSO login...")
    driver.get("https://mingle-portal.inforcloudsuite.com/v2/WARCK89PZ9V6G3KK_TST/")
    input("👉 Complete SSO login, then press ENTER to continue...")

    batch_rows: List[tuple] = []
    processed = 0

    logger.info("🚀 Starting Data Structure extraction")

    for ds in datastructures:
        try:
            source = fetch_ds_json(driver, ds)
            rows = extract_fields_from_json(source)

            if rows:
                batch_rows.extend(rows)
                logger.info(f"✔ {ds}: {len(rows)} fields extracted")
            else:
                logger.warning(f"⚠️  {ds}: no fields returned")

            processed += 1

            if len(batch_rows) >= BATCH_SIZE:
                save_batch(batch_rows)
                logger.info(f"💾 Saved batch of {len(batch_rows)} rows")
                batch_rows.clear()

        except Exception as e:
            logger.error(f"❌ {ds}: {e}")

    if batch_rows:
        save_batch(batch_rows)
        logger.info(f"💾 Saved final batch of {len(batch_rows)} rows")

    driver.quit()
    logger.info(f"🎉 Done. Data structures processed: {processed}")

# -------------------------------------------------------------------
# Entry point
# -------------------------------------------------------------------
if __name__ == "__main__":
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s  %(levelname)-8s  %(message)s",
        datefmt="%H:%M:%S"
    )
    run()
