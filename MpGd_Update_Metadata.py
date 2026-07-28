# MpGd_Update_Metadata.py
# -----------------------------------------------------------------------
# PURPOSE
#   Finds rows in m3ReferenceDescription where the description still
#   contains the stale "Reference field" label instead of the correct
#   "Reference fld" abbreviation, then queries the M3 MDP
#   SearchByConfig/RefFld endpoint via Selenium to retrieve the
#   authoritative description.  Updates are applied in batches and any
#   "Reference field" prefix in the returned text is normalised to
#   "Reference fld" before saving.
#
# INPUTS
#   - m3ReferenceDescription (SQLite) — rows with stale descriptions
#   - M3 MDP /SearchByConfig?operation=mdpSearch&searchType=RefFld
#     endpoint (via Selenium + SSO)
#
# OUTPUTS
#   - m3ReferenceDescription (SQLite) — referenceFieldDescription updated
#                                       in-place
#
# DEPENDENCIES
#   - config.py, Selenium / ChromeDriver, active M3 SSO browser session
# -----------------------------------------------------------------------

import json
import logging
import sqlite3
import tempfile
import time
from typing import List, Optional

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By

from config import get_sqlite_db_path

logger = logging.getLogger(__name__)

DB_PATH = get_sqlite_db_path()

BASE_URL = (
    "https://m3-cm3xprduse1b.m32.m3.us1.mprd.inforcloudsuite.com"
    "/mdp/SearchByConfig"
)

BATCH_SIZE = 50
SLEEP_SECONDS = 0.3

# -------------------------------------------------------------------
# Selenium setup (same pattern as MpGd_ExtractPrompts)
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
# -------------------------------------------------------------------
# DB utilities
# -------------------------------------------------------------------
def get_reference_field_names() -> List[str]:
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute("""
        SELECT DISTINCT referenceFieldName
        FROM m3ReferenceDescription
        WHERE referenceFieldName IS NOT NULL
        AND referenceFieldName <> ''
        AND referenceFieldDescription LIKE 'Reference field %'
        AND referenceFieldDescription not LIKE 'Reference fld %'
    """)
    rows = [row[0] for row in cur.fetchall()]
    conn.close()
    return rows


def update_description_batch(rows: List[tuple[str, str]]):
    """
    rows: [(description, referenceFieldName)]
    """
    if not rows:
        return

    # Ensure description starts with 'Reference fld' instead of 'Reference field'
    rows = [
        (desc.replace("Reference field", "Reference fld") if desc.startswith("Reference field") else desc, name)
        for desc, name in rows
    ]

    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.executemany("""
        UPDATE m3ReferenceDescription
           SET referenceFieldDescription = ?
         WHERE referenceFieldName = ?
    """, rows)
    conn.commit()
    conn.close()

# -------------------------------------------------------------------
# Selenium fetch
# -------------------------------------------------------------------
def fetch_reference_description(
    driver: webdriver.Chrome, ref_field: str
) -> Optional[str]:

    url = (
        f"{BASE_URL}"
        f"?operation=mdpSearch"
        f"&beSystem=M3BE_16.0_ALL"
        f"&searchType=RefFld"
        f"&searchComponent="
        f"&searchCategory="
        f"&searchInput={ref_field}"
        f"&programFilter="
        f"&tableFilter="
        f"&langId=GB"
        f"&returnLimit=500"
        f"&searchCounter=0"
    )

    driver.get(url)
    time.sleep(0.4)

    body = driver.find_element(By.TAG_NAME, "body")
    raw = body.get_attribute("textContent") or body.text

    payload = json.loads(raw)

    dataset = payload.get("dataset") or []
    if not dataset:
        return None

    return dataset[0].get("description")

# -------------------------------------------------------------------
# Runner
# -------------------------------------------------------------------
def run():
    fieldnames = get_reference_field_names()
    if not fieldnames:
        logger.info("✅ No reference fields found — nothing to do.")
        return

    logger.info(f"🔍 Found {len(fieldnames)} reference fields")

    driver = get_driver()

    logger.info("🔐 Opening browser for SSO login...")
    driver.get(
        "https://mingle-portal.inforcloudsuite.com/v2/WARCK89PZ9V6G3KK_TST/"
    )
    input("👉 Log in via SSO, then press ENTER here to continue...")

    logger.info("🚀 Updating reference field descriptions")

    batch_rows: List[tuple[str, str]] = []
    processed = 0
    updated = 0

    for field in fieldnames:
        try:
            description = fetch_reference_description(driver, field)
            processed += 1

            if description:
                batch_rows.append((description, field))
                updated += 1
                logger.info(f"✔ {field} → {description}")
            else:
                logger.warning(f"⚠️  {field}: no description returned")

            if len(batch_rows) >= BATCH_SIZE:
                update_description_batch(batch_rows)
                logger.info(f"💾 Saved batch of {len(batch_rows)}")
                batch_rows.clear()

            time.sleep(SLEEP_SECONDS)

        except Exception as e:
            logger.error(f"❌ {field}: {e}")

    if batch_rows:
        update_description_batch(batch_rows)
        logger.info(f"💾 Saved final batch of {len(batch_rows)}")

    driver.quit()

    logger.info(f"🎉 Done.  Processed: {processed}  Updated: {updated}")

# -------------------------------------------------------------------
# Run script
# -------------------------------------------------------------------
if __name__ == "__main__":
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s  %(levelname)-8s  %(message)s",
        datefmt="%H:%M:%S"
    )
    run()
