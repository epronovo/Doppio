# MpGd_Extract_Prompts.py
# -----------------------------------------------------------------------
# PURPOSE
#   For each M3 API field in m3Api2Table that does not yet have a prompt
#   entry, fetches the associated PromptProgram (the linked display/search
#   program) from the M3 foundation-rest cRef endpoint via Selenium.
#   Where a matching MI API exists in cmipgm, the API name is stored
#   alongside the prompt.  Results are persisted to m3Prompts in batches.
#
# INPUTS
#   - m3Api2Table  (SQLite) — source of field names to process
#   - cmipgm       (SQLite) — used to resolve PromptProgram → MI API name
#   - M3 foundation-rest /cRef{field}ext endpoint (via Selenium + SSO)
#
# OUTPUTS
#   - m3Prompts    (SQLite) — FieldName, PromptProgram, API
#
# DEPENDENCIES
#   - config.py, Selenium / ChromeDriver, active M3 SSO browser session
# -----------------------------------------------------------------------

import logging
import re
import sqlite3
import tempfile
import time
from typing import Optional, List

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By

from config import get_sqlite_db_path

logger = logging.getLogger(__name__)

DB_PATH = get_sqlite_db_path()
BASE_URL = (
    "https://m3-cm3xprduse1b.m32.m3.us1.mprd.inforcloudsuite.com"
    "/foundation-rest/src/MVX/util/cRef{field}ext"
)

APCALL_REGEX = re.compile(r'inter\.apCall\("([^"]+)"\)')

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
        CREATE TABLE IF NOT EXISTS m3Prompts (
            FieldName TEXT PRIMARY KEY,
            PromptProgram TEXT,
            API TEXT
        )
    """)
    conn.commit()
    conn.close()

def get_missing_fieldnames() -> List[str]:
    """
    Only return fields that do NOT already exist in m3Prompts
    """
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute("""
        SELECT DISTINCT t.FieldName
        FROM m3Api2Table t
        LEFT JOIN m3Prompts p
               ON p.FieldName = t.FieldName
        WHERE p.FieldName IS NULL
    """)
    rows = [row[0] for row in cur.fetchall()]
    conn.close()
    return rows

def save_prompt_batch(rows: List[tuple[str, Optional[str], Optional[str]]]):
    if not rows:
        return
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.executemany("""
        INSERT INTO m3Prompts (FieldName, PromptProgram, API)
        VALUES (?, ?, ?)
        ON CONFLICT(FieldName) DO UPDATE
            SET PromptProgram = excluded.PromptProgram,
                API = excluded.API
    """, rows)
    conn.commit()
    conn.close()

def get_api_from_cmipgm(prompt_program: str) -> Optional[str]:
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute("""
        SELECT MINM
        FROM cmipgm
        WHERE MINM = ? || 'MI'
        LIMIT 1
    """, (prompt_program,))
    row = cur.fetchone()
    conn.close()
    return row[0] if row else None

# -------------------------------------------------------------------
# Extraction
# -------------------------------------------------------------------
def extract_prompt_from_page(source: str) -> Optional[str]:
    match = APCALL_REGEX.search(source)
    return match.group(1) if match else None

def fetch_prompt(driver: webdriver.Chrome, field: str) -> Optional[str]:
    url = BASE_URL.format(field=field)
    driver.get(url)
    time.sleep(0.5)
    body = driver.find_element(By.TAG_NAME, "body")
    source = body.get_attribute("textContent") or body.text or driver.page_source
    return extract_prompt_from_page(source)

# -------------------------------------------------------------------
# Runner
# -------------------------------------------------------------------
def run():
    ensure_table()

    # 🔍 Get missing fields FIRST
    fieldnames = get_missing_fieldnames()
    if not fieldnames:
        logger.info("✅ No missing fields found — nothing to do.")
        return

    logger.info(f"Found {len(fieldnames)} missing fields to process")

    # 🚀 Start Selenium ONLY if needed
    driver = get_driver()

    logger.info("🔐 Opening browser for SSO login...")
    driver.get(
        "https://mingle-portal.inforcloudsuite.com/v2/WARCK89PZ9V6G3KK_TST/"
    )
    input("👉 Log in via SSO, then press ENTER here to continue...")

    logger.info("🚀 Starting prompt extraction")

    batch_rows: List[tuple[str, Optional[str], Optional[str]]] = []
    processed = 0

    for field in fieldnames:
        try:
            program = fetch_prompt(driver, field)
            api = get_api_from_cmipgm(program) if program else None

            batch_rows.append((field, program, api))
            processed += 1

            if program:
                logger.info(f"✔ {field} → {program}, API={api}")
            else:
                logger.warning(f"⚠️  {field}: no prompt program found")

            if len(batch_rows) >= BATCH_SIZE:
                save_prompt_batch(batch_rows)
                logger.info(f"💾 Saved batch of {len(batch_rows)}")
                batch_rows.clear()

        except Exception as e:
            logger.error(f"❌ {field}: {e}")

    if batch_rows:
        save_prompt_batch(batch_rows)
        logger.info(f"💾 Saved final batch of {len(batch_rows)}")

    driver.quit()
    logger.info(f"🎉 Done. New fields processed: {processed}")

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
