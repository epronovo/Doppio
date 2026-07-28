# MpGd_Extract_DocBits.py
# -----------------------------------------------------------------------
# PURPOSE
#   Continuously monitors queries/docbits/ for DocBits JSON export files.
#   When a file arrives, each program block is parsed for its ERP API,
#   transaction names (from the actions map), and field-to-DocBits-field
#   mappings.  Only fields that have a DocBits-backed value field are
#   persisted.  After processing, the source file is moved to an archive
#   sub-folder (with a timestamp suffix if a name collision exists) to
#   prevent reprocessing.
#
# INPUTS
#   - DocBits JSON export files dropped into queries/docbits/
#     Expected structure: list of blocks, each with "program", "actions",
#     and "mapping" keys.
#
# OUTPUTS
#   - m3Docbits     (SQLite) — API, TransactionName, FieldName, DocBitsField
#   - Processed JSON files archived to queries/docbits/archive/
#
# DEPENDENCIES
#   - config.py
# -----------------------------------------------------------------------

import json
import logging
import sqlite3
import time
from pathlib import Path
from datetime import datetime

from config import BASE_DIR, get_sqlite_db_path

logger = logging.getLogger(__name__)

# ------------------------------------------------------------
# Configuration
# ------------------------------------------------------------
QUERIES_DIR = BASE_DIR / "queries/docbits"
ARCHIVE_DIR = QUERIES_DIR / "archive"

DB_PATH = get_sqlite_db_path()
POLL_INTERVAL = 5  # seconds

# ------------------------------------------------------------
# Database setup
# ------------------------------------------------------------
def init_db():
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()

    cur.execute("""
        CREATE TABLE IF NOT EXISTS m3Docbits (
            API TEXT NOT NULL,
            TransactionName TEXT NOT NULL,
            FieldName TEXT NOT NULL,
            DocBitsField TEXT NOT NULL
        )
    """)

    cur.execute("""
        CREATE UNIQUE INDEX IF NOT EXISTS idx_m3Docbits
        ON m3Docbits (API, TransactionName, FieldName)
    """)

    conn.commit()
    return conn


# ------------------------------------------------------------
# JSON processing
# ------------------------------------------------------------
def process_docbits_file(json_path: Path, conn: sqlite3.Connection):
    with open(json_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    cur = conn.cursor()

    for block in data:
        program = block.get("program")
        actions = block.get("actions", {})
        mappings = block.get("mapping", [])

        transaction_names = actions.values()

        for mapping in mappings:
            field_name = mapping.get("erp_field_name")
            docbits_field = mapping.get("value_field_name")

            # Only persist DocBits-backed fields
            if not docbits_field:
                continue

            for txn in transaction_names:
                cur.execute("""
                    INSERT OR IGNORE INTO m3Docbits
                        (API, TransactionName, FieldName, DocBitsField)
                    VALUES (?, ?, ?, ?)
                """, (
                    program,
                    txn,
                    field_name,
                    docbits_field
                ))

    conn.commit()


# ------------------------------------------------------------
# Archive handling
# ------------------------------------------------------------
def archive_file(src: Path):
    ARCHIVE_DIR.mkdir(parents=True, exist_ok=True)

    dest = ARCHIVE_DIR / src.name

    # If file already exists in archive, append timestamp
    if dest.exists():
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        dest = ARCHIVE_DIR / f"{src.stem}_{ts}{src.suffix}"

    src.rename(dest)


# ------------------------------------------------------------
# Folder monitor
# ------------------------------------------------------------
def monitor_folder():
    conn = init_db()

    logger.info(f"Monitoring DocBits folder: {QUERIES_DIR}")
    logger.info(f"Archiving processed files to: {ARCHIVE_DIR}")

    while True:
        for json_file in QUERIES_DIR.glob("*.json"):
            try:
                logger.info(f"Processing {json_file.name}")
                process_docbits_file(json_file, conn)
                archive_file(json_file)
                logger.info(f"Archived {json_file.name}")
            except Exception as e:
                logger.error(f"Error processing {json_file.name}: {e}")

        time.sleep(POLL_INTERVAL)


# ------------------------------------------------------------
# Entry point
# ------------------------------------------------------------
if __name__ == "__main__":
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s  %(levelname)-8s  %(message)s",
        datefmt="%H:%M:%S"
    )
    monitor_folder()
