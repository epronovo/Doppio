# chc_spv.py
# -----------------------------------------------------------------------
# PURPOSE
#   Scans the input folder for .txt files (tab-delimited panel sequence
#   data: CSYSPV / CSYSOR / CSYVIU exports) and loads each file into the
#   doppio SQLite database as its own table, named after the file
#   (dropped and recreated on every run).
#
#   FILE LAYOUT CONVENTION
#     One file → one table, table name = file name (without extension)
#     Row 1    → column headers
#     Row 2+   → data rows
#     Columns are separated by a single tab character.
#
#   COLUMN TYPING
#     Each column's SQLite type is inferred from the actual cell values:
#       all integers (incl. integral floats)  → INTEGER
#       any non-integral float                → REAL
#       anything else / mixed                 → TEXT
#     Large numbers such as C9LMTS / CILMTS / CJLMTS (millisecond
#     timestamps) are stored as full integers — never in scientific
#     notation (e.g. 1.39548E+12).
#     In mixed TEXT columns, numeric cells are rendered without an
#     exponent as well.
#
# INPUTS
#   ./input/*.txt
#
# OUTPUTS
#   Tables in /Users/ericpronovost/sqlite/doppio.db (one per file)
#
# USAGE
#   python chc_spv.py [--input-folder PATH] [--db PATH] [--verbose]
# -----------------------------------------------------------------------

import argparse
import glob
import logging
import os
import re
import sqlite3
import sys
from datetime import date, datetime

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------
SCRIPT_DIR           = os.path.dirname(os.path.abspath(__file__))
DEFAULT_INPUT_FOLDER = os.path.join(SCRIPT_DIR, "input")
DEFAULT_DB_PATH      = "/Users/ericpronovost/sqlite/doppio.db"

log = logging.getLogger("chc_spv")


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def sanitize(name):
    """Make a safe SQLite identifier from a file or header name."""
    name = re.sub(r"\W+", "_", str(name).strip())
    return name or "UNNAMED"


def parse_value(s):
    """Parse a raw tab-delimited cell string into int / float / str / None."""
    if s is None:
        return None
    s = s.strip()
    if s == "":
        return None
    try:
        return int(s)
    except ValueError:
        pass
    try:
        return float(s)
    except ValueError:
        return s


def infer_type(values):
    """Infer the SQLite column type from the non-null cell values."""
    saw_int = saw_real = False
    for v in values:
        if isinstance(v, bool):
            saw_int = True
        elif isinstance(v, int):
            saw_int = True
        elif isinstance(v, float):
            if v.is_integer():
                saw_int = True
            else:
                saw_real = True
        else:
            return "TEXT"
    if saw_real:
        return "REAL"
    if saw_int:
        return "INTEGER"
    return "TEXT"          # column had no data


def convert(v, sql_type):
    """Convert a cell value for storage, avoiding scientific notation."""
    if v is None:
        return None
    if sql_type == "INTEGER":
        return int(v)
    if sql_type == "REAL":
        return float(v)
    # TEXT — render numbers in full, never scientific notation
    if isinstance(v, bool):
        return "1" if v else "0"
    if isinstance(v, int):
        return str(v)
    if isinstance(v, float):
        if v.is_integer():
            return str(int(v))
        return format(v, "f").rstrip("0").rstrip(".")
    if isinstance(v, (datetime, date)):
        return v.isoformat()
    return str(v)


def read_rows(path):
    """Yield tab-split rows from a text file, skipping blank lines."""
    with open(path, encoding="utf-8-sig", newline="") as f:
        for line in f:
            line = line.rstrip("\r\n")
            if line == "":
                continue
            yield line.split("\t")


# ---------------------------------------------------------------------------
# Core
# ---------------------------------------------------------------------------
def load_file(conn, path):
    """Load one tab-delimited file into its own table. Returns (table, row count)."""
    rows_iter = read_rows(path)
    try:
        header = next(rows_iter)
    except StopIteration:
        log.warning("File %s is empty — skipped", os.path.basename(path))
        return None, 0

    cols = [sanitize(h) for h in header if h != ""]
    ncol = len(cols)
    data = [
        [parse_value(row[i]) if i < len(row) else None for i in range(ncol)]
        for row in rows_iter
    ]

    # Infer each column's type from its actual values
    types = [
        infer_type(row[i] for row in data if row[i] is not None)
        for i in range(ncol)
    ]

    table = sanitize(os.path.splitext(os.path.basename(path))[0]).upper()
    col_defs = ", ".join(f'"{c}" {t}' for c, t in zip(cols, types))
    conn.execute(f'DROP TABLE IF EXISTS "{table}"')
    conn.execute(f'CREATE TABLE "{table}" ({col_defs})')

    placeholders = ", ".join("?" * ncol)
    insert = f'INSERT INTO "{table}" VALUES ({placeholders})'
    conn.executemany(
        insert,
        (
            tuple(convert(row[i], types[i]) for i in range(ncol))
            for row in data
        ),
    )

    for c, t in zip(cols, types):
        log.debug("  %s.%s → %s", table, c, t)
    return table, len(data)


def main():
    ap = argparse.ArgumentParser(description="Load panel sequence text files into doppio.db")
    ap.add_argument("--input-folder", default=DEFAULT_INPUT_FOLDER)
    ap.add_argument("--db", default=DEFAULT_DB_PATH)
    ap.add_argument("--verbose", action="store_true")
    args = ap.parse_args()

    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(levelname)s %(message)s",
    )

    files = sorted(
        f
        for f in glob.glob(os.path.join(args.input_folder, "*.txt"))
        if not os.path.basename(f).startswith("~$")
    )
    if not files:
        log.error("No .txt files found in %s", args.input_folder)
        sys.exit(1)

    os.makedirs(os.path.dirname(args.db), exist_ok=True)
    conn = sqlite3.connect(args.db)
    try:
        total = []
        for path in files:
            log.info("Processing %s", os.path.basename(path))
            table, n = load_file(conn, path)
            if table:
                log.info("  %-12s → %d rows", table, n)
                total.append((table, n))
        conn.commit()
    finally:
        conn.close()

    log.info("Done — %d table(s) loaded into %s", len(total), args.db)


if __name__ == "__main__":
    main()
