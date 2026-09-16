"""
ERP_Concur_Db - schema and connection helpers for the ERP -> Concur extract
picker.

Everything lives in one SQLite file, by default ~/sqlite/doppio.db, matching
ADP_Concur_Db.py and M3_Security_Db.py. An explicit --db wins, then the
ERP_CONCUR_DB environment variable, then the default.

The shape of the data follows the three files the SSIS package
ConcurExtracts/PurchaseOrder.dtsx writes - purchase order, vendor and PO
receipt - and nothing else. Each record keeps its own raw line verbatim in
`raw`, and the parsed columns exist only so the page can search and the
selection can be resolved. The extract replays `raw`; it never rebuilds a
line from the parsed columns. That is deliberate: these files carry CHAR
padding and quoting quirks that a round trip through Python would quietly
"fix", and a subset that differs from the source is no longer a regression
test of the source.

Two identities matter, because they are the two Concur matches that fail in
production:
  * a vendor is (vendor_code, vendor_address_code) - the exact pair a PO 200
    header is matched on (error 2000 when it misses);
  * a PO line is its External ID - the string a receipt's Line Item External
    ID must equal byte for byte (error 1001 when it misses).
Both are stored as they appear in the file, untrimmed.
"""
from __future__ import annotations

import json
import logging
import os
import sqlite3
from pathlib import Path

SQLITE_DIR = Path.home() / "sqlite"
DEFAULT_DB_PATH = str(SQLITE_DIR / "doppio.db")
DB_ENV_VAR = "ERP_CONCUR_DB"

# The three files, by the kind the parser decides from their content. The
# label is what the page calls them; `spec_keys` is what record types may
# legitimately appear in one.
FILE_KINDS = {
    "po": {
        "label": "Purchase Order",
        "records": {"300": "po_300", "400": "po_400", "210": "po_210",
                    "220": "po_220", "200": "po_200"},
    },
    "vendor": {
        "label": "Vendor",
        "records": {"100": "vendor_100", "200": "vendor_200"},
    },
    "receipt": {
        "label": "PO Receipt",
        "records": {"200": "receipt_200"},
    },
}

# The import results report Concur mails back after a run. Not one of the
# three files the dtsx writes and never a source for an extract - it is the
# answer to one, read to find out which orders have to go again. Kept apart
# from FILE_KINDS so nothing that walks "the three files" picks it up.
RESULTS_KIND = "results"
RESULTS_LABEL = "Import results"

# Every kind that can occupy a row in ERP_Concur_Files, for labelling.
ALL_KINDS = {**{k: v["label"] for k, v in FILE_KINDS.items()},
             RESULTS_KIND: RESULTS_LABEL}

SCHEMA = """
PRAGMA foreign_keys = ON;

-- -------------------------------------------------------------------- files
-- One row per file currently loaded. Dropping a file of a kind that is
-- already loaded replaces it (see replace_file) rather than adding a second:
-- there is exactly one purchase order file, one vendor file and one receipt
-- file in play at a time, which is how the SSIS package emits them.
CREATE TABLE IF NOT EXISTS ERP_Concur_Files (
    file_id       INTEGER PRIMARY KEY AUTOINCREMENT,
    kind          TEXT    NOT NULL,          -- po | vendor | receipt
    file_name     TEXT    NOT NULL,
    file_path     TEXT,
    sha256        TEXT,
    byte_size     INTEGER NOT NULL DEFAULT 0,
    line_count    INTEGER NOT NULL DEFAULT 0,
    record_counts TEXT,                      -- JSON: record type -> rows
    newline       TEXT,                      -- the line ending found in it
    trailing_nl   INTEGER NOT NULL DEFAULT 1,
    loaded_at     TEXT    NOT NULL DEFAULT (datetime('now'))
);

-- ----------------------------------------------------------------- settings
-- The vendor file's 100 Import Settings record - 4 fields, written by hand by
-- the dtsx rather than read from a proc. Kept because the extract has to
-- re-emit it ahead of the vendor rows or Concur rejects the file.
CREATE TABLE IF NOT EXISTS ERP_Concur_Settings (
    settings_key INTEGER PRIMARY KEY AUTOINCREMENT,
    file_id      INTEGER NOT NULL REFERENCES ERP_Concur_Files (file_id)
                 ON DELETE CASCADE,
    line_no      INTEGER NOT NULL,
    error_threshold TEXT,
    default_country_code TEXT,
    pay_method_type TEXT,
    raw          TEXT    NOT NULL
);

-- ------------------------------------------------------------------ vendors
CREATE TABLE IF NOT EXISTS ERP_Concur_Vendors (
    vendor_key   INTEGER PRIMARY KEY AUTOINCREMENT,
    file_id      INTEGER NOT NULL REFERENCES ERP_Concur_Files (file_id)
                 ON DELETE CASCADE,
    line_no      INTEGER NOT NULL,
    vendor_code  TEXT    NOT NULL,
    -- Position 16. Concur's Excel importer calls it Address Accounting Code;
    -- position 15 (vendor_address_id) is its Address Import Sync ID. Getting
    -- these two the wrong way round loads the vendor and fails every PO.
    vendor_address_code TEXT NOT NULL DEFAULT '',
    vendor_address_id   TEXT NOT NULL DEFAULT '',
    vendor_name  TEXT,
    currency     TEXT,
    payment_term_days TEXT,
    address1     TEXT,
    address2     TEXT,
    address3     TEXT,
    city         TEXT,
    state        TEXT,
    postal_code  TEXT,
    country_code TEXT,
    contact_email TEXT,
    field_count  INTEGER NOT NULL DEFAULT 0,
    raw          TEXT    NOT NULL
);

-- ------------------------------------------------------------- PO headers
-- `po_group` is the position of the header in the file, 1-based over headers.
-- The dtsx writes lines(300) -> 210 -> 220 -> header(200) for each PO, so the
-- records belonging to a header are the ones that precede it, and the group
-- number is the only link that does not depend on parsing an External ID.
CREATE TABLE IF NOT EXISTS ERP_Concur_PoHeaders (
    po_key       INTEGER PRIMARY KEY AUTOINCREMENT,
    file_id      INTEGER NOT NULL REFERENCES ERP_Concur_Files (file_id)
                 ON DELETE CASCADE,
    line_no      INTEGER NOT NULL,
    po_group     INTEGER NOT NULL,
    po_number    TEXT    NOT NULL,
    policy_external_id TEXT,
    currency_code TEXT,
    vendor_code  TEXT    NOT NULL DEFAULT '',
    vendor_address_code TEXT NOT NULL DEFAULT '',
    order_date   TEXT,
    payment_terms TEXT,
    tax          TEXT,
    shipping     TEXT,
    ledger_code  TEXT,
    entity_id    TEXT,                       -- Custom_1, position 38
    field_count  INTEGER NOT NULL DEFAULT 0,
    raw          TEXT    NOT NULL
);

-- --------------------------------------------------------------- PO lines
CREATE TABLE IF NOT EXISTS ERP_Concur_PoLines (
    line_key     INTEGER PRIMARY KEY AUTOINCREMENT,
    file_id      INTEGER NOT NULL REFERENCES ERP_Concur_Files (file_id)
                 ON DELETE CASCADE,
    po_key       INTEGER REFERENCES ERP_Concur_PoHeaders (po_key) ON DELETE CASCADE,
    line_no      INTEGER NOT NULL,
    po_group     INTEGER NOT NULL,
    external_id  TEXT    NOT NULL,
    line_number  TEXT,
    supplier_part_id TEXT,
    expense_type TEXT,
    account_code TEXT,
    description  TEXT,
    quantity     TEXT,
    unit_price   TEXT,
    uom          TEXT,
    entity_id    TEXT,                       -- Custom_1, position 33
    -- Set when the line looks like the charge branch of the UNION in
    -- p_concurinvoicepo_get_poline (TOTAL TAX / TOTAL FREIGHT / TOTAL OTHER
    -- against the hardcoded 5555 / 9999 / 4444 account codes) rather than a
    -- real PO line. Charges are the rows that collide on External ID.
    is_charge    INTEGER NOT NULL DEFAULT 0,
    field_count  INTEGER NOT NULL DEFAULT 0,
    raw          TEXT    NOT NULL
);

-- ------------------------------------------------------- PO 400 allocations
-- The spec workbook's own note on the PO - 400 tab says no proc in the
-- deployed script emits this record type; real production files say
-- otherwise, so treat that note as stale. line_key is nullable for two
-- different reasons: a truncated file (like po_key on every other child row
-- here), and a 400 that arrives with no 300 right after it - the record
-- carries no External ID or line number of its own, so position against the
-- immediately FOLLOWING 300 (confirmed against production data: a run of
-- 400s' Amounts match the next 300's Quantity x Unit Price, not the
-- previous one's) is the *only* link, same as a 210/220 against its header.
-- A run of 400s that runs into a 210/220/200 instead of a 300 is orphaned.
--
-- `synthetic` marks a row this tool wrote itself (see
-- ERP_Concur_Parse.synthesize_allocations) rather than one read out of the
-- source file - the one place in this whole database where that distinction
-- has to be kept, because every other table's contract is "raw is exactly
-- what was in the file." A synthetic row's raw is a real, spec-shaped 400
-- line (so the extract, the field-count and required-field checks, and
-- edit_record all treat it exactly like any other record) but it never
-- claims to be a real accounting split - see synthesize_allocations for why.
CREATE TABLE IF NOT EXISTS ERP_Concur_PoLineAllocations (
    alloc_key    INTEGER PRIMARY KEY AUTOINCREMENT,
    file_id      INTEGER NOT NULL REFERENCES ERP_Concur_Files (file_id)
                 ON DELETE CASCADE,
    po_key       INTEGER REFERENCES ERP_Concur_PoHeaders (po_key) ON DELETE CASCADE,
    line_key     INTEGER REFERENCES ERP_Concur_PoLines (line_key) ON DELETE CASCADE,
    line_no      INTEGER NOT NULL,
    po_group     INTEGER NOT NULL,
    amount       TEXT,
    synthetic    INTEGER NOT NULL DEFAULT 0,
    field_count  INTEGER NOT NULL DEFAULT 0,
    raw          TEXT    NOT NULL
);

-- ------------------------------------------------------- PO 210 / 220 rows
CREATE TABLE IF NOT EXISTS ERP_Concur_PoAddresses (
    addr_key     INTEGER PRIMARY KEY AUTOINCREMENT,
    file_id      INTEGER NOT NULL REFERENCES ERP_Concur_Files (file_id)
                 ON DELETE CASCADE,
    po_key       INTEGER REFERENCES ERP_Concur_PoHeaders (po_key) ON DELETE CASCADE,
    line_no      INTEGER NOT NULL,
    po_group     INTEGER NOT NULL,
    record_type  TEXT    NOT NULL,           -- 210 bill-to | 220 ship-to
    external_id  TEXT,
    name         TEXT,
    address1     TEXT,
    address2     TEXT,
    address3     TEXT,
    city         TEXT,
    state        TEXT,
    postal_code  TEXT,
    country_code TEXT,
    field_count  INTEGER NOT NULL DEFAULT 0,
    raw          TEXT    NOT NULL
);

-- ----------------------------------------------------------------- receipts
CREATE TABLE IF NOT EXISTS ERP_Concur_Receipts (
    receipt_key  INTEGER PRIMARY KEY AUTOINCREMENT,
    file_id      INTEGER NOT NULL REFERENCES ERP_Concur_Files (file_id)
                 ON DELETE CASCADE,
    line_no      INTEGER NOT NULL,
    po_number    TEXT    NOT NULL,
    line_item_external_id TEXT NOT NULL,
    goods_receipt_number  TEXT,
    delivery_slip_number  TEXT,
    uom          TEXT,
    received_quantity TEXT,
    received_date TEXT,
    is_deleted   TEXT,
    field_count  INTEGER NOT NULL DEFAULT 0,
    raw          TEXT    NOT NULL
);

-- ----------------------------------------------------------------- findings
-- Rebuilt from scratch on every parse, so it always describes the files as
-- they stand rather than accumulating history. `scope` + `ref_key` point at
-- the row the finding is about so the page can jump to it.
CREATE TABLE IF NOT EXISTS ERP_Concur_Findings (
    finding_key INTEGER PRIMARY KEY AUTOINCREMENT,
    severity    TEXT NOT NULL,               -- error | warning | info
    kind        TEXT NOT NULL,               -- a stable slug, for grouping
    scope       TEXT NOT NULL,               -- vendor|po|line|address|receipt|file
    ref_key     INTEGER,
    po_number   TEXT NOT NULL DEFAULT '',
    vendor_code TEXT NOT NULL DEFAULT '',
    label       TEXT NOT NULL DEFAULT '',    -- what the row is, for the list
    field       TEXT NOT NULL DEFAULT '',
    message     TEXT NOT NULL,
    detected_at TEXT NOT NULL DEFAULT (datetime('now'))
);

-- ---------------------------------------------------------------- selection
-- What is currently picked, as two sets held in one table: vendors by
-- vendor_code and purchase orders by po_number. Kept in the database rather
-- than the browser for the same reasons as ADP_Concur_Selection - it survives
-- a reload, every list filters on it with a join instead of a query string
-- carrying a hundred keys, and the extract and the screen read the same rows
-- so they cannot disagree.
--
-- `direct` is 1 when the user picked this thing itself and 0 when it was
-- pulled in by the other side (a vendor's orders, or an order's vendor). It
-- is what lets an un-pick be the exact inverse of a pick: dropping one PO of
-- a vendor you picked by hand leaves the vendor alone, while dropping the
-- last PO of a vendor that only came along for the ride drops it too.
CREATE TABLE IF NOT EXISTS ERP_Concur_Selection (
    kind     TEXT NOT NULL,                  -- vendor | po
    id       TEXT NOT NULL,                  -- vendor_code | po_number
    direct   INTEGER NOT NULL DEFAULT 1,
    reason   TEXT,
    added_at TEXT NOT NULL DEFAULT (datetime('now')),
    PRIMARY KEY (kind, id)
);

-- ------------------------------------------------------------------ results
-- One row per line of the import results report Concur returns for a run.
--
-- `record_id` is the report's Record Identifier, which is the 1-based line
-- number of the record in the file that was sent - the Info row that closes
-- the report carries the line count of the whole file as its identifier. That
-- is the only link back, because most messages name no purchase order: the
-- "sequence of the record types is invalid" pair names none at all, and only
-- the 5001 and 4002 messages carry one in their text.
--
-- So both are kept. `stated_po` is the order the message names, `resolved_po`
-- is the order whose records occupy that line of the purchase order file, and
-- `aligned` is whether the two agree where both are known. A report read
-- against the wrong run resolves to plausible-looking orders, and `aligned`
-- is what catches it.
CREATE TABLE IF NOT EXISTS ERP_Concur_Results (
    result_key   INTEGER PRIMARY KEY AUTOINCREMENT,
    file_id      INTEGER NOT NULL REFERENCES ERP_Concur_Files (file_id)
                 ON DELETE CASCADE,
    row_no       INTEGER NOT NULL,          -- row in the report, 1-based
    level        TEXT    NOT NULL,          -- Error | Warning | Info
    record_id    INTEGER,                   -- line number in the sent file
    message      TEXT    NOT NULL,          -- as it came, markup and all
    text         TEXT    NOT NULL,          -- the same with the <br/> taken out
    error_code   TEXT    NOT NULL DEFAULT '',
    error_text   TEXT    NOT NULL DEFAULT '',
    field_level  TEXT    NOT NULL DEFAULT '',
    field_code   TEXT    NOT NULL DEFAULT '',
    line_item_external_id TEXT NOT NULL DEFAULT '',
    stated_po    TEXT    NOT NULL DEFAULT '',   -- named in the message
    resolved_po  TEXT    NOT NULL DEFAULT '',   -- found at that line
    resolved_scope TEXT  NOT NULL DEFAULT '',   -- po | line | address | ''
    resolved_line_no INTEGER,
    aligned      INTEGER                        -- 1 agree, 0 disagree, NULL n/a
);

-- ----------------------------------------------------------------- extracts
-- One row per set of files written, so an extract can be explained and
-- re-downloaded after the fact.
CREATE TABLE IF NOT EXISTS ERP_Concur_Extracts (
    extract_key INTEGER PRIMARY KEY AUTOINCREMENT,
    stamp       TEXT NOT NULL,
    label       TEXT NOT NULL DEFAULT '',
    scope       TEXT NOT NULL DEFAULT '',    -- human-readable: what was in it
    n_vendor    INTEGER NOT NULL DEFAULT 0,
    n_po        INTEGER NOT NULL DEFAULT 0,
    n_line      INTEGER NOT NULL DEFAULT 0,
    n_allocation INTEGER NOT NULL DEFAULT 0,
    n_address   INTEGER NOT NULL DEFAULT 0,
    n_receipt   INTEGER NOT NULL DEFAULT 0,
    files       TEXT NOT NULL DEFAULT '[]',  -- JSON: [{kind,name,path,rows}]
    written_at  TEXT NOT NULL DEFAULT (datetime('now'))
);
"""

# Indexes are applied after the tables, and after migrate(), for the reason
# ADP_Concur_Db gives: CREATE TABLE IF NOT EXISTS does nothing to a table that
# already exists, so on an older database the table is old and the index is
# new, and the whole script would fail on a column that is not there yet.
INDEXES = """
CREATE INDEX IF NOT EXISTS ix_erpc_ven_code   ON ERP_Concur_Vendors (vendor_code);
CREATE INDEX IF NOT EXISTS ix_erpc_ven_pair   ON ERP_Concur_Vendors (vendor_code, vendor_address_code);
CREATE INDEX IF NOT EXISTS ix_erpc_po_num     ON ERP_Concur_PoHeaders (po_number);
CREATE INDEX IF NOT EXISTS ix_erpc_po_vendor  ON ERP_Concur_PoHeaders (vendor_code);
CREATE INDEX IF NOT EXISTS ix_erpc_po_group   ON ERP_Concur_PoHeaders (po_group);
CREATE INDEX IF NOT EXISTS ix_erpc_line_po    ON ERP_Concur_PoLines (po_key);
CREATE INDEX IF NOT EXISTS ix_erpc_line_ext   ON ERP_Concur_PoLines (external_id);
CREATE INDEX IF NOT EXISTS ix_erpc_alloc_po   ON ERP_Concur_PoLineAllocations (po_key);
CREATE INDEX IF NOT EXISTS ix_erpc_alloc_line ON ERP_Concur_PoLineAllocations (line_key);
CREATE INDEX IF NOT EXISTS ix_erpc_addr_po    ON ERP_Concur_PoAddresses (po_key);
CREATE INDEX IF NOT EXISTS ix_erpc_rec_po     ON ERP_Concur_Receipts (po_number);
CREATE INDEX IF NOT EXISTS ix_erpc_rec_ext    ON ERP_Concur_Receipts (line_item_external_id);
CREATE INDEX IF NOT EXISTS ix_erpc_res_po     ON ERP_Concur_Results (resolved_po);
CREATE INDEX IF NOT EXISTS ix_erpc_res_stated ON ERP_Concur_Results (stated_po);
CREATE INDEX IF NOT EXISTS ix_erpc_res_rec    ON ERP_Concur_Results (record_id);
CREATE INDEX IF NOT EXISTS ix_erpc_find_kind  ON ERP_Concur_Findings (kind);
CREATE INDEX IF NOT EXISTS ix_erpc_find_sev   ON ERP_Concur_Findings (severity);
CREATE INDEX IF NOT EXISTS ix_erpc_find_scope ON ERP_Concur_Findings (scope, ref_key);
"""

# Columns added after the first release. SQLite cannot add one through
# CREATE TABLE IF NOT EXISTS, and doppio.db is shared with three other tools
# so it is never a fresh database. Append here when SCHEMA above gains a
# column; naming one that is already there is harmless.
MIGRATIONS: list[tuple[str, str, str]] = [
    ("ERP_Concur_Extracts", "n_allocation", "INTEGER NOT NULL DEFAULT 0"),
    ("ERP_Concur_PoLineAllocations", "synthetic", "INTEGER NOT NULL DEFAULT 0"),
]

log = logging.getLogger("ERP_Concur_Db")


def migrate(conn: sqlite3.Connection) -> list[str]:
    """Add any columns an older database is missing, and report what was added."""
    applied = []
    for table, column, ddl in MIGRATIONS:
        if not conn.execute(
            "SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = ?",
            (table,)).fetchone():
            continue                       # a fresh database; SCHEMA built it
        try:
            conn.execute(f"ALTER TABLE {table} ADD COLUMN {column} {ddl}")
            applied.append(f"{table}.{column}")
        except sqlite3.OperationalError as exc:
            if "duplicate column" not in str(exc).lower():
                raise
    return applied


def resolve_db_path(db_path: str | None = None) -> str:
    """--db wins, then ERP_CONCUR_DB, then ~/sqlite/doppio.db."""
    if db_path:
        return str(Path(db_path).expanduser())
    env = os.environ.get(DB_ENV_VAR)
    if env:
        return str(Path(env).expanduser())
    return DEFAULT_DB_PATH


def connect(db_path: str | None = None) -> sqlite3.Connection:
    """Open the database, creating the folder and the schema if need be."""
    path = Path(resolve_db_path(db_path))
    path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(path))
    conn.row_factory = sqlite3.Row
    conn.executescript(SCHEMA)
    applied = migrate(conn)
    conn.executescript(INDEXES)
    conn.commit()
    if applied:
        log.info("Added %d column(s) to the existing schema: %s",
                 len(applied), ", ".join(applied))
    return conn


# ------------------------------------------------------------------- counts


def counts(conn: sqlite3.Connection) -> dict:
    """Row counts for the status bar - one query per table it names."""
    def n(sql: str, *args) -> int:
        return conn.execute(sql, args).fetchone()[0]
    return {
        "vendors": n("SELECT COUNT(*) FROM ERP_Concur_Vendors"),
        "vendor_codes": n("SELECT COUNT(DISTINCT vendor_code) FROM ERP_Concur_Vendors"),
        "po_headers": n("SELECT COUNT(*) FROM ERP_Concur_PoHeaders"),
        "po_lines": n("SELECT COUNT(*) FROM ERP_Concur_PoLines"),
        "po_charges": n("SELECT COUNT(*) FROM ERP_Concur_PoLines WHERE is_charge = 1"),
        "po_allocations": n("SELECT COUNT(*) FROM ERP_Concur_PoLineAllocations"),
        "po_allocations_synthetic": n(
            "SELECT COUNT(*) FROM ERP_Concur_PoLineAllocations WHERE synthetic = 1"),
        "po_addresses": n("SELECT COUNT(*) FROM ERP_Concur_PoAddresses"),
        "receipts": n("SELECT COUNT(*) FROM ERP_Concur_Receipts"),
        "findings": n("SELECT COUNT(*) FROM ERP_Concur_Findings"),
        "findings_warning": n(
            "SELECT COUNT(*) FROM ERP_Concur_Findings WHERE severity = 'warning'"),
        "findings_error": n(
            "SELECT COUNT(*) FROM ERP_Concur_Findings WHERE severity = 'error'"),
        "results": n("SELECT COUNT(*) FROM ERP_Concur_Results"),
        "results_errors": n(
            "SELECT COUNT(*) FROM ERP_Concur_Results WHERE level = 'Error'"),
        # Distinct orders that failed, however the report said so - the number
        # the "pick every failed order" button acts on.
        "results_failed_pos": n(
            "SELECT COUNT(DISTINCT CASE WHEN stated_po <> '' THEN stated_po "
            "  ELSE resolved_po END) FROM ERP_Concur_Results "
            "WHERE level = 'Error' AND (stated_po <> '' OR resolved_po <> '')"),
        "picked_vendors": n(
            "SELECT COUNT(*) FROM ERP_Concur_Selection WHERE kind = 'vendor'"),
        "picked_pos": n(
            "SELECT COUNT(*) FROM ERP_Concur_Selection WHERE kind = 'po'"),
    }


def files(conn: sqlite3.Connection) -> list[dict]:
    out = []
    for r in conn.execute("SELECT * FROM ERP_Concur_Files ORDER BY kind"):
        row = dict(r)
        row["record_counts"] = json.loads(row.get("record_counts") or "{}")
        row["label"] = ALL_KINDS.get(row["kind"], row["kind"])
        out.append(row)
    return out


def replace_file(conn: sqlite3.Connection, kind: str) -> int:
    """
    Forget whatever file of this kind is loaded, and everything parsed from
    it. ON DELETE CASCADE clears the record tables; the selection is left
    alone on purpose - re-dropping a corrected purchase order file should not
    throw away the POs that were picked from the last one. Anything picked
    that is no longer in the file simply stops resolving, and the Extract tab
    says so rather than silently shipping less than was asked for.
    """
    rows = conn.execute("SELECT file_id FROM ERP_Concur_Files WHERE kind = ?",
                        (kind,)).fetchall()
    for r in rows:
        conn.execute("DELETE FROM ERP_Concur_Files WHERE file_id = ?", (r[0],))
    return len(rows)


def clear_all(conn: sqlite3.Connection, selection: bool = True,
              commit: bool = True) -> dict:
    """Empty the tool out. Used by the Clear button and by the test harness."""
    done = {}
    for table in ("ERP_Concur_Findings", "ERP_Concur_Files"):
        done[table] = conn.execute(f"DELETE FROM {table}").rowcount
    if selection:
        done["ERP_Concur_Selection"] = conn.execute(
            "DELETE FROM ERP_Concur_Selection").rowcount
    if commit:
        conn.commit()
    return done


# ---------------------------------------------------------------- selection
#
# The two sets are kept consistent by propagate(): a picked vendor drags in
# every purchase order whose header names it, and a picked purchase order
# drags in the vendor its header names. That is the behaviour the tool exists
# for - "give me everything for this vendor" and "give me the vendors behind
# these orders" are the same operation read from either end.
#
# Matching a PO to a vendor is deliberately done on vendor_code alone here,
# not on the (vendor_code, vendor_address_code) pair Concur matches on. A
# header whose pair does not resolve is exactly the error-2000 case, and it
# still has to end up in the extract together with its vendor row so the
# mismatch can be looked at. The pair mismatch is reported as a finding
# instead.


def selection_ids(conn: sqlite3.Connection, kind: str) -> list[str]:
    return [r[0] for r in conn.execute(
        "SELECT id FROM ERP_Concur_Selection WHERE kind = ? ORDER BY id", (kind,))]


def _pick(conn: sqlite3.Connection, kind: str, ids: list[str],
          reason: str) -> None:
    """
    Upsert a by-hand pick. A row that is already there as a passenger is
    promoted to direct = 1 rather than left alone: asking for something by
    name is what protects it from being dropped when its last sibling goes.
    """
    conn.executemany(
        "INSERT INTO ERP_Concur_Selection (kind, id, direct, reason) "
        "VALUES (?, ?, 1, ?) ON CONFLICT (kind, id) DO UPDATE SET "
        "direct = 1, reason = excluded.reason",
        [(kind, i, reason) for i in ids])


def selection_add(conn: sqlite3.Connection, vendors: list[str] | None = None,
                  pos: list[str] | None = None, reason: str = "",
                  propagate: bool = True, commit: bool = True) -> dict:
    """
    Pick vendors and/or purchase orders, then pull in the other side.

    Returns what the selection looks like afterwards plus how many rows each
    side gained, counted by measuring the sets before and after rather than
    trusting an upsert's rowcount.
    """
    vendors = [v for v in (vendors or []) if v]
    pos = [p for p in (pos or []) if p]
    before_v = set(selection_ids(conn, "vendor"))
    before_p = set(selection_ids(conn, "po"))

    if vendors:
        _pick(conn, "vendor", vendors, reason or "picked")
    if pos:
        _pick(conn, "po", pos, reason or "picked")
    if propagate:
        propagate_selection(conn)

    after_v = set(selection_ids(conn, "vendor"))
    after_p = set(selection_ids(conn, "po"))
    if commit:
        conn.commit()
    return {"added_vendors": len(after_v - before_v),
            "added_pos": len(after_p - before_p),
            **selection_summary(conn)}


def selection_remove(conn: sqlite3.Connection, vendors: list[str] | None = None,
                     pos: list[str] | None = None, commit: bool = True) -> dict:
    """
    Un-pick vendors and/or purchase orders, and take their passengers with
    them.

    Dropping a vendor drops its orders, unless an order was picked by hand
    (direct = 1) or belongs to another picked vendor. Dropping an order drops
    its vendor only if that vendor was never picked by hand and has no other
    picked order left. So an un-pick undoes the pick it mirrors and nothing
    more.
    """
    vendors = [v for v in (vendors or []) if v]
    pos = [p for p in (pos or []) if p]
    before_v = set(selection_ids(conn, "vendor"))
    before_p = set(selection_ids(conn, "po"))

    if vendors:
        conn.executemany("DELETE FROM ERP_Concur_Selection "
                         "WHERE kind = 'vendor' AND id = ?",
                         [(v,) for v in vendors])
        # Their orders go too, except the ones that stand on their own.
        conn.execute(
            "DELETE FROM ERP_Concur_Selection WHERE kind = 'po' AND direct = 0 "
            "AND id IN (SELECT h.po_number FROM ERP_Concur_PoHeaders h "
            "           WHERE h.vendor_code IN (%s)) "
            "AND id NOT IN (SELECT h.po_number FROM ERP_Concur_PoHeaders h "
            "               JOIN ERP_Concur_Selection s ON s.kind = 'vendor' "
            "                AND s.id = h.vendor_code)"
            % ",".join("?" * len(vendors)), vendors)
    if pos:
        conn.executemany("DELETE FROM ERP_Concur_Selection "
                         "WHERE kind = 'po' AND id = ?", [(p,) for p in pos])
        # A vendor with nothing left pointing at it, that nobody asked for by
        # name, leaves with its last order.
        conn.execute(
            "DELETE FROM ERP_Concur_Selection WHERE kind = 'vendor' AND direct = 0 "
            "AND id NOT IN (SELECT h.vendor_code FROM ERP_Concur_PoHeaders h "
            "               JOIN ERP_Concur_Selection s ON s.kind = 'po' "
            "                AND s.id = h.po_number)")

    after_v = set(selection_ids(conn, "vendor"))
    after_p = set(selection_ids(conn, "po"))
    if commit:
        conn.commit()
    return {"removed_vendors": len(before_v - after_v),
            "removed_pos": len(before_p - after_p),
            **selection_summary(conn)}


def propagate_selection(conn: sqlite3.Connection) -> dict:
    """
    Close the selection over the vendor <-> purchase order link.

    Only a vendor picked by hand (direct = 1, picked from the Vendors tab)
    pulls in every order that names it. A vendor pulled in as a passenger of
    one order pick stays scoped to that order - it does not also sweep in the
    rest of that vendor's orders. Run to a fixed point (it converges in one
    pass over this data, but a directly-picked vendor reached through an
    order could in principle bring further orders), and every row it adds is
    marked direct = 0 with a reason naming what pulled it in.
    """
    added_v = added_p = 0
    for _ in range(10):
        cur = conn.execute(
            "INSERT INTO ERP_Concur_Selection (kind, id, direct, reason) "
            "SELECT 'po', h.po_number, 0, 'vendor ' || h.vendor_code "
            "  FROM ERP_Concur_PoHeaders h "
            "  JOIN ERP_Concur_Selection s ON s.kind = 'vendor' AND s.id = h.vendor_code "
            "   AND s.direct = 1 "
            " WHERE h.po_number NOT IN (SELECT id FROM ERP_Concur_Selection WHERE kind = 'po') "
            " GROUP BY h.po_number")
        n1 = cur.rowcount if cur.rowcount > 0 else 0
        cur = conn.execute(
            "INSERT INTO ERP_Concur_Selection (kind, id, direct, reason) "
            "SELECT 'vendor', h.vendor_code, 0, 'PO ' || MIN(h.po_number) "
            "  FROM ERP_Concur_PoHeaders h "
            "  JOIN ERP_Concur_Selection s ON s.kind = 'po' AND s.id = h.po_number "
            " WHERE h.vendor_code <> '' "
            "   AND h.vendor_code NOT IN (SELECT id FROM ERP_Concur_Selection WHERE kind = 'vendor') "
            " GROUP BY h.vendor_code")
        n2 = cur.rowcount if cur.rowcount > 0 else 0
        added_p += n1
        added_v += n2
        if not (n1 or n2):
            break
    return {"added_vendors": added_v, "added_pos": added_p}


def selection_clear(conn: sqlite3.Connection, commit: bool = True) -> int:
    n = conn.execute("DELETE FROM ERP_Concur_Selection").rowcount
    if commit:
        conn.commit()
    return n


def selection_summary(conn: sqlite3.Connection) -> dict:
    """
    What is picked, and what it resolves to in the files that are loaded.

    `unresolved_*` are ids that were picked from an earlier file and are not
    in the current one. They are reported rather than quietly dropped,
    because an extract that is short of what was asked for should say so.
    """
    def n(sql: str) -> int:
        return conn.execute(sql).fetchone()[0]
    picked_v = n("SELECT COUNT(*) FROM ERP_Concur_Selection WHERE kind = 'vendor'")
    picked_p = n("SELECT COUNT(*) FROM ERP_Concur_Selection WHERE kind = 'po'")
    return {
        "vendors": picked_v,
        "pos": picked_p,
        "vendor_rows": n(
            "SELECT COUNT(*) FROM ERP_Concur_Vendors v JOIN ERP_Concur_Selection s "
            "ON s.kind = 'vendor' AND s.id = v.vendor_code"),
        "po_rows": n(
            "SELECT COUNT(*) FROM ERP_Concur_PoHeaders h JOIN ERP_Concur_Selection s "
            "ON s.kind = 'po' AND s.id = h.po_number"),
        "line_rows": n(
            "SELECT COUNT(*) FROM ERP_Concur_PoLines l JOIN ERP_Concur_PoHeaders h "
            "ON h.po_key = l.po_key JOIN ERP_Concur_Selection s "
            "ON s.kind = 'po' AND s.id = h.po_number"),
        "address_rows": n(
            "SELECT COUNT(*) FROM ERP_Concur_PoAddresses a JOIN ERP_Concur_PoHeaders h "
            "ON h.po_key = a.po_key JOIN ERP_Concur_Selection s "
            "ON s.kind = 'po' AND s.id = h.po_number"),
        "receipt_rows": n(
            "SELECT COUNT(*) FROM ERP_Concur_Receipts r JOIN ERP_Concur_Selection s "
            "ON s.kind = 'po' AND s.id = r.po_number"),
        "unresolved_vendors": [r[0] for r in conn.execute(
            "SELECT id FROM ERP_Concur_Selection WHERE kind = 'vendor' AND id NOT IN "
            "(SELECT vendor_code FROM ERP_Concur_Vendors) ORDER BY id")],
        "unresolved_pos": [r[0] for r in conn.execute(
            "SELECT id FROM ERP_Concur_Selection WHERE kind = 'po' AND id NOT IN "
            "(SELECT po_number FROM ERP_Concur_PoHeaders) ORDER BY id")],
    }


def picked_clause(picked: str, column: str, kind: str) -> str:
    """
    The single definition of what "picked" means, for every list in the app.

    `picked` is 'only', 'not' or anything else for no filter; `column` is the
    qualified column holding the id (h.po_number, v.vendor_code, ...) and
    `kind` is which set to test it against. Returned as a clause that can be
    appended to a WHERE that already has a condition in it, so it composes
    with the search box and the other filters instead of competing with them.
    """
    if picked not in ("only", "not"):
        return ""
    op = "IN" if picked == "only" else "NOT IN"
    return (f" AND {column} {op} (SELECT id FROM ERP_Concur_Selection "
            f"WHERE kind = '{kind}')")
