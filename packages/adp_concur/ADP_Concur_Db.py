"""
ADP_Concur_Db - schema and connection helpers for the ADP -> Concur load.

Everything the capture holds lives in one SQLite file, by default
~/sqlite/doppio.db, matching M3_Security_Db.py, Sheet2Db.py and config.py.
An explicit --db argument wins, then the ADP_CONCUR_DB environment variable,
then the default.

The shape of the data follows Kelly's workbook: one sheet of raw ADP export
columns, six lookup tabs that map ADP values onto Concur values, and three
record layouts (305, 350, 360) that are built by referencing them.
"""
from __future__ import annotations

import logging
import os
import sqlite3
from pathlib import Path

# The database lives beside the other Doppio SQLite files.
SQLITE_DIR = Path.home() / "sqlite"
DEFAULT_DB_PATH = str(SQLITE_DIR / "doppio.db")

# Overrides the default without a command-line argument.
DB_ENV_VAR = "ADP_CONCUR_DB"

# The ADP export, worksheet header -> database column. Order matters: it is the
# order the Employees tab is shown in and the order a re-export writes.
# Columns A..AB of Kelly's sheet "1".
ADP_COLUMNS: list[tuple[str, str]] = [
    ("Payroll Company Code", "payroll_company_code"),                  # A
    ("Legal First Name", "legal_first_name"),                          # B
    ("Middle Initial", "middle_initial"),                              # C
    ("Legal Last Name", "legal_last_name"),                            # D
    ("Preferred or Chosen First Name", "preferred_first_name"),        # E
    ("Preferred or Chosen Last Name", "preferred_last_name"),          # F
    ("Reports To Legal Name", "reports_to_legal_name"),                # G
    ("Supervisor ID", "supervisor_id_raw"),                            # H
    ("Work Contact: Work Email", "work_email"),                        # I
    ("Personal Contact: Personal Email", "personal_email"),            # J
    ("Job Title Description", "job_title"),                            # K
    ("Hire Date", "hire_date"),                                        # L
    ("Termination Date", "termination_date"),                          # M
    ("Rehire Date", "rehire_date"),                                    # N
    ("Position Status", "position_status"),                            # O
    ("File Number", "file_number"),                                    # P
    ("Pay Frequency", "pay_frequency"),                                # Q
    ("Employee Type", "employee_type"),                                # R
    ("Business Unit Code", "business_unit_code"),                      # S
    ("Business Unit Description", "business_unit_desc"),               # T
    ("Home Department Description", "home_department_desc"),           # U
    ("Home Department Code", "home_department_code"),                  # V
    ("Location Description", "location_desc"),                         # W
    ("Language Description", "language_desc"),                         # X
    ("Language Code", "language_code"),                                # Y
    ("Legal / Preferred Address: Country Code", "legal_country_code"),  # Z
    ("Pay Grade Code", "pay_grade_code"),                              # AA
    ("Pay Grade Description", "pay_grade_desc"),                       # AB
]

# Columns the non-US roster carries instead of ADP's, because those people are
# maintained by hand on the '305 Non US Non SE' tab in Concur's own terms
# rather than being derived from an ADP report. Editable like the ADP columns.
NON_US_COLUMNS: list[tuple[str, str]] = [
    ("Ledger Code", "ledger_code"),
    ("Org Unit 2", "org_unit_2_raw"),
    ("Password", "password"),
]

# The two populations the workbook maintains separately. Each has its own
# derivation rules and its own extract, because they are two different loads
# into Concur, not one load with a filter on it.
ROSTERS = {
    "us": {"label": "US (from ADP)", "records": ["305", "350", "360"]},
    "non_us": {"label": "Non-US / non-SE", "records": ["305", "360"]},
}

# The derived columns Kelly builds with lookups, AC..AM. These are recomputed
# by ADP_Concur_Map.ADP_Concur_derive() rather than imported, so a map change
# takes effect without reloading the workbook - but they are stored so the
# Employees tab can show them and the export does not have to recompute.
DERIVED_COLUMNS: list[tuple[str, str]] = [
    ("SupervisorID Formatted", "supervisor_id"),        # AC
    ("Org Unit 1", "org_unit_1"),                       # AD
    ("Org Unit 2", "org_unit_2"),                       # AE
    ("Concur Profile", "concur_profile"),               # AF
    ("Travel Profile", "travel_profile"),               # AG
    ("Legal Country", "legal_country"),                 # AH
    ("Locale Code", "locale_code"),                     # AI
    ("Reimbursement Currency", "reimbursement_currency"),  # AJ
    ("Preferred Name", "preferred_name"),               # AK
    ("Status", "concur_status"),                        # AL
    ("Term Date", "term_date"),                         # AM
    ("Login ID", "login_id"),                           # not in the workbook
]

# What the employee editor is allowed to write. The derived columns stay out -
# they are rebuilt from the maps - and so does employee_key.
EMPLOYEE_EDITABLE_FIELDS = ([name for _, name in ADP_COLUMNS]
                            + [name for _, name in NON_US_COLUMNS])

# The three record types the flat file carries, in the order they are written.
RECORD_TYPES = ["305", "350", "360"]

SCHEMA = """
PRAGMA foreign_keys = ON;

-- ------------------------------------------------------------- employees
-- One row per person: the raw ADP columns, then the derived Concur values.
-- source is 'adp' for a row that came out of the workbook and 'manual' for
-- one keyed in here, which is how a branch that has not reached ADP yet gets
-- its people into the load.
CREATE TABLE IF NOT EXISTS ADP_Concur_Employees (
    employee_key            INTEGER PRIMARY KEY AUTOINCREMENT,
    file_number             TEXT    NOT NULL,
    payroll_company_code    TEXT,
    legal_first_name        TEXT,
    middle_initial          TEXT,
    legal_last_name         TEXT,
    preferred_first_name    TEXT,
    preferred_last_name     TEXT,
    reports_to_legal_name   TEXT,
    supervisor_id_raw       TEXT,
    work_email              TEXT,
    personal_email          TEXT,
    job_title               TEXT,
    hire_date               TEXT,
    termination_date        TEXT,
    rehire_date             TEXT,
    position_status         TEXT,
    pay_frequency           TEXT,
    employee_type           TEXT,
    business_unit_code      TEXT,
    business_unit_desc      TEXT,
    home_department_desc    TEXT,
    home_department_code    TEXT,
    location_desc           TEXT,
    language_desc           TEXT,
    language_code           TEXT,
    legal_country_code      TEXT,
    pay_grade_code          TEXT,
    pay_grade_desc          TEXT,
    -- derived
    supervisor_id           TEXT,
    org_unit_1              TEXT,
    org_unit_2              TEXT,
    concur_profile          TEXT,
    travel_profile          TEXT,
    legal_country           TEXT,
    locale_code             TEXT,
    reimbursement_currency  TEXT,
    preferred_name          TEXT,
    concur_status           TEXT,
    term_date               TEXT,
    login_id                TEXT,
    -- Carried rather than derived. The non-US roster is maintained by hand on
    -- its own tab in Concur's own terms, so these arrive already answered.
    ledger_code             TEXT,
    org_unit_2_raw          TEXT,
    password                TEXT,
    -- housekeeping
    roster                  TEXT    NOT NULL DEFAULT 'us',    -- 'us' | 'non_us'
    source                  TEXT    NOT NULL DEFAULT 'adp',   -- 'adp' | 'manual' | 'workbook'
    row_state               TEXT    NOT NULL DEFAULT 'unchanged',
    include_305             INTEGER NOT NULL DEFAULT 1,
    include_350             INTEGER NOT NULL DEFAULT 1,
    include_360             INTEGER NOT NULL DEFAULT 1,
    -- On by default like the three above: SAP wants Login ID changes carried
    -- by the 320 rather than the 305, so the normal population needs one.
    include_320             INTEGER NOT NULL DEFAULT 1,
    import_id               INTEGER,
    source_row              INTEGER,
    -- ADP sends one row per employment record, so a transfer arrives as the
    -- same File Number twice: terminated under the old payroll company and
    -- active under the new one. These say how many rows the person had and
    -- which one was taken.
    duplicate_rows          INTEGER NOT NULL DEFAULT 1,
    duplicate_note          TEXT,
    derived_at              TEXT,
    created_at              TEXT    NOT NULL DEFAULT (datetime('now')),
    modified_at             TEXT,
    UNIQUE (file_number)
);

-- ------------------------------------------------------------------ maps
-- The six lookup tabs. Each one keeps its own table rather than a generic
-- key/value pair table, because the columns differ and the editor shows them.

CREATE TABLE IF NOT EXISTS ADP_Concur_StatusMap (
    map_key         INTEGER PRIMARY KEY AUTOINCREMENT,
    position_status TEXT NOT NULL,
    concur_status   TEXT,
    row_state       TEXT NOT NULL DEFAULT 'unchanged',
    UNIQUE (position_status)
);

CREATE TABLE IF NOT EXISTS ADP_Concur_CountryMap (
    map_key      INTEGER PRIMARY KEY AUTOINCREMENT,
    adp_country  TEXT NOT NULL,      -- 'Legal / Preferred Address: Country Code'
    concur_country TEXT,             -- the two-character code Concur wants
    row_state    TEXT NOT NULL DEFAULT 'unchanged',
    UNIQUE (adp_country)
);

CREATE TABLE IF NOT EXISTS ADP_Concur_OrgMap (
    map_key              INTEGER PRIMARY KEY AUTOINCREMENT,
    business_unit_desc   TEXT NOT NULL,
    home_department_desc TEXT,
    home_department_code TEXT,
    org_unit_1           TEXT,       -- 'Concur Org 1 (and others)'
    org_unit_2           TEXT,       -- 'Concur / ERP Code Org 2'
    default_language     TEXT,       -- 'Default Org Language', e.g. 'en_'
    currency             TEXT,       -- 'Reimbursement Currency'
    row_state            TEXT NOT NULL DEFAULT 'unchanged',
    UNIQUE (business_unit_desc, home_department_code)
);

CREATE TABLE IF NOT EXISTS ADP_Concur_LanguageMap (
    map_key       INTEGER PRIMARY KEY AUTOINCREMENT,
    language_desc TEXT NOT NULL,     -- ADP 'Language Description'
    language_code TEXT,              -- ADP numeric code, informational
    adp_language  TEXT,              -- the 'xx_' stem the locale is built from
    row_state     TEXT NOT NULL DEFAULT 'unchanged',
    UNIQUE (language_desc)
);

CREATE TABLE IF NOT EXISTS ADP_Concur_SalaryMap (
    map_key        INTEGER PRIMARY KEY AUTOINCREMENT,
    pay_grade_code TEXT NOT NULL,
    pay_grade_desc TEXT,
    expense_map    TEXT,             -- 'Expense Map'  -> Concur Profile
    travel_map     TEXT,             -- 'Travel Map'   -> Travel Profile
    row_state      TEXT NOT NULL DEFAULT 'unchanged',
    UNIQUE (pay_grade_code)
);

-- Exceptions where ADP has no usable Supervisor ID: keyed on the employee's
-- own File Number, giving the supervisor to use instead. A row with a blank
-- supervisor is the top of the food chain and is left blank on purpose.
CREATE TABLE IF NOT EXISTS ADP_Concur_SupervisorMap (
    map_key           INTEGER PRIMARY KEY AUTOINCREMENT,
    file_number       TEXT NOT NULL,
    employee_name     TEXT,
    supervisor_name   TEXT,
    supervisor_id     TEXT,
    note              TEXT,
    row_state         TEXT NOT NULL DEFAULT 'unchanged',
    UNIQUE (file_number)
);

-- Country code -> country name, currency. The right-hand block of the Country
-- Map tab: a reference list keyed on the two-character Concur country code,
-- which is a different table from the ADP-country lookup on its left even
-- though the workbook keeps them side by side.
CREATE TABLE IF NOT EXISTS ADP_Concur_CountryRef (
    ref_key       INTEGER PRIMARY KEY AUTOINCREMENT,
    country_code  TEXT NOT NULL,
    country_name  TEXT,
    currency_code TEXT,
    currency_name TEXT,
    row_state     TEXT NOT NULL DEFAULT 'unchanged',
    UNIQUE (country_code)
);

-- Country code -> Concur locale. The right-hand block of the Language Map tab,
-- which the non-US sheet looks up by country rather than by language.
CREATE TABLE IF NOT EXISTS ADP_Concur_LocaleMap (
    locale_key   INTEGER PRIMARY KEY AUTOINCREMENT,
    country_code TEXT NOT NULL,
    locale_name  TEXT,
    locale_code  TEXT,
    row_state    TEXT NOT NULL DEFAULT 'unchanged',
    UNIQUE (country_code)
);

-- --------------------------------------------------------------- layouts
-- The 305 / 350 / 360 column layouts, read from the workbook so a new Concur
-- template can be adopted by dropping a new file rather than editing code.
CREATE TABLE IF NOT EXISTS ADP_Concur_Layouts (
    layout_key  INTEGER PRIMARY KEY AUTOINCREMENT,
    record_type TEXT    NOT NULL,    -- '305' | '350' | '360'
    position    INTEGER NOT NULL,    -- 1-based column number in the record
    column_ref  TEXT    NOT NULL,    -- 'A', 'AB', ... as the template shows it
    heading     TEXT,                -- the template's own heading text
    max_width   TEXT,                -- the width row, where the template has one
    UNIQUE (record_type, position)
);

-- --------------------------------------------------------------- imports
CREATE TABLE IF NOT EXISTS ADP_Concur_Imports (
    import_id    INTEGER PRIMARY KEY AUTOINCREMENT,
    file_name    TEXT    NOT NULL,
    file_path    TEXT,
    sheet_counts TEXT,                -- JSON: rows loaded per worksheet
    row_count    INTEGER NOT NULL DEFAULT 0,
    imported_at  TEXT    NOT NULL DEFAULT (datetime('now'))
);

-- ------------------------------------------------------------ exceptions
-- Rebuilt from scratch on every derive, so it always describes the data as it
-- stands rather than accumulating history.
CREATE TABLE IF NOT EXISTS ADP_Concur_Exceptions (
    exception_key INTEGER PRIMARY KEY AUTOINCREMENT,
    employee_key  INTEGER REFERENCES ADP_Concur_Employees (employee_key)
                  ON DELETE CASCADE,
    file_number   TEXT,
    employee_name TEXT,
    severity      TEXT NOT NULL,      -- 'error' | 'warning'
    field         TEXT,
    message       TEXT NOT NULL,
    detected_at   TEXT NOT NULL DEFAULT (datetime('now'))
);

-- --------------------------------------------------------------- selection
-- The people currently picked, for a scoped extract and for the 'only picked'
-- view on every tab. Held here rather than in the browser for three reasons:
-- a pilot selection is worth keeping overnight, every list can then filter on
-- it with a join instead of a query string carrying 150 keys, and the extract
-- and the screen can never disagree about who is in it.
-- `reason` records how each person got picked - a subtree, a filter, or by
-- hand - so a selection can be explained after the fact.
CREATE TABLE IF NOT EXISTS ADP_Concur_Selection (
    employee_key INTEGER PRIMARY KEY
                 REFERENCES ADP_Concur_Employees (employee_key) ON DELETE CASCADE,
    reason       TEXT,
    added_at     TEXT NOT NULL DEFAULT (datetime('now'))
);

-- ---------------------------------------------------------------- extracts
-- One row per flat file written, so a file can be traced back to what was in
-- it and re-downloaded.
CREATE TABLE IF NOT EXISTS ADP_Concur_Extracts (
    extract_key INTEGER PRIMARY KEY AUTOINCREMENT,
    file_name   TEXT    NOT NULL,
    file_path   TEXT,
    n_305       INTEGER NOT NULL DEFAULT 0,
    n_350       INTEGER NOT NULL DEFAULT 0,
    n_360       INTEGER NOT NULL DEFAULT 0,
    -- Always 0 on a 305/350/360 file and vice versa - the 320 is its own file,
    -- never merged with the other three. See ADP_Concur_export_320.
    n_320       INTEGER NOT NULL DEFAULT 0,
    scope       TEXT,                 -- 'active' | 'all'
    delimiter   TEXT,
    written_at  TEXT    NOT NULL DEFAULT (datetime('now'))
);

-- ------------------------------------------------------------ load results
-- What Concur said back. One ADP_Concur_Loads row per result workbook, one
-- ADP_Concur_Results row per message in it.
--
-- extract_key is how a message gets a name on it: Concur's Record Identifier
-- is a line number in the file that was sent, so without the extract it
-- points at nothing. It is nullable because a result can arrive for a file
-- this database no longer holds, and half a result is still worth reading.
CREATE TABLE IF NOT EXISTS ADP_Concur_Loads (
    load_key     INTEGER PRIMARY KEY AUTOINCREMENT,
    run_label    TEXT    NOT NULL,
    source_file  TEXT,
    extract_key  INTEGER,
    extract_name TEXT,
    n_error      INTEGER NOT NULL DEFAULT 0,
    n_warning    INTEGER NOT NULL DEFAULT 0,
    n_message    INTEGER NOT NULL DEFAULT 0,
    loaded_at    TEXT    NOT NULL DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS ADP_Concur_Results (
    result_key    INTEGER PRIMARY KEY AUTOINCREMENT,
    load_key      INTEGER NOT NULL,
    level         TEXT    NOT NULL,     -- Error | Warning | Info
    record_id     INTEGER NOT NULL,     -- the line number Concur quoted
    record_type   TEXT    NOT NULL DEFAULT '',   -- 100 | 305 | 350 | 360
    file_number   TEXT    NOT NULL DEFAULT '',
    employee_key  INTEGER,
    employee_name TEXT    NOT NULL DEFAULT '',
    category      TEXT    NOT NULL DEFAULT '',   -- the cause, for grouping
    field         TEXT    NOT NULL DEFAULT '',   -- the Concur field, when named
    message       TEXT    NOT NULL
);
"""

# Indexes are kept out of SCHEMA and applied after the migration below. An
# index names a column, and CREATE TABLE IF NOT EXISTS silently does nothing to
# a table that already exists - so on a database created by an earlier version
# the table is old, the index is new, and the whole script fails on a column
# that has not been added yet.
INDEXES = """
CREATE INDEX IF NOT EXISTS ix_adpc_emp_status ON ADP_Concur_Employees (position_status);
CREATE INDEX IF NOT EXISTS ix_adpc_emp_state  ON ADP_Concur_Employees (row_state);
CREATE INDEX IF NOT EXISTS ix_adpc_emp_source ON ADP_Concur_Employees (source);
CREATE INDEX IF NOT EXISTS ix_adpc_emp_bu     ON ADP_Concur_Employees (business_unit_desc);
CREATE INDEX IF NOT EXISTS ix_adpc_emp_roster ON ADP_Concur_Employees (roster);
CREATE INDEX IF NOT EXISTS ix_adpc_orgmap_dept ON ADP_Concur_OrgMap (home_department_code);
CREATE INDEX IF NOT EXISTS ix_adpc_orgmap_bu   ON ADP_Concur_OrgMap (business_unit_desc);
CREATE INDEX IF NOT EXISTS ix_adpc_exc_emp ON ADP_Concur_Exceptions (employee_key);
CREATE INDEX IF NOT EXISTS ix_adpc_exc_sev ON ADP_Concur_Exceptions (severity);
CREATE INDEX IF NOT EXISTS ix_adpc_res_load ON ADP_Concur_Results (load_key);
CREATE INDEX IF NOT EXISTS ix_adpc_res_emp  ON ADP_Concur_Results (file_number);
CREATE INDEX IF NOT EXISTS ix_adpc_res_cat  ON ADP_Concur_Results (category);
"""

# Every column added to ADP_Concur_Employees after the first release. SQLite
# cannot add a column through CREATE TABLE IF NOT EXISTS, so an existing
# database - and doppio.db is shared with the M3_Security tables, so it is
# never going to be a fresh one - needs them applied by hand.
#
# Append here whenever a column is added to SCHEMA above. Listing one that is
# already present is harmless: SQLite refuses it and the migration moves on.
MIGRATIONS: list[tuple[str, str, str]] = [
    ("ADP_Concur_Employees", "duplicate_rows", "INTEGER NOT NULL DEFAULT 1"),
    ("ADP_Concur_Employees", "duplicate_note", "TEXT"),
    ("ADP_Concur_Employees", "ledger_code", "TEXT"),
    ("ADP_Concur_Employees", "org_unit_2_raw", "TEXT"),
    ("ADP_Concur_Employees", "password", "TEXT"),
    ("ADP_Concur_Employees", "roster", "TEXT NOT NULL DEFAULT 'us'"),
    ("ADP_Concur_Employees", "include_320", "INTEGER NOT NULL DEFAULT 1"),
    ("ADP_Concur_Extracts", "n_320", "INTEGER NOT NULL DEFAULT 0"),
]


def migrate(conn: sqlite3.Connection) -> list[str]:
    """
    Add any columns an older database is missing, and report what was added.

    Run between the tables and the indexes. Each one is attempted and the
    'duplicate column name' refusal is taken as "already there" - which is the
    whole test, so there is no need to read the schema first.
    """
    applied = []
    for table, column, ddl in MIGRATIONS:
        if not conn.execute(
            "SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = ?",
            (table,)).fetchone():
            continue                      # a fresh database; SCHEMA built it
        try:
            conn.execute(f"ALTER TABLE {table} ADD COLUMN {column} {ddl}")
            applied.append(f"{table}.{column}")
        except sqlite3.OperationalError as exc:
            if "duplicate column" not in str(exc).lower():
                raise
    return applied


def resolve_db_path(db_path: str | None = None) -> str:
    """--db wins, then ADP_CONCUR_DB, then ~/sqlite/doppio.db."""
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
        # Said out loud rather than done silently: a schema change to a
        # database three other tools share is worth seeing in the log.
        logging.getLogger("ADP_Concur_Db").info(
            "Added %d column(s) to the existing schema: %s",
            len(applied), ", ".join(applied))
    return conn

def counts(conn: sqlite3.Connection) -> dict:
    """Row counts for the status bar - one query per table it names."""
    def n(sql: str, *args) -> int:
        return conn.execute(sql, args).fetchone()[0]

    # Per source, built from the registry rather than written out one key at a
    # time, so the per-source counts always add up to the total. They stopped
    # adding up once the non-US tabs started writing a third source that only
    # two of these lines knew about.
    per_source = {
        f"employees_{s}": n("SELECT COUNT(*) FROM ADP_Concur_Employees "
                            "WHERE source = ? AND row_state <> 'deleted'", s)
        for s in EMPLOYEE_SOURCES}

    return {
        **per_source,
        "employees": n("SELECT COUNT(*) FROM ADP_Concur_Employees WHERE row_state <> 'deleted'"),
        "employees_active": n("SELECT COUNT(*) FROM ADP_Concur_Employees "
                              "WHERE concur_status = 'Y' AND row_state <> 'deleted'"),
        "employees_changed": n("SELECT COUNT(*) FROM ADP_Concur_Employees "
                               "WHERE row_state IN ('new','modified')"),
        "org_map": n("SELECT COUNT(*) FROM ADP_Concur_OrgMap"),
        "status_map": n("SELECT COUNT(*) FROM ADP_Concur_StatusMap"),
        "country_map": n("SELECT COUNT(*) FROM ADP_Concur_CountryMap"),
        "language_map": n("SELECT COUNT(*) FROM ADP_Concur_LanguageMap"),
        "salary_map": n("SELECT COUNT(*) FROM ADP_Concur_SalaryMap"),
        "supervisor_map": n("SELECT COUNT(*) FROM ADP_Concur_SupervisorMap"),
        "country_ref": n("SELECT COUNT(*) FROM ADP_Concur_CountryRef"),
        "locale_map": n("SELECT COUNT(*) FROM ADP_Concur_LocaleMap"),
        "employees_us": n("SELECT COUNT(*) FROM ADP_Concur_Employees "
                          "WHERE roster = 'us' AND row_state <> 'deleted'"),
        "employees_non_us": n("SELECT COUNT(*) FROM ADP_Concur_Employees "
                              "WHERE roster = 'non_us' AND row_state <> 'deleted'"),
        "errors": n("SELECT COUNT(*) FROM ADP_Concur_Exceptions WHERE severity = 'error'"),
        "warnings": n("SELECT COUNT(*) FROM ADP_Concur_Exceptions WHERE severity = 'warning'"),
        "layouts": n("SELECT COUNT(DISTINCT record_type) FROM ADP_Concur_Layouts"),
        "picked": n("SELECT COUNT(*) FROM ADP_Concur_Selection s "
                    "JOIN ADP_Concur_Employees e ON e.employee_key = s.employee_key "
                    "WHERE e.row_state <> 'deleted'"),
    }

# --------------------------------------------------------------- selection

def selection_keys(conn: sqlite3.Connection) -> list[int]:
    """
    The picked employees, skipping anyone since deleted.

    Read through a join rather than straight off the table so a person left
    out of the load stops counting as picked without anyone having to
    remember to untick them.
    """
    return [r[0] for r in conn.execute(
        "SELECT s.employee_key FROM ADP_Concur_Selection s "
        "JOIN ADP_Concur_Employees e ON e.employee_key = s.employee_key "
        "WHERE e.row_state <> 'deleted' ORDER BY s.employee_key")]

def selection_summary(conn: sqlite3.Connection) -> dict:
    """How many are picked, and how they got there."""
    rows = [dict(r) for r in conn.execute(
        """
        SELECT s.reason, COUNT(*) AS n
          FROM ADP_Concur_Selection s
          JOIN ADP_Concur_Employees e ON e.employee_key = s.employee_key
         WHERE e.row_state <> 'deleted'
         GROUP BY s.reason ORDER BY n DESC
        """)]
    return {"total": sum(r["n"] for r in rows), "reasons": rows,
            "label": "; ".join(r["reason"] for r in rows if r["reason"])[:300]}

def selection_add(conn: sqlite3.Connection, keys: list[int], reason: str = "",
                  commit: bool = True) -> int:
    """
    Add people to the selection.

    Adding somebody already picked leaves their original reason alone - the
    first answer to 'why is this person in the pilot' is the interesting one.
    """
    if not keys:
        return 0
    before = conn.execute("SELECT COUNT(*) FROM ADP_Concur_Selection").fetchone()[0]
    conn.executemany(
        "INSERT OR IGNORE INTO ADP_Concur_Selection (employee_key, reason) "
        "VALUES (?, ?)", [(int(k), reason) for k in keys])
    after = conn.execute("SELECT COUNT(*) FROM ADP_Concur_Selection").fetchone()[0]
    if commit:
        conn.commit()
    return after - before

def selection_remove(conn: sqlite3.Connection, keys: list[int],
                     commit: bool = True) -> int:
    if not keys:
        return 0
    cur = conn.cursor()
    cur.execute(
        f"DELETE FROM ADP_Concur_Selection WHERE employee_key IN "
        f"({','.join('?' * len(keys))})", [int(k) for k in keys])
    if commit:
        conn.commit()
    return cur.rowcount

def selection_clear(conn: sqlite3.Connection, commit: bool = True) -> int:
    cur = conn.cursor()
    cur.execute("DELETE FROM ADP_Concur_Selection")
    if commit:
        conn.commit()
    return cur.rowcount

def picked_clause(picked: str, alias: str = "e") -> str:
    """
    The SQL fragment behind the 'only picked' view, shared by every list.

    'only' keeps the picked, 'not' keeps the rest, anything else keeps
    everyone. One place so the Employees tab, the record tabs and the
    exceptions can never disagree about what picked means.
    """
    if picked == "only":
        return (f" AND EXISTS (SELECT 1 FROM ADP_Concur_Selection s "
                f"WHERE s.employee_key = {alias}.employee_key)")
    if picked == "not":
        return (f" AND NOT EXISTS (SELECT 1 FROM ADP_Concur_Selection s "
                f"WHERE s.employee_key = {alias}.employee_key)")
    return ""

# The two places an employee can have come from, and what clearing each one
# actually costs. The asymmetry is the whole reason this is a choice rather
# than a button: an ADP row is a copy of something that still exists in the
# ADP report, so it comes back on the next drop. A hand-keyed person exists
# nowhere else - clearing them destroys the only copy.
# Every value the `source` column can hold. This is the registry that decides
# what a clear can reach and what the Extract tab offers, so a source missing
# from here is a source nobody can delete.
#
# 'workbook' was exactly that. The non-US 305 importer writes it, the Employees
# filter lists it as 'Non-US tab', and it was never added here - so 'clear
# employees' with both boxes ticked left 82 non-US people behind and the
# preview never mentioned them. Adding a source value means adding it in both
# places; there is no default that covers a name the registry has not seen.
EMPLOYEE_SOURCES = {
    "adp": {
        "label": "From ADP",
        "recoverable": True,
        "note": "Comes back the moment the workbook is dropped again.",
    },
    "workbook": {
        "label": "From the non-US tabs",
        "recoverable": True,
        "note": "Typed into the workbook's '305 Non US Non SE' and "
                "'360 Non US Non SE' tabs rather than fed by ADP. Comes back "
                "when the same workbook is dropped again - but only if that "
                "workbook still carries them.",
    },
    "manual": {
        "label": "Added by hand",
        "recoverable": False,
        "note": "Keyed in here and held nowhere else - no workbook will bring "
                "them back. This includes anyone created from the Fixes tab to "
                "repair a broken supervisor chain.",
    },
}

def clear_preview(conn: sqlite3.Connection,
                  sources: list[str] | tuple[str, ...] = ("adp",)) -> dict:
    """
    Exactly what a clear would remove, before anything is removed.

    Reported per source, with what goes with the people: their exceptions, and
    their rows in the selection. The maps are deliberately *not* touched by
    this - a correction filed in the Supervisor Map survives, which is the
    point of having put it there.
    """
    wanted = [s for s in sources if s in EMPLOYEE_SOURCES]
    out = {"sources": {}, "total": 0, "picked": 0, "unrecoverable": 0,
           "changed": 0}
    for source in EMPLOYEE_SOURCES:
        row = conn.execute(
            """
            SELECT COUNT(*) AS n,
                   COALESCE(SUM(row_state IN ('new','modified')), 0) AS changed,
                   COALESCE(SUM(EXISTS (SELECT 1 FROM ADP_Concur_Selection s
                                 WHERE s.employee_key = e.employee_key)), 0) AS picked
              FROM ADP_Concur_Employees e WHERE e.source = ?
            """, (source,)).fetchone()
        entry = {"label": EMPLOYEE_SOURCES[source]["label"],
                 "recoverable": EMPLOYEE_SOURCES[source]["recoverable"],
                 "note": EMPLOYEE_SOURCES[source]["note"],
                 "n": row["n"], "changed": row["changed"], "picked": row["picked"],
                 "selected": source in wanted}
        out["sources"][source] = entry
        if source in wanted:
            out["total"] += row["n"]
            out["picked"] += row["picked"]
            out["changed"] += row["changed"]
            if not EMPLOYEE_SOURCES[source]["recoverable"]:
                out["unrecoverable"] += row["n"]

    marks = ",".join("?" * len(wanted)) or "''"
    out["exceptions"] = conn.execute(
        f"""
        SELECT COUNT(*) FROM ADP_Concur_Exceptions x
          JOIN ADP_Concur_Employees e ON e.employee_key = x.employee_key
         WHERE e.source IN ({marks})
        """, wanted).fetchone()[0] if wanted else 0
    out["kept"] = conn.execute(
        f"SELECT COUNT(*) FROM ADP_Concur_Employees WHERE source NOT IN ({marks})",
        wanted).fetchone()[0] if wanted else conn.execute(
        "SELECT COUNT(*) FROM ADP_Concur_Employees").fetchone()[0]
    return out

def clear_employees(conn: sqlite3.Connection,
                    sources: list[str] | tuple[str, ...] = ("adp",),
                    commit: bool = True) -> dict:
    """
    Empty the employee table, by source, so the next drop starts from scratch.

    Naming the sources rather than a keep_manual flag, because clearing the
    hand-keyed people is a genuinely different decision from clearing the ADP
    ones and should have to be asked for by name. The default is 'adp' alone:
    a fresh ADP file is no reason to lose the people ADP does not have.

    The maps and the captured layouts are untouched - clear_maps() does those,
    and a correction filed in the Supervisor Map is meant to outlive the
    employees it was about.
    """
    wanted = [s for s in sources if s in EMPLOYEE_SOURCES]
    if not wanted:
        raise ValueError("Name at least one source: "
                         + ", ".join(EMPLOYEE_SOURCES))
    before = clear_preview(conn, wanted)

    marks = ",".join("?" * len(wanted))
    cur = conn.cursor()
    cur.execute(f"DELETE FROM ADP_Concur_Employees WHERE source IN ({marks})",
                wanted)
    removed = cur.rowcount
    # The selection cascades on the foreign key; the exceptions do not carry
    # one on every row, so they are swept explicitly.
    cur.execute("DELETE FROM ADP_Concur_Exceptions WHERE employee_key NOT IN "
                "(SELECT employee_key FROM ADP_Concur_Employees)")
    cur.execute("DELETE FROM ADP_Concur_Selection WHERE employee_key NOT IN "
                "(SELECT employee_key FROM ADP_Concur_Employees)")
    # Load results survive - what Concur said is history, and it carries the
    # file number and the name of its own - but the key that pointed at the
    # employee row is now dangling, and a dangling key opens the wrong editor.
    cur.execute("UPDATE ADP_Concur_Results SET employee_key = NULL "
                "WHERE employee_key IS NOT NULL AND employee_key NOT IN "
                "(SELECT employee_key FROM ADP_Concur_Employees)")
    if commit:
        conn.commit()
    return {"removed": removed, "sources": wanted, "kept": before["kept"],
            "unrecoverable": before["unrecoverable"],
            "picked_lost": before["picked"]}

# The log tables: what has happened, as opposed to what the load is made of.
# Kept separate from employees and maps because they are a different kind of
# thing to throw away - nothing downstream is built from them, and clearing
# them loses history rather than data.
HISTORY_PARTS = {
    "imports": {
        "label": "Workbook loads",
        "tables": ["ADP_Concur_Imports"],
        "note": "Every workbook dropped, with its per-sheet counts.",
    },
    "extracts": {
        "label": "Files written",
        "tables": ["ADP_Concur_Extracts"],
        "note": "The record of every flat file written. The files themselves "
                "stay on disk - this only forgets them, and a load result can "
                "no longer be matched to one by name.",
    },
    "results": {
        "label": "Concur load results",
        "tables": ["ADP_Concur_Results", "ADP_Concur_Loads"],
        "note": "Everything Concur has said back, and the Results tab with it.",
    },
}


def history_preview(conn: sqlite3.Connection) -> dict:
    """What each history part currently holds, for the confirm step."""
    out = {}
    for part, spec in HISTORY_PARTS.items():
        n = sum(conn.execute(f"SELECT COUNT(*) FROM {t}").fetchone()[0]
                for t in spec["tables"])
        out[part] = {"label": spec["label"], "note": spec["note"], "rows": n}
    return out


def clear_history(conn: sqlite3.Connection,
                  parts: list[str] | tuple[str, ...] = tuple(HISTORY_PARTS),
                  commit: bool = True) -> dict:
    """
    Empty the log tables, so 'start again' really does start again.

    Named parts rather than one switch, because forgetting which workbooks were
    loaded and throwing away what Concur said about the last run are different
    decisions - the second is the one somebody actually wants to keep.

    Child tables are listed before their parents in HISTORY_PARTS so the
    deletes go in an order that never orphans a row.
    """
    wanted = [p for p in parts if p in HISTORY_PARTS]
    if not wanted:
        raise ValueError("Name at least one part: " + ", ".join(HISTORY_PARTS))
    out = {}
    cur = conn.cursor()
    for part in wanted:
        removed = 0
        for table in HISTORY_PARTS[part]["tables"]:
            cur.execute(f"DELETE FROM {table}")
            removed += cur.rowcount
        out[part] = removed
    if commit:
        conn.commit()
    return out


def clear_maps(conn: sqlite3.Connection, commit: bool = True) -> dict:
    """Empty the six lookup tables and the captured layouts."""
    tables = ["ADP_Concur_StatusMap", "ADP_Concur_CountryMap", "ADP_Concur_OrgMap",
              "ADP_Concur_LanguageMap", "ADP_Concur_SalaryMap",
              "ADP_Concur_SupervisorMap", "ADP_Concur_CountryRef",
              "ADP_Concur_LocaleMap", "ADP_Concur_Layouts"]
    out = {}
    cur = conn.cursor()
    for t in tables:
        cur.execute(f"DELETE FROM {t}")
        out[t] = cur.rowcount
    if commit:
        conn.commit()
    return out

if __name__ == "__main__":
    import argparse

    ap = argparse.ArgumentParser(description="Create the ADP_Concur_* schema.")
    ap.add_argument("--db", default=None, help=f"SQLite path (default {DEFAULT_DB_PATH})")
    args = ap.parse_args()

    c = connect(args.db)
    print(f"Schema ready in {resolve_db_path(args.db)}")
    for t in c.execute(
        "SELECT name FROM sqlite_master WHERE type='table' "
        "AND name LIKE 'ADP_Concur_%' ORDER BY name"
    ):
        print("  ", t[0])
    c.close()
