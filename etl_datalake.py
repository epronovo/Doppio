#!/usr/bin/env python3
"""
etl_datalake.py - Mini ETL routine for Infor Data Lake -> SQLite

Runs every 10 minutes. Stores the last run date/time in the LastRun table
and picks up from that date on restart. If the table is not found, the
current date/time is used.

Steps:
  1. List .ionapi files in IONAPI_DIR; the user picks an environment and a
     polling interval on the status webpage, then clicks Start.
  2. Ping the server; stop the routine (until restarted from the webpage)
     if not active.
  3. If ping is OK, display the current build number of the Infor Data Lake.
  4. Get the dl_ids to process from the last run date (dataobjects list API).
  5. Get the details (download) for each dl_id.
  6. Process the file: table name = dl_document_name, columns = the column
     names in the download. Create the table if it does not exist and insert.
  7. Serve a local status webpage (http://WEB_HOST:WEB_PORT/) with Start/Stop
     controls and a live-editable polling interval, so the routine can be
     driven and watched from a browser.

Uses only the Python standard library.
"""

import glob
import io
import json
import os
import shutil
import sqlite3
import sys
import tempfile
import threading
import time
import urllib.parse
import urllib.request
import zipfile
import zlib
from collections import deque
from datetime import datetime, timedelta, timezone
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

# The zip files dropped on the status page are M3 Grid Access binary table
# exports (the same format m3_unpacker.py reads). Reuse its parser so the
# on-wire decoding stays in one place. m3_unpacker.py lives next to this file.
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import m3_unpacker as m3

# ---------------------------------------------------------------- settings
DB_PATH = "/Users/ericpronovost/sqlite/etl.db"
IONAPI_DIR = "/Users/ericpronovost/Doppio/ionapi"
DEFAULT_ENV = "DOPPIO_DEM"
ROUTINE_NAME = "etl_datalake"
INTERVAL_SECONDS = 60  # 10 minutes
PAGE_RECORDS = 500
WEB_HOST = "127.0.0.1"
WEB_PORT = 8787
LOG_MAX_LINES = 300
M3_CONO = "001"  # company used for the MNS120MI.Get metadata lookup (table keys)


def etl_utcnow_iso():
    """Current UTC time as ISO-8601 with milliseconds, e.g. 2026-07-02T14:00:00.000Z"""
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%S.%f")[:-3] + "Z"


def etl_iso_plus_seconds(seconds):
    return (datetime.now(timezone.utc) + timedelta(seconds=seconds)).strftime(
        "%Y-%m-%dT%H:%M:%S.%f")[:-3] + "Z"


# --------------------------------------------------------------- shared state
# Guards STATUS/LOG_LINES, which the web status page reads from its own thread.
_STATE_LOCK = threading.Lock()
STATUS = {
    "environment": None,
    "state": "waiting-for-selection",
    "available_environments": [],
    "default_environment": DEFAULT_ENV,
    "poll_interval_seconds": INTERVAL_SECONDS,
    "cycle_count": 0,
    "cycle_start": None,
    "last_cycle_end": None,
    "next_run_at": None,
    "build": None,
    "objects_total": 0,
    "objects_done": 0,
    "rows_loaded_last_cycle": 0,
    "rows_loaded_total": 0,
    "last_error": None,
    "since_override": None,
}
LOG_LINES = deque(maxlen=LOG_MAX_LINES)

# Set while the routine should be actively cycling; cleared by Stop (or a
# ping failure) to pause it. The webpage's Start/Stop buttons toggle this.
_RUN_EVENT = threading.Event()
# Set by the Now button to run a single cycle immediately, regardless of
# whether continuous mode (_RUN_EVENT) is on.
_NOW_EVENT = threading.Event()
# Set by Stop to interrupt a cycle that's already in flight.
_ABORT_EVENT = threading.Event()
# Set while a dropped zip is being loaded into the database. Blocks Start and
# further uploads so a bulk load and the tracking loop never write at once.
_LOADING_EVENT = threading.Event()
# Object IDs (dl_ids) the user asked to force-load via the webpage. Consumed by
# the run loop on its own thread/connection; only accepted while running.
_FORCE_EVENT = threading.Event()
_FORCE_QUEUE = deque()
_SELECTED_CFG = None
# One-shot override for the next cycle's "since" cutoff; cleared once used.
_SINCE_OVERRIDE = None
MIN_POLL_INTERVAL_SECONDS = 5


class EtlPingFailed(Exception):
    """Raised when the data lake ping check fails; pauses the run loop."""


def etl_log(msg):
    """Print to the console and record the line for the web status page."""
    print(msg)
    line = f"[{etl_utcnow_iso()}] {msg}"
    with _STATE_LOCK:
        LOG_LINES.append(line)


def etl_set_status(**kwargs):
    with _STATE_LOCK:
        STATUS.update(kwargs)


def etl_status_snapshot():
    with _STATE_LOCK:
        return dict(STATUS), list(LOG_LINES)


# ------------------------------------------------------- step 1: environment
def etl_list_environments():
    files = sorted(glob.glob(os.path.join(IONAPI_DIR, "*.ionapi")))
    if not files:
        etl_log(f"No .ionapi files found in {IONAPI_DIR}")
        sys.exit(1)
    return [os.path.splitext(os.path.basename(f))[0] for f in files]


def etl_load_environment(name):
    path = os.path.join(IONAPI_DIR, f"{name}.ionapi")
    with open(path) as fh:
        return json.load(fh)


def etl_start_run(name, interval_seconds):
    """Called from the web handler when Start is pressed."""
    global _SELECTED_CFG
    cfg = etl_load_environment(name)
    with _STATE_LOCK:
        _SELECTED_CFG = cfg
    etl_set_status(environment=name, poll_interval_seconds=interval_seconds,
                    state="starting", last_error=None)
    etl_log(f"Started via web: environment={name}, poll interval={interval_seconds}s")
    _RUN_EVENT.set()


def etl_stop_run():
    """Called from the web handler when Stop is pressed."""
    _RUN_EVENT.clear()
    _ABORT_EVENT.set()
    etl_set_status(state="stopping")
    etl_log("Stop requested via web.")


def etl_trigger_now(name):
    """Called from the web handler when Now is pressed; runs one cycle
    immediately without waiting for the poll interval or turning on
    continuous mode (unless continuous mode is already on)."""
    global _SELECTED_CFG
    cfg = etl_load_environment(name)
    with _STATE_LOCK:
        _SELECTED_CFG = cfg
    etl_set_status(environment=name, last_error=None)
    etl_log(f"Immediate run requested via web: environment={name}")
    _NOW_EVENT.set()


def etl_update_interval(interval_seconds):
    etl_set_status(poll_interval_seconds=interval_seconds)
    etl_log(f"Poll interval changed to {interval_seconds}s via web.")


def etl_normalize_since(raw):
    """Accepts an ISO-ish datetime (e.g. from an HTML datetime-local input,
    with no timezone) and returns it in the same Zulu format used elsewhere.
    A bare date/time with no offset is treated as UTC."""
    dt = datetime.fromisoformat(raw.strip().replace("Z", "+00:00"))
    dt = dt.replace(tzinfo=timezone.utc) if dt.tzinfo is None else dt.astimezone(timezone.utc)
    return dt.strftime("%Y-%m-%dT%H:%M:%S.%f")[:-3] + "Z"


def etl_set_since_override(value):
    """Called from the web handler; overrides the "since" cutoff used to
    pick up data objects for just the next cycle, then reverts to the
    normal LastRun-based progression."""
    global _SINCE_OVERRIDE
    with _STATE_LOCK:
        _SINCE_OVERRIDE = value
    etl_set_status(since_override=value)
    if value:
        etl_log(f"Next run will pick up from: {value} (override via web)")
    else:
        etl_log("Since override cleared via web.")


def etl_authenticate(cfg):
    """OAuth2 resource-owner (service account) grant using the .ionapi file."""
    token_url = cfg["pu"] + cfg["ot"]
    data = urllib.parse.urlencode({
        "grant_type": "password",
        "username": cfg["saak"],
        "password": cfg["sask"],
        "client_id": cfg["ci"],
        "client_secret": cfg["cs"],
    }).encode()
    req = urllib.request.Request(token_url, data=data, method="POST",
                                 headers={"Content-Type": "application/x-www-form-urlencoded"})
    with urllib.request.urlopen(req, timeout=60) as resp:
        tok = json.loads(resp.read().decode())
    return tok["access_token"]


def etl_base_url(cfg):
    return f"{cfg['iu'].rstrip('/')}/{cfg['ti']}/DATAFABRIC/datalake/v2"


def etl_m3_base_url(cfg):
    return (f"{cfg['iu'].rstrip('/')}/{cfg['ti']}/M3/m3api-rest/v2/execute"
            f"?maxrecs=100&extendedresult=true&righttrim=true&cono={M3_CONO}")


def etl_post_json(url, token, payload):
    """POST a JSON body with bearer auth; returns the parsed JSON response."""
    data = json.dumps(payload).encode()
    headers = {
        "accept": "application/json; charset=UTF-8",
        "Content-Type": "application/json; charset=UTF-8",
        "Authorization": f"Bearer {token}",
    }
    req = urllib.request.Request(url, data=data, method="POST", headers=headers)
    with urllib.request.urlopen(req, timeout=120) as resp:
        return json.loads(resp.read().decode())


def etl_get_table_keys(cfg, token, table):
    """MNS120MI.Get on "{table}00" (the unique-key file variant) returns the
    file's key columns in KEY1-KEY9 (unused slots are empty strings), named
    as "{2-char file mnemonic}{4-char generic field}" (e.g. "CTCONO" for
    CSYTAB). The Data Lake NDJSON records drop that 2-char file prefix (the
    same row is just "CONO"), so strip it here to match the actual columns."""
    payload = {
        "program": "MNS120MI",
        "transactions": [{"transaction": "Get", "record": {"FILE": f"{table}00"}}],
    }
    data = etl_post_json(etl_m3_base_url(cfg), token, payload)
    records = data.get("results", [{}])[0].get("records", [])
    if not records:
        return []
    rec = records[0]
    raw_keys = [rec[f"KEY{i}"] for i in range(1, 10) if rec.get(f"KEY{i}", "").strip()]
    return [k[2:] for k in raw_keys]


def etl_get(url, token, accept="application/json", deflate=False):
    """GET with bearer auth; returns raw bytes (deflate-decompressed if needed)."""
    headers = {"accept": accept, "Authorization": f"Bearer {token}"}
    if deflate:
        headers["Accept-Encoding"] = "deflate"
    req = urllib.request.Request(url, headers=headers)
    with urllib.request.urlopen(req, timeout=120) as resp:
        body = resp.read()
        enc = resp.headers.get("Content-Encoding", "")
    if enc == "deflate" or deflate:
        for wbits in (zlib.MAX_WBITS, -zlib.MAX_WBITS, zlib.MAX_WBITS | 16):
            try:
                return zlib.decompress(body, wbits)
            except zlib.error:
                continue
    return body


# ------------------------------------------------------------- step 2: ping
def etl_ping(base, token):
    try:
        body = etl_get(f"{base}/ping", token, accept="*/*").decode().strip()
        etl_log(f"Ping: {body}")
        return body.strip('"').upper() == "OK"
    except Exception as e:
        etl_log(f"Ping failed: {e}")
        return False


# ---------------------------------------------------------- step 3: version
def etl_show_version(base, token):
    body = etl_get(f"{base}/version", token, accept="*/*").decode().strip()
    try:
        info = json.loads(body)
    except json.JSONDecodeError:
        info = None
    if isinstance(info, dict):
        build = info.get("build") or info.get("buildNumber") or info.get("version") or body
    elif isinstance(info, str):
        build = info
    else:
        build = body
    etl_log(f"Infor Data Lake build: {build}")
    return build


# ------------------------------------------------------------ LastRun table
def etl_get_last_run(conn):
    cur = conn.execute(
        "SELECT name FROM sqlite_master WHERE type='table' AND name='LastRun'")
    if cur.fetchone() is None:
        conn.execute("CREATE TABLE LastRun (RoutineName TEXT PRIMARY KEY, LastRun TEXT)")
        conn.commit()
        return etl_utcnow_iso()  # table not found -> use current date/time
    row = conn.execute(
        "SELECT LastRun FROM LastRun WHERE RoutineName=?", (ROUTINE_NAME,)).fetchone()
    return row[0] if row else etl_utcnow_iso()


def etl_save_last_run(conn, when):
    conn.execute(
        "INSERT INTO LastRun (RoutineName, LastRun) VALUES (?, ?) "
        "ON CONFLICT(RoutineName) DO UPDATE SET LastRun=excluded.LastRun",
        (ROUTINE_NAME, when))
    conn.commit()


# ----------------------------------------------------- step 4: list objects
def etl_list_dataobjects(base, token, since):
    """Return all data objects with dl_document_date >= since (paged)."""
    flt = urllib.parse.quote(f'dl_document_date ge "{since}"')
    objects, page = [], 1
    while True:
        url = (f"{base}/dataobjects?filter={flt}"
               f"&sort={urllib.parse.quote('dl_document_date:asc')}"
               f"&page={page}&records={PAGE_RECORDS}")
        data = json.loads(etl_get(url, token).decode())
        fields = data.get("fields", [])
        objects.extend(fields)
        if len(objects) >= data.get("numFound", 0) or not fields:
            break
        page += 1
    return objects


def etl_get_object_meta(base, token, dl_id):
    """Look up one data object's metadata by dl_id (used by the force-load
    feature to discover the object's target table, dl_document_name). Returns
    the metadata dict or None if the id isn't found."""
    flt = urllib.parse.quote(f'dl_id eq "{dl_id}"')
    url = f"{base}/dataobjects?filter={flt}&page=1&records=1"
    data = json.loads(etl_get(url, token).decode())
    fields = data.get("fields", [])
    return fields[0] if fields else None


# ------------------------------------------------- step 5: download details
def etl_download_details(base, token, dl_id):
    """Download one data object; returns a list of dict records (NDJSON)."""
    raw = etl_get(f"{base}/dataobjects/{dl_id}", token,
                  accept="application/octet-stream", deflate=True)
    records = []
    for line in raw.decode("utf-8", errors="replace").splitlines():
        line = line.strip()
        if line:
            records.append(json.loads(line))
    return records


# ---------------------------------------------------------- step 6: process
def etl_sql_type(value):
    if isinstance(value, bool):
        return "INTEGER"
    if isinstance(value, int):
        return "INTEGER"
    if isinstance(value, float):
        return "REAL"
    return "TEXT"


def etl_sql_value(value):
    if isinstance(value, bool):
        return int(value)
    if isinstance(value, (dict, list)):
        return json.dumps(value)
    return value


def etl_table_columns(conn, table):
    return [r[1] for r in conn.execute(f'PRAGMA table_info("{table}")')]


def etl_view_exists(conn, view):
    cur = conn.execute(
        "SELECT name FROM sqlite_master WHERE type='view' AND name=?", (view,))
    return cur.fetchone() is not None


def etl_index_exists(conn, index):
    cur = conn.execute(
        "SELECT name FROM sqlite_master WHERE type='index' AND name=?", (index,))
    return cur.fetchone() is not None


def etl_ensure_variation_unique(conn, table, existing_columns):
    """Every table carries a variationNumber; enforce one row per
    variationNumber so re-downloading the same version (e.g. after a
    since override re-pulls an overlapping date range) doesn't duplicate
    rows in the raw table."""
    if "variationNumber" not in existing_columns:
        return
    index = f"{table}_variationNumber_uq"
    if etl_index_exists(conn, index):
        return
    try:
        conn.execute(f'CREATE UNIQUE INDEX "{index}" ON "{table}" ("variationNumber")')
        conn.commit()
    except sqlite3.IntegrityError as e:
        etl_log(f"  Could not enforce unique variationNumber on {table}: {e}")


def etl_ensure_latest_view(conn, table, cfg, token):
    """Create the "{table}00" view (latest row per unique key) the first
    time this table is seen, using the key columns from MNS120MI.Get."""
    view = f"{table}00"
    if etl_view_exists(conn, view):
        return
    try:
        keys = etl_get_table_keys(cfg, token, table)
    except Exception as e:
        etl_log(f"  Could not look up keys for {table}: {e}")
        return
    if not keys:
        etl_log(f"  No unique keys found for {table}; skipping {view} view.")
        return

    # SQLite silently treats an unrecognized quoted identifier as a string
    # literal instead of erroring, which would collapse PARTITION BY into a
    # single constant partition (one row for the whole table) -- so refuse
    # to build the view unless every key is an actual column.
    existing = etl_table_columns(conn, table)
    missing = [k for k in keys if k not in existing]
    if missing:
        etl_log(f"  Keys {missing} not found in {table}'s columns; skipping {view} view.")
        return

    partition = ", ".join(f'"{k}"' for k in keys)
    cols = ", ".join(f'"{c}"' for c in existing)
    conn.execute(f'''
        CREATE VIEW IF NOT EXISTS "{view}" AS
        SELECT {cols} FROM (
            SELECT *,
                   ROW_NUMBER() OVER (
                       PARTITION BY {partition}
                       ORDER BY variationNumber DESC
                   ) AS rn
            FROM "{table}"
        )
        WHERE rn = 1
    ''')
    conn.commit()
    etl_log(f"  Created view {view} (keys: {', '.join(keys)})")


def etl_load_records(conn, table, records, cfg, token):
    if not records:
        return 0
    table = "".join(c for c in table if c.isalnum() or c == "_")

    existing = etl_table_columns(conn, table)
    if not existing:  # create table from the columns of the first record
        cols = ", ".join(f'"{k}" {etl_sql_type(v)}' for k, v in records[0].items())
        conn.execute(f'CREATE TABLE "{table}" ({cols})')
        existing = etl_table_columns(conn, table)

    for rec in records:  # add any columns not yet in the table
        for k, v in rec.items():
            if k not in existing:
                conn.execute(f'ALTER TABLE "{table}" ADD COLUMN "{k}" {etl_sql_type(v)}')
                existing.append(k)

    etl_ensure_variation_unique(conn, table, existing)

    inserted = 0
    for rec in records:
        keys = list(rec.keys())
        sql = (f'INSERT OR IGNORE INTO "{table}" ({", ".join(f_quote(keys))}) '
               f'VALUES ({", ".join("?" * len(keys))})')
        cur = conn.execute(sql, [etl_sql_value(rec[k]) for k in keys])
        inserted += cur.rowcount
    conn.commit()

    etl_ensure_latest_view(conn, table, cfg, token)
    return inserted


def etl_force_load_records(conn, table, records, cfg, token):
    """Load records from a user-forced Object ID. Unlike the normal cycle (which
    appends a new version row per change), a forced object UPDATES the existing
    records in place: for each incoming row the current row(s) for the same
    business key are replaced, so re-forcing an object never leaves duplicate
    rows behind. Business keys come from MNS120MI -- the same keys the {table}00
    latest-row view uses. If the keys can't be resolved, falls back to the
    normal append load."""
    if not records:
        return 0
    table = "".join(c for c in table if c.isalnum() or c == "_")

    existing = etl_table_columns(conn, table)
    if not existing:  # create table from the columns of the first record
        cols = ", ".join(f'"{k}" {etl_sql_type(v)}' for k, v in records[0].items())
        conn.execute(f'CREATE TABLE "{table}" ({cols})')
        existing = etl_table_columns(conn, table)

    for rec in records:  # add any columns not yet in the table
        for k, v in rec.items():
            if k not in existing:
                conn.execute(f'ALTER TABLE "{table}" ADD COLUMN "{k}" {etl_sql_type(v)}')
                existing.append(k)

    try:
        keys = etl_get_table_keys(cfg, token, table)
    except Exception as e:
        etl_log(f"  Could not look up keys for {table}: {e}")
        keys = []
    keys = [k for k in keys if k in existing]
    if not keys:
        etl_log(f"  No key columns for {table}; appending forced rows instead of updating.")
        return etl_load_records(conn, table, records, cfg, token)

    loaded = 0
    for rec in records:
        cols = list(rec.keys())
        # Replace any existing row(s) for this business key, then insert the
        # fresh one -- i.e. update in place rather than adding a new version row.
        where = " AND ".join(f'"{k}" = ?' for k in keys)
        conn.execute(f'DELETE FROM "{table}" WHERE {where}',
                     [etl_sql_value(rec.get(k)) for k in keys])
        sql = (f'INSERT INTO "{table}" ({", ".join(f_quote(cols))}) '
               f'VALUES ({", ".join("?" * len(cols))})')
        conn.execute(sql, [etl_sql_value(rec[k]) for k in cols])
        loaded += 1
    conn.commit()

    etl_ensure_latest_view(conn, table, cfg, token)
    etl_log(f"  {table}: updated {loaded} record(s) in place (forced load, keys: {', '.join(keys)})")
    return loaded


def f_quote(keys):
    return [f'"{k}"' for k in keys]


# ---------------------------------------------- initial load from a dropped zip
#
# The status page lets the user drop a zip of M3 Grid Access binary table
# exports (one file per table, plus a TABLE_INFO companion) while the routine is
# stopped. Each table is decoded with m3_unpacker, given the four trailing audit
# fields the Data Lake feed carries, then loaded as a fresh baseline: the target
# table is truncated first (the zip is the most accurate current version) and
# every row is inserted with variationNumber 0. The Data Lake loop then keeps
# tracking changes on top of that baseline going forward.

# States in which no cycle is running and no load is in flight, so a zip may be
# loaded. Everything else means the routine is busy.
_UPLOAD_OK_STATES = {"waiting-for-selection", "stopped", "error", "idle"}


def etl_upload_allowed():
    """True only while the routine is inactive (not cycling, not sleeping, not
    already loading a zip). Uploads are refused otherwise so a bulk load never
    races the tracking loop."""
    if _RUN_EVENT.is_set() or _NOW_EVENT.is_set() or _LOADING_EVENT.is_set():
        return False
    status, _ = etl_status_snapshot()
    return status["state"] in _UPLOAD_OK_STATES


def etl_parse_m3_binary(data, table_info_counts, table_name):
    """Decode one M3 binary table export (bytes) into (field_names, records)
    using m3_unpacker. fill=True so inherited values and numeric defaults are
    materialised — i.e. the full current snapshot of every row."""
    fields, types, widths, scales, start = m3.parse_field_defs(data)
    data_section = data[start:]
    expected = (table_info_counts or {}).get(table_name)
    if expected == 0:
        return fields, []
    if m3.detect_format(data_section) == "ffe0":
        records = m3.parse_records_ffe0_style(data_section, len(fields))
    else:
        records = m3.parse_records_bitmap_style(
            data_section, len(fields), field_types=types, field_widths=widths,
            field_scales=scales, expected_count=expected, fill=True,
        )
        if expected and len(records) > expected:
            records = records[:expected]
    return fields, records


def etl_build_accounting_entity(cono, divi):
    """CONO (a 3-digit company) rendered as a zero-padded alpha string, e.g.
    1 -> "001", joined to DIVI with "_" when a division is present, else CONO
    alone."""
    cono = str(cono or "").strip()
    cono3 = cono.zfill(3) if cono else ""
    divi = str(divi or "").strip()
    return f"{cono3}_{divi}" if divi else cono3


def etl_load_table_full(conn, table, raw_fields, records):
    """Truncate `table` and fully load it from a parsed M3 export.

    The binary field names carry M3's 2-char file mnemonic (e.g. IDCONO on
    CIDMAS); the Data Lake feed drops that prefix (CONO), so strip it here too
    so the baseline and the ongoing feed share one set of column names. Each
    row then gets the four trailing audit fields:
        accountingEntity  CONO (3-digit alpha) [+ "_" + DIVI if present]
        variationNumber   0    (baseline version)
        timestamp         current UTC time, e.g. 2026-05-04T07:48:30.015Z
        deleted           0
    """
    table = "".join(c for c in table if c.isalnum() or c == "_")
    if not records:
        etl_log(f"  {table}: 0 records in zip; skipped")
        return 0

    cols = [f[2:] if len(f) > 2 else f for f in raw_fields]
    has_divi = "DIVI" in cols
    ts = etl_utcnow_iso()

    dict_records = []
    for row in records:
        rec = dict(zip(cols, row))
        rec["accountingEntity"] = etl_build_accounting_entity(
            rec.get("CONO"), rec.get("DIVI") if has_divi else "")
        rec["variationNumber"] = 0
        rec["timestamp"] = ts
        rec["deleted"] = 0
        dict_records.append(rec)

    all_cols = list(dict_records[0].keys())

    existing = etl_table_columns(conn, table)
    if not existing:  # first time we see this table -> create it
        coldefs = ", ".join(
            f'"{k}" {etl_sql_type(dict_records[0][k])}' for k in all_cols)
        conn.execute(f'CREATE TABLE "{table}" ({coldefs})')
        existing = etl_table_columns(conn, table)
    else:
        for k in all_cols:  # add any columns the existing table is missing
            if k not in existing:
                conn.execute(
                    f'ALTER TABLE "{table}" ADD COLUMN "{k}" '
                    f'{etl_sql_type(dict_records[0][k])}')
                existing.append(k)

    # A baseline sets every row's variationNumber to 0, which collides with the
    # one-row-per-variationNumber unique index the tracking loop may have built.
    # Drop it before loading; etl_load_records re-attempts it on the next cycle
    # (and simply logs if the zeros keep it from being recreated).
    uq = f"{table}_variationNumber_uq"
    if etl_index_exists(conn, uq):
        conn.execute(f'DROP INDEX "{uq}"')

    conn.execute(f'DELETE FROM "{table}"')  # truncate: zip is the source of truth

    inserted = 0
    for rec in dict_records:
        keys = list(rec.keys())
        sql = (f'INSERT INTO "{table}" ({", ".join(f_quote(keys))}) '
               f'VALUES ({", ".join("?" * len(keys))})')
        conn.execute(sql, [etl_sql_value(rec[k]) for k in keys])
        inserted += 1
    conn.commit()
    etl_log(f"  {table}: truncated and loaded {inserted} row(s)")
    return inserted


def etl_load_zip(zip_bytes):
    """Load every table in a dropped zip as a fresh baseline. Returns a list of
    (table_name, rows_loaded). Runs on its own SQLite connection (the web
    handler's thread) — safe because uploads are only accepted while the
    tracking loop is idle."""
    tmp = tempfile.mkdtemp(prefix="etl_zip_")
    try:
        with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
            zf.extractall(tmp)

        ti_counts, table_paths = {}, []
        for root, _, files in os.walk(tmp):
            for fn in files:
                path = os.path.join(root, fn)
                if fn == "TABLE_INFO":
                    ti_counts = m3.parse_table_info(path)
                elif not fn.startswith(".") and not fn.endswith(".csv"):
                    table_paths.append(path)

        if not table_paths:
            raise ValueError("no table files found in zip")
        if ti_counts:
            etl_log(f"TABLE_INFO: {ti_counts}")

        conn = sqlite3.connect(DB_PATH)
        results = []
        try:
            for path in sorted(table_paths):
                table = os.path.basename(path)
                with open(path, "rb") as fh:
                    data = fh.read()
                try:
                    fields, records = etl_parse_m3_binary(data, ti_counts, table)
                    n = etl_load_table_full(conn, table, fields, records)
                    results.append((table, n))
                except Exception as e:
                    etl_log(f"  {table}: FAILED ({e})")
                    results.append((table, -1))
        finally:
            conn.close()
        return results
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


# --------------------------------------------------------- step 7: status page
_STATUS_PAGE_HTML = """<!doctype html>
<html>
<head>
<meta charset="utf-8">
<title>ETL Data Lake Status</title>
<style>
  body { font-family: -apple-system, Helvetica, Arial, sans-serif; margin: 2rem; background: #111; color: #eee; }
  h1 { font-size: 1.2rem; }
  table { border-collapse: collapse; margin-bottom: 1.5rem; }
  td { padding: 2px 10px 2px 0; vertical-align: top; }
  td.label { color: #888; }
  .state { display: inline-block; padding: 2px 10px; border-radius: 4px; font-weight: bold; }
  .state-idle { background: #274; }
  .state-sleeping { background: #245; }
  .state-error { background: #722; }
  .state-processing, .state-authenticating, .state-pinging,
  .state-listing-objects, .state-checking-version, .state-starting,
  .state-waiting-for-selection, .state-stopping, .state-loading-zip { background: #552; }
  #log { background: #000; color: #9f9; font-family: monospace; font-size: 0.85rem;
         padding: 1rem; height: 400px; overflow-y: auto; white-space: pre-wrap; }
  #controls { margin-bottom: 1.5rem; display: flex; align-items: center; gap: 8px; flex-wrap: wrap; }
  #controls select, #controls input, #controls button { font-size: 1rem; padding: 4px 8px; }
  #controls input[type=number] { width: 6rem; }
  #controls button:disabled { opacity: 0.4; }
  #upload { margin-bottom: 1.5rem; }
  #dropArea { border: 2px dashed #555; border-radius: 6px; padding: 1.2rem;
              text-align: center; color: #aaa; background: #181818; cursor: pointer; }
  #dropArea.dragover { border-color: #9f9; color: #9f9; background: #1c221c; }
  #dropArea.disabled { opacity: 0.4; cursor: not-allowed; }
  #dropArea input[type=file] { display: none; }
  #uploadRow { margin-top: 8px; display: flex; align-items: center; gap: 10px; flex-wrap: wrap; }
  #uploadRow button { font-size: 1rem; padding: 4px 8px; }
  #uploadRow button:disabled { opacity: 0.4; }
  #zipName { color: #9cf; }
  #uploadMsg { color: #cc9; }
  #tabBar { display: flex; gap: 4px; margin-bottom: 1rem; border-bottom: 1px solid #333; }
  .tabBtn { font-size: 1rem; padding: 6px 14px; background: #1a1a1a; color: #aaa;
            border: 1px solid #333; border-bottom: none; border-radius: 6px 6px 0 0;
            cursor: pointer; }
  .tabBtn.active { background: #222; color: #eee; }
  .tabPanel { display: none; }
  .tabPanel.active { display: block; }
</style>
</head>
<body>
<h1>ETL Data Lake &rarr; SQLite</h1>
<div id="tabBar">
  <button class="tabBtn active" data-tab="main">Main</button>
  <button class="tabBtn" data-tab="loaddb">Load DB</button>
</div>
<div id="tab-main" class="tabPanel active">
  <div id="controls">
    <label>Environment <select id="envSelect"></select></label>
    <label>Poll interval (s) <input type="number" id="intervalInput" min="5" step="1"></label>
    <button id="applyIntervalBtn">Apply interval</button>
    <label>Since override (UTC) <input type="datetime-local" id="sinceInput" step="1"></label>
    <button id="applySinceBtn">Apply since</button>
    <button id="clearSinceBtn">Clear since</button>
    <button id="startBtn">Start</button>
    <button id="nowBtn">Now</button>
    <button id="stopBtn">Stop</button>
  </div>
  <table id="fields"></table>
  <div id="log"></div>
</div>
<div id="tab-loaddb" class="tabPanel">
  <div id="controls">
    <label>Force object ID <input type="text" id="objectIdInput" placeholder="dl_id" size="28"></label>
    <button id="loadObjectBtn">Load object</button>
  </div>
  <div id="upload">
    <div id="dropArea">
      Drop a table export <b>.zip</b> here, or <u>click to browse</u>
      <input type="file" id="zipInput" accept=".zip,application/zip">
    </div>
    <div id="uploadRow">
      <button id="uploadBtn" disabled>Load zip into database</button>
      <span id="zipName"></span>
      <span id="uploadMsg"></span>
    </div>
    <div style="color:#777;font-size:0.85rem;margin-top:4px;">
      Only available while the routine is stopped. Loading truncates each table in
      the zip and reloads it as the current baseline; tracking resumes on Start.
    </div>
  </div>
</div>
<script>
document.querySelectorAll(".tabBtn").forEach(btn => {
  btn.onclick = () => {
    document.querySelectorAll(".tabBtn").forEach(b => b.classList.remove("active"));
    document.querySelectorAll(".tabPanel").forEach(p => p.classList.remove("active"));
    btn.classList.add("active");
    document.getElementById(`tab-${btn.dataset.tab}`).classList.add("active");
  };
});
function fmt(v) { return v === null || v === undefined ? "-" : v; }

const RUNNING_STATES = ["starting", "authenticating", "pinging", "checking-version",
                         "listing-objects", "processing", "sleeping", "stopping",
                         "loading-zip"];

let controlsPopulated = false;
let selectedZip = null;
let uploading = false;

function populateControls(s) {
  const select = document.getElementById("envSelect");
  const envs = s.available_environments || [];
  select.innerHTML = envs.map(e =>
    `<option value="${e}" ${e === (s.environment || s.default_environment) ? "selected" : ""}>${e}</option>`
  ).join("");
  document.getElementById("intervalInput").value = s.poll_interval_seconds || 60;
}

async function postForm(path, body) {
  await fetch(path, {
    method: "POST",
    headers: {"Content-Type": "application/x-www-form-urlencoded"},
    body: body
  });
  refresh();
}

document.getElementById("startBtn").onclick = () => {
  const env = document.getElementById("envSelect").value;
  const interval = document.getElementById("intervalInput").value;
  postForm("/start", `environment=${encodeURIComponent(env)}&interval=${encodeURIComponent(interval)}`);
};
document.getElementById("nowBtn").onclick = () => {
  const env = document.getElementById("envSelect").value;
  postForm("/now", `environment=${encodeURIComponent(env)}`);
};
document.getElementById("stopBtn").onclick = () => postForm("/stop", "");
document.getElementById("loadObjectBtn").onclick = () => {
  const obj = document.getElementById("objectIdInput").value.trim();
  if (!obj) return;
  postForm("/loadobject", `object_id=${encodeURIComponent(obj)}`);
  document.getElementById("objectIdInput").value = "";
};
document.getElementById("applyIntervalBtn").onclick = () => {
  const interval = document.getElementById("intervalInput").value;
  postForm("/interval", `interval=${encodeURIComponent(interval)}`);
};
document.getElementById("applySinceBtn").onclick = () => {
  const since = document.getElementById("sinceInput").value;
  postForm("/since", `since=${encodeURIComponent(since)}`);
};
document.getElementById("clearSinceBtn").onclick = () => {
  document.getElementById("sinceInput").value = "";
  postForm("/since", "since=");
};

// ---- zip upload (drag/drop or browse) ----
const dropArea = document.getElementById("dropArea");
const zipInput = document.getElementById("zipInput");

function setSelectedZip(file) {
  if (dropArea.classList.contains("disabled")) return;
  selectedZip = file || null;
  document.getElementById("zipName").textContent = selectedZip ? selectedZip.name : "";
  document.getElementById("uploadMsg").textContent = "";
  refresh();
}

dropArea.onclick = () => { if (!dropArea.classList.contains("disabled")) zipInput.click(); };
zipInput.onchange = () => setSelectedZip(zipInput.files[0]);
dropArea.addEventListener("dragover", e => {
  e.preventDefault();
  if (!dropArea.classList.contains("disabled")) dropArea.classList.add("dragover");
});
dropArea.addEventListener("dragleave", () => dropArea.classList.remove("dragover"));
dropArea.addEventListener("drop", e => {
  e.preventDefault();
  dropArea.classList.remove("dragover");
  if (dropArea.classList.contains("disabled")) return;
  if (e.dataTransfer.files && e.dataTransfer.files.length) setSelectedZip(e.dataTransfer.files[0]);
});

document.getElementById("uploadBtn").onclick = async () => {
  if (!selectedZip || uploading) return;
  uploading = true;
  document.getElementById("uploadMsg").textContent = "Loading… this can take a while for large tables.";
  refresh();
  try {
    const buf = await selectedZip.arrayBuffer();
    const res = await fetch("/upload", {
      method: "POST",
      headers: {"Content-Type": "application/zip", "X-Filename": selectedZip.name},
      body: buf,
    });
    const data = await res.json();
    if (data.ok) {
      const parts = data.results.map(r => `${r.table}: ${r.rows < 0 ? "FAILED" : r.rows + " rows"}`);
      document.getElementById("uploadMsg").textContent = "Loaded — " + parts.join(", ");
      selectedZip = null;
      zipInput.value = "";
      document.getElementById("zipName").textContent = "";
    } else {
      document.getElementById("uploadMsg").textContent = "Error: " + (data.error || "upload failed");
    }
  } catch (err) {
    document.getElementById("uploadMsg").textContent = "Error: " + err;
  } finally {
    uploading = false;
    refresh();
  }
};

function updateDashboard(s, log) {
  const running = RUNNING_STATES.includes(s.state);
  document.getElementById("envSelect").disabled = running;
  document.getElementById("startBtn").disabled = running || uploading;
  document.getElementById("stopBtn").disabled = !running;

  // Force-load a single object is only available while the routine is running
  // (continuous mode), and never during a zip load.
  const etlRunning = running && s.state !== "loading-zip";
  document.getElementById("objectIdInput").disabled = !etlRunning;
  document.getElementById("loadObjectBtn").disabled = !etlRunning;

  // Uploads are only allowed while the routine is inactive.
  const canUpload = !running && !uploading;
  dropArea.classList.toggle("disabled", !canUpload);
  document.getElementById("uploadBtn").disabled = !canUpload || !selectedZip;

  const rows = [
    ["Environment", fmt(s.environment)],
    ["State", `<span class="state state-${(s.state||"").replace(/\\s+/g,"-")}">${fmt(s.state)}</span>`],
    ["Build", fmt(s.build)],
    ["Poll interval (s)", fmt(s.poll_interval_seconds)],
    ["Cycle #", fmt(s.cycle_count)],
    ["Cycle started", fmt(s.cycle_start)],
    ["Last cycle ended", fmt(s.last_cycle_end)],
    ["Next run at", fmt(s.next_run_at)],
    ["Next run since", s.since_override ? `${fmt(s.since_override)} (override)` : "automatic"],
    ["Objects", `${fmt(s.objects_done)} / ${fmt(s.objects_total)}`],
    ["Rows (last cycle)", fmt(s.rows_loaded_last_cycle)],
    ["Rows (total)", fmt(s.rows_loaded_total)],
    ["Last error", fmt(s.last_error)],
  ];
  document.getElementById("fields").innerHTML = rows.map(
    ([label, value]) => `<tr><td class="label">${label}</td><td>${value}</td></tr>`
  ).join("");

  const logEl = document.getElementById("log");
  const atBottom = logEl.scrollTop + logEl.clientHeight >= logEl.scrollHeight - 20;
  logEl.textContent = log.join("\\n");
  if (atBottom) logEl.scrollTop = logEl.scrollHeight;
}

async function refresh() {
  const res = await fetch("/status.json");
  const data = await res.json();
  const s = data.status;

  if (!controlsPopulated) {
    populateControls(s);
    controlsPopulated = true;
  }
  updateDashboard(s, data.log);
}

refresh();
setInterval(refresh, 2000);
</script>
</body>
</html>
"""


class _StatusHandler(BaseHTTPRequestHandler):
    def log_message(self, fmt, *args):
        pass  # keep the console clean; etl_log() carries the real status output

    def _respond(self, status, payload, content_type):
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(payload)))
        self.end_headers()
        self.wfile.write(payload)

    def do_GET(self):
        if self.path.startswith("/status.json"):
            status, log = etl_status_snapshot()
            payload = json.dumps({"status": status, "log": log}).encode()
            self._respond(200, payload, "application/json")
        else:
            self._respond(200, _STATUS_PAGE_HTML.encode(), "text/html; charset=utf-8")

    def _read_params(self):
        length = int(self.headers.get("Content-Length", 0))
        return urllib.parse.parse_qs(self.rfile.read(length).decode())

    def do_POST(self):
        if self.path.startswith("/start"):
            self._handle_start()
        elif self.path.startswith("/stop"):
            self._handle_stop()
        elif self.path.startswith("/now"):
            self._handle_now()
        elif self.path.startswith("/interval"):
            self._handle_interval()
        elif self.path.startswith("/since"):
            self._handle_since()
        elif self.path.startswith("/upload"):
            self._handle_upload()
        elif self.path.startswith("/loadobject"):
            self._handle_load_object()
        else:
            self._respond(404, b"not found", "text/plain")

    def _handle_start(self):
        params = self._read_params()
        name = (params.get("environment") or [""])[0]
        raw_interval = (params.get("interval") or [""])[0]

        if _LOADING_EVENT.is_set():
            self._respond(400, b'{"ok": false, "error": "loading a zip"}', "application/json")
            return
        status, _ = etl_status_snapshot()
        if status["state"] not in ("waiting-for-selection", "stopped", "error"):
            self._respond(400, b'{"ok": false, "error": "already running"}', "application/json")
            return
        if name not in etl_list_environments():
            self._respond(400, b'{"ok": false, "error": "unknown environment"}', "application/json")
            return
        try:
            interval_seconds = max(MIN_POLL_INTERVAL_SECONDS, int(raw_interval))
        except ValueError:
            interval_seconds = status["poll_interval_seconds"]

        etl_start_run(name, interval_seconds)
        self._respond(200, b'{"ok": true}', "application/json")

    def _handle_stop(self):
        etl_stop_run()
        self._respond(200, b'{"ok": true}', "application/json")

    def _handle_now(self):
        params = self._read_params()
        name = (params.get("environment") or [""])[0]
        if name not in etl_list_environments():
            self._respond(400, b'{"ok": false, "error": "unknown environment"}', "application/json")
            return
        etl_trigger_now(name)
        self._respond(200, b'{"ok": true}', "application/json")

    def _handle_interval(self):
        params = self._read_params()
        raw_interval = (params.get("interval") or [""])[0]
        try:
            interval_seconds = max(MIN_POLL_INTERVAL_SECONDS, int(raw_interval))
        except ValueError:
            self._respond(400, b'{"ok": false, "error": "invalid interval"}', "application/json")
            return
        etl_update_interval(interval_seconds)
        self._respond(200, b'{"ok": true}', "application/json")

    def _handle_since(self):
        params = self._read_params()
        raw = (params.get("since") or [""])[0].strip()
        if not raw:
            etl_set_since_override(None)
            self._respond(200, b'{"ok": true}', "application/json")
            return
        try:
            since = etl_normalize_since(raw)
        except ValueError:
            self._respond(400, b'{"ok": false, "error": "invalid date"}', "application/json")
            return
        etl_set_since_override(since)
        self._respond(200, b'{"ok": true}', "application/json")

    def _handle_load_object(self):
        """Queue a pasted object id (dl_id) for immediate load. Only accepted
        while the routine is running (continuous mode)."""
        if not _RUN_EVENT.is_set():
            self._respond(409, b'{"ok": false, "error": "routine is not running; '
                          b'start it before forcing an object load"}', "application/json")
            return
        params = self._read_params()
        dl_id = (params.get("object_id") or [""])[0].strip()
        if not dl_id:
            self._respond(400, b'{"ok": false, "error": "missing object id"}', "application/json")
            return
        etl_queue_force_object(dl_id)
        self._respond(200, b'{"ok": true}', "application/json")

    def _handle_upload(self):
        """Accept a dropped zip (raw bytes as the POST body) and load each table
        as a fresh baseline. Refused unless the routine is inactive."""
        if not etl_upload_allowed():
            self._respond(409, b'{"ok": false, "error": "routine is active; '
                          b'stop it before loading a zip"}', "application/json")
            return
        length = int(self.headers.get("Content-Length", 0))
        if length <= 0:
            self._respond(400, b'{"ok": false, "error": "empty upload"}', "application/json")
            return

        body = self.rfile.read(length)
        prev_state, _ = etl_status_snapshot()
        prev_state = prev_state["state"]
        _LOADING_EVENT.set()
        etl_set_status(state="loading-zip", last_error=None)
        fname = self.headers.get("X-Filename", "upload.zip")
        etl_log(f"Loading dropped zip: {fname} ({length} bytes)")
        try:
            results = etl_load_zip(body)
            payload = json.dumps({
                "ok": True,
                "results": [{"table": t, "rows": n} for t, n in results],
            }).encode()
            loaded = sum(n for _, n in results if n >= 0)
            etl_log(f"Zip load complete: {len(results)} table(s), {loaded} row(s) total.")
            self._respond(200, payload, "application/json")
        except Exception as e:
            etl_log(f"Zip load failed: {e}")
            payload = json.dumps({"ok": False, "error": str(e)}).encode()
            self._respond(400, payload, "application/json")
        finally:
            _LOADING_EVENT.clear()
            etl_set_status(state=prev_state)


def etl_start_web_server():
    server = ThreadingHTTPServer((WEB_HOST, WEB_PORT), _StatusHandler)
    threading.Thread(target=server.serve_forever, daemon=True).start()
    etl_log(f"Status page: http://{WEB_HOST}:{WEB_PORT}/")
    return server


# -------------------------------------------------------------- one cycle
def etl_run_cycle(cfg, conn):
    global _SINCE_OVERRIDE
    cycle_start = etl_utcnow_iso()
    etl_set_status(state="authenticating", cycle_start=cycle_start, last_error=None)

    token = etl_authenticate(cfg)          # step 1 (auth)
    base = etl_base_url(cfg)

    etl_set_status(state="pinging")
    if not etl_ping(base, token):          # step 2
        etl_log("Server is not active. Stopping the routine.")
        raise EtlPingFailed("Server is not active")

    etl_set_status(state="checking-version")
    build = etl_show_version(base, token)  # step 3
    etl_set_status(build=build)

    since = etl_get_last_run(conn)
    with _STATE_LOCK:
        override, _SINCE_OVERRIDE = _SINCE_OVERRIDE, None
    if override:
        since = override
    etl_set_status(since_override=None)
    etl_log(f"Picking up from: {since}" + (" (manual override)" if override else ""))

    etl_set_status(state="listing-objects")
    objects = etl_list_dataobjects(base, token, since)   # step 4
    etl_log(f"Data objects to process: {len(objects)}")
    etl_set_status(objects_total=len(objects), objects_done=0)

    total = 0
    etl_set_status(state="processing")
    for obj in objects:                                  # steps 5 & 6
        if _ABORT_EVENT.is_set():
            etl_log("Stop requested; halting current cycle early.")
            break
        dl_id = obj["dl_id"]
        table = obj["dl_document_name"]
        try:
            recs = etl_download_details(base, token, dl_id)
            n = etl_load_records(conn, table, recs, cfg, token)
            total += n
            etl_log(f"  {dl_id} -> {table}: {n} rows")
        except Exception as e:
            etl_log(f"  {dl_id} -> {table}: FAILED ({e})")
        finally:
            with _STATE_LOCK:
                STATUS["objects_done"] += 1

    etl_save_last_run(conn, cycle_start)
    with _STATE_LOCK:
        STATUS["rows_loaded_last_cycle"] = total
        STATUS["rows_loaded_total"] += total
        STATUS["cycle_count"] += 1
        STATUS["last_cycle_end"] = etl_utcnow_iso()
        STATUS["state"] = "idle"
    etl_log(f"Cycle complete. {total} rows loaded. LastRun set to {cycle_start}")


# --------------------------------------------------- force-load a single object
def etl_queue_force_object(dl_id):
    """Called from the web handler; queues a dl_id for the run loop to load out
    of band. Only meaningful while the routine is running."""
    with _STATE_LOCK:
        _FORCE_QUEUE.append(dl_id)
    _FORCE_EVENT.set()
    etl_log(f"Force load requested for object {dl_id}")


def etl_process_forced_objects(cfg, conn):
    """Drain the force queue: for each dl_id, authenticate, resolve its table
    from the object metadata, download it and load it via etl_force_load_records
    -- which UPDATES the existing records in place (replacing each row by its
    business key) rather than appending new version rows. Runs on the run loop's
    thread/connection so it never races the tracking writes."""
    while True:
        with _STATE_LOCK:
            if not _FORCE_QUEUE:
                _FORCE_EVENT.clear()
                return
            dl_id = _FORCE_QUEUE.popleft()

        etl_set_status(state="processing")
        try:
            token = etl_authenticate(cfg)
            base = etl_base_url(cfg)
            meta = etl_get_object_meta(base, token, dl_id)
            if not meta:
                etl_log(f"Force load: object {dl_id} not found.")
                etl_set_status(last_error=f"object {dl_id} not found")
                continue
            table = meta.get("dl_document_name")
            if not table:
                etl_log(f"Force load: object {dl_id} has no document name.")
                etl_set_status(last_error=f"object {dl_id} has no table name")
                continue
            recs = etl_download_details(base, token, dl_id)
            n = etl_force_load_records(conn, table, recs, cfg, token)
            with _STATE_LOCK:
                STATUS["rows_loaded_total"] += n
            etl_log(f"Force load: {dl_id} -> {table}: {n} rows")
        except Exception as e:
            etl_log(f"Force load {dl_id} FAILED: {e}")
            etl_set_status(last_error=str(e))


def etl_run_loop(conn):
    """Waits for Start or Now (via the webpage), runs cycles until Stop is
    pressed or a ping failure pauses the routine, then waits again.

    Now runs a single cycle immediately: if continuous mode is off it goes
    straight back to "stopped" afterward instead of sleeping/looping.
    """
    while True:
        while not (_RUN_EVENT.is_set() or _NOW_EVENT.is_set()):
            time.sleep(0.3)

        one_shot = _NOW_EVENT.is_set() and not _RUN_EVENT.is_set()
        _NOW_EVENT.clear()
        _ABORT_EVENT.clear()

        with _STATE_LOCK:
            cfg = _SELECTED_CFG
        etl_log(f"\n===== ETL cycle started {etl_utcnow_iso()} =====")
        try:
            etl_run_cycle(cfg, conn)
        except EtlPingFailed as e:
            etl_log(f"Cycle error: {e}")
            etl_set_status(state="error", last_error=str(e))
            _RUN_EVENT.clear()
            continue
        except Exception as e:
            etl_log(f"Cycle error: {e}")
            etl_set_status(state="error", last_error=str(e))

        if one_shot or not _RUN_EVENT.is_set():
            etl_set_status(state="stopped")
            continue

        interval = etl_status_snapshot()[0]["poll_interval_seconds"]
        etl_set_status(state="sleeping", next_run_at=etl_iso_plus_seconds(interval))
        etl_log(f"Sleeping up to {interval} seconds...")
        slept = 0
        while slept < interval and _RUN_EVENT.is_set() and not _NOW_EVENT.is_set():
            if _FORCE_EVENT.is_set():   # user pasted an object id to force-load
                etl_process_forced_objects(cfg, conn)
                etl_set_status(state="sleeping")
            time.sleep(1)
            slept += 1
            interval = etl_status_snapshot()[0]["poll_interval_seconds"]  # pick up live changes

        if not _RUN_EVENT.is_set():
            etl_set_status(state="stopped")


def main():
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
    envs = etl_list_environments()
    etl_set_status(available_environments=envs, default_environment=DEFAULT_ENV)
    etl_start_web_server()
    etl_log(f"Status page: http://{WEB_HOST}:{WEB_PORT}/ -- use Start/Stop there to control the routine.")

    conn = sqlite3.connect(DB_PATH)
    try:
        etl_run_loop(conn)
    except KeyboardInterrupt:
        etl_log("\nStopped by user.")
    finally:
        conn.close()


if __name__ == "__main__":
    main()
