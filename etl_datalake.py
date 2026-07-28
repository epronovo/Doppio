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
import json
import os
import sqlite3
import sys
import threading
import time
import urllib.parse
import urllib.request
import zlib
from collections import deque
from datetime import datetime, timedelta, timezone
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

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


def f_quote(keys):
    return [f'"{k}"' for k in keys]


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
  .state-waiting-for-selection, .state-stopping { background: #552; }
  #log { background: #000; color: #9f9; font-family: monospace; font-size: 0.85rem;
         padding: 1rem; height: 400px; overflow-y: auto; white-space: pre-wrap; }
  #controls { margin-bottom: 1.5rem; display: flex; align-items: center; gap: 8px; flex-wrap: wrap; }
  #controls select, #controls input, #controls button { font-size: 1rem; padding: 4px 8px; }
  #controls input[type=number] { width: 6rem; }
  #controls button:disabled { opacity: 0.4; }
</style>
</head>
<body>
<h1>ETL Data Lake &rarr; SQLite</h1>
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
<script>
function fmt(v) { return v === null || v === undefined ? "-" : v; }

const RUNNING_STATES = ["starting", "authenticating", "pinging", "checking-version",
                         "listing-objects", "processing", "sleeping", "stopping"];

let controlsPopulated = false;

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

function updateDashboard(s, log) {
  const running = RUNNING_STATES.includes(s.state);
  document.getElementById("envSelect").disabled = running;
  document.getElementById("startBtn").disabled = running;
  document.getElementById("stopBtn").disabled = !running;

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
        else:
            self._respond(404, b"not found", "text/plain")

    def _handle_start(self):
        params = self._read_params()
        name = (params.get("environment") or [""])[0]
        raw_interval = (params.get("interval") or [""])[0]

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
