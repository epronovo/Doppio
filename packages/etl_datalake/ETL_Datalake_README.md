# etl_datalake — Infor Data Lake → SQLite

A small polling routine that pulls data objects out of the Infor Data Lake and
loads them into a local SQLite database, with a status webpage to drive and
watch it — Start / Stop / Now, a live-editable poll interval, a "since"
override, forcing a single object to reload, and dropping a table-export
`.zip` in as a fresh baseline.

Unlike `packages/m3_security` and `packages/adp_concur`, this one uses **only
the Python standard library** — no Flask, no `requests`. The status page is
served by a plain `http.server.ThreadingHTTPServer`; that is deliberate, not
an oversight, so keep it dependency-free if you touch it.

| File | Role |
|------|------|
| `etl_datalake.py` | Everything — polling loop, M3 Data Lake calls, SQLite loads, the status webpage |
| `templates/ETL_Index.html` | The status page (read once at import time via a plain file read, not a template engine) |

## Quick start

```bash
python etl_datalake.py                 # http://127.0.0.1:8787/
```

No `pip install` needed — the whole thing runs on the standard library plus
`m3_unpacker.py` (see below).

## Two things this package deliberately does NOT own

- **`ionapi/`** — the shared repo-root folder of `.ionapi` tenant-credential
  files that every Infor automation tool in this repo reads (M3 Security,
  ADP↔Concur, MIG sync, and this routine). `IONAPI_DIR` is computed as two
  directories up from this file, not a package-local copy.
- **`m3_unpacker.py`** — stays at the repo root. It parses the M3 Grid Access
  binary table format used both by the Data Lake `.zip` exports this routine
  can load and by `m3_repacker.py`, `m3_unpacker_db.py`, `m3_repacker_db.py`,
  `build_brazil_csv.py` and `load_brazil_db.py` — five other root scripts, so
  it is not specific to this package and was left where it is. `sys.path` is
  adjusted at import time to find it there.

`DB_PATH` (`~/sqlite/etl.db`) is its own database, separate from the shared
`~/sqlite/doppio.db` the other three packages use — this routine's tables are
raw Data Lake dumps, not a capture/edit workflow, so there was never a reason
to share the file.

## What it does

1. Lists the `.ionapi` files in `ionapi/`; pick an environment and a poll
   interval on the status page, then click **Start**.
2. Pings the Data Lake server; the routine stops itself (until restarted from
   the page) if the ping fails.
3. Shows the current Data Lake build number once the ping succeeds.
4. Lists the data objects since the last run (or since the "since" override,
   for one cycle only) via the dataobjects list API.
5. Downloads each object's details.
6. Loads it: table name = `dl_document_name`, columns = whatever the download
   carries. The table is created if it doesn't exist yet.
7. Repeats every poll interval while running; **Now** runs one cycle
   immediately without turning on continuous mode.

**Force object ID**, on the Load DB tab, queues a single `dl_id` to reload out
of band while the routine is running — it *updates* existing rows in place by
business key rather than appending new version rows, unlike the normal cycle.

**Drop a table-export `.zip`**, also on the Load DB tab, loads every table in
it as a fresh baseline — each table in the zip is truncated and reloaded.
Only available while the routine is stopped, so a bulk load and the tracking
loop never write at the same time.

## Branding

The status page uses the same doppiogroup.com look as the other three
packages — Montserrat, white/`#F7F7F7` backgrounds, `#D4000E` red as the
primary accent — with the state pills (idle/sleeping/processing/error) kept
on the same green/blue-gray/amber/red semantic mapping used in
`packages/m3_security`'s tables.
