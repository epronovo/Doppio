# MIG_* — M3 migration sync tools

Nine tools for moving M3 configuration between tenants during a migration —
panel views, partner media, partner reference data, sort options, sort
orders, translation data, standard field-group generation, export-log
summaries, and a live-count migration review — wrapped in one small Flask
app instead of nine separate `input()`-prompting command line scripts.

All routines are prefixed `MIG_` so they group together in the folder.

| File | Role |
|------|------|
| `MIG_Api.py` | Self-contained auth/HTTP module — `Tenant`, `authenticate()`, `post_to_m3()`, upload/process helpers |
| `MIG_App.py` | Flask front end — nine tabs, tenant connect, background jobs |
| `MIG_SyncPanelViews.py` | Panel view (CSYSPV) diff + EVS100 export |
| `MIG_SyncPartnerMedia.py` | Partner media (CRS949) diff + direct write |
| `MIG_SyncPartnerRef.py` | Partner reference (CRS945) diff + direct write |
| `MIG_SyncSortingOptions.py` | Sort option (CRS021) diff + EVS100 export |
| `MIG_SyncSortingOrders.py` | Sort order (CRS022) diff + EVS100 export |
| `MIG_SyncTranslData.py` | Translation data (MBMTRN/MBMTRD) extract + EVS100 export + direct push |
| `MIG_GenerateFieldGroups.py` | CMS005MI.GenStandard + MNS320 job poll |
| `MIG_ExportSummary.py` | `.log` → `.xlsx` export-log summariser (no M3 calls) |
| `MIG_MigrationReview.py` | Live M3 record counts written into a review workbook |
| `templates/MIG_Index.html` | The single-page UI |

## Quick start

```bash
pip install flask requests openpyxl xlsxwriter
python MIG_App.py                       # http://127.0.0.1:5059
```

`--host`, `--port` (default 5059) and `--debug` are all available on the
command line, same as `M3_Security_App.py` (5057) and `ADP_Concur_App.py`
(5058) — 5059 is simply the next free port in the sequence.

## Deliberate scope decisions — read before changing anything here

**No persistent database.** Unlike `packages/m3_security`, this package does
not capture the sync data into SQLite. Every "Fetch & Diff" run is a live,
in-memory SOURCE ↔ DEST comparison, exactly like the original CLI scripts —
nothing survives a page reload, there is no edit-then-push-later workflow,
and there is no audit-log table of past runs. If you connect SOURCE and DEST
tenants, run a diff, and then refresh the page, you start over. This was a
scope decision made deliberately, not an oversight — mig_sync is a thin web
wrapper around nine already-correct scripts, not a new persistent app.

For the same reason, `MIG_Api.post_to_m3()` does **not** call an
`APIBatchLogger` / write any sqlite audit row the way the original
`InforMI.post_to_m3()` did. The CLI tools logged every M3 call to a local
sqlite table for later inspection; this app intentionally does not — adding
that back would mean adding the very persistence layer this package was
scoped to avoid.

**Two shared repo-root folders — not package-local, do not copy them in:**

- `ionapi/` — the same `.ionapi` tenant-credential files every Infor
  automation tool in this repo reads (M3 Security, ADP↔Concur, and MIG all
  point at the one folder). `MIG_Api.DEFAULT_IONAPI_DIR` computes this as
  `Path(__file__).resolve().parents[2] / "ionapi"` — two levels up from
  `packages/mig_sync/` to the repo root.
- `evs100/ToProcess/` — a live pickup folder. An external process outside
  this repo watches it and moves whatever it consumes into
  `evs100/Complete/`. The four routines that write an EVS100-format `.xlsx`
  (Panel Views, Sorting Options, Sorting Orders, Transl Data) write there —
  `MIG_App.EVS100_TO_PROCESS`, computed the same two-levels-up way. Pointing
  either of these at a package-local copy breaks the real M3 import pickup
  or leaves every other Infor tool unable to find its credentials.

**Shared SQLite lookup DB (read-only here).** `MIG_ExportSummary` and
`MIG_MigrationReview` read table descriptions out of `~/sqlite/doppio.db`
(tables `m3tables` / `m3TableCols`), populated by an unrelated tool — out of
scope for this package. Override the path with the `MIG_SYNC_DB` environment
variable, the same convention `M3_SECURITY_DB` uses in `m3_security`.

Everything else — uploaded logs/workbooks, generated report/summary/error
`.xlsx` files — is package-local under `input/mig_sync/…` and
`output/mig_sync/…`, matching the `input/<pkg>/` / `output/<pkg>/`
convention already used by `m3_security` and `adp_concur`.

## Tenant connect

One shared control, `POST /api/tenant/connect`, used by every tab that needs
a live M3 connection (every tab except Export Summary). Pick an `.ionapi`
file from `GET /api/ionapi`, give a company (and division, where relevant),
and the app authenticates immediately and hands back a `tenant_id` — the
access token itself never reaches the browser. Tabs that compare two tenants
(Panel Views, Partner Media, Partner Ref, Sorting Options, Sorting Orders,
Transl Data) connect SOURCE and DEST independently, so they can be different
`.ionapi` files, different companies, or the same tenant twice. Field Groups
and Migration Review need exactly one tenant. There is no "reuse the last
answer for an hour" caching the CLI scripts had — every session connects
fresh, on purpose, since two tabs may need two different tenants at once.

## Panel Views

*"Syncs panel view column configurations from a SOURCE tenant to a
DESTINATION tenant."* Reads `CSYSPV` from both tenants via
`EXPORTMI.Select` (optionally filtered to one `PGNM`), finds views whose
`C9PARA` column layout differs (matched on program + interactive column +
version), and exports the differing views to an EVS100-format workbook —
a `DelPanelVersion` sheet plus an `ImportView` sheet. `C9PARA` is sanitized
before comparison and before export: byte values that are not a valid field
name, digit, or the handful of punctuation characters CRS020MI expects are
replaced with `'0'`, so a view already imported once (where those bytes were
already replaced) compares equal to its SOURCE original on the next run
instead of flagging a false difference forever. Upload the exported file to
DEST and trigger `EVS100MI.ImportFile` from the same tab once it's ready.

## Partner Media

*"Syncs partner media data (CRS949) from a SOURCE tenant to a DESTINATION
tenant."* No EVS100 file here — this one writes directly. Records that exist
in DEST but differ are deleted (`CRS949MI.DltPartnerMedia`) and re-added;
records missing from DEST are added outright. Which "Add" transaction to use
is routed by `MEDC`: `MAIL` → `AddPartnerEmail`, `MBMEVENT` → `AddPartnerMBM`;
any other `MEDC` value is skipped entirely (neither deleted nor added) and
called out separately in both the preview and the summary. A summary
workbook is always written after Apply, whether or not anything needed
changing, listing every Dlt/Add call made with its outcome.

## Partner Ref

*"Syncs partner reference data (CRS945) from a SOURCE tenant to a
DESTINATION tenant."* Matches on division + partner document reference key,
compares `TX15`/`TX40`, and calls `CRS945MI.UpdPartnerRef` for records that
differ or `CRS945MI.AddPartnerRef` for records missing from DEST. Unlike
Partner Media, this one only writes an error workbook when something
actually failed.

## Sorting Options

*"Syncs sort options from a SOURCE tenant to a DESTINATION tenant."* Reads
`CRS021MI.LstSrtOpt` from both tenants and finds `(FILE, SOPT)` pairs present
in SOURCE but missing from DEST (M3's own default `SOPT='JD'` is always
excluded — never migrated). Custom sort options in the `U`/`V`/`X` ranges
(`U1`-`U9`, `V1`-`V9`, `X1`-`X9`) go through `AddSrtOpt` + `ActSrtOpt`;
anything else is a standard sort option, generated once per file via
`CrtStdSrtOpt` regardless of how many standard `SOPT`s that file was missing.
Exports to an EVS100 workbook with up to three sheets, upload and process
the same way as Panel Views.

## Sorting Orders

*"Syncs sort orders from a SOURCE tenant to a DESTINATION tenant."* Reads
`CRS022MI.LstSortOrder` (optionally filtered to one `PGNM`; `CMS100` and
`LISTMI` programs are always excluded) and finds records missing from DEST
(`AddSortOrder`) or present but different (`ChgSortOrder`) — keyed on
program + file + sort option + query type, since one `(PGNM, FILE, SOPT)`
can have a row per query type. Every changed record's field-by-field diff
(`diff_fields()`) is preserved in the plan so the results table shows exactly
which columns moved, not just that something did.

## Transl Data

*"Extracts translation data (MBMTRN/MBMTRD) from a SOURCE tenant and writes
it to an EVS100-compatible Excel file for import into a DESTINATION
tenant."* Translation data is not scoped to a company or division, so both
SOURCE and DEST connect with `global_scope: true`, which drops `&cono=`/
`&divi=` from the M3 API URL entirely instead of sending them blank — the
same effect the original CLI got from a private `_global_url()` helper,
reached here through `MIG_Api.Tenant(global_scope=True)` instead. Header
rows (`MBMTRN`) are joined to detail rows (`MBMTRD`) on `IDTR`; headers are
deduplicated so `AddTranslation` gets one row per unique translation ID while
`AddTranslData` gets one row per detail. Besides the EVS100 export, this tab
can also push `CRS881MI.AddTranslation` / `AddTranslData` directly at DEST —
an error workbook is written only if some records failed.

## Field Groups

*"Runs CMS005MI.GenStandard (ST02=1) against a single tenant and waits for
the background job to complete."* One tenant, one button. Submits
`CMS005MI.GenStandard` and then polls `MNS320MI.Get` for the returned `BJNO`
every 60 seconds (`POLL_INTERVAL` — this is the real MNS320 polling cadence,
not a shortened demo value) until the job either drops out of the queue
(completed normally) or reports a `STAT` outside `{00, 20}` (ended in error).
This can run for several minutes; the job's phase text updates every poll so
the tab doesn't look stuck, but there is no percentage — the wait is
open-ended.

## Export Summary

*"Reads `.log` files ... and produces a detailed Excel workbook summary for
each file."* The only tab that never talks to M3. Drop one or more M3
IES/export component `.log` files (or use **Scan folder** to reprocess
whatever is already sitting in `input/mig_sync/export_summary/`) and get back
a multi-sheet workbook per log: an overview, tables with data (ranked),
empty tables, every table with company/division-scoping flags looked up from
`~/sqlite/doppio.db`, errors/warnings, and the tables the export's own
exclusion-filter regex left out (parsed and expanded from the log header,
not just shown as raw regex).

## Migration Review

Watches for uploaded workbooks containing a `DTA_FIN_MVX` sheet, queries a
live M3 record count for every table listed (`EXPORTMI` `count(#) from
<table>`, fetched in parallel), and writes back five columns: the live
count, the difference against the workbook's own count, a percent-difference
formula, and the table's description + maintained-by, looked up from
`~/sqlite/doppio.db`. A four-row summary block classifies every row as a
matching count, a protected/remarked row, an acceptable difference, or a
questionable one. The reviewed copy lands in
`output/mig_sync/migration_review/reviewed/` with a `_REVIEWED` suffix; the
original moves to `output/mig_sync/migration_review/processed/`. Upload
workbooks, then use **Scan folder** (needs a connected tenant) to process
whatever is waiting in `input/mig_sync/migration_review/`.
