# FMG — Servers & Interfaces Map

A running list of where everything lives. Not a technical spec — the deep detail lives in
`ngusmsrv0019/SFTP-Interfaces-Technical-Reference.md` and
`ngusmsrv0031/ConcurImport/ConcurImport-Technical-Reference.md`. The client-facing view of the same
inventory lives in `FMG_Integration_Management_Tracker.xlsx` (see §6).

**Last updated:** 2026-09-15

---

## 1. Servers at a glance

| Server | Domain name | What it is | Key locations |
|---|---|---|---|
| **NGUSMSRV0031** | `ngusmsrv0031.ng.bg1857.net` | **Development + SQL box.** SQL Server, SSMS, Visual Studio | `C:\Development` — VS / SSIS projects<br>`F:\common\ConnectionStrings.config` — global connection config |
| **NGUSMSRV0019** | `ngusmsrv0019.ng.bg1857.net` | **File transfer server.** Runs the scheduled SFTP jobs via WinSCP + PowerShell | `E:\` — the whole SFTP tree (mirrored in this repo under `sftp/`) |

Anything else in the table below is a system we *talk to*, not a box we own.

### NGUSMSRV0031 — the dev/SQL server

- **SSMS** is installed here — this is where you query the databases.
- **Visual Studio** is installed here — SSIS projects live in `C:\Development`.
- **`F:\common\ConnectionStrings.config`** is the shared connection config used across projects.
- Databases seen in the SSIS projects on this instance:
  - `DW_BGI_OpData` — data warehouse (operational data)
  - `DW_BGI_OpStage` — data warehouse staging
  - `DW_BGI_MasterData` — master data
  - `App_Concur` — Concur PO extract stored procs
  - `Concur_Invoice_Worktables` — Concur invoice work tables

### NGUSMSRV0019 — the SFTP / file transfer server

- All partner file transfer runs here, via **WinSCP** driven by **PowerShell** scheduled tasks.
- Everything sits on the **`E:\`** drive. The tree cannot be moved off `E:\` without editing scripts —
  paths are hard-coded.
- Layout:
  - `E:\Applications\<Partner>\` — working folders, per-environment config, the scripts that run
  - `E:\Scripts\WinScpFTP\<Partner>\` — the shared connection library + connection config
  - `E:\KeysEncryptions\<Partner>\` — SSH keys, PGP keys, passphrase files
- Each interface has **two config files per environment**: one for paths/routing (under `Applications`),
  one for the connection itself (under `Scripts\WinScpFTP`). Both must exist or the run fails.

---

## 2. Interfaces

| Interface | Direction | Partner / system | What it does | Runs on | Status |
|---|---|---|---|---|---|
| **HSBC** | Both ways | HSBC bank | AP payments out (ACH / check / wire), acknowledgements back | NGUSMSRV0019 | **Live** — every 3–10 min |
| **HighRadius** | Inbound | HighRadius | AR cash application — remittance advice into M3 and JBA | NGUSMSRV0019 | Configured; stage only so far |
| **Concur Import** | Inbound | SAP Concur | Pull Concur extract files, then load them via SSIS | NGUSMSRV0019 (pull) + NGUSMSRV0031 (SSIS) | Stage only |
| **Concur Extracts** | Outbound | SAP Concur | Purchase Order / vendor extract files out to Concur | NGUSMSRV0031 (SSIS) | PO extract built; SFTP side is **empty scaffolding** |
| **BGI DW — M3** | Internal | Infor M3 | Load M3 data into the data warehouse | NGUSMSRV0031 | Live |
| **BGI DW — Aurora** | Internal | JBA / AS400 (Aurora) | Load JBA data into the data warehouse | NGUSMSRV0031 | Live |
| **HighRadius Extracts** | Outbound | HighRadius | AR customer / open item / cash files out | NGUSMSRV0031 (SSIS) | Live |
| **JBA AR Autoposting** | Internal | JBA / I-series | Pulls the HighRadius 820/823 data into JBA so Shared Services can apply cash (10/AIAS1) | I-series | **Live** — daily, 12:15 & 12:16 ET |
| **JBA Concur T&E posting** | Internal | JBA / I-series | Posts the Concur T&E batch landed in PLP81U to the GL via 14/GLP | I-series | Live — schedule TBC |

---

## 3. Where each interface lives

### HSBC — AP payments
- **Runs on:** NGUSMSRV0019
- **Scripts:** `E:\Applications\HSBC\_scripts\` — three of them, on separate schedules:
  - `HSBC-US_prepare.ps1` — every 3 min, picks up payment files from the ERP share
  - `HSBC-US_send.ps1` — every 5 min, encrypts and uploads
  - `HSBC-US_get.ps1` — every 10 min, downloads acknowledgements and decrypts them
- **Config:** `E:\Applications\HSBC\configuration\{production,stage}.json`
- **Connection:** `E:\Scripts\WinScpFTP\HSBC\`
- **Keys:** `E:\KeysEncryptions\HSBC\` — SSH key plus PGP keys
- **Runs as:** service account `NG\M0042`
- **Files come from the ERP at:** `\\RAUSMSRV0001.ra.bg1857.net\ASR\AP-Bank\HSBC\out` (production),
  `\\RAUSMSRV0002...` (stage)
- **Note:** this is the **only** interface that uses PGP encryption. The others are plaintext.

### HighRadius — AR cash application
- **Runs on:** NGUSMSRV0019
- **Scripts:** `E:\Applications\HighRadius\` + `E:\Scripts\WinScpFTP\HighRadius\`
- **Keys:** `E:\KeysEncryptions\HighRadius\`
- **Outbound extract files** are produced by the `HighRadiusExtracts` SSIS project and land on a
  `\\...\HighRadius\stage\out\` share (ARCASH / ARCUST / ARMAST)

### Concur Import — inbound
- **Pull side runs on:** NGUSMSRV0019 — `E:\Applications\ConcurImport\`, connection under
  `E:\Scripts\WinScpFTP\Concur\`, key in `E:\KeysEncryptions\Concur\`
- **Load side runs on:** NGUSMSRV0031 — the `ConcurImport` SSIS project
- **Two pipelines:**
  - **Travel & Expense** — lands the Concur SAE extract in SQL, then fans it out to each business
    unit's ERP (JBA, M3, SAP, Syteline)
  - **Invoice** — lands invoice header/detail, enriches, posts to JBA and M3
- **Only stage is configured** on the SFTP side — there is no Concur production connection config yet.

### Concur Extracts — outbound
- **Built on:** NGUSMSRV0031 — the `ConcurExtracts` SSIS project (`PurchaseOrder.dtsx`)
- Generates the PO, PO receipt and vendor import files into `C:\temp\ConcurInvoice\`
- **The SFTP side does not exist yet.** `E:\Applications\ConcurExtracts\` on NGUSMSRV0019 is an
  empty folder structure — no scripts, no config, no send function.

### JBA / I-series legs (AR autoposting, Concur T&E posting)

Source: interface workbook from **Greg French**, 2026-09-14 (`source-documents/`). These are the
**downstream halves** of interfaces we already had — the part that runs on the AS/400, which nothing in
this repo covers.

- **Server:** `10.192.0.5` / **RackSquared**. See the open question in §5 — our other notes say
  `172.16.101.19` for the same box.
- **Libraries:** `PRDMODSF4` (production), `QUAMODSF4` (test) — the Aurora qualifiers.
- **AR autoposting** — HighRadius data lands as `HRADIUS820` / `HRADIUS823` around **11:45 AM ET**;
  jobs **`IHRAD820` (12:15)** and **`IHRAD823` (12:16)** pull it in; Shared Services applies cash with
  **`10/AIAS1`**. The 12:15 jobs have a silent dependency on the 11:45 load having happened.
- **Concur T&E** — the batch in `PLP81U` is posted via **`14/GLP`**. Bad GL accounts arrive from Concur
  routinely, so Shared Services runs the **`CONCURCHK1`** and **`CONCURCHK2`** queries in **`JBAQRY`** to
  find them first. Schedule unknown to Greg — get it from Shared Services.
- **Contacts:** Greg French, Ray Schut (I-series); Kevin Alexander, Dean Amo, Kim Diorio (interfaces);
  Catherine Lee (IT contact for JBA application support).

Greg also supports JBA application areas well outside integration work — backups, DR, year-end, audits,
Spoolit, ICS Forms, source change management, and the Cyrus One → RackSquared migration. That list is on
the **JBA Interfaces & Support** tab of the tracker; it is single-person knowledge the TSA transition has
to land somewhere.

### SQL Agent schedules (NGUSMSRV0031)

Source: job export provided 2026-09-15 (`source-documents/jobs.csv`). Every SSIS job runs from the
SSISDB catalog folder **`\SSISDB\NG_Packages\<Project>`** on `ngusmsrv0031.ng.bg1857.net`. All are daily.

| Job | Package | Runs | Interface |
|---|---|---|---|
| `_08:00_BGIDW_ASR_M3_Daily` | `BGIDW_ASR_M3_Daily.dtsx` | 08:00 | DW — M3 load |
| `_08:00_BGIDW_Aurora_Daily` | `BGIDW_Aurora_Daily.dtsx` | **08:30** (name says 08:00) | DW — Aurora load |
| `_11:30_HighRadius-EDISendToJBA` | `EDISendToJBA.dtsx` | 11:30 | HighRadius 820/823 → JBA |
| `_ConcurExtracts-PurchaseOrder` | `PurchaseOrder.dtsx` | 20:10 | Concur PO/vendor/receipt extract |
| `_ConcurImport_T&E` | `TravelAndExpense_Incoming.dtsx` | **no schedule attached** | Concur T&E |
| `_ConcurImport-Invoices` | `InvoiceImport.dtsx` | **no schedule attached** | Concur invoices |

Plus two housekeeping jobs: `SSIS Server Maintenance Job` (daily 00:00 — purges SSISDB operation records
outside the retention window, which is what limits how far back execution history can be read) and
`syspolicy_purge_history` (daily 02:00).

Three things worth knowing:

- **The HighRadius chain now has a full clock.** Job starts 11:30 → data on the I-series ~11:45 →
  `IHRAD820`/`IHRAD823` at 12:15/12:16. About 45 minutes of headroom, and no alarm if the load is late.
- **`_ConcurExtracts-PurchaseOrder` runs nightly into a dead end.** The SFTP send half does not exist, so
  it has been writing files to `C:\temp\ConcurInvoice` with nothing to ship them.
- **Nothing on NGUSMSRV0019 appears in this export.** The HSBC, HighRadius and Concur file transfers run
  under Windows Task Scheduler on that box — still a separate, unconfirmed inventory.

### Data warehouse loads (BGI_DW_M3, BGIDW_Aurora)
- **Run on:** NGUSMSRV0031
- Pull from the AS/400 (JBA) and Infor M3, load into `DW_BGI_OpData` / `DW_BGI_OpStage` /
  `DW_BGI_MasterData`
- Daily and hourly packages in each project

---

## 4. Systems we connect to

| System | Where | Used by |
|---|---|---|
| HSBC SFTP | `ecom-sftp.fguk-prd2.hsbc.com` (prod) / `...-pprd2...` (stage), port 10022 | HSBC interface |
| HighRadius SFTP | `sftp-cloud13.highradius.com` (prod) / `uatsftp-turtle13...` (UAT) | HighRadius |
| SAP Concur SFTP | `mft-us2.concursolutions.com` | Concur Import / Extracts |
| JBA — AS/400 (Aurora, BI) | `172.16.101.19` — but G. French gives `10.192.0.5` / RackSquared (§5) | DW loads, Concur invoice posting, AR autoposting, Concur T&E posting |
| ASRDW — AS/400 (Brazil) | `172.16.28.10` | Concur T&E distribution |
| Infor M3 | Azure SQL (`ngeumsmi0002...database.windows.net`) and `172.16.228.53` | DW loads, M3 XML posting |
| Raymond ERP share | `\\RAUSMSRV0001` / `\\RAUSMSRV0002` | HSBC payments, Concur T&E |
| SAP (Maenner) file drop | `\\EDDES0010.ed.bg1857.net` | Concur T&E distribution |
| Synventive file drop | `\\SYUSS1000.SY.BG1857.net` | Concur T&E distribution |
| Syteline | `BAUSS9100` / `BAUSS9000` | Concur T&E distribution |

---

## 5. Open questions / to confirm

- **JBA I-series address.** G. French gives `10.192.0.5` (RackSquared); our notes and the SSIS connection
  strings use `172.16.101.19`. The Cyrus One → RackSquared migration is the likely explanation — confirm
  which address is current and whether anything still points at the old one.
- **PNC → HighRadius.** G. French: *some* customer payments still reach HighRadius from **PNC**, not HSBC.
  That is a second bank on the receivables side and it has no row in the inventory yet — find out what
  still flows that way and who owns it.
- **Concur bad GL accounts.** A standing data-quality defect, worked around manually with the
  `CONCURCHK1/2` queries. Root cause is in the Concur GL mapping, not in JBA — worth a tactical item.
- **`ng` vs `bg` server names.** The SSIS projects reference both `ngusmsrv0031` and `bgusmsrv0031`,
  and both `ngusmsrv0019` and `bgusmsrv0019` (plus `bgusmsrv0019-te`). Looks like an old domain and a
  new one — need to confirm which is current and whether the old names still resolve.
- **`bgusmsrv0031`** shows a `Concur_SAE` database that doesn't appear on `ngusmsrv0031`. Where does
  it live now?
- **HighRadius extract output path** points at `\\BGUSMSRV0019...\HighRadius\stage\out` while the
  SFTP side runs on NGUSMSRV0019 — confirm these line up.
- Also seen but not yet placed: `ngusmsrv0030`, `bgusmsrv0030`, `bguss0012`, `bguss0014`, `bguss0029`
  (older data warehouse hosts).
- ~~Where SSIS packages are actually deployed~~ — **answered 2026-09-15** by the SQL Agent job export:
  SSISDB catalog on `ngusmsrv0031.ng.bg1857.net`, folder **`\SSISDB\NG_Packages\<Project>`**. See §3.
- **Which SSISDB environment each job binds to.** The job export shows no environment reference, so
  whether a job runs under the Development, Stage or Production configuration is unconfirmed. Matters most
  for `_ConcurExtracts-PurchaseOrder`, which writes to a local `C:\temp` path.
- **Unscheduled packages — dead or missing a schedule?** `BGIDW_M3_Daily`, `_026`, `_Test`, both M3 hourly
  packages, `BGIDW_Aurora_Hourly/_Orders/_History`, `HighRadiusExtracts.dtsx` and
  `HighRadiusInvoiceExtract.dtsx` have no SQL Agent job at all.
- **`_ConcurExtracts-PurchaseOrder` output.** It runs nightly at 20:10 but nothing ships the files — see
  §3. Find out what has piled up in `C:\temp\ConcurInvoice` and how long it has been running.
- Scheduled task inventory on NGUSMSRV0019 — cadences below are inferred from log timestamps, not
  read off Task Scheduler.

---

## 6. Adding a new interface — what to update

Three things get updated every time an interface is added or materially changed. Do all three;
they are read by different people.

| # | Document | What goes in it |
|---|---|---|
| 1 | **This file** (`Servers-and-Interfaces.md`) | The engineer's map. Add a row to §2, a subsection in §3, and any new endpoint to §4. |
| 2 | **`FMG_Integration_Management_Tracker.xlsx`** | The client-facing tracker (SOW Doppio ID 30127, owner Denise Meyer). Add a row to **Strategic Integration Inventory** with the next `STRAT-nnn` id, and a row to **Integration Monitoring Log** once it runs. Tactical requests go in the **Tactical Integration Work Log**. |
| 3 | The relevant **technical reference** | `SFTP-Interfaces-Technical-Reference.md` for anything on NGUSMSRV0019, `ConcurImport-Technical-Reference.md` for the Concur SSIS estate — or start a new one for a new project. |

Rough order to capture for a new interface: what it does → which server runs it → where the code
lives → where the config lives → where the data lands → who/what it talks to → live or not →
how you would know it failed.

**Tracker column notes** (the workbook's own Legend tab is authoritative):

- *Integration ID* — `STRAT-nnn`, sequential. Cite it from the Monitoring Log.
- *Source Code Repo Status* — `Migrated` means it is in the client DevOps repo. The SSIS projects
  carry a `build.yaml`; the PowerShell tree on NGUSMSRV0019 has no repo evidence yet.
- *Documentation Status* — `Complete` means a technical reference covers it, not that a README exists.
- *Notes / Risks* — carry the concrete defect over from the technical reference; that column is what
  gets read in the status meeting.
