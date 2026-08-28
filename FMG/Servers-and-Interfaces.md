# FMG — Servers & Interfaces Map

A running list of where everything lives. Not a technical spec — the deep detail lives in
`sftp/SFTP-Interfaces-Technical-Reference.md` and `ConcurImport/ConcurImport-Technical-Reference.md`.

**Last updated:** 2026-08-26

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
| JBA — AS/400 (Aurora, BI) | `172.16.101.19` | DW loads, Concur invoice posting |
| ASRDW — AS/400 (Brazil) | `172.16.28.10` | Concur T&E distribution |
| Infor M3 | Azure SQL (`ngeumsmi0002...database.windows.net`) and `172.16.228.53` | DW loads, M3 XML posting |
| Raymond ERP share | `\\RAUSMSRV0001` / `\\RAUSMSRV0002` | HSBC payments, Concur T&E |
| SAP (Maenner) file drop | `\\EDDES0010.ed.bg1857.net` | Concur T&E distribution |
| Synventive file drop | `\\SYUSS1000.SY.BG1857.net` | Concur T&E distribution |
| Syteline | `BAUSS9100` / `BAUSS9000` | Concur T&E distribution |

---

## 5. Open questions / to confirm

- **`ng` vs `bg` server names.** The SSIS projects reference both `ngusmsrv0031` and `bgusmsrv0031`,
  and both `ngusmsrv0019` and `bgusmsrv0019` (plus `bgusmsrv0019-te`). Looks like an old domain and a
  new one — need to confirm which is current and whether the old names still resolve.
- **`bgusmsrv0031`** shows a `Concur_SAE` database that doesn't appear on `ngusmsrv0031`. Where does
  it live now?
- **HighRadius extract output path** points at `\\BGUSMSRV0019...\HighRadius\stage\out` while the
  SFTP side runs on NGUSMSRV0019 — confirm these line up.
- Also seen but not yet placed: `ngusmsrv0030`, `bgusmsrv0030`, `bguss0012`, `bguss0014`, `bguss0029`
  (older data warehouse hosts).
- Where SSIS packages are actually deployed (SSISDB catalog, which server, which folder) — not
  recorded anywhere in the repo.
- Scheduled task inventory on NGUSMSRV0019 — cadences below are inferred from log timestamps, not
  read off Task Scheduler.

---

## 6. Adding to this doc

Drop new facts straight into the tables above. Rough order to capture for a new interface:
what it does → which server runs it → where the code lives → where the config lives →
where the data lands → who/what it talks to → live or not.
