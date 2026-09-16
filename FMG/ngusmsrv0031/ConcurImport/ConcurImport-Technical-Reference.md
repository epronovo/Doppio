# ConcurImport — SSIS Technical Reference

**Solution:** `ConcurImport.sln` · **Project:** `ConcurImport.dtproj` (SSIS Project Deployment Model)
**Target:** SQL Server 2019 · **Designer version:** 16.0.5467.0 (VS 2019/2022 SSDT) · **Package format:** 8
**Protection level:** `EncryptSensitiveWithPassword` (project password: `SSIS`, per `build.yaml`)
**Source path:** `FMG/ConcurImport`
**Documented from source as of:** 2026-08-21

---

## 1. What this project does

ConcurImport is the integration layer between **SAP Concur** and Barnes Group's ERP estate. It runs two
independent pipelines:

| Pipeline | Trigger file | Purpose |
|---|---|---|
| **Travel & Expense** | `CONCURSAEEXTRACT*.TXT` (Concur Standard Accounting Extract, pipe-delimited) | Land the SAE extract into SQL, then fan it out to each business unit's ERP — AS/400 JBA, Infor M3, SAP, and Syteline file drops |
| **Invoice** | `invoice_header.txt` + `invoice_detail.txt` (pipe-delimited) | Land Concur Invoice header/detail into work tables, enrich them, then post to JBA (AS/400 tables) and M3 (Infor OAGIS XML BODs) |

Both pipelines follow the same shape:

```
detect file → land raw data in SQL → transform/enrich via stored procedures
   → distribute to each ERP → log status → email a summary
```

Almost all real logic lives in **SQL Server stored procedures** and in **C# Script Tasks**. The SSIS
packages are orchestration and transport; they contain very little transformation logic of their own.

---

## 2. Package inventory

Seven `.dtsx` files are present in the project folder and listed in the project manifest. All seven are
marked `EntryPoint="1"`, but only three are true entry points in practice.

| Package file | Package name (`ObjectName`) | Role | File mtime |
|---|---|---|---|
| `TravelandExpense_Incoming.dtsx` | `TravelAndExpense_Incoming` | **Entry point** — T&E ingest | 2026-03-03 |
| `TravelAndExpense_Distribution.dtsx` | `TravelAndExpense_Distribution` | Child of Incoming — T&E fan-out | 2026-08-18 |
| `InvoiceImport.dtsx` | `Invoice_Import` | **Entry point** — Invoice ingest + post (current) | 2026-08-18 |
| `InvoiceImportStandard.dtsx` | `InvoiceImportStandard` | **Entry point** — Invoice "standard extract" variant (mostly disabled) | 2026-04-08 |
| `InvoicePost_Aurora.dtsx` | `InvoicePost_Aurora` | Child — post invoices to JBA Aurora AS/400 | 2026-08-07 |
| `InvoicePost_Raymond.dtsx` | `InvoicePost_Raymond` | Child — post invoices to JBA Raymond AS/400 | 2026-08-18 |
| `InvoicePost_ASRM3.dtsx` | `InvoicePost_ASRM3` | Child — generate Infor M3 XML BODs | 2026-08-07 |

*Note: the package file name and its internal package name differ in two cases —
`TravelandExpense_Incoming.dtsx` contains `TravelAndExpense_Incoming` (capital A), and `InvoiceImport.dtsx`
contains `Invoice_Import` (underscore). Project references and `Execute Package` tasks use the **file**
name.*

### 2.1 Missing packages — read this first

The build output under `obj/{Development,Stage,Production}/` contains **four additional packages whose
source files no longer exist in the project folder**:

- `InvoicePost_Syteline.dtsx`
- `InvoicePost_Syteline9.01.dtsx`
- `InvoicePost_Syteline10.dtsx`
- `InvoicePost_SytelineAsia.dtsx`

The `obj/*/ConcurImport.dtproj` manifests from the last three builds all list **11 packages**; the current
source `.dtproj` lists **7**. The Syteline packages were removed from the project (project file modified
2026-08-14) but their source `.dtsx` files are gone, not just unreferenced.

**Consequence:** `InvoiceImportStandard` still contains `Execute Package` tasks pointing at
`InvoicePost_Syteline.dtsx` and `InvoicePost_Syteline9.01.dtsx` via project reference. Those references
cannot resolve in the current project. They sit inside the disabled `Invoice Posting` container, so a build
may still succeed, but the project as it stands cannot reproduce the deployed `.ispac` files.

Corroborating evidence that the removal was incomplete: `ConcurImport.dtproj` still carries per-configuration
overrides for `InvoicePost_Syteline10::FlatFile_EGOEM_Header` and `InvoicePost_Syteline10::FlatFile_EGOEM_Detail`
(pointing at `\\BAUSS0930.ba.bg1857.net\SL10FileServer\EGOEM\{TRN,PRD}\Concur\`), for a package that is no
longer in the manifest.

If Syteline posting is still live in production, the source for those four packages needs to be recovered
(from `obj/Production/*.dtsx`, from the deployed `.ispac`, or from source control) before any redeploy.
`obj/Production/` and `obj/Stage/` are the most likely recovery points — they hold full, buildable `.dtsx`
copies of all four.

### 2.2 Build artifacts on disk

| Path | Built (mtime) | Contains |
|---|---|---|
| `bin/Development/ConcurImport.ispac` | 2026-06-08 | 11 packages |
| `bin/Production/ConcurImport.ispac` | 2026-06-18 | 11 packages |
| `bin/Stage/ConcurImport.ispac` | 2026-08-03 | 11 packages |

Production's `.ispac` is the oldest of the three. Verify against what is actually deployed to SSISDB before
assuming these reflect production.

---

## 3. Project-level configuration

### 3.1 Project parameters

Defined in `Project.params`; values below are the checked-in defaults, overridden per configuration.

| Parameter | Type | Purpose |
|---|---|---|
| `ConnectionStringsFile` | String | UNC path to a shared `.config` file holding every connection string. **This is the linchpin of the whole connection model** — see §4. |
| `EmailNotificationService` | String | REST endpoint for outbound email: `https://apps.onebarnes.com:7443/BGI_Svc_Notification/api/Notification/SendEmail` |
| `EmailNotificationTo` | String | Distribution list for all notifications |
| `Environment` | String | Free-text environment label used in email subjects/bodies |

### 3.2 Build configurations

Three configurations: **Development**, **Stage**, **Production**. 45 parameter overrides are stored in
`ConcurImport.dtproj`. The values that actually differ between environments:

| Parameter | Development | Stage | Production |
|---|---|---|---|
| `Project::ConnectionStringsFile` | `\\bgusmsrv0031…\common\ConnectionStrings.config` | same as Dev | `\\bgusmsrv0030…\common\ConnectionStrings.config` |
| `Project::EmailNotificationTo` | `kalexander@onebarnes.com` | same as Dev | `DL_Concur_TE_Distribution@bginc.com` |
| `Project::Environment` | `Development` | `Stage` | `Production` |

**AS/400 library qualifiers** — the pattern is `{env-prefix}{system}.{file}`:

| Parameter group | Development | Stage | Production |
|---|---|---|---|
| Aurora (`BI_PLP9xU`, `PLP8xU`, `PLP9xU`) | `QUAMODSF4.*` | `QUAMODSF4.*` | `PRDMODSF4.*` |
| Raymond (`BR_PLP8xU`, `BR_PLP9xU`) | `PTFMODF35.*` | `PTFMODF35.*` | `PRDMODF35.*` |

Note that **Development and Stage point at the same AS/400 libraries** for both Aurora and Raymond. A
"Stage" run therefore writes to the same non-production AS/400 files a "Development" run does.

**File paths** follow three shapes:

- **Development** — everything redirected under `C:\temp\…` (local scratch)
- **Stage** — everything redirected under `\\ngusmsrv0019.ng.bg1857.net\ConcurImport\stage\…`, including
  simulated "out server" folders (`…\OutServers\EDDES0010.ed.bg1857.net\…`) so no real downstream server is
  touched
- **Production** — real UNC targets on the owning business unit's server

One exception worth noting: `InvoicePost_ASRM3::ASRM3_Xml_Directory` in **Production** points at
`\\RAUSMSRV0001.RA.BG1857.NET\asr\Concur_Invoice\IN` while its **backup** directory points at
`\\ngusmsrv0019.ng.bg1857.net\ConcurImport\production\ConcurInvoice\import\ASRM3XML` — a different server
from the Dev/Stage pattern. Intentional or not, it is inconsistent with every other pairing.

### 3.3 Build pipeline

`build.yaml` (Azure DevOps):

```yaml
trigger: all branches except master
pool:   Azure Pipelines / windows-2022
steps:
  - CopyFiles@2       ConcurImport → $(Build.ArtifactStagingDirectory)
  - SSISBuild@1       projectPath: staging dir, projectPassword: SSIS
  - PublishBuildArtifacts@1   drop
name: $(SourceBranchName)_$(BuildDefinitionName)_$(Date:yyyyMMdd)$(Rev:.r)
```

The trigger **excludes** `master`, so merges to master do not build. There is no release/deploy stage in
this file — deployment to SSISDB happens elsewhere (or manually).

---

## 4. The connection model

This project does **not** use SSIS project connection managers, and it does not use SSIS environments or
parameterised connection strings in the usual way. Instead:

1. Every package declares SSIS **package-level** connection managers whose `ConnectionString` property is
   bound by expression to a **user variable named `CNN_<Something>`** — e.g.
   `ConnectionString = @[User::CNN_SQL_OLEDB_Concur_SAE]`.
2. Every package's first task is a Script Task called **`GetConnectString`**, which:
   - reads `$Project::ConnectionStringsFile`,
   - opens it with `ConfigurationManager.OpenMappedExeConfiguration`,
   - iterates every variable in the package whose name starts with `CNN_`,
   - strips the `CNN_` prefix, looks up that name in `<connectionStrings>`, and assigns the value.
3. Every downstream task therefore resolves its connection at runtime. All connection managers and most
   containers carry `DelayValidation="True"` so SSIS does not try to validate before `GetConnectString` runs.

```
Project::ConnectionStringsFile  ──▶  GetConnectString (C# Script Task)
                                          │
                                          ├─▶ User::CNN_SQL_OLEDB_Concur_SAE
                                          ├─▶ User::CNN_AS400_OLEDB_BI_PCSRVR2
                                          └─▶ …                     ▲
                                                                    │ expression binding
                                          Package Connection Managers
```

`GetConnectString` swallows all exceptions and returns `""` for a name it cannot find, and the task still
reports `Success`. A missing or misnamed entry in `ConnectionStrings.config` produces an empty connection
string and a downstream failure whose message points at the wrong place.

Two packages — `InvoicePost_ASRM3` and `TravelAndExpense_Incoming` — have **no SSIS connection managers at
all**. Their `CNN_*` variables are consumed directly by ADO.NET code inside Script Tasks.

### 4.1 Logical connections used

| `CNN_` variable / CM name | Kind | Target |
|---|---|---|
| `CNN_SQL_OLEDB_Concur_SAE` | OLE DB | SQL Server DB `Concur_SAE` (T&E work/staging) |
| `CNN_Concur_SAE` | ADO.NET | Same DB, used from Script Tasks |
| `CNN_SQL_OLEDB_Concur_Invoice_Worktables_Integrated` | OLE DB | SQL Server DB `Concur_Invoice_Worktables` |
| `CNN_Concur_Invoice_Worktables` | ADO.NET | Same DB, used from Script Tasks |
| `CNN_SQLMI_ADO_DW` | ADO.NET | Data-warehouse SQL MI hosting `App_Concur` |
| `CNN_AS400_OLEDB_BI_PCSRVR2` | OLE DB | AS/400 — **Aurora** (Barnes Industrial), catalog `DB2L` |
| `CNN_AS400_OLEDB_BR_ASRDW` | OLE DB | AS/400 — **Raymond**, catalog `S781CCD0` |
| `CNN_SQL_OLEDB_BA_SL_App` | OLE DB | Syteline app DB — **declared but unused** (see §9) |

### 4.2 Security note

The `.dtsx` files carry **default (design-time) values for the `CNN_*` variables that include live-looking
AS/400 host IPs, user IDs and plaintext passwords**, and a SQL data source name. These are overwritten at
runtime by `GetConnectString`, so they are dead values functionally — but they are readable in plain text in
the repository, and the package `ProtectionLevel` (`EncryptSensitiveWithPassword`) does not protect them
because SSIS does not classify a String variable as sensitive.

Recommended: blank those variable defaults, and rotate the AS/400 credentials if they are still valid.

---

## 5. Travel & Expense pipeline

### 5.1 `TravelAndExpense_Incoming` — ingest

Package parameter: `FOLDER_ConcurSAE` (watch folder).
Variables: `CNN_Concur_SAE`, `varConcurImportFile`.

**Control flow**

```
GetConnectString
   └─▶ Travel and Expense File Exists          (Script Task)
          ├─ on Failure ──▶ File Does Not Exists Notification  (Script Task, email)
          └─ on Success ──▶ Foreach Loop Container
                              ├─ Store Concur Raw Data   (Script Task)
                              └─▶ Archive file           (Script Task)
                            └─▶ Travel and Expense Distribution  (Execute Package)
```

| Task | Detail |
|---|---|
| **Travel and Expense File Exists** | Counts files matching `CONCURSAEEXTRACT*.TXT` (lower-cased) in `FOLDER_ConcurSAE`, top directory only. Returns `Failure` if zero — the failure is the branch signal, not an error. |
| **Foreach Loop Container** | `ForEachFileEnumerator`, folder `= @[$Package::FOLDER_ConcurSAE]`, filespec `"CONCURSAEEXTRACT*.TXT"`, non-recursive, fully-qualified filename → `User::varConcurImportFile`. |
| **Store Concur Raw Data** | See below. |
| **Archive file** | Moves the processed file to `<FOLDER_ConcurSAE>\save\`, deleting any same-named file already there. |
| **Travel and Expense Distribution** | `Execute Package Task`, project reference → `TravelAndExpense_Distribution.dtsx`. Runs **once, after the loop**, not per file. |

**`Store Concur Raw Data`** (C#, ~750 lines) parses the Concur SAE file:

- **Header row** (line 1, `|`-delimited) → `p_insert_concur_raw_file`, with (0-based field positions):
  - `@batch_id` = field `[4]`
  - `@offset_batch_id` = field `[4]` **+ 1700** — the offset that turns a Concur batch id into "our" batch id
  - `@record_count` = field `[2]`, `@journal_amount` = field `[3]`
  - `@raw_file_name` = the file's name, `@raw_file_path` = `<its directory, lower-cased>\save`
  - Header parsing has **no length guards** — a short header row throws `IndexOutOfRangeException`, which is
    caught and logged, and the task still succeeds with nothing loaded.
  - returns `@raw_file_id UNIQUEIDENTIFIER` (OUTPUT)
- **Detail rows** (lines 2+) → mapped to a `Concur_Raw_Files_Data` model across **400 positional columns**,
  then written with `SqlBulkCopy` to table **`Concur_Raw_Files_Data`**.
- Column semantics follow the Concur SAE layout: employee, report, report-entry, credit-card transaction,
  cash advance, allocation (`allocation_custom_1..20`), tax/VAT reclaim, travel request, banking, and a
  scattering of `future_use_columnNNN` placeholders.
- Date-ish fields are normalised to `yyyy-MM-dd` strings; money fields to nullable decimals.

**Error handling:** the parse loop, the DataTable fill, and `InsertFile` each wrap in `try/catch` that only
calls `Dts.Log(...)` and continues. A malformed row, or a failed insert, produces a log line and a silent
partial load — the task still reports Success.

### 5.2 `TravelAndExpense_Distribution` — fan-out

The largest package (6 MB of XML, 26 package parameters, 9 Script Tasks, 7 data flows).

**Top-level control flow**

```
GetConnectString ──▶ Get Batch ID's to Process ──▶ Foreach Batch ID
```

`Get Batch ID's to Process` calls `Concur_SAE.dbo.p_get_concur_batches_to_process`, loads the result into a
`DataTable`, and assigns it to `User::batch_id_list` (Object). The Foreach container is a
`ForEachADOEnumerator` over rows of the first table, mapping:

| Index | Variable |
|---|---|
| 0 | `User::batch_id` |
| 1 | `User::concur_extract_file` |

**Inside the loop** — the order is fixed by precedence constraints:

```
Populate table Concur Extract Data
   ├─ batch_count == 0 ─▶ Batch Times Insert Empty ─▶ File Does Not Exists Notification
   └─ batch_count  > 0 ─▶ Batch Times Insert Occurence ─▶ Archive Extract Data
                              └─▶ Mexico ─▶ Maenner ─▶ JBA_Aurora ─▶ JBA_Raymond
                                     ─▶ M3_Raymond ─▶ Synventive ─▶ Wrap Up
```

Every arrow after `Archive Extract Data` is an **on-Completion** constraint (`Value=2`), so one division
failing does **not** stop the others.

| Task | SQL |
|---|---|
| Populate table Concur Extract Data | `execute concur_sae.dbo.p_concur_extract_data_insert ?` → `batch_count` |
| Batch Times Insert Empty / Occurence | `execute concur_sae.dbo.p_concur_extract_batch_times_insert ?,?` |
| Archive Extract Data | `execute concur_sae.dbo.p_concur_extact_data_archive ?` *(note the typo `extact` in the proc name — it is spelled that way in the package)* |

#### The per-division pattern

Each division sequence repeats the same three-step shape:

1. **`<Div> Extract Distribution Status`** — `p_concur_extract_distribution_process_status ?, '<Div>'`,
   result column `complete` → `User::complete`. The next step only runs when `@[User::complete] == "N"`,
   which makes each division **individually idempotent** across reruns of the same batch.
2. **The extract** — a data flow.
3. **`Log <Div> Extract`** — `concur_sae.dbo.p_concur_extract_distribution_log_upsert ?, '<Div>', 'Y'`,
   marking that division done for the batch.

#### Division detail

**Mexico** → AS/400 Aurora (`PRDMODSF4`)

| Data flow | Source | Destination |
|---|---|---|
| Extract Mexico T&E Headers to PLP91U | `concur_sae.dbo.p_concur_extract_data_jba_mexico_invoice_header ?` | `$Package::PLP91U` |
| Extract Mexico T&E Details to PLP92U | `concur_sae.dbo.p_concur_extract_data_jba_mexico_invoice_details ?` | `$Package::PLP92U` |

Headers → details, in that order. Mexico writes to the **invoice** files (PLP91/92), not the T&E files
(PLP81) the JBA divisions use.

**Maenner** → SAP, via five per-division text files

- Data flow `Extract Meanner Data to File Server` *(sic)*: Flat File Source `FlatFile_Concur_Extract`
  (the raw SAE file, 400 columns, UTF-8, `|`-delimited, 1 header row skipped) → Conditional Split:

  | Output | Condition | Destination file |
  |---|---|---|
  | Division 0400 | `[Column 0] == "DETAIL" && [Column 9] == "0400"` | `SAP_P11_BGA\Concur0400.txt` |
  | Division 0401 | `… [Column 9] == "0401"` | `SAP_P11_MADE\Concur0401.txt` |
  | Division 0404 | `… [Column 9] == "0404"` | `SAP_P11_MAUS\Concur0404.txt` |
  | Division 0501 | `… [Column 9] == "0501"` | `SAP_P11_FODE\Concur0501.txt` |
  | Division 0502 | `… ([Column 9] == "0502" \|\| [Column 9] == "AT01")` | `SAP_P11_FOCN\Concur0502.txt` |

  There is **no default output** — rows matching no division are dropped silently.

- Script Task **`Rename Maenner Files`**: for every variable named `FlatFile_Maenner_*`, if the file exists
  and is non-empty, it (a) copies the **first line of the original SAE extract** to the top of the file as a
  header row, then (b) renames it to `<name>_yyyyMMddHHmmss.txt`. Empty files are left in place under their
  base name and are not renamed — so a stale zero-byte `Concur0400.txt` will persist.

**JBA_Aurora** (`AS400_OLEDB_BI_PCSRVR2`) and **JBA_Raymond** (`AS400_OLEDB_BR_ASRDW`) — identical logic,
different library:

```
<Div> Extract Distribution Status          (complete == "N")
   └─▶ Check if Batch Already Exists       SELECT COUNT(1) … FROM <PLP81U> WHERE BTCH81 = <batch_id>
          ├─ jba_batch_count > 0 ─▶ Batch ID Already Exists Notification ─▶ Log <Div> Extract
          └─ jba_batch_count == 0 ─▶ Extract <Div> Data To AS400   (data flow)
                 └─▶ Get Count Of Records Inserted
                        ├─ extract == insert ─▶ Log <Div> Extract
                        └─ extract != insert ─▶ Batch Row Count Doesn't Match Notification
                                                   └─▶ Move Rows From PLP81U to PLP81EU
                                                          └─▶ Delete Rows from PLP81U
```

- The data flow is `exec concur_sae.dbo.p_concur_extract_data_jba_{aurora|raymond} ?` → **Row Count**
  (`User::jba_batch_extract_row_count`) → OLE DB Destination into `PLP81U`.
- **Row-count mismatch is self-healing:** rows are copied into the error file `PLP81EU` (25 columns:
  `BTCH81, BDAT81, CONO81, RCUR81, CNTY81, PNME81, SCUR81, CEXR81, RDIR81, ACCT81, TTYP81, JAMT81, ACM181,
  ACM281, CACT81, LACT81, FNAM81, LNAM81, CODE81, PDAT81, PTIM81, EDAT81, ETIM81, DOCR81, GLBT81`) and then
  deleted from `PLP81U`, leaving the batch clean for a rerun. **`Log <Div> Extract` is not reached on this
  path** — its constraints come only from the equal-count branch and from the "already exists" branch — so a
  mismatched batch stays marked incomplete and will be retried on the next run.
- All four SQL statements are built by `SqlStatementSource` expressions concatenating the library parameter
  and `@[User::batch_id]` — i.e. **string-concatenated SQL**, not parameterised. `batch_id` originates from
  a database proc, so injection risk is low, but a non-numeric value would produce invalid SQL.

**M3_Raymond** → CSV drop

- `execute p_concur_extract_data_m3_raymond ?` → Flat File Destination `FlatFile_M3_Raymond`.
- 19 columns, comma-delimited, `"` qualifier, header row written.
- Filename expression appends `MMddyyyyHHmmss` before `.csv`:
  `…\Concur910_088_MMddyyyyHHmmss.csv`

**Synventive** → up to 13 CSV drops

Source `execute p_concur_extract_data_synventive ?`, then a Conditional Split on `Allocation_Custom_1`:

| Value | Route |
|---|---|
| `110` | Multicast → live folder + archive folder, each through a further P-Card split |
| `120` | live + archive |
| `215` | live + archive (Mexico subfolder) |
| `610` | live only |
| `710` | Multicast → live + archive, each through a P-Card split |
| `0430` | Entity file, live only |

The nested `PCard or Not …` splits route on `Report_Policy_Name == "US P-Card Policy"`, sending P-Card rows
to `…PCard_.csv` and everything else to the plain file. Every Synventive filename carries the same
`MMddyyyyHHmmss` suffix expression.

Note the format drift across these destinations: 110/120/610/0430 use a 19-column layout, 710 uses a
20-column layout with different column names, and 215 uses a 37-column layout. `HeaderRowsToSkip` and
`ColumnNamesInFirstDataRow` are also inconsistent between siblings (e.g. `FlatFile_Synventive_120` skips 0,
`FlatFile_Synventive_110` skips 1). Treat each Synventive file as its own contract.

**Wrap Up**

```
Insert Batch Totals    exec p_concur_extract_insert_batch_totals ?
   └─▶ Distribution Summary Notification   (Script Task)
          └─▶ Mark Batch As Complete       exec p_concur_extract_complete_batch ?
```

`Distribution Summary Notification` calls `p_concur_extract_results_summary @batch_id`, which returns two
result sets — a header block (batch date, Concur batch id, our batch id, date imported, raw file name) and a
per-division breakdown (records, division, journal amount) — renders them as an HTML table, totals the
journal amounts, and POSTs to the notification service with subject
**"Concur Travel Expense Distribution Summary"** from `donotreply@barnesgrp.net`.

The whole method is wrapped in a `try/catch` that writes to `Debug.WriteLine` and then unconditionally sets
`TaskResult = Success`. **A failed summary email is invisible**, and because `Mark Batch As Complete` is
constrained on Completion, the batch is closed either way.

---

## 6. Invoice pipeline

### 6.1 `Invoice_Import` — the live invoice package

Parameters: `BI_PLP91U`, `BI_PLP92U`, `BR_PLP91U`, `BR_PLP92U`, `FlatFile_InvoiceHeader`,
`FlatFile_InvoiceDetail`, `Folder_InvoiceImport`.

**Control flow**

```
GetConnectString
   └─(Completion)─▶ Check If files exist
                       ├─(Failure)─▶ File Does Not Exists Notification ─(Completion)─▶ Fail Package
                       └─(Success)─▶ Import Invoice Files
                                        └─(Success)─▶ Update Records
                                               └─(Success)─▶ Invoice Posting
                                                      └─(Completion)─▶ Wrap Up
```

Note the first constraint is **on Completion, not Success** — if `GetConnectString` were to fail, the
package would still proceed with empty connection strings.

`Check If files exist` is a Script Task that returns `Failure` unless **both** `invoice_header.txt` and
`invoice_detail.txt` are present. `Fail Package` is `raiserror('Concur invoice files do not exist',16,1)`.

**Import Invoice Files** (sequence)

| Step | Detail |
|---|---|
| Truncate Invoice Tables | `execute p_archive_invoice_tables` then `truncate table Invoice_Header` / `truncate table Invoice_Detail` |
| Populate Invoice Header | Flat File Source → Derived Column → OLE DB Destination `[dbo].[Invoice_Header]` |
| Populate Invoice Detail | Flat File Source → Script Component → Derived Column → OLE DB Destination `[dbo].[Invoice_Detail]`, fast-load with `CHECK_CONSTRAINTS` |

Precedence from Truncate is on **Completion** (`Value=2`) — a failed truncate does not stop the load.

*Flat file layouts* (both pipe-delimited, `"` qualifier, CP1252, no header row):

- **Header, 21 fields:** `PCNO91, CONO91, VEND91, VSEQ91, PONO91, TYPE91, INDT91, INVN91, TINV91, CURC91,
  TFRT91, TTAX91, TFUL91, TMAT91, TOTH91, TFUT91, AINV91, PYMT91, DESC91, PTYP91, CIDN91`
- **Detail, 20 fields:** `PCNO92, CONO92, BOLN92, INVL92, PONO92, ExtraCol1, POLS92, QTYI92, IPRC92,
  IAMT92, QTYR92, ITEM92, ACCT92, DEPT92, CIDN92, ExtraCol2, ExtraCol3, POLN92, LineAmount, LineQuantity`

  Positions 6, 16 and 17 are placeholders (`ExtraCol1..3`) — the file has columns the load ignores.

*Derived Column expressions* — these are the only real transformations in the package:

| Output | Expression | Meaning |
|---|---|---|
| `PDAT91` / `PDAT92` | `"1" + yy + mm + dd` cast to numeric/int | AS/400 CYYMMDD date (century digit `1` = 2000s) |
| `PTIM91` / `PTIM92` | `HHmmss` cast to int | AS/400 time |
| `EDAT9x`, `ETIM9x` | `0` | Edit date/time, not yet edited |
| `APBT91`, `GLBT91` | `0` | AP / GL batch numbers, assigned later by JBA |
| `VCHR91` | `"0"` | Voucher number placeholder |
| `CODE91` / `CODE92` | `"10"` | Editing status code — "new/unedited" |
| `drcCIDN91` / `dcCIDN92` | `(DT_NUMERIC,7,0)(3000000 + CIDN9x)` | **Catalyst ID offset by 3,000,000** |
| `drcINVN91` | `REPLACE(INVN91," ","")` | Strip spaces from invoice number |
| `drcPONO9x` | `SUBSTRING(PONO9x,1,255)` | Truncate PO number |
| `drcDESC91` | `SUBSTRING(DESC91,1,50)` | Truncate description |
| `ITEM92_Trun` | `SUBSTRING(ITEM92,1,255)` | Truncate item |
| `dcRecptQty` | `QTYI92` | Copy of invoiced qty used as received qty |

The `3000000 +` offset on the Catalyst ID is the equivalent of the `+1700` batch offset in the T&E pipeline
— a namespace separator so Concur-originated records cannot collide with natively-keyed ones.

**Update Records** (sequence) — pure stored-procedure work, in this order:

```
Set Vendor              exec p_concur_invoice_header_set_vendor
   └─▶ Remove Leading Zeros    exec p_concur_invoice_header_removing_leading_zero
          └─▶ Update Surcharges and Set Po Numbers   [DISABLED]
                 └─▶ Update Concur Invoice Records   exec p_concur_invoice_update_records
```

`Update Surcharges and Set Po Numbers` (`p_concur_invoice_update_surcharges_set_po_numbers`) is **disabled**.
Because the constraint chain runs *through* it, and a disabled task reports Success, `Update Concur Invoice
Records` still runs. Worth confirming whether the surcharge logic was retired deliberately or parked.

**Invoice Posting** (sequence of sequences, each chained on Completion)

| Sub-sequence | Contents |
|---|---|
| JBA Aurora | two data flows **(both disabled)** + `Execute Package → InvoicePost_Aurora.dtsx` |
| JBA Raymond | two data flows **(both disabled)** + `Execute Package → InvoicePost_Raymond.dtsx` |
| M3_Raymond | `Execute Package → InvoicePost_ASRM3.dtsx` |

The disabled data flows are the previous in-line implementation, left in place after the logic was factored
out into the child packages. They duplicate exactly what the child packages now do.

**Wrap Up**

```
Create Count Batch    select count_batch = cast(datediff("n",'01/01/2007',getdate()) as int)  → User::count_batch
   └─▶ Save Counts               exec p_concur_invoice_save_import_counts ?
          └─▶ Distribution Summary Notification   (Script Task)
                 └─▶ Archive Files                (Script Task)
```

- `count_batch` is minutes since 2007-01-01 — a monotonically increasing run identifier. It overflows `int`
  around the year 6091, so that is not a concern.
- `Distribution Summary Notification` calls `p_concur_invoice_get_batch_counts` and emails the result.
- `Archive Files` moves every `*.txt` in `Folder_InvoiceImport` to `…\processed\`, deleting any same-named
  file already there.

### 6.2 `InvoiceImportStandard` — the parked variant

Structurally a superset of `Invoice_Import`, targeting the Concur **standard** extract format
(`extract_payment_gl_*.txt`) rather than the header/detail pair. **Almost all of it is disabled:**

Disabled containers: `Foreach Loop Container`, `Import Invoice Files`, `Invoice Posting` (and within it
`JBA Raymond`), `Update Records`, `Wrap Up`.

What still runs: `GetConnectString` → `Check If files exist` → (on failure) notification + `Fail Package`.
Since the only enabled downstream path is the failure path, **this package currently does nothing useful and
fails when the files are missing**.

Its distinguishing content, should it be revived:

- `Store Concur Raw Data` parses a **255-column** standard-extract layout (vs 400 for SAE), calls
  `p_insert_standard_raw_file`, and bulk-copies into **`Concur_Standard_Raw_Files_Data`**. Header parsing is
  defensive here (`firstLineCol.Length > n` guards) where the SAE version is not.
- `Invoice Posting` additionally references `InvoicePost_Syteline.dtsx` and `InvoicePost_Syteline9.01.dtsx`
  — **the missing packages from §2.1**.

### 6.3 `InvoicePost_Aurora` / `InvoicePost_Raymond`

Minimal packages. Identical shape, different AS/400:

```
GetConnectString ──▶ Sequence Container
                        Extract Header To PLP91U  ──▶  Extract Detail To PLP92U
```

| Package | Header source proc | Detail source proc | Destinations |
|---|---|---|---|
| `InvoicePost_Aurora` | `p_concur_invoice_header_jba_aurora` | `p_concur_invoice_detail_jba_aurora` | `$Package::BI_PLP91U`, `$Package::BI_PLP92U` on `AS400_OLEDB_BI_PCSRVR2` |
| `InvoicePost_Raymond` | `p_concur_invoice_header_jba_raymond` | `p_concur_invoice_detail_jba_raymond` | `$Package::BR_PLP91U`, `$Package::BR_PLP92U` on `AS400_OLEDB_BR_ASRDW` |

Both read from `Concur_Invoice_Worktables`. Destination access mode 1 = "table name from variable", so the
target library/file comes from the package parameter and therefore from the build configuration.

Neither package validates row counts or checks for an existing batch the way the T&E JBA divisions do.

### 6.4 `InvoicePost_ASRM3` — Infor M3 XML generation

Parameters: `ASRM3_Xml_Directory` (drop folder), `ASRM3_Xml_BackupDirectory`.
Variables: `CNN_Concur_Invoice_Worktables`, `CNN_SQLMI_ADO_DW`.
No SSIS connection managers, no data flows — one Script Task does everything.

```
GetConnectString ──▶ Create XML Document
```

**`Create XML Document`** (~650 lines of C#):

1. `GetAccountCodes()` — `App_Concur.dbo.p_concur_get_account_codes` on the **DW** connection → list of
   account codes that mark a line as "charge-only".
2. `p_get_concur_invoice_headers_m3_raymond` on the **work-tables** connection → one `Invoice` per row,
   carrying both the JBA-style fields (`PCNO91`, `CONO91`, `VEND91`, `PONO91`, `TINV91`, …) and M3-specific
   ones (`M3INVOICEDATE`, `M3CONO`, `M3DIVI`, `contact`, `minpoline`).
3. For each header, `GetInvoiceDetails()` — `p_get_concur_invoice_details_m3_raymond`
   `@parentCompany, @companyNumber, @poNumber, @catalystIDNumber` → the `PLP92`-shaped lines.
4. `SetChargers()` — for each line, `App_Concur.dbo.p_concur_get_po_chargers`
   `@po_no, @po_line, @item, @amount, @accountcode, @minpo_line`:
   - charges returned with `po_line == "0"` become **header** charges, and the originating line is flagged
     `delete_from_line_details`;
   - other charges become **line** charges; the originating line is flagged for deletion only if its
     account number is in the account-code list from step 1.
5. `CreateXMLDocument()` branches on `BatchType` (`PTYP91`):
   - `"M"` → **`MatchScenario`** (PO-matched invoice)
   - anything else → **`ExceptionDirectScenario`** (direct/GL-coded invoice)

**Output document** — Infor OAGIS `LoadSupplierInvoice` BOD:

- Namespace `http://schema.infor.com/InforOAGIS/2`,
  schemaLocation `…/2.13.0/InforOAGIS/BODs/LoadSupplierInvoice.xsd`
- `releaseID="9.2"`, `versionID="2.13.0"`, `systemEnvironmentCode="production"`, `languageCode="GB"`
- `ApplicationArea/Sender`: `LogicalID = lid://infor.m3.m3`, `ComponentID = M3BE`,
  `ConfirmationCode = OnError`
- `DataArea/Load/AccountingEntityID = {M3CONO}_{M3DIVI}`, action `Replace`

| | Match scenario (`M`) | Exception/Direct scenario |
|---|---|---|
| Lines emitted | `details.Where(delete_from_line_details == false)` | all details |
| Line content | `PurchaseOrderReference` (line no + PO no), `Quantity`, `UnitPrice`, `Item` | `ExtendedAmount` + `UserArea` accounting dimensions |
| Accounting dims | — | `Dim1 = AccountNumber`, `Dim2 = DepartmentCode`, `Dim3..Dim6 = ""`, **`Dim7 = "AP50200"` hard-coded** |
| UserArea props | `scenario = PTYP91`, `ext_ref_number = CIDN91` | `scenario = PTYP91`, `AccountingLine = "True"` |
| Header charges | emitted as `<Charge type= sequence=>` | not emitted |
| Line charges | emitted as extra `<SupplierInvoiceLine>` with `<InvoiceCharge>` | not emitted |

**File naming:** `{M3CONO}_{M3DIVI}_{invoiceNumber}_{PTYP91}_{yyyyMMddHHmmss}.xml`, where the invoice number
is stripped to letters and digits. Written to `ASRM3_Xml_Directory`, then copied to
`ASRM3_Xml_BackupDirectory`.

**Known defects in this task:**

- `new XDeclaration("1.0", "UTF-9", null)` — **`UTF-9` is not an encoding.** It appears in both scenario
  methods. Depending on how `XDocument.Save` resolves it, the emitted declaration is either wrong or the
  save throws (and the throw is swallowed — see next point).
- `MatchScenario` wraps everything in `try/catch { Debug.WriteLine }`; `ExceptionDirectScenario` has no
  try/catch at all. `GetInvoiceDetails` and the header read also swallow exceptions. A partial or missing
  XML file will not fail the task or the parent package.
- `BODID` is the **same hard-coded GUID** (`0a02464c-…`) on every document. If M3 or the middleware treats
  BODID as a de-duplication key, this is a live risk.
- `FileBackup` swallows its exception entirely (`catch (Exception) { }`), so a failed backup copy is silent.
- `GetAccountCodes` and `GetCharges` do `throw;` on error, which *will* fail the task — inconsistent with
  everything around them.

---

## 7. Notifications

All email goes through one REST endpoint (`$Project::EmailNotificationService`) via `HttpClient` POST of a
JSON `{ emailTo, subject, body, emailFrom }`, TLS 1.2 forced, from `donotreply@barnesgrp.net`.

| Script Task | Package(s) | Fires when |
|---|---|---|
| `File Does Not Exists Notification` | Incoming, Distribution, Invoice_Import, Standard | Expected input file absent |
| `Batch ID Already Exists Notification` | Distribution | `BTCH81` already present in the target AS/400 file |
| `Batch Row Count Doesn't Match Notification` | Distribution | Extract row count ≠ inserted row count |
| `Distribution Summary Notification` | Distribution, Invoice_Import, Standard | End of a successful run — HTML summary table |

Every one of these tasks sets `TaskResult = Success` regardless of the HTTP response. Notification delivery
is best-effort and unmonitored.

---

## 8. Stored procedures referenced

Grouped by database, as invoked from the packages.

**`Concur_SAE`**

| Procedure | Called from |
|---|---|
| `p_get_concur_batches_to_process` | Distribution — Get Batch ID's to Process |
| `p_insert_concur_raw_file` | Incoming — Store Concur Raw Data |
| `p_concur_extract_data_insert` | Distribution — Populate table Concur Extract Data |
| `p_concur_extract_batch_times_insert` | Distribution — Batch Times Insert Empty / Occurence |
| `p_concur_extact_data_archive` *(sic)* | Distribution — Archive Extract Data |
| `p_concur_extract_distribution_process_status` | Distribution — every division status gate |
| `p_concur_extract_distribution_log_upsert` | Distribution — every division log step |
| `p_concur_extract_data_jba_aurora` | Distribution — JBA_Aurora data flow |
| `p_concur_extract_data_jba_raymond` | Distribution — JBA_Raymond data flow |
| `p_concur_extract_data_jba_mexico_invoice_header` | Distribution — Mexico headers |
| `p_concur_extract_data_jba_mexico_invoice_details` | Distribution — Mexico details |
| `p_concur_extract_data_m3_raymond` | Distribution — M3 CSV |
| `p_concur_extract_data_synventive` | Distribution — Synventive CSVs |
| `p_concur_extract_insert_batch_totals` | Distribution — Wrap Up |
| `p_concur_extract_results_summary` | Distribution — summary email (2 result sets) |
| `p_concur_extract_complete_batch` | Distribution — Wrap Up |

**`Concur_Invoice_Worktables`**

| Procedure | Called from |
|---|---|
| `p_archive_invoice_tables` | Invoice_Import / Standard — Truncate Invoice Tables |
| `p_concur_invoice_header_set_vendor` | Update Records |
| `p_concur_invoice_header_removing_leading_zero` | Update Records |
| `p_concur_invoice_update_surcharges_set_po_numbers` | Update Records **(disabled)** |
| `p_concur_invoice_update_records` | Update Records |
| `p_concur_invoice_save_import_counts` | Wrap Up |
| `p_concur_invoice_get_batch_counts` | Wrap Up — summary email |
| `p_insert_standard_raw_file` | Standard — Store Concur Raw Data |
| `p_concur_invoice_header_jba_aurora` / `_detail_jba_aurora` | InvoicePost_Aurora |
| `p_concur_invoice_header_jba_raymond` / `_detail_jba_raymond` | InvoicePost_Raymond |
| `p_get_concur_invoice_headers_m3_raymond` | InvoicePost_ASRM3 |
| `p_get_concur_invoice_details_m3_raymond` | InvoicePost_ASRM3 |

**`App_Concur` (DW / SQL MI)**

| Procedure | Called from |
|---|---|
| `App_Concur.dbo.p_concur_get_account_codes` | InvoicePost_ASRM3 |
| `App_Concur.dbo.p_concur_get_po_chargers` | InvoicePost_ASRM3 |

**Tables written directly**

| Table | Writer | Method |
|---|---|---|
| `Concur_Raw_Files_Data` | Incoming — Store Concur Raw Data | `SqlBulkCopy` |
| `Concur_Standard_Raw_Files_Data` | Standard — Store Concur Raw Data | `SqlBulkCopy` |
| `[dbo].[Invoice_Header]` | Invoice_Import — Populate Invoice Header | OLE DB Destination |
| `[dbo].[Invoice_Detail]` | Invoice_Import — Populate Invoice Detail | OLE DB fast load |
| `PLP81U` / `PLP81EU` | Distribution — JBA divisions | OLE DB Destination / INSERT…SELECT |
| `PLP91U` / `PLP92U` | Mexico, InvoicePost_Aurora, InvoicePost_Raymond | OLE DB Destination |

---

## 9. Dead and orphaned objects

| Object | Package | Status |
|---|---|---|
| `FlatFile_Brazil_48` connection manager | Distribution | Declared with a 37-column layout and a `yymmdd` filename expression, **used by no component**. There is no Brazil division sequence. |
| `SQL_OLEDB_BA_SL_App` connection manager + `CNN_SQL_OLEDB_BA_SL_App` variable | Invoice_Import, Standard | Declared and populated by `GetConnectString`, **used by no task**. Left over from Syteline posting. |
| `Extract Header/Detail To PLP91/PLP92` data flows | Invoice_Import (JBA Aurora, JBA Raymond) | Disabled; superseded by the child packages. |
| `Update Surcharges and Set Po Numbers` | Invoice_Import, Standard | Disabled. |
| Entire body of `InvoiceImportStandard` | — | Disabled (see §6.2). |
| `InvoicePost_Syteline*.dtsx` (×4) | — | **Source missing** (see §2.1). |
| `varImportFile` | Standard | Only consumed by the disabled Foreach loop. |

---

## 10. Operational notes

**Entry points.** Schedule `TravelAndExpense_Incoming` and `Invoice_Import`. Do **not** schedule
`TravelAndExpense_Distribution`, `InvoicePost_*` — they are called as child packages. `InvoiceImportStandard`
should not be scheduled in its current state.

**Reruns.**

- *T&E:* safe. Each division is gated on `p_concur_extract_distribution_process_status` returning `N`, so a
  rerun of the same batch skips divisions already logged. The JBA divisions additionally check `PLP81U` for
  the batch id and self-clean on a row-count mismatch.
- *Invoice:* **not idempotent.** `Truncate Invoice Tables` archives then truncates unconditionally, and
  `Archive Files` moves the source files to `\processed`. A rerun after a partial failure needs the source
  files restored from `\processed` first.

**Where files land.**

| Flow | Input | Archive |
|---|---|---|
| T&E | `<FOLDER_ConcurSAE>\CONCURSAEEXTRACT*.TXT` | `<FOLDER_ConcurSAE>\save\` |
| Invoice | `<Folder_InvoiceImport>\invoice_{header,detail}.txt` | `<Folder_InvoiceImport>\processed\` |
| M3 XML | — | `ASRM3_Xml_Directory` (live) + `ASRM3_Xml_BackupDirectory` (copy) |

**Failure modes that do not raise an error.** This is the most important operational fact about the
project. The following all log-and-continue:

1. `GetConnectString` — a missing entry yields `""`, task still succeeds.
2. `Store Concur Raw Data` — row parse errors, DataTable fill errors, and the header insert are each caught
   and only `Dts.Log`ged. Partial loads are silent.
3. All four notification tasks — HTTP failures are swallowed.
4. `Distribution Summary Notification` (T&E) — the entire method is in a catch-all that still returns
   Success, and `Mark Batch As Complete` runs on Completion regardless.
5. `InvoicePost_ASRM3` — header read, detail read, `MatchScenario`, and `FileBackup` all swallow exceptions.
   A batch can produce zero XML files and report success.
6. Division-to-division constraints in `TravelAndExpense_Distribution` are all on **Completion**, by design —
   one ERP being down does not block the rest, but it also does not surface.

**What to watch instead of package status:** the `p_concur_extract_distribution_log_upsert` table (divisions
not marked `Y` for a batch), row counts in `PLP81U` vs the extract count, the `\save` and `\processed`
folders, and the presence/timestamp of the per-division output files.

**Logging.** Script Tasks call `Dts.Log(...)` with a consistent
`{PackageName} - Entring/Exiting {MethodName}` convention. No SSIS log providers are configured in the
packages, so this lands wherever the SSISDB/agent execution logging is set up.

---

## 11. Findings summary

Ordered roughly by risk.

1. **Four Syteline package sources are missing** while the project and deployed artifacts still reference
   them. The project cannot currently reproduce its own production build. (§2.1)
2. **Plaintext AS/400 credentials in the `.dtsx` files** as default variable values, unprotected by the
   package protection level. (§4.2)
3. **Pervasive swallowed exceptions** mean partial loads, missing XML files and failed notifications all
   report success. (§10)
4. **`UTF-9` encoding typo** in the M3 XML declaration, in both scenario methods. (§6.4)
5. **Hard-coded `BODID`** — the same GUID on every M3 document. (§6.4)
6. **Hard-coded `AccountingLine_Dim7 = "AP50200"`** in the M3 exception/direct scenario. (§6.4)
7. **Invoice pipeline is not rerunnable** without manually restoring files from `\processed`. (§10)
8. **String-concatenated SQL** against the AS/400 in the JBA division tasks. (§5.2)
9. **Maenner conditional split has no default output** — unmatched divisions are dropped silently. (§5.2)
10. **Synventive flat-file layouts are inconsistent** across siblings (19 / 20 / 37 columns, differing
    header settings). (§5.2)
11. **Dev and Stage configurations share the same AS/400 libraries.** (§3.2)
12. **`build.yaml` excludes `master`** from the trigger and has no deploy stage. (§3.3)
13. Dead objects: `FlatFile_Brazil_48`, `SQL_OLEDB_BA_SL_App`, four disabled data flows, and the
    almost-entirely-disabled `InvoiceImportStandard`. (§9)
14. Misspellings baked into names that cannot be fixed without coordinated changes:
    `p_concur_extact_data_archive`, `Extract Meanner Data to File Server`, `Extract Detal To Syteline`,
    `Log Raymond_Aurora Extract` (in the JBA_Raymond sequence).

---

## Appendix A — Task type census

| Package | Script Tasks | Execute SQL | Data Flows | Execute Package | Containers |
|---|---|---|---|---|---|
| `TravelAndExpense_Incoming` | 5 | 0 | 0 | 1 | 1 Foreach |
| `TravelAndExpense_Distribution` | 9 | 26 | 7 | 0 | 1 Foreach + 7 Sequence |
| `Invoice_Import` | 5 | 8 | 6 | 3 | 7 Sequence |
| `InvoiceImportStandard` | 7 | 8 | 10 | 5 | 1 Foreach + 9 Sequence |
| `InvoicePost_Aurora` | 1 | 0 | 2 | 0 | 1 Sequence |
| `InvoicePost_Raymond` | 1 | 0 | 2 | 0 | 1 Sequence |
| `InvoicePost_ASRM3` | 2 | 0 | 0 | 0 | 0 |

## Appendix B — Package dependency graph

```
TravelAndExpense_Incoming
   └── TravelAndExpense_Distribution

Invoice_Import
   ├── InvoicePost_Aurora
   ├── InvoicePost_Raymond
   └── InvoicePost_ASRM3

InvoiceImportStandard            (entire posting container disabled)
   ├── InvoicePost_Aurora
   ├── InvoicePost_Raymond
   ├── InvoicePost_ASRM3
   ├── InvoicePost_Syteline        ← source missing
   └── InvoicePost_Syteline9.01    ← source missing
```

## Appendix C — Where things live

| Concern | File / location |
|---|---|
| Connection strings | `ConnectionStringsFile` UNC `.config` — **not in this repo** |
| Environment overrides | `ConcurImport.dtproj` → `Configurations` |
| Project parameters | `ConcurImport/Project.params` |
| Build definition | `build.yaml` (Azure DevOps) |
| Business logic | SQL stored procedures in `Concur_SAE`, `Concur_Invoice_Worktables`, `App_Concur` — **not in this repo** |
| Parsing / XML / email logic | C# inside `<ScriptProject>` nodes in the `.dtsx` files |
| Last-built artifacts | `ConcurImport/bin/{Development,Stage,Production}/ConcurImport.ispac` |

> The stored procedures carry most of the actual business rules — account mapping, vendor resolution,
> division routing, batch selection. This document covers the orchestration layer only. A complete picture
> requires the `Concur_SAE`, `Concur_Invoice_Worktables`, and `App_Concur` database sources.
