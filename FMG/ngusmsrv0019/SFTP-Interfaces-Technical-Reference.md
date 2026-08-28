# SFTP Interfaces — Technical Reference

**Source path:** `FMG/sftp`
**Runtime host:** `NGUSMSRV0019` (utility server), Windows Server, drive `E:\`
**Runtime:** Windows PowerShell + WinSCP .NET assembly (`C:\Program Files (x86)\WinSCP\WinSCPnet.dll`) + GnuPG (`gpg`)
**Documented from source and logs as of:** 2026-08-21

> **Secrets:** this folder contains private keys, key passphrases and one hard-coded password. None of that
> material is reproduced in this document. Where a secret exists it is named and located, not quoted.
> See §8.

---

## 1. What this is

`FMG/sftp` is a working copy of `E:\` on **NGUSMSRV0019**, the utility server that runs Barnes Group's
scheduled file-transfer jobs. It holds **four interfaces**, each moving files between an external partner's
SFTP server and internal ERP file shares:

| Interface | Direction | Partner | Business function | Status |
|---|---|---|---|---|
| **HSBC** | Bidirectional | HSBC (`ecom-sftp.*.hsbc.com`) | **AP payments out** (ACH / check / wire), acknowledgements back | **Live** — runs every 3–10 min |
| **HighRadius** | Inbound | HighRadius (`sftp-cloud13.highradius.com`) | **AR cash application** — remittance advice into M3 and JBA | Configured; last stage run 2026-08-12 |
| **Concur** | Inbound | SAP Concur (`mft-us2.concursolutions.com`) | Pull Concur extract files for the ConcurImport SSIS project | Stage only; last run 2026-08-17, no files |
| **ConcurExtracts** | — | — | Placeholder for a Concur **Purchase Order** extract | **Empty scaffolding** — no scripts, no config, no data |

Three of the four connect to a partner SFTP endpoint via WinSCP. HSBC is the only one that also *sends*,
and the only one that uses PGP.

---

## 2. Folder layout and the two-layer pattern

```
E:\  (= FMG/sftp)
├── Applications\          per-interface working directories, config, orchestrator scripts
│   ├── ConcurImport\
│   ├── ConcurExtracts\
│   ├── HSBC\
│   └── HighRadius\
├── Scripts\WinScpFTP\     shared connection libraries + connection config
│   ├── Concur\
│   ├── HSBC\
│   └── HighRadius\
└── KeysEncryptions\       SSH private keys, PGP keys, passphrase files
    ├── Concur\
    ├── HSBC\
    └── HighRadius\
```

Every interface follows the same two-layer shape:

```
Task Scheduler
     │
     ▼
Applications\<App>\_scripts\<Script>.ps1        ── the ORCHESTRATOR
     │   • takes -Environment stage|production
     │   • reads Applications\<App>\configuration\<env>.json   (paths, routing)
     │   • Start-Transcript to a local .log
     │   • dot-sources ↓ by ABSOLUTE PATH
     ▼
Scripts\WinScpFTP\<App>\<App>Connect.ps1        ── the CONNECTION LIBRARY
         • reads Scripts\WinScpFTP\<App>\config\<env>\config.json  (host, user, key)
         • builds WinSCP.SessionOptions, opens session
         • writes a WinSCP session log alongside itself
```

**Two config files per interface per environment** — application paths live under `Applications\`,
connection credentials live under `Scripts\WinScpFTP\`. They are read independently by the two layers, so a
`-Environment` value must exist in *both* places or the run fails.

**Dot-sourcing is by hard-coded absolute path** (`. "E:\Scripts\WinScpFTP\HSBC\HSBC-US.ps1"`), inside the
function body, so the library is re-parsed on every call. The tree cannot be relocated off `E:\` without
editing every orchestrator.

---

## 3. Interface: HSBC (AP payments)

The only live, high-frequency interface. Three independent scheduled scripts.

### 3.1 The three scripts and their cadence

Derived from the PowerShell transcripts (`Start time:` markers), all running as **`NG\M0042`** on
**NGUSMSRV0019**:

| Script | Cadence | Runs recorded | First → last observed |
|---|---|---|---|
| `HSBC-US_prepare.ps1` | **every 3 min** | 31,288 | 2026-04-09 → 2026-08-21 |
| `HSBC-US_send.ps1` | **every 5 min** | 18,340 | 2026-04-16 → 2026-08-21 |
| `HSBC-US_get.ps1` | **every 10 min** | 9,883 | 2026-04-07 → 2026-08-21 |

### 3.2 End-to-end flow

```
    ERP (Raymond M3 / Hyson JBA)
            │  writes payment files
            ▼
  \\RAUSMSRV000{1,2}.RA.bg1857.net\ASR\AP-Bank\HSBC\out
            │
            │  [prepare]  every 3 min — Move-Item
            ▼
  E:\Applications\HSBC\<env>\                    ← staging / data directory
            │
            │  [send]  every 5 min
            │    1. gpg --encrypt → <name>_yyyyMMdd_HHmmss.pgp
            │    2. Move original → processed\<yyyy-MM-dd>\
            │    3. WinSCP PutFiles → HSBC "/"
            │    4. Move .pgp → processed\<yyyy-MM-dd>\
            ▼
      HSBC SFTP  (port 10022)
            │
            │  [get]  every 10 min — WinSCP GetFiles from "/Inbox/"
            ▼
  E:\Applications\HSBC\<env>\acknowledgments\
            │
            │    1. gpg -d → decrypted_<name>
            │    2. classify by filename → ERP folder
            │    3. locate the ORIGINAL payment file in processed\
            │    4. robocopy /MOV both to E:\shared\hsbc\<Erp>\<env>\<yyyy-MM-dd>\
            │    5. robocopy /MOV encrypted ack → acknowledgments\Decrypted\<yyyy-MM-dd>\
            ▼
  E:\shared\hsbc\{HysonJBA|ASRaymondM3|Unknown}\<env>\<date>\
```

### 3.3 `HSBC-US_prepare.ps1`

The simplest script. For each path in `App_ApOutFileLocations`, moves **every file** (no filter) into
`App_DataDirectory`.

| Environment | Source share |
|---|---|
| stage | `\\RAUSMSRV0002.RA.bg1857.net\ASR\AP-Bank\HSBC\out` |
| production | `\\RAUSMSRV0001.RA.bg1857.net\ASR\AP-Bank\HSBC\out` |

If the share is unreachable it prints `Path not accessible` and exits 0 — an unreachable ERP share is
indistinguishable from "no payments today".

### 3.4 `HSBC-US_send.ps1`

**`Encrypt-Files`** — picks up files with extension `.dat`, `.xml`, `.txt` or **empty** from the data
directory and encrypts each to `<BaseName>_<yyyyMMdd_HHmmss>.pgp`:

```
gpg --encrypt --recipient "cmbitconnectdigitalmappingvalidation@hsbc.co.in" --output <pgp> <file>
```

The recipient address is **hard-coded and identical for stage and production**. The original is moved to
`processed\<yyyy-MM-dd>\` only after a successful encrypt.

**`Send-Files`** — uploads every `*.pgp` in the data directory to remote path `/`, and on a `$true` result
moves it to `processed\<yyyy-MM-dd>\`.

A dead `old_main` function remains in the file, containing the earlier single-pass version of the same
logic. It is never called.

### 3.5 `HSBC-US_get.ps1` — the routing engine

This is where all the real logic lives. After downloading `/Inbox/` and decrypting each `.txt`/`.xml` with
`gpg -d`, each acknowledgement is classified:

| # | Test on `$file.Name` | ERP folder | How the original payment file is found |
|---|---|---|---|
| 1 | `-like "*JBA_*"` | `HysonJBA` | Regex `ACK[12]PSRV3\.PC\d+\.JBA_(?<type>[A-Z]+)_(?<id>[A-F0-9]+)\.` → search `processed\` recursively for `JBA_<type>_<id>*.xml` |
| 2 | `-like "*RaymondM3_*"` | `ASRaymondM3` | Regex `ACK[12]PSRV3\.PC\d+\.(?<erp>RaymondM3)_(?<ref>\d+)_(?<seq>\d+)\.` → search for `RaymondM3_*_<ref>_<seq>.xml` |
| 3 | `-match '^decrypted_ACK[12]STDK'` | by NACHA company id | reads the decrypted body, finds `PC000078445(\d{9})`, maps the 9-digit id → ERP, then scans the first 5 lines of candidate files for a **reversed company id with a leading 5** |
| 4 | anything else | `Unknown` | not attempted |

**The NACHA company-id mapping (branch 3):**

| Company id in ACK | ERP folder | Search company id (reversed + leading `5`) | File pattern |
|---|---|---|---|
| `797025162` | `HysonJBA` | `5261520797` | `JBA_ACH_*.DAT` |
| `797025154` | `ASRaymondM3` | `5451520797` | `RaymondM3_ACH_*.txt` |

The reversal trick reflects how the company id appears in the NACHA batch header of the outbound file.

Everything then moves with **`robocopy … /MOV /COPY:DAT /R:1 /W:1`** rather than `Move-Item` — the
`Move-Item` calls are still present but commented out. Robocopy exit codes `< 8` are treated as success
(correct — robocopy uses 0–7 for informational results).

### 3.6 Observed file naming

**Outbound payment files** (production, from the ERPs):

| Pattern | Payment type | Format |
|---|---|---|
| `JBA_ACH_<hex>.<yyyyMMddHHmmssfff>.DAT` | ACH | NACHA fixed-width |
| `JBA_CHECK_<hex>.<timestamp>.xml` | Check | ISO 20022 `pain.001` |
| `JBA_WIRE_<hex>.<timestamp>.xml` | Wire | ISO 20022 `pain.001` |
| `RaymondM3_ACH_<ref>_<yyyy-MM-dd_HH-mm-ss-fff>.txt` | ACH | NACHA |
| `RaymondM3_CH2_<ref>_<seq>_<ts>_<ts>.xml` | Check | ISO 20022 |
| `RaymondM3_WIR_<ref>_<seq>_<ts>_<ts>.xml` | Wire | ISO 20022 |

Stage archives show an older, retired naming convention (`BI.JBA.ACH.<hex>…`, `HSBC_ACH_pain001_<test>_manager_…`,
`ASR_ACH_…`) from the March–April 2026 onboarding period.

**Inbound acknowledgements** — all prefixed `ACK`, second segment is the HSBC customer id
(`PC000078445` production, `PC000031734` stage):

| Prefix | Count on disk | Handled by branch |
|---|---|---|
| `ACK1PSRV3` | 201 | 1 / 2 (when JBA_ or RaymondM3_ appears in the name) |
| `ACK1STDK` | 128 | **none — see §9.1** |
| `ACKMT999` | 113 | **none** |
| `ACK2PSRV3` | 98 | 1 / 2 |
| `ACK2STDK` | 37 | **none — see §9.1** |
| `ACK2INTPSR3` | 20 | **none** |

### 3.7 Volume

Production `processed\` holds dated folders from 2026-05-01 onward; the largest days are 2026-06-23 (338
files) and 2026-06-24 (304 files) — a backfill or reprocessing event. Typical days are 2–32 files. Each
payment appears twice (original + `.pgp`).

---

## 4. Interface: HighRadius (AR cash application)

Inbound only. One script, one library.

```
Task Scheduler
     ▼
Applications\HighRadius\_scripts\GetFiles.ps1 -Environment <env>
     │
     ├─ Get_Files  → Scripts\WinScpFTP\HighRadius\HighRadiusConnect.ps1
     │                 GetFiles -RemotePath $config.App_GetRemotePath
     │                          -SavePath   $config.App_DataDirectory
     │
     └─ Process_Files — split filename on "_", switch on token[1]:
            ├─ CHECK | EDI  → SendToFileShare -Erp token[2]
            │                    "M3" → Copy-Item to App_GetM3SaveLocation
            │                           then Move-Item to processed\<yyyy-MM-dd>\
            │                    else → Write-Error "Unknown Erp File System"
            ├─ AURORA       → SendToJBA
            │                    Copy-Item to App_GetJBASaveLocation
            │                    then Move-Item to processed\<yyyy-MM-dd>\
            └─ default      → Write-Error "Unknown File Type"
```

### 4.1 File naming and payloads

Observed in `stage\in\processed\2026-08-12\`:

| File | token[1] | token[2] | Route | Payload |
|---|---|---|---|---|
| `HRCAA-6305_CHECK_M3_20260806_0550_39_958.XML` | `CHECK` | `M3` | file share | Infor OAGIS `LoadM3EDICustomerRemittanceAdvice` (releaseID 9.2, versionID 2.10.0) |
| `HRCAA-6305_EDI_M3_20260806_0547_52_674.XML` | `EDI` | `M3` | file share | same BOD type |
| `HRCAA-7093_AURORA_BarnesEDI820_20260807_1012_44_921.EDI` | `AURORA` | — | JBA | **X12 820** payment order / remittance advice |
| `HRCAA-7093_AURORA_BarnesEDI823_20260811_0947_21_728.EDI` | `AURORA` | — | JBA | **X12 823** lockbox |

The X12 files name HSBC as the bank (`N1*BK*HSBC*13*…`) — so HighRadius is the receivables side of the same
banking relationship HSBC handles on the payables side.

### 4.2 Connection differences by environment

`HighRadiusConnect.ps1` branches on `$Environment`:

| | production | stage |
|---|---|---|
| Auth | **SSH private key** (`_ProdHighRadius.ppk`) + passphrase | **Password** (`App_KeyPassPhrase` used as `Password`) |
| Port | `10022` | WinSCP default (`22`) |
| Host | `sftp-cloud13.highradius.com` | `uatsftp-turtle13.highradius.com` |
| User | `BarnesGroupProd` | `UATBaresGroup` *(sic — "Bares", partner-side typo)* |
| Host key | `ssh-ed25519 255 …` | `ecdsa-sha2-nistp256 256 …` |

`GetFiles` in this library passes `-Remove $true` to `session.GetFiles` — **downloaded files are deleted
from the partner server**. (HSBC and Concur pass `$false` and leave them.)

`_scripts\GetFiles.txt` is a stub skeleton of the same script with an empty `Get_Files` function — an
abandoned scratch copy, not used.

---

## 5. Interface: Concur (inbound extracts)

```
Applications\ConcurImport\_scripts\ConcurGetFiles.ps1 -Environment <env>
     │  reads Applications\ConcurImport\<env>\configuration\concurimport.json
     │        → App_IncomingSavePath
     ▼
Scripts\WinScpFTP\Concur\ConcurConnect.ps1
     GetFiles → Get-SftpFiles, RemotePath hard-coded "/out/", port 22
```

Downloads every file from `/out/` on `mft-us2.concursolutions.com` into
`E:\Applications\ConcurImport\<env>`. There is **no post-processing** — the orchestrator's `Process_Files`
call is commented out. Downstream handling is the SSIS project's job (see §7).

`App_KeyPassPhrase` is read from config **in clear text** (unlike HSBC, which at least has an
encrypted-file mechanism, even if bypassed).

**Last run:** 2026-08-17 18:55. Five runs total in the transcript. The WinSCP session log shows the session
opened, `ls /out/` returned only `..`, and zero files were downloaded across all five runs.

---

## 6. Interface: ConcurExtracts

```
Applications\ConcurExtracts\
├── production\ConcurPurchaseOrder\    (empty)
└── stage\ConcurPurchaseOrder\         (empty)
```

Directories only. No `_scripts`, no `configuration`, no `Scripts\WinScpFTP\ConcurExtracts`, no key material,
no data, no logs. This is scaffolding for a planned **outbound** Concur Purchase Order extract that was never
built — or whose implementation lives somewhere outside this tree.

---

## 7. Relationship to the ConcurImport SSIS project

The `ConcurImport` interface here is the **file-delivery half** of the SSIS project documented separately in
`ConcurImport-Technical-Reference.md`. The two meet on disk:

| SSIS parameter (Stage config) | Directory in this tree |
|---|---|
| `TravelAndExpense_Incoming::FOLDER_ConcurSAE` = `\\ngusmsrv0019…\ConcurImport\stage\ConcurTravelExpense` | `Applications\ConcurImport\stage\ConcurTravelExpense\` |
| `Invoice_Import::Folder_InvoiceImport` = `…\ConcurImport\stage\ConcurInvoice\import` | `Applications\ConcurImport\stage\ConcurInvoice\import\` |
| `TravelAndExpense_Distribution::FlatFile_Maenner_*` = `…\stage\ConcurTravelExpense\OutServers\EDDES0010…\SAP_P11_*` | `…\OutServers\EDDES0010.ed.bg1857.net\SAP_P11_{BGA,MADE,MAUS,FODE,FOCN}\` |
| `TravelAndExpense_Distribution::FlatFile_Synventive_*` = `…\OutServers\SYUSS0001…\ConcurExchange[Log]\` | `…\OutServers\SYUSS0001.sy.bg1857.net\ConcurExchange{,Log}\{,Asia,Mexico}\` |
| `InvoicePost_ASRM3::ASRM3_Xml_BackupDirectory` (prod) = `…\production\ConcurInvoice\import\ASRM3XML` | `Applications\ConcurImport\production\ConcurInvoice\import\ASRM3XML\` |

Every one of these directories is **currently empty** — the folder structure is intact but no files are
staged. Combined with `/out/` on the Concur SFTP being empty and only five recorded runs of
`ConcurGetFiles.ps1`, the Concur SFTP pull looks like it was set up and tested in August 2026 but is not yet
on a schedule.

**Server roles for reference:**

| Server | Role |
|---|---|
| `NGUSMSRV0019` | Utility server — runs everything in this tree (`E:\`) and hosts the SSIS file-drop shares |
| `NGUSMSRV0030` | Production SQL |
| `NGUSMSRV0031` | Development SQL |

Note that the SSIS project's `ConnectionStringsFile` parameter points at `bgusmsrv0030` (production) and
`bgusmsrv0031` (dev/stage) — the `BG`-prefixed names for the same 0030/0031 pair.

---

## 8. Credentials and key material

### 8.1 Inventory

| Path | What it is |
|---|---|
| `KeysEncryptions\Concur\Dev\DevConcurSFTP_private.ppk` | **SSH private key** — Concur dev/stage |
| `KeysEncryptions\Concur\Dev\DevConcurSFTP_public` | matching public key |
| `KeysEncryptions\Concur\Dev\concur_exported_key_pair.key` | **exported PGP key pair** |
| `KeysEncryptions\Concur\Dev\concursolutions.asc` | Concur PGP key |
| `KeysEncryptions\Concur\Dev\keyphrase.txt` | 8 bytes — **plaintext passphrase** |
| `KeysEncryptions\HSBC\SSH\id_rsa_hsbcbank_private.ppk` | **SSH private key** — HSBC, used by *both* stage and production |
| `KeysEncryptions\HSBC\SSH\id_rsa_hsbcbank.pub` | matching public key |
| `KeysEncryptions\HSBC\hsbc_publickey.asc`, `hsbc_publickey_prod.asc` | Barnes PGP public keys registered with HSBC |
| `KeysEncryptions\HSBC\theirs\{,production\}` | HSBC's SFTP host keys and PGP public keys (dated 202408) |
| `KeysEncryptions\HighRadius\_ProdHighRadius.ppk` | **SSH private key** — HighRadius production |
| `Scripts\WinScpFTP\{HSBC,HighRadius}\config\<env>\passphrase.txt` | 1,054-byte hex blob — `ConvertTo-SecureString -Key`-encrypted |
| `Scripts\WinScpFTP\{HSBC,HighRadius}\config\<env>\passphrase2.txt` | 362-byte hex blob — same mechanism |
| `Applications\HSBC\configuration\passphrase.txt` | identical to the `passphrase.txt` above |

### 8.2 Connection configuration reference

| Interface | Env | Host | Port | User | Auth |
|---|---|---|---|---|---|
| Concur | stage | `mft-us2.concursolutions.com` | 22 | `t1209783x4lw` | key + plaintext passphrase in config |
| Concur | production | **— no config file —** | | | |
| HSBC | stage | `ecom-sftp.fguk-pprd2.hsbc.com` | 10022 | `CT000032270_36098` | key + passphrase via `Get-SecretValue` |
| HSBC | production | `ecom-sftp.fguk-prd2.hsbc.com` | 10022 | `CT000011501_88233` | key + passphrase via `Get-SecretValue` |
| HighRadius | stage | `uatsftp-turtle13.highradius.com` | 22 | `UATBaresGroup` | **password** in config |
| HighRadius | production | `sftp-cloud13.highradius.com` | 10022 | `BarnesGroupProd` | key + passphrase in config |

Host key fingerprints are pinned in every config (`App_SSHKey`) — good practice, and correctly done.

### 8.3 Security findings

Ordered by severity. **No secret values are reproduced here.**

1. **`Get-SecretValue` in `HSBC-US.ps1` returns a hard-coded plaintext password.** The function accepts
   `$PhraseLocation` and `$FileKey`, but the two lines that decrypt the passphrase file are commented out
   and the function ends with a literal string assignment. Both parameters are ignored. This password is the
   HSBC SSH key passphrase and it is in the script in clear text.

2. **The encrypted passphrase mechanism, even when re-enabled, is weak.** `File_Key` in the HSBC configs is
   `[1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16]` — a sequential AES key stored next to the ciphertext.
   Anyone with read access to the config directory can decrypt `passphrase*.txt`.

3. **Plaintext credentials in JSON configs.** Concur stage (`App_KeyPassPhrase`), HighRadius stage
   (`App_KeyPassPhrase` used as an SFTP password) and HighRadius production (`App_KeyPassPhrase`) all store
   their secret directly in `config.json`.

4. **Private keys are committed alongside the code.** Four `.ppk` files, one exported PGP key pair and one
   plaintext `keyphrase.txt` live in `KeysEncryptions\`. If this tree is ever in source control or backed up
   off-host, the keys travel with it.

5. **The same HSBC private key serves stage and production.** `App_KeyPath` is
   `id_rsa_hsbcbank_private.ppk` in both configs, even though the usernames and hosts differ.

6. **`passphrase.txt` is byte-identical across HSBC stage, HSBC production, HighRadius stage, HighRadius
   production and `Applications\HSBC\configuration\`** (verified by hash). `passphrase2.txt` differs
   stage-vs-production but is shared between HSBC and HighRadius. One secret protects several relationships.

7. **HighRadius production `App_KeyPassPhrase` is a 16-character prefix of its own host key fingerprint.**
   Whether deliberate or a copy-paste error, it is not a suitable passphrase.

**Recommended remediation order:** rotate the HSBC key passphrase and remove the hard-coded value → move all
secrets to Windows Credential Manager, DPAPI (`ConvertTo-SecureString` with no `-Key`, machine-scoped), or a
vault → separate stage and production key material → move `KeysEncryptions\` out of any path that gets
copied or committed.

---

## 9. Defects

### 9.1 The NACHA acknowledgement branch can never fire

In `HSBC-US_get.ps1` the loop variable `$file` is the **encrypted downloaded file**; `$outputFile` is the
decrypted one:

```powershell
$outputFile = Join-Path $file.DirectoryName ("decrypted_" + $file.BaseName + $_extension)
gpg … -o $outputFile -d $file.FullName
…
elseif ($file.Name -match '^decrypted_ACK[12]STDK') {   # ← tests $file, not $outputFile
```

`$file.Name` is `ACK1STDK.J…PC….TXT` — it never starts with `decrypted_`. Branches 1 and 2 use `-like
"*JBA_*"` / `"*RaymondM3_*"`, which match on a substring and so are unaffected; only this branch is anchored
with `^decrypted_`.

**Confirmed in the logs:** 30 `ACK1STDK` and 25 `ACK2STDK` files hit `WARNING: Unknown ERP type … Moving to
fallback`. Every NACHA (ACH) acknowledgement is landing in `E:\shared\hsbc\Unknown\<env>\<date>\` instead of
`HysonJBA` or `ASRaymondM3`, and the matching outbound ACH file is never moved out of `processed\`.

**Fix:** test `(Split-Path $outputFile -Leaf)` — or simply `$file.Name -match '^ACK[12]STDK'`.

### 9.2 Two acknowledgement families have no branch at all

`ACKMT999` (SWIFT MT999 free-format acknowledgements, 113 files) and `ACK2INTPSR3` (20 files) match none of
the three tests. Log evidence: **112 `ACKMT999` and 7 `ACK2INTPSR3` warnings.** A third payment stream whose
files are named `KM<mmddyyyy>-<ts>` also falls through (~15 warnings across `ACK1PSRV3`/`ACK2PSRV3`).

Totals across the retained log: **227 `Unknown ERP type` warnings** and **232 `Original file not found`
warnings**.

### 9.3 HighRadius production config is broken

`Applications\HighRadius\configuration\production.json`:

```json
"App_ProcessedPath": "E:\\Applications\HighRadius\\production\\processed\\",
                                        ↑ single backslash — invalid JSON escape
"App_GetSaveLocation": "\\\\RAUSMSRV0001.RA.bg1857.net\\ASR\\AR-Bank\\HSBC\\in"
```

Three problems in one file:

- **`\H` is not a valid JSON escape.** Depending on the PowerShell edition, `ConvertFrom-Json` either throws
  or silently drops the backslash, yielding `E:\ApplicationsHighRadius\production\processed\`.
- **`App_GetM3SaveLocation` and `App_GetJBASaveLocation` are missing.** `Process_Files` reads both. In
  production, `Copy-Item -Destination $null` would fail for every file.
- **`App_GetSaveLocation`** is a key no script reads — apparently a rename that was applied to production
  and not to stage, or vice versa.
- `App_DataDirectory` is `…\production\` where stage is `…\stage\in\`, despite `production\in\` and
  `production\out\` both existing on disk.

**HighRadius has almost certainly never run successfully in production.** The only processed files on disk
are from a stage run on 2026-08-12.

### 9.4 Concur has no production configuration

`Scripts\WinScpFTP\Concur\config\production\` exists but is **empty**. `ConcurGetFiles.ps1 -Environment
production` passes its own `-ValidateSet` check and finds `Applications\ConcurImport\production\configuration\concurimport.json`
fine, then fails inside `GetFiles` when `Get-Content` cannot find the connection config.

### 9.5 Stale dead code

| Location | Issue |
|---|---|
| `HSBC-US.ps1` top of file | A module-scope `$sessionOptions` block with the **pre-prod** host and user hard-coded, referencing `E:\Keys_Encryptions\HSBC\SSH\hsbc_private_key.ppk` — a folder name (`Keys_Encryptions`) and file name that do not exist. It is never used (each function builds its own options) but executes on every dot-source and is actively misleading. |
| `HSBC-US_get.ps1` | `old-main` — the previous implementation, never called. |
| `HSBC-US_send.ps1` | `old_main` — same. |
| `HighRadius\_scripts\GetFiles.txt` | Empty-bodied copy of `GetFiles.ps1`. |
| Both HSBC scripts | Commented-out `Move-Item` calls next to their `robocopy` replacements. |

### 9.6 Historical errors now resolved

Worth knowing so they are not re-reported:

- **396 `Decryption failed` messages** in `get_log.log` are **75 distinct files, all from 2026-04-16**, each
  retried ~9 times before being cleared manually. All carry the stage customer id `PC000031734`. No current
  failures; the acknowledgements inbox is empty.
- **`Get-SecretValue is not recognized`** (2 occurrences, HighRadius) — an older `HighRadiusConnect.ps1`
  called it; the current version reads `App_KeyPassPhrase` directly.
- **`Missing ')' in function parameter list` at `ConcurConnect.ps1:73`** — the `$Environment` parameter was
  added without a comma. Fixed; the last runs connect cleanly.

### 9.7 Smaller items

- `HSBC-US_prepare.ps1` moves **every** file from the ERP out-share with no extension filter, including any
  partially-written file the ERP is still creating. There is no stability check (size-unchanged, lock test).
- `Encrypt-Files` includes files with an **empty extension**; `stage\processed\2026-04-02\230326_ACH` is an
  example that was encrypted and sent.
- All three HSBC scripts share a single append-mode transcript each, never rotated:
  `prepare.log` is **30 MB**, `hsbc-us-get.log` (WinSCP session log) is **99 MB**, `get_log.log` 11 MB,
  `send_log.log` 16 MB. At current growth the WinSCP log alone adds ~1 MB/day.
- The PGP recipient in `HSBC-US_send.ps1` is the same for stage and production, so **stage traffic is
  encrypted to the production HSBC key**.
- `Send-Files` catches its exception with `Write-Host` (not `Write-Error`), so a failed upload does not
  surface as an error in the transcript.
- Neither `SendToFileShare` nor `SendToJBA` in HighRadius verifies the copy before `Move-Item`s the source to
  `processed\` — though `-ErrorAction Stop` on `Copy-Item` covers the common case.

---

## 10. Operations

### 10.1 Running a script by hand

```powershell
E:\Applications\HSBC\_scripts\HSBC-US_prepare.ps1  -Environment production
E:\Applications\HSBC\_scripts\HSBC-US_send.ps1     -Environment production
E:\Applications\HSBC\_scripts\HSBC-US_get.ps1      -Environment production
E:\Applications\HighRadius\_scripts\GetFiles.ps1   -Environment stage
E:\Applications\ConcurImport\_scripts\ConcurGetFiles.ps1 -Environment stage
```

`-Environment` is mandatory and validated against `stage|production`.

### 10.2 Where to look when something is wrong

| Symptom | First place to look |
|---|---|
| Payments not reaching HSBC | `Applications\HSBC\_scripts\send_log.log` (tail), then whether `.pgp` files are piling up in `Applications\HSBC\<env>\` |
| Payments not leaving the ERP | `prepare.log` for `Path not accessible`; then the `\\RAUSMSRV000x\ASR\AP-Bank\HSBC\out` share |
| Acknowledgements not reaching the ERP folders | `get_log.log` for `Unknown ERP type` / `Original file not found`; then `E:\shared\hsbc\Unknown\` |
| SFTP connection problems | the WinSCP session logs: `Scripts\WinScpFTP\<App>\<app>-{get,send}.log` — full protocol traces |
| Concur extracts missing | `Applications\ConcurImport\_scripts\concurget.log`; note there is no production config (§9.4) |

### 10.3 Idempotency and reruns

| Interface | Safe to rerun? |
|---|---|
| HSBC prepare | Yes — moves whatever is present. |
| HSBC send | Yes — operates on whatever is in the data directory; already-sent files have been moved to `processed\`. |
| HSBC get | Yes — a failed decrypt leaves the encrypted file in `acknowledgments\` for the next run (explicit `continue`, not `throw`). |
| HighRadius get | **No.** The library passes `-Remove $true`; a file downloaded and then lost before processing cannot be re-fetched from the partner. |
| Concur get | Yes — files are left on the partner server (`-Remove $false`). |

### 10.4 Scheduled task inventory (inferred)

Not present in this tree — Task Scheduler definitions live on NGUSMSRV0019. From the transcripts:

| Task | Account | Interval | Evidence |
|---|---|---|---|
| HSBC prepare, production | `NG\M0042` | 3 min | 31,288 starts, continuous since 2026-04-09 |
| HSBC send, production | `NG\M0042` | 5 min | 18,340 starts since 2026-04-16 |
| HSBC get, production | `NG\M0042` | 10 min | 9,883 starts since 2026-04-07 |
| HighRadius get | `BG\AKAFG` | — | 2026-08-12 only, manual |
| Concur get | `BG\AKAFG` | — | 5 runs on 2026-08-17, manual |

Only the HSBC trio is actually scheduled. `NG\M0042` is the service account; `BG\AKAFG` is a person running
things interactively.

---

## 11. Findings summary

| # | Finding | Where |
|---|---|---|
| 1 | Hard-coded plaintext SSH key passphrase in `Get-SecretValue`; both parameters ignored | §8.3.1 |
| 2 | NACHA acknowledgement branch tests the wrong variable — **every ACH ack is misfiled** (55 confirmed) | §9.1 |
| 3 | HighRadius production config has invalid JSON and is missing both destination keys — production has never worked | §9.3 |
| 4 | `File_Key` is `[1..16]`; the "encrypted" passphrase files are trivially decryptable | §8.3.2 |
| 5 | Private keys, PGP key pair and a plaintext passphrase committed alongside the code | §8.3.4 |
| 6 | `ACKMT999`, `ACK2INTPSR3` and the `KM*` stream are unhandled — 227 total `Unknown ERP type` warnings | §9.2 |
| 7 | One HSBC private key and one `passphrase.txt` shared across stage, production and two partners | §8.3.5–6 |
| 8 | Concur has no production connection config | §9.4 |
| 9 | Plaintext passwords in three `config.json` files | §8.3.3 |
| 10 | Stage traffic is PGP-encrypted to the production HSBC recipient | §9.7 |
| 11 | `prepare` moves files with no filter and no write-completion check | §9.7 |
| 12 | Logs never rotate — 99 MB WinSCP log, 30 MB transcript, growing ~1 MB/day | §9.7 |
| 13 | Misleading dead code: stale module-scope session options with a non-existent key path, plus three unused functions | §9.5 |
| 14 | HighRadius deletes files from the partner on download — no re-fetch possible | §10.3 |
| 15 | `ConcurExtracts` is empty scaffolding | §6 |

---

## Appendix A — Configuration key reference

**`Applications\<App>\configuration\<env>.json`** (application layer)

| Key | Used by | Meaning |
|---|---|---|
| `App_Environment` | — | Label only; not read by any script |
| `App_DataDirectory` | HSBC, HighRadius | Working directory for files in flight |
| `App_ProcessedPath` | HSBC, HighRadius | Archive root; scripts append `\<yyyy-MM-dd>\` |
| `App_AcknowledgmentPath` | HSBC get | Download target for `/Inbox/` |
| `App_AcknowledgmentDecryptedPath` | HSBC get | Archive root for consumed encrypted acks |
| `App_KeyPassPhraseLocation` | HSBC get | Path to `passphrase.txt` — **passed to `Get-SecretValue`, which ignores it** |
| `App_ApOutFileLocations` | HSBC prepare | Array of ERP out-shares to drain |
| `App_GetRemotePath` | HighRadius | Remote directory to list and download |
| `App_GetM3SaveLocation` | HighRadius | Destination for `CHECK`/`EDI` + `M3` files — **absent in production** |
| `App_GetJBASaveLocation` | HighRadius | Destination for `AURORA` files — **absent in production** |
| `App_GetSaveLocation` | — | Present in HighRadius production; **read by nothing** |
| `App_IncomingSavePath` | Concur | Download target |

**`Scripts\WinScpFTP\<App>\config\<env>\config.json`** (connection layer)

| Key | Meaning |
|---|---|
| `App_HostName` | SFTP host |
| `App_UserName` | SFTP user |
| `App_SSHKey` | Pinned host key fingerprint → `SshHostKeyFingerprint` |
| `App_KeyPath` | Path to the `.ppk` private key (empty string = password auth) |
| `App_KeyPassPhrase` | Concur / HighRadius: the secret, in clear text |
| `App_KeyPassPhraseLocation` | HSBC: path to the encrypted passphrase file |
| `File_Key` | HSBC: 16-byte AES key for `ConvertTo-SecureString` — currently `[1..16]` |

## Appendix B — File and script index

| File | Lines | Purpose |
|---|---|---|
| `Applications\HSBC\_scripts\HSBC-US_prepare.ps1` | 45 | Drain ERP out-share → data directory |
| `Applications\HSBC\_scripts\HSBC-US_send.ps1` | 168 | PGP-encrypt and upload payments |
| `Applications\HSBC\_scripts\HSBC-US_get.ps1` | 365 | Download, decrypt, classify and route acknowledgements |
| `Scripts\WinScpFTP\HSBC\HSBC-US.ps1` | 212 | HSBC WinSCP send/get library + `Get-SecretValue` |
| `Applications\HighRadius\_scripts\GetFiles.ps1` | 140 | Download and route remittance advice |
| `Scripts\WinScpFTP\HighRadius\HighRadiusConnect.ps1` | 102 | HighRadius WinSCP get library |
| `Applications\ConcurImport\_scripts\ConcurGetFiles.ps1` | 39 | Download Concur extracts |
| `Scripts\WinScpFTP\Concur\ConcurConnect.ps1` | 90 | Concur WinSCP get library |
| `Applications\HighRadius\_scripts\GetFiles.txt` | 40 | Abandoned stub |

> This document covers the transfer layer. What the ERPs do with the delivered files — how JBA and M3
> consume remittance advice, how payments are generated into `AP-Bank\HSBC\out`, and what
> `E:\shared\hsbc\<Erp>\` feeds — is outside this tree.
