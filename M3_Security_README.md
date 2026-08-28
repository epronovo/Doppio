# M3_Security_* — Infor OS / M3 security capture

Captures the Infor OS security exports — users, security roles and functional
security roles — into `doppio.db`, lets you edit them in a browser, and writes
them back out in the inbound format so changes can be pushed into M3.

All routines are prefixed `M3_Security_` so they group together in the folder.

| File | Role |
|------|------|
| `M3_Security_Db.py` | Schema + connection helpers (`~/sqlite/doppio.db`) |
| `M3_Security_Import.py` | Parses and loads the CSV exports (kind detected from the header) |
| `M3_Security_Export.py` | Writes the CSVs back out in the original format |
| `M3_Security_App.py` | Flask front end — six tabs, editing, export, live M3 |
| `M3_Security_M3Api.py` | Live M3 calls over the ION REST API |
| `M3_Security_Bod.py` | Builds the SyncSecurityRoleMaster event sent when a role is deleted |
| `templates/M3_Security_Index.html` | The single-page UI |

## Quick start

```bash
pip install flask requests              # only external dependencies
python M3_Security_App.py               # http://127.0.0.1:5057
```

Drop the CSVs on the page. They are saved to `input/m3_security/`, parsed into
`~/sqlite/doppio.db`, and appear under the **IFS Users**, **IFS Security Roles**
and **IFS Functional Security Role** tabs. Drop all the functional-role files at
once — IFS caps that export at 75 rows, and they merge without duplicating.

Command line equivalents:

```bash
python M3_Security_Db.py                                    # create the tables
python M3_Security_Import.py "ExportFile_....csv" "Security Role_....csv" \
                             "Functional Security Role_....csv"
python M3_Security_Export.py --tenant ZFQP353QZYV89ZHG_TST  # changed records only
python M3_Security_Export.py --tenant ZFQP353QZYV89ZHG_TST --scope full
python M3_Security_Export.py --tenant ZFQP353QZYV89ZHG_TST --kind functional
```

Every routine takes `--db PATH`; the `M3_SECURITY_DB` environment variable works
too. Default is `~/sqlite/doppio.db`, matching `Sheet2Db.py` and `config.py`.

## The tenant

The tenant is folded into the users export file name and is stored on every row,
so several tenants can live in the database side by side:

```
ExportFile_DOPPIO_DEM_0ea904a0-a76b-46df-b69d-b099d609fa4c_639226844002985993.csv
           └─ tenant ─┘ └──────────── guid ───────────────┘ └──── .NET ticks ────┘
```

The Security Role and Functional Security Role exports have no tenant of their
own. Either add it to the file name, or just select the tenant in the front end
before dropping them:

```
Security Role_ZFQP353QZYV89ZHG_TST_08_17_2026_15_07_05.6295484.csv
Functional Security Role_ZFQP353QZYV89ZHG_TST_08_18_2026_16_45_50.6657692.csv
```

Both guid forms are accepted in the users name — dashed
(`0ea904a0-a76b-46df-...`, the way IFS writes it) or with the dashes stripped —
and whichever form came in is the one written back out.

The tenant is read by dropping the label IFS puts in front (`ExportFile`,
`Security Role`, `Functional Security Role`) and the timestamp: whatever is left
is the tenant. That works for `DOPPIO_DEM` as well as `ZFQP353QZYV89ZHG_TST` —
it does not depend on the tenant containing digits or looking unusual. For a
file with no recognisable label, the parser falls back to scanning for an
`<alphanumeric id>_<ENV>` token anywhere in the name.

**If the name carries no tenant, the tenant selected in the front end is used.**
That is the common case for the Security Role and Functional Security Role
exports, which IFS names without one — drop them and they load into whichever
tenant is picked in the header. The drop zone says which tenant that is, and the
result line marks the file `(selected)` so it is obvious where it went. A name
that *does* carry a tenant always wins, so a mixed drop still lands correctly.
With no tenant selected and none in the name, the load is refused rather than
guessed.

On the command line, `--tenant` overrides every file, and `--tenant-fallback`
does what the front end does — fills the gap only for files with no tenant of
their own.

One consequence worth knowing: an export normally reuses the name of the file it
came from, but if that name had no tenant in it a fresh name is generated
instead, so the file you send back is always self-describing.

File kind is decided from the header row, not the name — a renamed file still
lands in the right tables.

## Tables

| Table | Contents |
|-------|----------|
| `M3_Security_Users` | One row per user: identity, status, locale, `SecurityAccessProfiles`, tenant |
| `M3_Security_UserRoles` | The `SecurityRoleN` / `FunctionalSecurityRoleN` columns unpivoted to one row per role held, position preserved |
| `M3_Security_Roles` | Role master — `Name`, `Description` |
| `M3_Security_RoleAssignments` | Role → email assignments from the Security Role export |
| `M3_Security_Imports` | Audit trail of every file loaded, plus the file-name pieces needed to rebuild an export |
| `M3_Security_FunctionRoles` | SES400 function ↔ role authorisations, per company/division |
| `M3_Security_FunctionRoleStatus` | Per SES400 role, whether MNS405 has it |
| `M3_Security_M3RoleDefs` | MNS405 role definitions as M3 holds them |
| `M3_Security_M3Members` / `M3Users` / `M3Log` | Live M3 membership, the USID↔email map, and every write sent to M3 |
| `M3_Security_FunctionalRoles` | Functional security role master — `Name`, `Description` |
| `M3_Security_FunctionalRoleMembers` | The security roles inside each functional role |

Each load runs inside one transaction. The **Security Role** file is a full
refresh per tenant — re-loading it replaces that tenant's roles and assignments.
The other two merge.

**The users file merges on `PersonId`.** Drop a fresh export of users you already
hold and every field it fills is refreshed — name, title, status, UPN, last
login, the generic properties — while the user keeps its key, its `row_state`
and, above all, **its roles**. The `SecurityRoleN` / `FunctionalSecurityRoleN`
columns are read only for users this database has never seen, because role work
done here (a strip, a role added by hand) is pending local state that has not
reached IFS yet, and a file load must not undo it. Two more rules follow from how
much the real exports differ: a column the file leaves out, or carries empty, is
left as stored rather than blanked — clearing a field on purpose is a job for the
user editor; and a user held here but absent from the file is left alone, not
deleted. Use **Clear all** for a true from-scratch reload.

**The Functional Security Role file also merges.** IFS caps that
export at 75 rows per file, so a real capture arrives as several files. Each one
is added to what is already held, and a functional role or a role-inside-a-role
that is already there is skipped rather than duplicated — drop all the files at
once and the counts tell you what was new and what was a duplicate.

Every editable table carries `row_state`: `unchanged`, `modified`, `new` or
`deleted`. Deletes are soft — the row stays in the database and is simply left
out of the export, so you can restore it.

## What gets exported

**Exports carry only the records you changed.** A user is included when the
record was edited or added, or when the roles they hold were edited. A security
role is included when its name, description or membership changed.

Two rules keep a delta from being read as a revocation:

* a changed user is written with their **complete** role list, not just the roles
  that were touched;
* a changed role is written with its **complete** current membership.

So editing one member of a 15-person role produces a file containing that role
and all 16 of its members — and nothing else.

Records flagged for removal are **not exported at all**, in either scope. The
delete flag only keeps them out of the file; removals are handled outside this
tool.

`Export changes` is the default button on both tabs and shows the pending count;
it is disabled when there is nothing to send. `Full export` writes every record,
for re-baselining or handing over a snapshot. Files land in
`output/m3_security/changes/` and `output/m3_security/full/` so a delta is never
confused with a snapshot.

## Sorting and paging

Every list sorts by clicking a column heading: once for ascending, again for
descending, a third time back to the list's natural order. **Sorting happens in
SQL, not in the page**, so it orders the whole result set — page 3 of a
descending sort is the third page of that sort, not the third page re-shuffled.
An unknown sort key falls back to the default order rather than erroring, and
only declared columns ever reach the query.

The window is one screen tall: the header, tabs and toolbars stay put, the table
body is the only thing that scrolls, and the pager sits under it where it is
always reachable. Rows per page (25 to 500) is on the pager itself. The drop
zone is collapsed to a single line — click *file names* to see the naming
conventions again.

## Clearing a capture

Each IFS tab has a **Clear all** button that empties what that tab owns for the
selected tenant, so the next drop starts from scratch:

| Tab | Clears | Leaves |
|-----|--------|--------|
| IFS Users | `M3_Security_Users`, `M3_Security_UserRoles` | the role and functional-role masters |
| IFS Security Roles | `M3_Security_Roles`, `M3_Security_RoleAssignments` | the roles users hold, and the MNS405 capture |
| IFS Functional Security Role | `M3_Security_FunctionalRoles`, `M3_Security_FunctionalRoleMembers` | the `FunctionalSecurityRoleN` entries on the users |

Each one also drops that kind's rows from `M3_Security_Imports`, so the file
history and the "reuse the source file name" behaviour reset with it.

You get a preview of exactly what goes — a row count per table plus the import
history — and a warning if any of it was edited and never exported, because
those edits go too. Typing `CLEAR` applies it. Nothing is sent to M3, and the
live captures (SES400, MNS405, MNS410) are never touched.

Reasons to reach for it: the users and functional imports merge, so a wrong set
of files can only be undone this way; and a re-import refreshes the rows but
keeps a tenant you no longer want, or edits you would rather abandon.

The tenant picker is fed by every table that carries a tenant, so clearing one
capture never makes a tenant disappear while its other captures are still held.

### Marking a push as done

Once a delta has actually been pushed into M3, hit **Mark pushed** (or run the
export with `--mark-pushed`). That clears the change flags so the next delta
starts from a clean baseline, and purges the rows that were flagged for removal.
Until you do, the same records keep appearing in every delta.

## Export fidelity

Both scopes are written to match the inbound files exactly:

* **Users** — no BOM, CRLF, minimal quoting, 20 base columns then
  `SecurityRole1..1080`, `FunctionalSecurityRole1..46`, `SecurityAccessProfiles`.
  Column widths default to the widths of the file that was imported and grow
  automatically if someone is given more roles than the original file had slots for.
  **Tenants do not all export the same columns.** One that does not send
  `GenericProperty_LanguageOrigin` and friends is loaded anyway, those columns
  stored as null, and the export gives back exactly the columns that came in —
  the file's own header is recorded on the import so the round trip stays
  column-for-column. Only `PersonId` is genuinely required, since it is the key.
  The load reports how many columns were absent.
* **Security roles** — UTF-8 BOM, CRLF, `Name,Description,EmailId` header, and the
  trailing comma the source file puts on every data row.

Rows keep their original file order; new rows are appended in place.

A delta uses the same column layout as a full export — same 1,147 columns, same
header — so M3 sees an identical file shape either way.

Verified against the 17 Aug 2026 CHC exports: re-exporting an unedited load
with `--scope full`
reproduces both files byte for byte, apart from **35 duplicate rows the source
files contain** — 18 users had the same security role listed twice, and there
were 17 repeated `role,email` pairs (mostly `IFSApplicationAdmin`). These are
deduplicated on load, which is almost certainly what you want before pushing back
into M3. Every user's role *set* is identical, and re-importing an export and
exporting again is byte-for-byte stable.

## File naming on export

By default the front end generates a fresh name using the same convention —
the original guid is preserved and the `.NET` ticks / timestamp are regenerated,
so an export is never confused with the file it came from. Untick **fresh file
name** (or drop `--new-name` on the CLI) to reuse the imported name instead.

Exports land in `output/m3_security/` and download through the browser.

## Live M3

`M3_Security_M3Api.py` calls M3 over the ION REST API to answer two questions:
which of the captured role names actually exist in M3, and who holds them.

The `.ionapi` file is chosen automatically — every file in `ionapi/` is read and
the one whose `ti` equals the tenant wins. `ZFQP353QZYV89ZHG_TST` resolves to
`CHC_TST.ionapi` with no prompting, which is what lets the Flask app call it.
Authentication is the same service-account password grant `InforMI.py` uses.
Calls run with `cono=001&divi=100` by default, matching your working curl;
override with `--company` / `--division` on either the API module or the app.
Records are sent 500 to a call — see **Bulk calls and progress** below.

Programs used, with signatures taken from `claude.db` (`cmitrn` / `cmifld`):

| Call | Purpose |
|------|---------|
| `MNS405MI/Lst` | Every role definition in M3 → `ROLL, TX40, TX15, TXID, ROLT` |
| `MNS405MI/Get` | One role, used to confirm a role exists before touching it |
| `MNS405MI/Add` | Creates a role definition |
| `MNS410MI/Lst` | Role per user — passing `ROLL` alone lists everyone holding it |
| `MNS410MI/Add` | Connects one user to one role |
| `MNS410MI/Dlt` | Removes one user from one role |
| `MNS150MI/LstUserData` | `USID → EMAL` map, so M3 members line up with the CSV emails |
| `SES400MI/Lst` | Function authorisations → `FNID, ROLL, CONO, DIVI, STAT, TXID` |
| `SES400MI/Upd` | Changes one function authorisation (status) |
| `SES400MI/Dlt` | Removes one function authorisation |

### Company and division

`cono` and `divi` are **not** assumed. They used to be hard-coded to the
`001` / `100` of the original curl example, which is wrong on any tenant that
numbers its companies differently — the call comes back
`403 Forbidden` because the service account has no authority for a company that
is not its own.

Unset, they are left off the request entirely and M3 uses the service account's
own default company and division, from its MNS150 user record. That is the right
answer most of the time, and it is the default.

When a tenant does need a specific company, pin it: the **Company / division**
button beside the tenant picker shows what is being sent, and **Read from M3**
lists what the tenant actually has (`MNS095MI/Lst` for companies,
`MNS100MI/LstDivisions` for the divisions under each). Click a row, save, and
the choice is stored per tenant in `M3_Security_M3.json`:

```json
{
  "DOPPIO_DEM": {"cono": "100", "divi": ""},
  "_default":   {"cono": "",    "divi": ""}
}
```

A blank means "do not send it". Precedence is `--company` / `--division` on the
command line, then the file, then nothing.

```bash
# what does this tenant actually have?
python M3_Security_M3Api.py --tenant DOPPIO_DEM --companies
```

SES400 rows keep the `CONO` / `DIVI` they were read with, so updating or
deleting an authorisation always sends that row's own company rather than a
global one.

### Highlighting

**Check M3** on the Security Roles tab reads the role list, flags every captured
role that exists there, and pulls each matched role's membership and the user
map. Roles in M3 get a green marker, an `M3` pill, their M3 description and
their live member count; the rest are shown as `not in M3` — that is how the
Infor OS / IFS-only roles (`IFS-ReadOnlyUser`, `MingleEnterprise`, and so on)
separate themselves from real M3 roles. The dropdown filters to **In M3**,
**Not in M3** or **Not yet checked**.

Opening a role shows an M3 panel: its live members as `USID → email`, with
anyone who is in M3 but not in the capture highlighted, plus counts of who is
only on one side.

Check M3 also pre-selects the roles MNS405 does not have — see below.

### Bulk calls and progress

`m3api-rest` takes an array of transactions in one body, so every bulk
operation posts **up to 500 records per call** rather than one call per record
(`--batch-size` on both the API module and the app). Creating 928 roles is 2 HTTP
calls instead of 928; rebuilding membership on the CHC data was 3 calls instead
of 1,831. Check M3 batches its `MNS410MI/Lst` reads the same way, so it is 3
calls rather than one per role.

Batching is a transport detail, not a loss of detail. The response `results`
array lines up positionally with the transactions that went in, so every record
still gets its own outcome and its own audit row — successes and failures inside
the same batch are attributed correctly. A short `results` array is reported as
*No result returned* rather than counted as success, and a batch that fails to
post at all is attributed to each record in it. Audit rows are written one
`executemany` per batch, and each batch commits once.

The long operations run on a worker thread and the page polls them, so you get a
**progress bar** with the phase and record count instead of a frozen button. The
CLI prints the same progress as a bar. Previews stay synchronous — they only
read the database.

### Rebuilding M3 from the capture

Two buttons on the selection bar push the captured data back into M3. Both work
off the same role selection, both preview first, and both log every call.

**Create in M3** — `MNS405MI/Add` for each selected role:

| Field | Value |
|-------|-------|
| `ROLL` | the role name |
| `TX40` | the description, cut to 40 characters |
| `TX15` | the description, cut to 15 characters |

A role is skipped, with the reason shown, when the name will not fit M3's `ROLL`
field or is not a plain M3 role id:

* **longer than 10 characters** — `ROLL` is `A10`
* **contains a hyphen** — Infor OS style names, not M3 roles
* **starts with an asterisk** — reserved
* **already in M3** — nothing to create

A role with no description falls back to its own name, since `TX40` is mandatory.
Descriptions over 40 characters are flagged in the preview so you can see what
got cut. Type `CREATE` to run it.

On the CHC test data that works out to 928 of the 1,206 not-in-M3 roles created,
278 skipped — 156 for a hyphen, 120 for length, 2 for a leading asterisk.

The same call is available from the
[IFS Functional Security Role](#ifs-functional-security-role) tab, where the
selection is functional roles and it creates the security roles inside them.

**Add members in M3** — `MNS410MI/Add` for each captured member of the selected
roles. `ROLL` is the role name; `USID` comes from **matching the member's email
address against the M3 user list** that Check M3 pulls from `MNS150MI`. An
address M3 does not know is reported, never guessed at. Members are skipped when:

* the role is not in M3 yet — create it first
* no M3 user has that email address
* the user already holds the role in M3

Type `ADD` to run it. Re-running is safe: everything already in place is skipped
as *already holds the role*, so a second pass only retries what failed.

The natural rebuild order is **Check M3 → Create in M3 → Add members in M3**.

### Clearing the roles M3 no longer has

**Check M3** finishes by pre-selecting every captured role that MNS405 does not
know about — the stale IFS-only roles. Tick and untick whatever you like from
there: the selection is a checkbox per row, survives paging and filtering, and
**Select all not in M3** rebuilds it at any time. **Remove members from selected**
then previews the whole set — every role, whether it is in M3, how many members
and user-export links it would lose — and typing `REMOVE` runs it.

This is a **local** operation. Nothing is sent to M3, because roles M3 does not
have cannot be deleted there. What it does is strip the membership in the
capture, so the next changes export carries those roles with no members and
pushing that into Infor OS removes the assignments. The users export is updated
in step: everyone who held one of those roles loses it and lands in the users
delta.

Stripping the last member leaves a blank-`EmailId` row behind, which is exactly
how the inbound file represents a member-less role — so a cleared role still
appears in the export rather than disappearing from it.

If you include roles that *do* exist in M3, they are stripped locally like the
rest; the bulk button never issues `MNS410MI/Dlt`. Deleting members in M3 stays a
one-role-at-a-time operation through the role's M3 panel, below.

One gap worth knowing: the Security Roles tab is built from the Security Role
export, so a role that appears only in the users export (`AttributeServiceCaller`
in the CHC data) has no row there and cannot be selected. Edit those on the user
record.

### Removing all members of a role, in M3

The panel's **Preview removal** button re-reads the membership from M3 and lists
the exact USIDs a real run would delete. Nothing has been sent at that point.
To go through with it you type the role name into the confirm box; only then are
the `MNS410MI/Dlt` calls issued, one per member.

Every call — preview, success and failure — is written to `M3_Security_M3Log`
with its payload, response and a shared `run_id`. Partial failures are reported
per user and do not stop the rest. The local capture is only marked as cleared
when every delete succeeded, so a partial run never leaves the database claiming
something M3 did not do.

From the command line:

```bash
python M3_Security_M3Api.py --list-ionapi
python M3_Security_M3Api.py --tenant ZFQP353QZYV89ZHG_TST --probe
python M3_Security_M3Api.py --tenant ZFQP353QZYV89ZHG_TST --sync-roles --with-members
python M3_Security_M3Api.py --tenant ZFQP353QZYV89ZHG_TST --remove-members APDIVRW
python M3_Security_M3Api.py --tenant ZFQP353QZYV89ZHG_TST --remove-members APDIVRW --commit

# rebuild: create every captured role M3 does not have, then its members
python M3_Security_M3Api.py --tenant ZFQP353QZYV89ZHG_TST --create-roles
python M3_Security_M3Api.py --tenant ZFQP353QZYV89ZHG_TST --create-roles --commit
python M3_Security_M3Api.py --tenant ZFQP353QZYV89ZHG_TST --add-members
python M3_Security_M3Api.py --tenant ZFQP353QZYV89ZHG_TST --add-members --commit
```

`--create-roles` and `--add-members` take role names; with none they act on every
captured role not in M3 / every role in M3 respectively.

`--probe` authenticates and reads a handful of roles — run it first to confirm
the tenant, the .ionapi file and the company/division are right.

## The SES400 Function tab

A third tab, fed entirely from M3 rather than the CSVs. **Load from M3** reads
`SES400MI/Lst`, which returns one row per function + role authorisation, and
stores it in `M3_Security_FunctionRoles`. Only the identifying columns are
requested — SES400 also carries 99 option flags and 24 function-key flags per
row, which would make the response enormous and are not needed here.

The same pass reads `MNS405MI/Lst` and records, per role, whether MNS405 has it.
That means the tab does not depend on the Security Roles tab having been checked
first.

**By role** (the default) lists each role with how many functions it is
authorised to; clicking one shows those functions. **By function** lists each
function with how many roles hold it; clicking one shows those roles. Clicking a
row in the detail panel flips the view to the other side of the pairing, so you
can walk role → function → role without going back to the toolbar.

`CONO` and `DIVI` are stored per row and shown, with a filter for a specific
company/division pair. There is also a filter for roles missing from MNS405, and
a search box.

### Changing the status

`STAT` is `A2` and takes two values:

| Code | Meaning |
|------|---------|
| `10` | Preliminary |
| `20` | Active |

Pick one next to **Set status** and it applies to the current selection through
`SES400MI/Upd`. The preview shows every row with its current status and where it
would land; typing `UPDATE` runs it. Anything already at the chosen status is
listed as such and **not sent** — no point spending a call writing a value that
is already there. A status outside the two codes is rejected before anything
leaves the machine.

The status is also a column in the detail panel (Preliminary rows are
highlighted) and a filter in the toolbar, showing how many authorisations sit at
each value.

**Only the key fields and `STAT` are sent.** SES400 rows also carry 99 option
flags and 24 function-key flags; leaving them out of the `Upd` leaves them as
they are, so a status change does not disturb the permissions on the
authorisation. That is standard M3 `Upd` behaviour, but it is worth confirming
on one record the first time you run it against a real tenant.

Rows M3 accepts are updated locally; rows it refuses keep their old value and
are reported, so the tab stays honest about what actually changed.

From the command line:

```bash
python M3_Security_M3Api.py --tenant ZFQP353QZYV89ZHG_TST \
    --update-function-roles CRS610:ADMIN OIS100:APAUDRO --set-function-status 20
# add --commit to perform it
```

### Deleting authorisations

Checkboxes work at two levels, and both feed one selection:

* a tick on a **detail row** picks that single authorisation — one function,
  one role, one company/division;
* a tick on an **aggregate row** stands for every authorisation underneath it,
  so ticking a role selects all of its functions (the keys are fetched from the
  server, not guessed from the page).

**Select all in view** takes everything matching the current filter, not just
the page on screen. The selection survives paging and drill-down, and is cleared
when you flip between by-role and by-function, change tenant, or reload SES400 —
in each of those cases the keys it held would no longer mean what you picked.

**Delete selected from SES400** previews first: every row it would remove, with
`FNID`, `ROLL`, `CONO` and `DIVI`. Typing `DELETE` runs it, batched through
`SES400MI/Dlt` like every other bulk call. Each delete targets the exact row
that was read from M3, keyed on all four fields.

Rows M3 accepts are removed from the local copy; rows it refuses are **left in
place** and reported per row, so the tab keeps showing what is really still
there. Roles left with no authorisations at all drop out of the by-role view.
Only the authorisation is removed — the role itself stays in MNS405.

From the command line:

```bash
python M3_Security_M3Api.py --tenant ZFQP353QZYV89ZHG_TST \
    --delete-function-roles CRS610:GHOSTROLE PPS200:ORPHAN2:001:200
# add --commit to perform it
```

### Roles that SES400 references but MNS405 does not have

These are supposed to pre-exist and often do not. Every such role is flagged in
red on both views, counted in the header, and **Add missing roles to MNS405**
creates them with `MNS405MI/Add`. SES400 carries no description, so `TX40` and
`TX15` are both the role name (`TX15` cut to its 15-character field).

No name filtering is applied here, unlike the Security Roles tab: these names
already exist in M3 as valid role ids, and `ROLL` is `A10` on both sides so the
length rule cannot bite. Anything M3 refuses is reported per role rather than
skipped in advance. Preview first, then type `CREATE`; roles that succeed flip to
*MNS405* immediately, and the ones that fail stay flagged so you can see what M3
objected to.

From the command line:

```bash
python M3_Security_M3Api.py --tenant ZFQP353QZYV89ZHG_TST --sync-functions
python M3_Security_M3Api.py --tenant ZFQP353QZYV89ZHG_TST --add-missing-roles
python M3_Security_M3Api.py --tenant ZFQP353QZYV89ZHG_TST --add-missing-roles --commit
```

## The tabs

The six tabs:

| Tab | Source | Writes |
|-----|--------|--------|
| **IFS Users** | the users CSV | export only |
| **IFS Security Roles** | the Security Role CSV | export, plus M3 role/member calls |
| **IFS Functional Security Role** | the Functional Security Role CSV | export only |
| **SES400 Function** | M3 `SES400MI` | Upd (status), Dlt, plus MNS405 Add for missing roles |
| **MNS405 Roles** | M3 `MNS405MI` | Add, Upd, Dlt |
| **MNS410 Roles Per User** | M3 `MNS410MI` | Add, Upd, Dlt |

The first three are the IFS side and are driven by the CSV exports. The last
three read from M3 and write straight back to it — nothing there depends on a
CSV having been loaded, beyond knowing which tenant you are on.

## IFS Functional Security Role

A functional security role is a named bag of security roles. The tab lists them
with how many security roles each contains and how many users hold it, and
search covers the name, the description, and "containing security role X".

Click one to edit its description or rename it, and to add or remove the
security roles inside it as chips with autocomplete over every role name known
to the tenant — the CSV capture, MNS405, and whatever the functional export
already referenced. Duplicates are refused. **+ Functional Role** creates an
empty one; **Delete** flags it for removal, which keeps it out of the export.

Checkbox selection drives two bulk actions.

**Remove security roles** works the same way as the bulk strip on the Users tab:
the preview lists every security role the selection contains with how many
functional roles hold it, each ticked, and unticking one keeps it. Typing
`REMOVE` applies it. Nothing is sent to M3 — the functional roles are flagged as
changed, and importing the next **Export changes** file into IFS is what actually
applies it.

**Create security roles in M3** is the one thing on this tab that does talk to
M3. It takes the security roles inside the selected functional roles and creates
them with `MNS405MI/Add` — the same call, preview and `CREATE` confirm as
[Create in M3](#rebuilding-m3-from-the-capture) on the security roles tab, which
is where it used to live. Select one functional role and you get its whole set in
one pass, which is the point: standing up `ADMIN` or `AUDITOR` in a fresh
environment is a single action rather than hunting its members down role by role.

A security role held by several of the selected functional roles is created
**once**, and the preview carries an extra *Inside* column saying which
functional roles asked for it. The same skip rules apply — over 10 characters,
a hyphen, a leading asterisk, already in M3 — and a role the Security Role export
never carried is still created, from its name, since inside a functional role a
name is all there is.

"Already in M3" is read from the **MNS405 definitions** when they have been
loaded, so it is right even if **Check M3** was never run against this capture,
and a role this app creates is recorded there too — run it twice and the second
pass reports everything as already in M3 rather than failing against M3.

### Starting again

See [Clearing a capture](#clearing-a-capture) — **Clear all** on this tab empties
the functional roles and their members. It matters most here, because these
files arrive in batches and the import merges them.

The export matches the inbound file byte for byte, including one detail worth
knowing: a functional role with no security roles in it is written as **two
fields with no trailing commas** — `FSR_OVEX,OVEX` — which is exactly how IFS
writes it. Rows that do carry a security role get all four fields with the empty
`EmailId` trailing comma.

## MNS405 Roles

**Load from M3** reads `MNS405MI/Lst` into `M3_Security_M3RoleDefs` — the role
definitions as M3 actually holds them, separate from the CSV capture in
`M3_Security_Roles`. Each row shows `ROLL`, `TX40`, `TX15`, the role type, and
how many users hold it and how many SES400 functions reference it.

Click a role to edit `TX40`, `TX15` or `ROLT` and save through `MNS405MI/Upd`;
**+ Role** adds one through `MNS405MI/Add`. `ROLL` is the key, so it is
read-only when editing — create a new role to change it. **Fields left blank on
an update are not sent**, so they keep their current value rather than being
cleared.

### Deleting a role

Selection drives **Delete selected from MNS405** (`MNS405MI/Dlt`). A delete does
three things:

1. **Deletes the role in M3** through `MNS405MI/Dlt`.
2. **Takes the role off the captured IFS users** — every `SecurityRoleN` /
   `FunctionalSecurityRoleN` entry holding it is dropped and those users are
   flagged as changed, so they land in the next users delta. Deleting `APDIVRW`
   on the CHC data took it off 13 user records.
3. **Writes a SyncSecurityRoleMaster event** with `SecurityRoleMasterStatusCode` `Deleted`.

So the working order is: delete the roles here → **IFS Users → Export changes**
→ import that into IFS, at which point nobody holds the role → the event
deletes it downstream. The preview tells you how many IFS user records each role will be
taken off before you commit.

SES400 authorisations are not touched — if M3 refuses the delete because the
role is still authorised to functions, clear those on the SES400 tab first.

### The SyncSecurityRoleMaster event

`M3_Security_Bod.py` writes the M3BODProcessor `EventData` document — the same
thing M3 publishes from the `CMNROL` smart rule. It needs only the tenant and
the role, so there is no LogicalID or accounting entity to configure.

Three things about M3's delete event are worth knowing, because none of them
are what you would guess from the shape of the document:

* **`<Operation>` stays `UPDATE`.** M3 does not send `DELETE` here. A role
  delete and a role change carry the same `Operation`.
* What marks it as a delete is **`SecurityRoleMasterStatusCode` = `Deleted`**
  plus the rule name ending **`CMNROL_DELETE_SecurityRoleMaster`**.
* The delete event carries **no `OldValue` elements**, where a change event
  carries one on every tracked field.

`currentProgram` / `startProgram` record where the change came from: `MNS405`
for the M3 program, `MNS405Fnc` / `MNS405MI` when it came through the API. The
generator defaults to the `MNS405` pair, matching M3's delete event.

`keyValue` is derived: `KRROLL,` plus the role padded to the 10-character `ROLL`
field with spaces written as `+`, so `AP ADMIN` becomes `KRROLL,AP+ADMIN++`.

`Sequence` is a local counter kept per tenant, since M3's own counter cannot be
read from outside. Set it to the last number you saw and each document takes the
next one. `TrackingId`, `EventId` and `SentTimestamp` are fresh per document.

Everything is optional and lives in `M3_Security_Bod.json`, set in the **event
settings** panel on the MNS405 tab or on the command line:

```bash
python M3_Security_Bod.py --tenant ZFQP353QZYV89ZHG_TST \
    --set-owner EPRONOVOST --set-sequence 261542
python M3_Security_Bod.py --tenant ZFQP353QZYV89ZHG_TST --role "AP ADMIN" \
    --tx15 "AP ADMIN" --tx40 "AP ADMIN ERIC" --print
```

Per-tenant overrides in that file: `owner`, `sequence`, `version`,
`application`, `environment`, `operation`, `current_program`, `start_program`
and `old_values`.

One `.xml` per successfully deleted role lands in `output/m3_security/bods/`,
listed in the result panel with a download link. Roles M3 refused do not get
one. Untick **write event XML** on the selection bar to delete without one —
`M3_Security_Bod.py` can build it afterwards, so that stays recoverable.
Nothing is transmitted by this tool.

Verified against a real M3 delete event: the generated document matches it
element for element, in the same order, with the same `OldValue` treatment —
only `TrackingId`, `EventId`, `SentTimestamp`, `Sequence` and the timestamps in
`LMTS` / `RGTM` differ, as they must.

## MNS410 Roles Per User

**Load from M3** pulls every role-per-user row in one blank `MNS410MI/Lst`, then
filters and pages locally. Rows show the role, the user, the M3 email for that
`USID` where known, and the valid-from / valid-to dates. A role that MNS405 does
not have is flagged, same as on the SES400 tab.

Click a row to change `FVDT` / `VTDT` through `MNS410MI/Upd`; **+ Role per
user** adds one through `MNS410MI/Add` (the role has to exist in MNS405 first).
Dates left blank are not sent and keep their current value. Selection drives
**Delete selected from MNS410** (`MNS410MI/Dlt`), which takes the user out of
the role and leaves the role itself alone.

Every write on both tabs previews first and needs a typed `CREATE`, `UPDATE` or
`DELETE`, runs batched with a progress bar, and lands in `M3_Security_M3Log`.

### The IFS tabs in detail

* **IFS Users tab** — search across name / email / UPN / PersonId, filter by
  status, by change state, by "holding role X", or by whether the user holds any
  functional role; click a row to edit identity, status, locale, security roles
  and functional roles as chips with autocomplete over every known role name.
  Checkbox selection drives the bulk strip below.

### Removing functional roles in bulk

Tick users on the IFS Users tab — or **Select all in view**, which takes
everything matching the current filter rather than the page on screen — and
**Remove functional roles** clears the `FunctionalSecurityRoleN` block on them.
The `SecurityRoleN` block is left alone.

The preview lists every functional role the selection holds with how many users
hold it, each ticked. Untick any you want to keep and only the rest are removed,
so the same control does "strip everything" and "strip these particular roles".
Typing `REMOVE` applies it.

Nothing is sent to M3. The users are flagged as changed, so the next
**Export changes** on that tab carries them without those roles — importing that
into IFS is what actually removes them. Pair the "holding a functional role"
filter with Select all in view to scope the operation to just the users it can
affect.

On the CHC data that works out to 394 of the 439 users holding 502 functional
role assignments across 47 roles; clearing them left all 27,412 security role
assignments untouched and produced a 394-row users delta with no functional
entries in it.
* **IFS Security Roles tab** — search by name or description, filter by whether the
  role exists in M3, select roles with the checkbox column; click a role to edit
  its description, rename it (the rename cascades to every assignment and every
  user record that holds it), add or remove members, and see its live M3
  membership.
* `+ User` and `+ Role` create records; new users get a generated PersonId.
* **Export changes (N)** / **Full export** / **Mark pushed** sit on both tabs; the
  count on the first button is what the delta will contain right now.
