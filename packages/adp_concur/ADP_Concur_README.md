# ADP_Concur_* — ADP and UKG → SAP Concur employee load

Takes the two HR exports Kelly's workbook carries — ADP and UKG — holds them in
`doppio.db`, applies the eight lookup tabs, and writes one 305 / 350 / 360 /
700 flat file for SAP Concur to collect.

**Two systems of record, one file.** The 15 September workbook retired the
US / non-US split: the company is no longer divided by geography but by which
HR system holds a person, and that decides their derivations and which record
types they produce — ADP writes 305 / 350 / 360 / 700, UKG has no 350 tab and
writes 305 / 360 / 700. Both go into a single extract with one 100 record,
sorted so that an approver is always written above anyone reporting to them.

All routines are prefixed `ADP_Concur_` so they group together in the folder,
the same way the `M3_Security_*` set does.

| File | Role |
|------|------|
| `ADP_Concur_Db.py` | Schema + connection helpers (`~/sqlite/doppio.db`) |
| `ADP_Concur_Import.py` | Loads the workbook — the ADP and UKG sheets, the maps, the layouts |
| `ADP_Concur_Map.py` | The derivations, the Login ID rule, the record builders |
| `ADP_Concur_Hierarchy.py` | The supervisor chain — walks it, and finds where it breaks |
| `ADP_Concur_Fix.py` | Regroups the exceptions by what would fix them, and applies it |
| `ADP_Concur_Export.py` | Writes the flat file into the outbound folder |
| `ADP_Concur_Result.py` | Reads Concur's load result back and puts names to it |
| `ADP_Concur_App.py` | Flask front end — nine tabs, editing, extract |
| `templates/ADP_Concur_Index.html` | The single-page UI |
| `ADP_Concur_Config.json` | The Login ID rule, the scopes, the file shape |

## Quick start

```bash
cd ~/Doppio/packages/adp_concur
pip install flask openpyxl xlrd         # only external dependencies
python ADP_Concur_App.py                # http://127.0.0.1:5058
```

`xlrd` is only needed to read the `.xls` result files Concur sends back, and it
is the one dependency that is easy to miss because nothing needs it until the
first result arrives. Without it the app runs normally, says so on the drop
zone, and warns once at startup — and the alternative needs no install at all:
open the result in Excel and Save As `.xlsx`, which the app also reads.

Drop the workbook on the page. It is saved to `input/adp_concur/`, parsed into
`~/sqlite/doppio.db`, and every tab fills. Command line equivalents:

```bash
python ADP_Concur_Db.py                                   # create the tables
python ADP_Concur_Import.py "Concur Draft ... for Load.xlsx"
python ADP_Concur_Map.py                                  # re-derive, list the exceptions
python ADP_Concur_Export.py --dry-run                     # count without writing
python ADP_Concur_Export.py                               # the one file Concur wants
python ADP_Concur_Export.py --held-back                   # who is being left out, and why
python ADP_Concur_Result.py --load "Run-20.xls"           # read Concur's answer back

python ADP_Concur_Hierarchy.py --chain 211837             # one person's chain to the top
python ADP_Concur_Hierarchy.py --subtree 211766           # everyone under a manager
python ADP_Concur_Hierarchy.py --problems --stats         # broken links, shape of the tree
python ADP_Concur_Hierarchy.py --map                      # validate the Supervisor Map
python ADP_Concur_Fix.py                                  # everything wrong, and what would fix it
python ADP_Concur_Fix.py --errors-only

# a pilot load: one manager's whole organisation and nobody else
python ADP_Concur_Export.py --under 211766 --dry-run
python ADP_Concur_Export.py --under 211766
```

Every routine takes `--db PATH`; the `ADP_CONCUR_DB` environment variable works
too. Default is `~/sqlite/doppio.db`, matching `M3_Security_Db.py`, `Sheet2Db.py`
and `config.py`.

### Upgrading an existing database

`doppio.db` is shared with the `M3_Security_*` tables, so it is never a fresh
file — and `CREATE TABLE IF NOT EXISTS` does nothing at all to a table that
already exists. A version that adds a column therefore has to add it by hand,
which is what `MIGRATIONS` in `ADP_Concur_Db.py` is: a list of the columns
added since the first release, applied on open and reported in the log.

    Added 22 column(s) to the existing schema: ADP_Concur_Employees.invoice_limit, ...

The indexes are deliberately kept out of `SCHEMA` and run *after* that step. An
index names a column, so on an older database the table is old, the index is
new, and running them together fails on a column that has not been added yet —
which is exactly what `no such column: roster` was. Nothing is dropped or
rebuilt: existing rows keep their data and take the column's default, so a
person loaded before UKG existed stays on source `adp`.

The `roster` column is still in the schema and nothing reads it. Dropping a
column is the one migration SQLite will not reliably do on an old file, and a
dead column costs nothing; `SOURCES` is the axis now.

**When adding a column to `SCHEMA`, add it to `MIGRATIONS` in the same edit.**
Listing one that is already there is harmless — SQLite refuses it and the
migration moves on — so erring towards listing it is right.

## What the workbook carries

Nothing depends on the tab names, and the 15 September workbook proved it:
every sheet was renamed — `1` became `ADP`, the record tabs grew names like
`305 ADP Employee (All)` — and nothing broke, because nothing was ever reading
the names. Each sheet is recognised by what its header row says, and the header
row is found by looking for it rather than assumed to be row 1. That matters
for the ADP sheet, which has three rows of Kelly's notes above the real
headings, and for UKG, whose headings are on row 1.

| Sheet | Recognised by | Goes to |
|-------|---------------|---------|
| the ADP export | `Payroll Company Code` + `File Number` + `Position Status` | `ADP_Concur_Employees`, source `adp` |
| the UKG export | `Employee Number` + `Salary Grade` + `Pay Group` | `ADP_Concur_Employees`, source `ukg` |
| Status Map | `Position Status` + `Concur Status` | `ADP_Concur_StatusMap` |
| Country Map | `ADP Country` | `ADP_Concur_CountryMap`, and its country block to `ADP_Concur_CountryRef` |
| Org Map | `Business Unit Description` + `Home Department Code` | `ADP_Concur_OrgMap` |
| Language Map | `Language` + `ADP Language` | `ADP_Concur_LanguageMap`, and its locale block to `ADP_Concur_LocaleMap` |
| Salary Map | `Pay Grade Code` + `Expense Map` | `ADP_Concur_SalaryMap` |
| Supervisor Map | `Exception Employee Id` + `Supervisor ID` | `ADP_Concur_SupervisorMap` |
| Role Assignment Map | `Role` + `Assign Role Automatically` | `ADP_Concur_RoleMap` |
| Invoice Exception Map | `Invoice Exception Employee` + `Invoice Access Value` | `ADP_Concur_InvoiceMap` |
| 305 / 350 / 360 / 700 | first cell reads `Trx Type (nnn)` | `ADP_Concur_Layouts` |
| 400 and 320 | first cell reads `Trx Type (400)` / `(320)` | nothing — named as skipped |

**Both exports land in the same table.** They describe the same people in
different words, so the raw columns are mapped onto one shared set of names on
the way in — UKG's Employee Number is a File Number, its Salary Grade is a Pay
Grade Code, its Employment Status is a Position Status — and `source` records
which system sent the row. Only the derivations and the record types differ
after that, which is what lets one file carry both.

UKG sends one name field where ADP sends three (`Shi, HanBing`), so it is split
on the comma the way the workbook's `FIND`/`LEFT`/`MID` does, with one
difference: a name with no comma keeps the whole string as the last name rather
than becoming `#VALUE!`.

**400 and 320 are recognised and skipped.** The workbook carries them, nothing
in scope reads them, and they are named in the load report rather than reported
as unrecognised — "ignored on purpose" and "not understood" are different
things, and a sheet that quietly vanishes is how a record type gets forgotten.

**The lookup tabs are a full refresh** — dropping a workbook replaces them with
what it carries, because the maps are the workbook's job. **The employees merge
on File Number**, so a fresh ADP cut refreshes the people already held without
disturbing the ones keyed in by hand. Anyone held here and absent from the file
is left alone rather than deleted; a partial cut of ADP is a normal thing to be
handed. **Start again** on the Extract tab is the from-scratch reload.

Only the raw columns are loaded. The derived ones are recomputed from the maps,
so a workbook carrying stale lookup results never puts them in the database.

The four record templates are captured as layouts rather than data, which is
what makes a new Concur template a matter of dropping a new file: the extract
takes its width, its column order and its field-width row from whatever the
workbook shows.

**How wide a record is comes from the numbered row, not the last cell with text
in it.** Both 700 tabs carry a note in the column after the last field — *Only
load users who can work with invoices* — with no field number and no width.
Counting headings would make that note a seventeenth field, and since Concur
requires every field to be represented, every 700 in the file would have gone
out one delimiter too wide.

## The derivations

Every formula on Kelly's sheet has a function in `ADP_Concur_Map.py` with the
formula it replaces quoted above it. In workbook order:

| Col | Value | Rule |
|-----|-------|------|
| AC | SupervisorID Formatted | characters 4–9 of ADP's `Supervisor ID`; where ADP has none, the Supervisor Map on the employee's own File Number |
| AD | Org Unit 1 | Org Map on Business Unit Description |
| AE | Org Unit 2 | Org Map on Home Department Code |
| AF | Concur Profile | Salary Map on Pay Grade Code → Expense Map |
| AG | Travel Profile | Salary Map on Pay Grade Code → Travel Map |
| AH | **Invoice Approval Limit** | Salary Map on Pay Grade Code → Invoice Approval |
| AI | Legal Country | Country Map |
| AJ | Locale Code | Language Map on Language Description, else the BU's default language, then the country appended — `en_` + `US` |
| AJ | Reimbursement Currency | Org Map on Business Unit Description |
| AK | Preferred Name | blank unless ADP has a preferred first name |
| AL | Status | Status Map on Position Status |
| AM | Term Date | the Termination Date as `yyyymmdd`, only when the status is `N` |
| — | Login ID | see below — this one is not in the workbook |

A lookup that misses writes the workbook's own text — `BU Description Not
Mapped` and friends — so a value that fails here is recognisable to anyone who
has been working in the spreadsheet.

Editing a map re-derives every employee straight away, so the Employees and
305 / 350 / 360 tabs can never disagree with the maps about what a lookup
returns.

## Three differences from the workbook, on purpose

Re-deriving Kelly's 190 rows reproduces her 305, 350 and 360 tabs exactly,
apart from three things — and in each case the workbook is the one that is
wrong. They are worth knowing about because two of them are the very problems
her email is asking after.

**The supervisors.** `=IF(ISBLANK(H5), VLOOKUP(...), MID(H5,4,6))` never
consults the Supervisor Map, because ADP writes an *empty string* rather than a
blank cell — and `ISBLANK("")` is FALSE, so the formula takes `MID("",4,6)` and
returns nothing. Every one of the Supervisor Map's exceptions is being thrown
away. Reading it properly fills the supervisor on **25 records**, which is what
the 305 approver field and both 360 approver fields are built from. That is the
"see what happens with the supervisors and hierarchies" question, answered.

**The Org Map lookups.** `VLOOKUP(V5,'Org Map'!C:E,3)` leaves the fourth
argument off, so Excel does an *approximate* match and silently returns the row
above the one it wanted whenever the codes are not sorted — and the Org Map's
department codes are not sorted, they restart at `000002` when the 088 block
begins. Ten records come out as `Home Department Code Not Mapped` when the code
is right there in the map, and three land on the wrong business unit entirely
(`0088` where the employee is `0096`). Matching exactly fixes both; a code the
map genuinely does not carry is reported rather than mapped to its neighbour.

**Duplicate employees.** ADP sends one row per employment record, so an
internal transfer arrives as the same File Number twice — terminated under the
old payroll company and active under the new one. **15 people account for 24
extra rows** in this cut. Concur wants one profile per employee and it wants
the live one, so the rows are ranked: a live Position Status beats a terminated
one, then no Termination Date beats having one, then the later rehire and hire
dates, then having a supervisor, then the later row in the file. Every rejected
row is written onto the employee as a warning saying exactly which rows were
seen and which was taken, so the choice is visible rather than silent. Left as
the spreadsheet has it, those 24 rows would each load as their own Concur
record.

## The Login ID

This is the piece the email is asking about, and it is the one thing that is
not in the workbook.

Concur login IDs are unique across **every** entity on the platform, not just
yours, so a plain work email address usually collides with one that already
exists somewhere. The rule is configurable rather than guessed, on the
**Extract** tab or in `ADP_Concur_Config.json`:

| Setting | Effect |
|---------|--------|
| `sources` | Fields tried in order; the first with a value wins |
| `suffix` | Appended whole — `.fmg` gives `a.user@onebarnes.com.fmg` |
| `prefix` | Prepended whole |
| `replace_domain` | Swaps everything after the `@` |
| `bare_domain` | Added when the chosen source has no `@` — a File Number |
| `lowercase` | Off by default, so the extract matches the workbook |

Out of the box nothing is applied and the login IDs come out as the work email
exactly as the spreadsheet has them. The panel shows a worked example as you
type. The two shapes SAP normally asks for are a suffix on the real address
(`suffix: ".fmg"`) or a dedicated domain (`replace_domain: "fmgconcur.com"`) —
set whichever they confirm and the whole extract follows it in one save.

**34 people in this cut have no work email address at all**, so they have no
login ID and are held out of the file. Adding `personal_email` or
`file_number` to `sources` — with a `bare_domain` for the latter — brings every
one of them in; leaving it alone keeps them out until someone decides. Nothing
is invented on their behalf either way.

## Exceptions

Rebuilt from scratch on every derive, so the list always describes the data as
it stands rather than accumulating history. **Errors** are held out of the
extract; **warnings** load anyway and are worth a look.

| Check | Severity |
|-------|----------|
| No File Number | error — Concur has no Employee ID to key on |
| No Login ID | error |
| A `... Not Mapped` value | error, or warning if the person is terminated |
| No work email address | warning |
| More than one ADP row | warning, naming the rows and the one taken |
| No supervisor and not terminated | warning |
| Supervisor not in this load, self-supervision, or a loop | error — Concur cannot resolve the approver |
| Supervisor is in the load but held out of the 305 | warning |
| Approver comes from the other HR system | warning — worth knowing, not fixing; one file, and the approver is written first |
| Approver is inactive | warning — an active employee cannot be approved by somebody who has left |
| Passwords run as an unbroken fill series | warning |
| Inactive with no Termination Date | warning |
| Non-US country code | warning |

That covers the three groups the email names: bad or incomplete ADP data,
terminated people whose mapping no longer resolves — which is why an unmapped
value on a terminated person is only a warning — and the missing supervisors.
Each one can be turned off in `block_on` to let it through and see what Concur
says.

## The hierarchy

One derived field carries the whole org chart: `supervisor_id`, another
employee's File Number. It is what Concur routes approvals through — the 305's
Expense Report Approver and both of the 360's approver fields are that number —
so a link pointing at nobody is not cosmetic. It is an approval chain that ends
in mid-air, and Concur rejects the record.

The **Hierarchy** tab is the tree, with the chain beside it. Click anyone and
the right-hand panel shows their **complete chain up**, rendered top-down the
way an org chart reads, with a line at the bottom saying how the chain *ended*:

| Ending | Means |
|--------|-------|
| `top` | the last person has no supervisor at all — a genuine top of the chain, which is what the Supervisor Map's "Top of food chain" note describes |
| `broken` | the last person names a supervisor no employee in this load has |
| `cycle` | the chain came back to somebody it had already visited |
| `deep` | it ran past 40 levels without resolving |

Underneath it: the direct reports, and how many people sit below in total. The
same chain appears in the employee editor, under **Chain up**, so it is visible
on the record it affects rather than only on its own tab.

Nothing is cached. The tree is derived from `supervisor_id` on demand, because
one Supervisor Map edit rearranges it and a stored copy would be wrong the
instant somebody saved.

### Filtering the tree

Three filters on the toolbar, and they combine: **status** (active, inactive,
or both), **source** (ADP, UKG, added by hand, or all three) and **picked**.
*Active only* is the one to reach for when reviewing hierarchies — the
leavers are noise, and Kelly's note about only loading the active employees
is the same instinct. *Source* is for the days ADP and UKG need to be looked
at separately — the two systems produce different record types (UKG has no
350 tab) and their own supervisor data, so a chain that looks broken may just
be crossing a boundary the filter can isolate.

Filtering prunes the tree but never breaks it. A manager who fails the filter
is kept when somebody below them passes it, because a tree with the middle cut
out of it is not a tree — and a leaver who still has active people under him is
exactly that case, where hiding him would orphan his whole branch. Those
pass-through managers are dimmed and marked **chain only**, and the toolbar
says how many there are: *inactive only* on this workbook reads

> showing 23 of 40 on screen (17 kept to hold the branches up)

so the 23 real leavers and the 17 managers holding them up never get confused.
The stats beside it — 13 roots, 4 levels deep — describe the whole hierarchy
and deliberately do not move when a filter is on, because those are facts about
the org chart rather than about the view.

**The chain-up panel is never filtered.** It is the truth about who approves
this person, and a filtered chain would be a lie — if their manager is a leaver,
that is precisely what you need to see. The direct reports *do* follow the
filter, so the panel and the tree agree, but the panel says what it hid:
*Direct reports — 18 of 19*, with *1 hidden by the filter* underneath. A
filter that quietly reports three reports when there are five is the one thing
this must never do.

### Where it breaks today

Three people in this cut report to somebody who is not in the load, and the tab
says so in a banner and marks each of them as a red root:

* **Finch, Jeffrey** and **Hibbitts, Brandon** both report to `207199` —
  ADP names him as *Erdogan, Mehmet*, who is not in this extract. Very likely
  one of the non-US people.
* **Oleksa, Laura** reports to `TBDTUN`, which is not a File Number at all —
  it is a placeholder somebody typed into the Supervisor Map where
  *Adewalure, Babatunde*'s ID should be.

These are errors by default and held out of the extract, because Concur cannot
resolve an approver it does not have. `block_on.broken_supervisor` turns them
into warnings if you would rather send them and see what happens. Two related
checks run alongside: reporting to yourself, and a supervisor who *is* in the
load but is being held out of it — the same problem arriving a step later.

The rest of the shape is worth a look too: **13 roots for 166 people**, of
which 10 are single people with no supervisor and no reports. Whether those are
genuine tops or simply people ADP has no manager for is a question for
Wednesday.

## Validating the Supervisor Map

The Supervisor Map is the one map whose rows can be wrong in a way that reading
them will not show, because both ends have to resolve to an employee. It is
also the map that matters most: it overrides ADP, so a wrong row silently
rewrites somebody's approval routing and nothing else in the load will argue.

So the Maps tab checks every row against the employees and marks it. Six
verdicts:

| Verdict | Means |
|---------|-------|
| **ok** | Both ends resolve. |
| **top of chain** | Supervisor deliberately blank — Teresa Bair's row. |
| **no such supervisor** | The supervisor it names is not in the load. Concur gets an approver that does not exist. |
| **never applies** | The row is keyed on a File Number this load does not have, so it fires for nobody — and whoever wrote it thinks it did. |
| **points at itself** | Keyed on the same person it names as supervisor. |
| **name mismatch** | Both ends resolve, but a name in the row disagrees with the employee it points at — usually a row keyed on the wrong File Number, quietly rewriting the wrong person's approver. |

A banner over the grid says how many did not hold up and why; a clean map says
so in green. Name comparison is loose — punctuation, spacing and word order are
all forgiven — so only a genuinely different name is reported.

Two of these are invisible from the employee side and only this check finds
them: **never applies** and **name mismatch**. The first is a dead row nothing
reports because it never fires; the second looks perfectly healthy from every
other angle.

On the workbook as it stands, 2 of the 7 rows name a supervisor who is not in
the load — `207199` (Erdogan, Mehmet) and `TBDTUN` (Adewalure, Babatunde).

### Creating the supervisor from the map

A row marked **no such supervisor** gets a **Create** button, but only when the
ID is actually a File Number — `TBDTUN` does not get one, because creating an
employee called TBDTUN would turn a placeholder into a fact. Clicking it goes
to that person's card on the Fixes tab with the create form already open and
filled in, rather than offering a second, thinner create form here.

The map improves that form: where a Supervisor Map row is what put the ID on
somebody, the map's own **Supervisor Name** is used for the new record's name
in preference to ADP's *Reports To*. Whoever wrote the map row knew who they
meant; ADP's copy may be stale or blank.

The Fixes tab also carries the two verdicts that have no employee to hang off:
a dead row can be removed or have its employee created, and a mismatched row
can be rewritten deliberately or dropped back to whatever ADP sent.

## Extracting a selection

The whole file is rarely what you want first. Pick people and the extract
carries only them — that is the pilot load, one branch or one manager's
organisation, small enough to look at every record Concur sends back.

A selection is one set shared by every tab, so the two ways of building it
combine:

* **Hierarchy tab** — click somebody, then **Select this organisation**. That
  is them and everyone underneath at any depth, resolved on the server rather
  than from the tree on screen, so a map edit between rendering and clicking
  cannot scope the file to a hierarchy that no longer exists. **Select just
  this person** takes one.
* **Employees tab** — tick rows, **Select this page**, or **Select all in
  view**, which takes everything matching the current filter rather than the
  page on screen. **Drop those in view** is its opposite, for pruning a pilot
  rather than building one. Ticks stay put as you page and filter.
* **The employee editor** — one button, for when you are already looking at
  somebody. **Drop this organisation** on the Hierarchy tab removes a whole
  branch the same way it added one.

The selection lives in `ADP_Concur_Selection`, not in the browser. That buys
three things: it is still there tomorrow morning, every list can filter on it
with a join instead of a query string carrying 150 keys, and the extract and
the screen cannot disagree about who is in it — they read the same table. Each
row records *how* that person got picked, so a selection can be explained after
the fact rather than just counted.

### Seeing only what is picked

Every tab that lists people has a **picked** control, and they share one
setting — flip it on the Employees tab and the record tabs follow.

| Tab | *Only picked* shows |
|-----|---------------------|
| **Employees** | just the picked, with *Only not picked* as well — that one is how you check what a pilot is leaving behind |
| **305 / 350 / 360** | the records a selection-scoped extract would write, built the same way, before it is a file |
| **Exceptions** | what is wrong *inside* the pilot, counts and all — 15 missing login IDs in this organisation rather than 34 across the company |
| **Hierarchy** | the pilot's own org chart — and it combines with the status filter, so *picked + active* is the org chart of what a pilot load would actually create |

Because it is a filter like any other, it composes: *only picked* plus
*blocking errors* answers "what will stop my pilot loading", which is usually
the question.

The pruned tree keeps the managers above each picked person even when they are
not picked themselves — a tree with the middle cut out of it is not a tree.
Those pass-through managers are dimmed and marked *not picked*, so it stays
obvious who is actually in the pilot and who is only holding the branch up.
Picking Greg Kemenah's organisation gives a tree of 78: the 77 who are in it,
under Teresa Bair who is not.

**Extract the selection** jumps to the Extract tab with *This file covers* set
to *Only what is picked*, and previews. Preview builds and counts the file
without writing; the held-back list is scoped too, so it names people inside
the selection rather than the hundred who were never chosen.

The written file is named **`FMG_Concur_Employee_Selection_{stamp}.txt`** —
deliberately not the full-load name. The outbound folder is a Concur pickup,
and a partial file that looks exactly like a complete one is the kind of thing
that gets loaded by accident at four in the afternoon. The row in
`ADP_Concur_Extracts` records the count and the label, so a file can be traced
back to what was picked.

From the command line, `--under` takes a whole organisation and `--only` takes
named people; both repeat, and they can be combined:

```bash
python ADP_Concur_Export.py --under 211766 --under 207588 --only 209181
```

## Fixing what is wrong

The exception list says what is wrong one employee at a time, which is the
wrong grain for actually clearing it. Two people blocked by the same absent
supervisor are **one missing person**, not two problems; eleven people in an
unmapped department are **one missing Org Map row**. The **Fixes** view on the
Exceptions tab regroups everything by what would fix it, worst first, and each
card carries the remedies with the answer already filled in as far as the data
allows. *Every exception* switches back to the raw list.

On the workbook as it stands that is 8 cards covering 37 errors.

| What is wrong | What the card offers |
|---------------|----------------------|
| A supervisor nobody in the load answers to | **Create them**, from a prefilled record; or point their reports at somebody else; or leave those reports with no supervisor |
| A Supervisor Map row that never applies, or whose names disagree | Create the employee it is about, remove the row, or set the supervisor deliberately — see [Validating the Supervisor Map](#validating-the-supervisor-map) |
| A value the maps do not carry | **Add the map row**, keyed on the value that is missing, with the rest borrowed from a sibling row |
| ADP left the column blank | **Fill it in**, or **leave them out of the load** |
| Nobody has a Login ID | **Change the rule** — one decision, not thirty-four records |
| Reports to themselves, or a loop | Point them at somebody else |
| No work email | Open the record; there is nothing to guess from |

### Creating the missing person

This is the one Kelly's email runs into. `207199` is the supervisor of both
Jeffrey Finch and Brandon Hibbitts, and he is not in the extract — ADP names
him in *Reports To Legal Name* as **Erdogan, Mehmet**, and that is the only
reason the new record can be created with a real name on it.

The form arrives filled in. The name is split out of `Erdogan, Mehmet`; the
File Number is the ID his reports are already pointing at; and the business
unit, department, location, country, pay frequency and payroll company are
taken from **the people who report to him**, on the reasoning that a manager
almost always shares a business unit with their reports. Every one of those
guesses is drawn dashed and grey, so the fields that are certain and the fields
that need checking are distinguishable at a glance. Pay grade is the highest
among the reports, which is the least wrong starting point rather than a claim.

He is created as a **manual** record — the same thing as anyone else ADP does
not have yet — so a later ADP cut takes him over cleanly on the File Number.

### The report after each fix

Every fix re-derives and comes back with the error and warning counts on either
side of it, plus anything still outstanding on the record it touched. That
matters more than it sounds: creating Mehmet Erdogan clears the two broken
chains and *introduces one new error*, because a person invented from guesses
has no email address and therefore no Login ID. The card says so on the spot —

> 37 → 36 errors, 56 → 58 warnings.
> Still outstanding on the record this touched:
> **207199** No Login ID — ADP has neither a work nor a personal email address…

— rather than reporting success and letting it turn up unexplained later. Work
the list top to bottom and the extract clears; on this workbook the 37 errors
go to zero once the Login ID rule has a fallback, Erdogan exists, Oleksa is
re-pointed and Carla Cunningham's blank record is either finished or dropped.

### Corrections that survive the next import

A correction is only worth making if it is still there after the next ADP cut.
So fixes never write a derived column, and they never quietly patch a value the
import will overwrite. They write to what the derive *reads from*:

* **supervisor changes go into the Supervisor Map**, keyed on the employee.
  The import replaces the employee's own ADP columns and leaves the map alone,
  so the correction sticks.
* **unmapped values go into the six maps**, which is where they belonged.
* **the Login ID rule goes into the config.**
* **a blank ADP column is written onto the employee** — which used to be
  purely a stopgap, overwritten the moment ADP sent a fresh cut. It isn't any
  more: editing a field on the Employees tab records which field actually
  changed in `overridden_fields`, and the next import leaves exactly those
  columns alone instead of putting ADP's value back — see "Edits that survive
  the next import" below. It is still worth fixing at source when that's
  possible, since a protected field stops following ADP entirely until the
  protection is cleared.

**One deliberate change came with this.** The Supervisor Map now *wins* over
ADP's own Supervisor ID rather than only filling in when ADP is blank. A table
of exceptions that can fill a hole but cannot correct a wrong value is no use
for the thing it is most needed for. On the workbook as it stands this changes
nothing — of the seven map rows only Moavero also has an ADP value, and both
give `006769`.

The supervisor picker only offers people who are actually in the load, and the
server refuses an ID that is not, so the fix for a broken chain cannot itself
create a broken chain.

### Edits that survive the next import

Opening a row on the Employees tab shows every ADP (or UKG) field as an
ordinary editable box, for any employee regardless of source — Save writes
straight to the same columns the import fills. Left alone, that would be
pointless: the next ADP cut would silently put its own value back over the
edit, since the merge in `ADP_Concur_Import.py` used to overwrite every
column the sheet carries.

So Save also works out which fields the edit actually *changed* — comparing
the new value against what was already on the row, not just what the form
happened to submit, since the form resends every ADP field whether it was
touched or not — and adds only those names to the employee's
`overridden_fields` column. The next import checks that list before writing
each column: a protected field keeps the edited value, everything else still
refreshes from ADP as usual. The edit panel says which fields are currently
protected and offers **Clear protection** to drop them, so a field can be
handed back to ADP's control once whatever prompted the edit is fixed at the
source.

This is scoped to *fields*, not the whole record — editing Job Title does not
also freeze Position Status, so a termination ADP reports the next day still
lands even on a row with a hand-edited title.

## People ADP does not have

**+ Employee** on the Employees tab keys someone in by hand, for a branch that
has not gone live on ADP yet. The record is marked `manual`, which keeps it
clear of the ADP merge and lets a later ADP cut take it over cleanly — the File
Number is the key either way, so when ADP finally sends that person they simply
become an ADP row with no duplication and nothing to clean up. Hand-keyed
people survive a clear unless they are named in it — see
[Starting again](#starting-again).

Every employee also carries three include flags, so one person can be held out
of the 305, the 350 or the 360 without touching anything else.

## Starting again

**Start again** on the Extract tab clears employees by where they came from,
because those are two different decisions with very different costs:

| Source | Clearing it costs |
|--------|-------------------|
| **From ADP** | Nothing. The rows are a copy of what the ADP report still holds, so they come back the moment the workbook is dropped again. |
| **From the non-US tabs** | Nothing, as long as you still have the workbook. These people are typed into the `305 Non US Non SE` and `360 Non US Non SE` tabs rather than fed by ADP, so they come back on the next drop of a workbook that still carries them. |
| **Added by hand** | Everything. They were keyed in here and are held nowhere else — including anyone created from the Fixes tab to repair a broken supervisor chain. No workbook will bring them back. |

Tick any of them. The panel says exactly what would go before anything goes
— how many employees, how many exceptions with them, how many are in the
current selection, and how many would remain — and it only turns red when the
choice is the irreversible one:

> **167 employees** would go · 94 exceptions with them · 77 of them are in the
> current selection · 0 employees would remain
> **1 of them cannot be recovered.** They were keyed in here and exist nowhere
> else — dropping the workbook again will not bring them back.

Typing `CLEAR` arms the button, the same gate the `M3_Security_*` tool puts on
its own clear. Naming no source at all is refused rather than quietly falling
back to a default, which is the sort of default that deletes 166 people.

**The third source was missing, and that is worth recording.** The non-US
importer has always written `source = 'workbook'`, and the Employees filter has
always listed it — but `EMPLOYEE_SOURCES`, the registry that decides what a
clear can actually reach, only knew about `adp` and `manual`. So "start again"
with every box ticked removed 166 people and silently left 82 behind, and the
preview never mentioned them. Both lists are now built from that one registry,
and so are the per-source counts, which had stopped adding up to the total for
the same reason. Adding a source value means adding it to the registry; there
is no default that covers a name it has not seen.

### Clearing the history

Underneath the sources is the rest of what "start again" ought to mean. Three
log tables record what has happened rather than what the load is made of, and
none of them were being cleared by anything:

| | Holds |
|---|---|
| **Workbook loads** | Every workbook dropped, with its per-sheet counts |
| **Files written** | The record of every flat file written |
| **Concur load results** | Everything Concur has said back, and the Results tab with it |

Nothing is derived from any of them, so clearing them loses the record and not
the data, and the extract files themselves stay on disk — this only forgets
them. The one consequence worth knowing: without the **Files written** row, a
load result can no longer be matched to the file it answers, so its record types
come out blank. The names survive, because those come from the employee IDs
Concur quotes rather than from the join.

Load results outlive a clear of the employees they are about, which is
deliberate — what Concur said is history. The key pointing at the employee row
is nulled when that employee goes, so a stale key can never open the wrong
editor; the file number and name stored on the row keep the result readable.

**The maps are not touched.** Clearing employees leaves the six lookup tables
and the captured layouts alone, which is the point of having filed a supervisor
correction in the Supervisor Map: it outlives the employees it was about, and
still applies after the next import. **Clear maps and layouts** beside it is
what throws those away, behind its own `CLEAR`.

One employee at a time, the editor has both: **Leave out of the load** is the
soft delete that keeps the row and simply stops writing it, and **Delete
outright** appears only on hand-keyed records — for an ADP row it would be
pointless, since the next drop returns them.

## One file, both systems

The workbook holds two HR exports and they are the same company, so they are
one load into Concur rather than two.

| Source | Comes from | Records | Derivations |
|--------|-----------|---------|-------------|
| **ADP** | the `ADP` sheet, through the maps | 305, 350, 360, 700 | thirteen lookups |
| **UKG** | the `UKG` sheet | 305, 360, 700 — there is no 350 tab for them | four lookups; the rest arrive answered |

Both carry the same Login ID rule (`<work email>.new.uat`), so the mechanism
for per-source overrides is still there and empty. One 100 record sits on the
front of the one file.

UKG is not a variant of the ADP rules — it is a different system describing the
same people differently, so it gets its own derivation. Half of what ADP looks
up, UKG already answers: the supervisor is a bare employee number, the org
units are codes on the record, the status is a code. What is left is four
lookups.

| Value | Rule on the UKG sheet |
|-------|----------------------|
| Org Unit 1 | the Site Location Code, straight through |
| Org Unit 2 | the Org Level 3 Code, straight through |
| Concur / Travel Profile | the Salary Map on a **trimmed** Salary Grade |
| Legal country, currency | the Country Map on the three-letter Country Code — UKG has no business unit, so the Org Map route ADP uses is not open to it |
| Locale | the derived two-character country in the Language Map's locale block, else `en_` + country |
| Invoice access | the Invoice Exception Map **only** — no UKG employee gets invoice access by pay grade |

Three of UKG's fourteen derived columns are deliberately empty because its own
derived block leaves them empty: Local Code, Preferred Name and Term Date have
no formula at all. The 305 tab computes its own locale from the country rather
than reading the blank one, which is why locale is filled here and the other
two are not.

### The sort is what made one file possible

Concur resolves an approver against what it already holds, so a record naming
somebody who has not been created yet loads without them. That is why the two
roster files had to be loaded in a particular order, and why twelve people who
reported across the line drew warnings every run.

The 305 records are now written **active first, then down the supervisor tree**
— every approver above everyone reporting to them. On the current cut that is
456 approver/report pairs in the file, **26 of them crossing between ADP and
UKG**, and not one approver written after their report.

Status is the outer key, as asked, so the whole tree is walked and then split
in place by a stable partition — which keeps the hierarchy inside each half.
The one pairing it cannot honour is an active person reporting to an inactive
one: the approver lands in the second half, below them. There is exactly one of
those, and it is not hidden — it now raises its own warning on the Exceptions
tab, because an approver who has left is a problem in the data rather than in
the sort. Burying the leaver mid-tree to make the file look ordered would only
have made it harder to see.

Anyone the walk cannot reach — people inside a supervisor cycle — is appended
in name order rather than dropped. Losing a record to keep a sort tidy would be
the worse bug by far.

### Invoice access, and the 700

The 700 record is new, and it is a payment request approval authority: one per
person who may approve invoices, carrying the limit they may approve up to.

Two derived values drive it, and they read in this order:

1. **Invoice Approval Limit** — the Salary Map's new `Invoice Approval` column,
   looked up on the pay grade.
2. **Invoice Access** — the Invoice Exception Map first, by name. For ADP,
   anything not named falls back to *has a limit above zero*. For UKG there is
   no fallback at all: nobody gets invoice access by grade, only by name.

That asymmetry is the sheets', not a simplification — it is why a UKG employee
on the same grade as an ADP one can come out differently. Invoice Access also
drives the 305's Invoice User and Invoice Approver flags and four flags on the
360, so somebody without it is loaded without Concur Invoice rather than given
it and left unable to use it.

**The 700 is not written for anybody without access.** Both tabs build the
whole row as `=IF(AccessIsY, ..., "")`, and an empty row is not a record, so
they are filtered out rather than written blank. On the current cut that is 37
of 495 people — 33 from ADP, 4 from UKG.

### What the new workbook needed

* **`Multiple Payrolls`** is a pivot drill-down of the same people, so it is
  ADP-shaped and would load as a second source. The fullest cut wins, and on a
  tie the `Details for ...` marker Excel writes into A1 of a generated
  drill-down settles it. The load reports which sheet it took and why.
* **The Country Map absorbed its own second block.** It used to be two tables
  on one tab; the new sheet carries the code, the name, the currency code and
  the currency name in one row, which is where UKG gets its currency.
* **Column matching now tries an exact heading before a prefix.** Prefix
  matching is what lets `Employee ID` find `Employee  ID (Cannot be changed
  using this record ...)`, but UKG's sheet is full of headings that start with
  each other — `Job` begins `Job Code`, `Job Family`, `Job Role` and `Job
  Type`; `Pay Group` begins `Pay Group Code`. Asking for `Job` returned the job
  code. Each term now gets an exact scan before a prefix scan, one term at a
  time in the order the caller listed them, because the caller's order is it
  saying which heading it would rather have — the Country Map asks for `Legal /
  Preferred Address` before `ADP Country`, and a global exact-first pass handed
  back the value column instead of the key.

## The 100 record

SAP requires one Import Settings record as the first line of every file. It is
not per-employee and it is not in the workbook, so it comes from the config and
is written by the extract:

```
100,0,TEXT,UPDATE,en,Y,Y
```

| # | Field | Default here | Why |
|---|-------|--------------|-----|
| 2 | Error Threshold | `0` | SAP says enter zero |
| 3 | Password Generation | `TEXT` | use the password on the 305. FMG's proven file uses `SSO`, so the passwords are ignored |
| 4 | Existing Record Handling | `UPDATE` | writes only the non-blank fields and never overwrites an existing password, which is the safe default for a repeated load. `REPLACE` overwrites the record wholesale |
| 5 | Language Code | `en` | the language of any localised text in the file |
| 6 | Validate Expense Group | `Y` | SAP's default |
| 7 | Validate Payment Group | `Y` | SAP's default |

The two fields with a fixed vocabulary are checked before the file is written,
so a typo is caught here rather than at the far end of a load — `Password
Generation is 'PLAINTEXT'; SAP accepts EMPID, LOGINID, SSO, TEXT.` The Extract
tab shows the line as it will actually be written, which is the only form in
which a wrong field is obvious.

## The extract

One file carries all three record types, one record per line, every position
present whether or not it has a value — which is what *All fields must be
represented* on the 350 and 360 tabs means.

Who is in it comes from three places: the per-employee include flags, the
configured scope, and the exception list. Scope is always `all` - active and
terminated people alike - except the 700, which is scoped to `invoice` (only
people with invoice access) because it exists for that reason alone. The one
other choice is `off`, which leaves a record type out of the extract
altogether — the Extract tab's "350 covers" dropdown has it for the days
every 350 Concur has seen gets rejected on its Travel Class Name and the
simplest fix is to stop sending it, without touching the per-employee flags or
the map.

`Preview` builds the whole file and counts it without writing anything, and
lists everyone being held back with the reason. `Write extract` drops it in the
outbound folder — `output/adp_concur/` unless `outbound_dir` names the folder
SAP Concur actually collects from — and records it in `ADP_Concur_Extracts` so
a file can be traced back to what was in it.

Delimiter, quoting, line ending and file name are all configurable. The default
is comma-delimited, CRLF, quoted only where a value contains the delimiter, and
`FMG_Concur_Employee_{stamp}.txt`. Record order defaults to all 305s, then all
350s, then all 360s: Concur creates the profile from the 305, and the 350 and
360 attach to a profile that has to exist already. `by_employee` groups a
person's three records together instead.

## Where the 305 / 350 / 360 columns come from

`FIELD_MAP` in `ADP_Concur_Map.py` is read straight out of the workbook's own
cell formulas, so that dictionary *is* the specification of the extract — a
position that is not in it is one the template leaves empty. The record tabs in
the front end show only the positions the map fills by default, because a
137-column table of mostly-empty cells is not readable; **Show all 137
positions** gives the full width.

One thing to check with SAP while you are in there: the 305's **Custom 3
Expense Profile** (column X) is pointed at the *Travel* map, not the Expense
map, so everyone comes out `General` / `VIP` rather than `Default` / `Grade 20`
/ `Officers`. The Concur Profile column (AF) is computed in the workbook and
then never used by any of the three tabs. That is reproduced faithfully here
rather than quietly corrected — if it is a slip, changing the one line in
`FIELD_MAP` fixes every record.

## Tables

| Table | Contents |
|-------|----------|
| `ADP_Concur_Employees` | One row per person: the raw ADP columns, the derived Concur values, the source (`adp` / `manual`), the include flags and the duplicate note |
| `ADP_Concur_OrgMap` | Business unit + department → org units, default language, currency |
| `ADP_Concur_StatusMap` | Position Status → `Y` / `N` |
| `ADP_Concur_CountryMap` | ADP country → the two-character code |
| `ADP_Concur_LanguageMap` | Language description → the `xx_` locale stem |
| `ADP_Concur_SalaryMap` | Pay grade → expense and travel profiles |
| `ADP_Concur_SupervisorMap` | The supervisor exceptions, keyed on the employee |
| `ADP_Concur_Layouts` | The 305 / 350 / 360 column layouts as the template shows them |
| `ADP_Concur_Selection` | Who is currently picked, and how each of them got picked |
| `ADP_Concur_Imports` | Every workbook loaded, with the per-sheet counts |
| `ADP_Concur_Exceptions` | Rebuilt on every derive — what is wrong and how badly |
| `ADP_Concur_Extracts` | Every flat file written, with its counts |
| `ADP_Concur_RoleMap` | The Role Assignment Map — reference only; nothing derives from it yet |
| `ADP_Concur_InvoiceMap` | The Invoice Exception Map — who gets Concur Invoice by name |
| `ADP_Concur_Loads` | Every Concur result read back, and which extract it answers |
| `ADP_Concur_Results` | One row per message Concur returned, with the person it belongs to |

Each load runs inside one transaction.

## The tabs

| Tab | Shows |
|-----|-------|
| **Employees** | Everyone, searchable, filterable by status, source, business unit, problem and whether they are picked; click a row to edit |
| **305 / 350 / 360 / 700** | The records exactly as the extract will write them, under the template's own headings |
| **Hierarchy** | The org tree, filterable by status and selection; one person's complete chain up and down, and where the chain breaks |
| **Maps** | The six lookup tables, editable in place — every edit re-derives; the Supervisor Map is validated row by row |
| **Exceptions** | **Fixes** — everything wrong, grouped by what would fix it, with the fix on the card — or every error and warning as a raw list |
| **Extract** | The Login ID rule, the file shape, preview, write, and clearing employees by source |
| **Results** | What Concur said back about a load — by cause or by person, with every message tied to a named employee |

Clicking an employee opens the editor: the exceptions against that person, the
include flags, the ADP fields, the derived values read-only beside them, and
the three records the person produces rendered as they will appear in the file.

## The look

The page is styled after doppiogroup.com: a near-black chrome bar carrying the
Doppio wordmark, white content on hairline `#e6e7e9` rules, Poppins for
headings and Nunito Sans for body, 8px cards and 6px buttons, and the Doppio
red `#d40814`.

That red forced one decision worth writing down. On the site red means *act* —
every button is red. In here red already meant *this is broken*. So red is
given back to actions, and the two states are told apart by treatment rather
than hue: a solid red fill is a button, a pale red wash under a deep red rule
is an error. Picked rows get a warmer, much lighter tint (`#fffafa`) with a red
left rule, so seventy-seven selected people read as chosen rather than as
seventy-seven problems.

The fonts load from Google Fonts and fall back to the system stack, so the app
still renders correctly on a machine with no network.

## What the new workbook changed

Dropping the 3 September cut in place of the 31 August one:

* **`TBDTUN` is gone.** Laura Oleksa's supervisor is now `203011`, and
  Babatunde Adewalure is on the non-US tab — so that chain resolves.
* **`207199` is gone too.** Mehmet Erdogan is on the non-US tab as well, which
  is why the two US employees pointing at him could never be found: he was
  never missing, he was in a population the old workbook did not carry.
* **The Login ID question is answered, at least for UAT.** Both 305 tabs build
  them as `=<work email>&".new.uat"`. `.new.uat` reads like a UAT-only value,
  so it is worth confirming before a production load.
* **12 people reported across the roster line.** That was a load-order problem
  while the extract was two files; it is not one any more. The current cut has
  26 such pairs and the sort handles all of them.
* **The passwords on the non-US tab are a fill series.** 82 people, 82
  different passwords, `Welcome01` to `Welcome82` with no gaps — that is Excel
  autofill, not a password policy, and the 100 record's `TEXT` setting is what
  would send them to Concur. Every one of them carries a warning.
* **4 people on the "Non US Non SE" tab are US or SE** — 3 `US` and 1 `SE` —
  which contradicts the tab's own name. Worth a look before the split is
  trusted.
* **`Sheet2`** is a pivot of the duplicate payroll companies and is ignored.

## Checked against a file that actually loaded

`305 360 import FMG 07.24.26.txt` is a real July load. Comparing it against
what this produces, for the 123 US people present in it, this load and Kelly's
current 305 tab:

**Structure matches exactly** — 137 fields on a 305, 35 on a 360, 7 on the 100,
CRLF, no BOM, comma-delimited, and the 360's ten populated positions are the
same ten. The **100 record now matches byte for byte**: `100,0,SSO,UPDATE,EN,N,N`
is what that file carried, and those are the defaults here — a proven load
beats a documentation default, so `TEXT`/`en`/`Y`/`Y` was wrong and is gone.

**Every value difference is explained**, and they fall into three groups.

*Two were bugs on my side*, both from building the field map off the August
workbook and not re-reading the September one. Both are fixed:

| Column | Was | Now |
|--------|-----|-----|
| `AP` Custom 21 Concur Expense Group Hierarchy | the pay grade code | Org Unit 1 — which is what the July file has too |
| `F` Login ID | the plain work email | the work email plus `.new.uat`, which **both** 305 tabs now append |

*The rest of my differences from Kelly's tab are the three documented
corrections doing their job.* Every one traces to the duplicate-payroll rule
picking the live employment record where her tab shows the terminated one, the
Org Map exact match, or the Supervisor Map being honoured. On the 123 people,
agreement with her tab is 100% on twelve columns and never below 107/116 on the
rest.

Plus one of hers: her Login ID formula is `=<email>&".new.uat"`, and on a blank
email that produces the literal string **`.new.uat`** as somebody's Login ID —
**41 rows on her 305 tab**. This holds those people out instead.

*The remainder is the template moving on since July*, which is expected — but
three of them are worth a question, because July populated a column the current
tab leaves empty:

| Column | July | Now |
|--------|------|-----|
| `J` Ctry Code | `US` | empty |
| `BJ` Employee ID of the Invoice Approver | the approver, same as `BG` | empty |
| `CI` Custom 22 Concur Invoice Group Hierarchy | the org unit | empty |

Everything else in that group is the template gaining fields rather than losing
them — `C` Middle Name, `X` Custom 3, `Z` Custom 5 and `CE` Future Use 2 are all
populated now and were not in July — plus data that has simply changed since
(supervisors, statuses, and ADP's pay grade codes going from `H0008` to `H08`).

One more: **the July file carried no 350 records at all**, and the workbook has
since grown a 350 tab. `records_by_source` in the config decides that, so
dropping `"350"` from the `adp` list goes back to exactly what loaded before.

## What Concur said back

`Employee_p0010945e24e Run-18` is the result of loading the US file: 369
records in, **132 errors and 256 warnings** out. Every question left open by
the section above got answered, and three of the answers changed the code.

Concur's result is a spreadsheet of Level, Record Identifier and Message, and
the Record Identifier is a **line number in the file you sent** — so on its own
it is unreadable, because "Record 139" is nobody. `ADP_Concur_Result.py` joins
it back through the extract to the record type and the person; the Results tab
is that join, and [The Results tab](#the-results-tab) explains how to use it.

### The four things that failed

**Every 350 record — all 118.** `Rule class 'General' is invalid`. The 350's
Travel Class Name comes from the Salary Map's Travel column, whose values are
`General`, `Senior Leadership` and `VIP`, and this tenant has no rule classes
by those names. Nothing in the code can fix that: either Concur gains the rule
classes or the 350 leaves the load. The July file carried none, which is the
strongest hint available.

**13 records rejected for a missing Ctry Code.** Column J had no formula on the
US 305 tab, so it went out empty for all 132 — and only the 13 people Concur
was *creating* failed, because `UPDATE` does not re-require it for anybody who
already exists. A gap that only shows on new starters is exactly the kind that
survives seventeen runs. **Fixed**: J is now filled from `legal_country`, the
same derived value that column I has been gluing onto the end of the locale
code all along. `en_US` in I beside an empty J was always a contradiction.

**Custom 3 was reading the wrong Salary Map column.** The map has C `Expense
Map` (`Default` / `Grade 20` / `Officers`) and D `Travel Map` (`General` /
`Senior Leadership` / `VIP`). The US 305 tab's `='1'!AG5` is `VLOOKUP(...,4)` —
the Travel column — in a field headed *Custom 3 Expense Profie*. The non-US tab
uses `VLOOKUP(...,3)`, the Expense column. That disagreement was reproduced
rather than reconciled while it was still a question for SAP; Concur answered it
by rejecting the Travel value 120 times. **Fixed**: both rosters read the
Expense column, and `FIELD_MAP_BY_ROSTER` is empty again.

**One employee could not be placed in the expense hierarchy.** Carla Cunningham
(007633), carrying `BU Description Not Mapped` into `Segment1`. She reached the
file because unmapped values used to be downgraded to a warning for terminated
people — my reasoning was that a leaver's stale BU is nobody's problem. Concur
validates the hierarchy node before it looks at whether the person is active,
so being a leaver bought nothing. **Fixed**: an unmapped value is an error
whoever it belongs to. The status is still named in the message, because for a
leaver the cheap fix is usually dropping them from the load rather than mapping
a BU that no longer exists.

### The warnings, and what they tell you about UPDATE

`Custom 5` drew 120 warnings — the Org Unit 2 department codes do not resolve
to list items, and July left Custom 5 blank. Whether the tenant is missing the
list or the column is unwanted is a question for SAP, so it is answered in
config rather than in code: `blank_fields` on the Extract tab empties a column
whatever the field map says.

The rest were all predicted here: six people whose Concur Login ID differs from
the `.new.uat` we send (a 305 cannot change a Login ID — that needs a 320),
three cross-roster approvers, and `207199` again.

Two **circular reference** warnings looked at first like a hierarchy bug and are
not. Our data is clean — Bair (207709) has a blank supervisor exactly as the
Supervisor Map says. The loop is inside Concur: we write `BY` BI Manager Employee
ID blank, and `UPDATE` never clears a field, so Bair's profile is still carrying
a manager from an earlier load.

That is the general lesson, and it applies to more than one column: **under
`UPDATE`, every field left blank is a field that can never be corrected.**
Right now that is BI Manager, Invoice Approver and Custom 22. If this load is
meant to be the system of record for the reporting chain, `BY` should be
populated from the supervisor rather than left empty.

## The Results tab

Drop the `.xls` Concur sends back onto the same drop zone as the workbook —
which file it is is decided by what is inside it, because Concur names its
results `Employee_p0010945e24e Run-18.xls` and nobody keeps that.

The tab shows one load two ways. **By cause** groups the messages by what went
wrong, each with the people named, one message quoted verbatim (that is what
gets pasted into a support ticket) and what to do about it. **By person** turns
it inside out: who to go and look at, worst first, clicking through to the
editor. Either way, **Pick these people** hands a group straight to the
selection, so "thirteen missing a country code" becomes thirteen ticked rows on
Employees and a retry file, without anybody copying a file number.

**The join is checked, not assumed.** Matching a result to an extract on line
count alone is weak — re-export the same people and you get another file of
exactly the same size, and one wrong join silently puts the wrong names on every
error in the run. So the extract's employee IDs are compared against the ones
Concur quotes in its messages, and the extract that agrees most wins. On Run-18
that is 383 agreements against 4 disagreements; pointed deliberately at the
non-US file it is 0 against 245, and at a US file one line shorter, 45 against
338. Both are refused.

The tolerance is not zero, because Concur is not exact either: four Run-18
messages carried line 369 — the last line of the file — while naming two people
who sit hundreds of lines earlier. Where a message says who it is about, that is
who it is about; the line number only supplies the record type.

## What the 15 September workbook changed, and what it revealed

The rename was the easy part — nothing read the tab names. The real changes:

* **Two systems of record instead of two geographies.** `roster` is gone;
  `source` decides everything. 134 people from ADP, 361 from UKG.
* **The 700 record**, and the two derived values behind it.
* **Custom 5 is now Org Unit 1**, not the department code. The workbook changed
  it after Concur rejected the department codes as invalid list items — so the
  fix we were waiting on arrived in the sheet.
* **Custom 3 reads the Expense column on both tabs.** The transposition that
  had the ADP tab reading the Travel Map is gone, which is what Concur's Run-20
  result had already told us.

Three things the sheet is doing that are worth raising with Kelly:

* **The UKG 305's `Active (Y/N)` column sends `A`, not `Y`.** It is
  `=UKG!BJ2`, and `BJ` is `=Y2` — the Employment Status *Code*. Concur wants
  `Y` or `N`. All 361 UKG records would carry an invalid Active flag. This app
  puts UKG's Employment Status through the Status Map like any other source and
  sends `Y`, so the extract is right and the sheet is not.
* **`Leave of absence` is not in the Status Map.** ADP says `Leave`, UKG says
  `Leave of absence`, and the map has only the ADP wording — so two UKG people
  come out `Position Status Not Mapped` and are held out of the file. One map
  row fixes it.
* **Three Invoice Exception Map entries cannot be matched by the sheet's own
  VLOOKUP.** Two file numbers carry trailing spaces (`'203011   '`,
  `'213174   '`) and one is stored as a number rather than text (`212705`), so
  Excel misses all three and quietly falls through to "no access". They are
  trimmed on the way in here, which is why Babatunde Adewalure, Robert Schortz
  and Bonnetta Warren get the invoice access the map says they should have and
  the workbook does not give them.

### Checked against the workbook's own answers

Every derived value was compared against the cells the workbook itself
computed — 134 ADP rows across thirteen columns, 361 UKG rows across nine.
Everything matches except four things, all deliberate:

| Difference | Count | Why |
|---|---|---|
| `org_unit_2`, `supervisor_id` — workbook `0`, here blank | 56 | `=AI2` on an empty cell is `0` in Excel. An artifact, not a value. |
| `invoice_access` for 203011 and 212705 | 2 | The trim fix above. |
| `invoice_limit` for 203011 | 1 | Follows from their access being `Y`. |

## Still open

* **The Login ID shape.** Set once SAP confirms it. Everything else is ready.
* **`207199` and `TBDTUN`.** Two supervisor IDs nobody in this load matched —
  worth re-checking now that UKG is loaded, since UKG is where several of the
  previously missing approvers turned out to live.
* **Brandon Hibbitts' Supervisor ID reads `000207199N`** — ten characters, and
  not the three-letter-prefix shape every other row uses. `MID(H,4,6)` happens
  to pull `207199` out of it correctly, but by luck rather than by rule. Worth
  asking ADP whether that column is reliably formatted.
* **13 roots.** Ten people have no supervisor and no reports at all. Real tops,
  or ADP simply having no manager on the record?
* **The 34 people with no work email.** A `sources` fallback brings them in.
* **The Country Map now holds eleven countries**, so most of the world
  resolves. Anything outside it still comes out `Country Not Mapped`, and that
  is a map row rather than code.
* **400 and 320.** The workbook carries three tabs this load ignores — two
  near-identical `400 ADP Exp Processor Non AP` and a `320 One Time`. They are
  recognised and named as skipped rather than silently dropped. If either is
  wanted, the second question is which of the two 400 tabs is authoritative.

Opened by Run-18, and none of them answerable from here:

* **The 350 records.** Every one was rejected on its Travel Class Name. Either
  Concur gains rule classes called `General`, `Senior Leadership` and `VIP`, or
  `"350"` comes out of `records_by_source` — which is what July did.
* **Custom 5, and Custom 3 after the fix.** Custom 5 does not resolve to a list
  item and July left it blank; Custom 3 now sends `Default` rather than the
  rejected `General`, but whether *that* is a list item is unconfirmed. Both are
  one line in `blank_fields` if the answer is to stop sending them.
* **BI Manager.** Left blank, so under `UPDATE` a stale manager in Concur can
  never be corrected — which is where both circular-reference warnings came
  from. Populating `BY` from the supervisor is a one-line change; it is not made
  because it decides that this load owns the reporting chain, and that is a
  decision rather than a fix. `BJ` Invoice Approver and `CI` Custom 22 are the
  same question, and July populated both.
* **The six differing Login IDs.** Expected while `.new.uat` is in play, but
  they will need a 320 record or User Administration before a production run.
