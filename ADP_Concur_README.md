# ADP_Concur_* — ADP → SAP Concur employee load

Takes the ADP report Bonetta built, holds it in `doppio.db`, applies the six
lookup tabs Kelly worked out in Excel, and writes the 305 / 350 / 360 flat file
for SAP Concur to collect.

All routines are prefixed `ADP_Concur_` so they group together in the folder,
the same way the `M3_Security_*` set does.

| File | Role |
|------|------|
| `ADP_Concur_Db.py` | Schema + connection helpers (`~/sqlite/doppio.db`) |
| `ADP_Concur_Import.py` | Loads the workbook — ADP sheet, six maps, three layouts |
| `ADP_Concur_Map.py` | The derivations, the Login ID rule, the record builders |
| `ADP_Concur_Hierarchy.py` | The supervisor chain — walks it, and finds where it breaks |
| `ADP_Concur_Fix.py` | Regroups the exceptions by what would fix them, and applies it |
| `ADP_Concur_Export.py` | Writes the flat file into the outbound folder |
| `ADP_Concur_App.py` | Flask front end — seven tabs, editing, extract |
| `templates/ADP_Concur_Index.html` | The single-page UI |
| `ADP_Concur_Config.json` | The Login ID rule, the scopes, the file shape |

## Quick start

```bash
pip install flask openpyxl              # only external dependencies
python ADP_Concur_App.py                # http://127.0.0.1:5058
```

Drop the workbook on the page. It is saved to `input/adp_concur/`, parsed into
`~/sqlite/doppio.db`, and every tab fills. Command line equivalents:

```bash
python ADP_Concur_Db.py                                   # create the tables
python ADP_Concur_Import.py "Concur Draft ... for Load.xlsx"
python ADP_Concur_Map.py                                  # re-derive, list the exceptions
python ADP_Concur_Export.py --dry-run                     # count without writing
python ADP_Concur_Export.py                               # write the flat file
python ADP_Concur_Export.py --held-back                   # who is being left out, and why

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

## What the workbook carries

Nothing depends on the tab names — a re-cut of the report will not keep them.
Each sheet is recognised by what its header row says, and the header row is
found by looking for it rather than assumed to be row 1. That matters for the
ADP sheet, which has three rows of Kelly's notes above the real headings.

| Sheet | Recognised by | Goes to |
|-------|---------------|---------|
| the ADP export | `Payroll Company Code` + `File Number` + `Position Status` | `ADP_Concur_Employees` |
| Status Map | `Position Status` + `Concur Status` | `ADP_Concur_StatusMap` |
| Country Map | `ADP Country` | `ADP_Concur_CountryMap` |
| Org Map | `Business Unit Description` + `Home Department Code` | `ADP_Concur_OrgMap` |
| Language Map | `Language` + `ADP Language` | `ADP_Concur_LanguageMap` |
| Salary Map | `Pay Grade Code` + `Expense Map` | `ADP_Concur_SalaryMap` |
| Supervisor Map | `Exception Employee Id` + `Supervisor ID` | `ADP_Concur_SupervisorMap` |
| 305 / 350 / 360 | first cell reads `Trx Type (nnn)` | `ADP_Concur_Layouts` |

**The lookup tabs are a full refresh** — dropping a workbook replaces them with
what it carries, because the maps are the workbook's job. **The employees merge
on File Number**, so a fresh ADP cut refreshes the people already held without
disturbing the ones keyed in by hand. Anyone held here and absent from the file
is left alone rather than deleted; a partial cut of ADP is a normal thing to be
handed. **Start again** on the Extract tab is the from-scratch reload.

Only the ADP columns are loaded. The derived ones are recomputed from the maps,
so a workbook carrying stale lookup results never puts them in the database.

The three record templates are captured as layouts rather than data, which is
what makes a new Concur template a matter of dropping a new file: the extract
takes its width, its column order and its field-width row from whatever the
workbook shows.

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
| AH | Legal Country | Country Map |
| AI | Locale Code | Language Map on Language Description, else the BU's default language, then the country appended — `en_` + `US` |
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

Two filters on the toolbar, and they combine: **status** (active, inactive, or
both) and **picked**. *Active only* is the one to reach for when reviewing
hierarchies — the leavers are noise, and Kelly's note about only loading the
active employees is the same instinct.

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
* **a blank ADP column is written onto the employee** — and this one is
  explicitly a stopgap. The card says so: the next import overwrites it with
  whatever ADP says by then, so it needs fixing at source as well.

**One deliberate change came with this.** The Supervisor Map now *wins* over
ADP's own Supervisor ID rather than only filling in when ADP is blank. A table
of exceptions that can fill a hole but cannot correct a wrong value is no use
for the thing it is most needed for. On the workbook as it stands this changes
nothing — of the seven map rows only Moavero also has an ADP value, and both
give `006769`.

The supervisor picker only offers people who are actually in the load, and the
server refuses an ID that is not, so the fix for a broken chain cannot itself
create a broken chain.

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
| **Added by hand** | Everything. They were keyed in here and are held nowhere else — including anyone created from the Fixes tab to repair a broken supervisor chain. No workbook will bring them back. |

Tick either or both. The panel says exactly what would go before anything goes
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

**The maps are not touched.** Clearing employees leaves the six lookup tables
and the captured layouts alone, which is the point of having filed a supervisor
correction in the Supervisor Map: it outlives the employees it was about, and
still applies after the next import. **Clear maps and layouts** beside it is
what throws those away, behind its own `CLEAR`.

One employee at a time, the editor has both: **Leave out of the load** is the
soft delete that keeps the row and simply stops writing it, and **Delete
outright** appears only on hand-keyed records — for an ADP row it would be
pointless, since the next drop returns them.

## The extract

One file carries all three record types, one record per line, every position
present whether or not it has a value — which is what *All fields must be
represented* on the 350 and 360 tabs means.

Who is in it comes from three places: the per-employee include flags, the
configured scope, and the exception list. Kelly's note about the 350 and 360
holding everyone is the `scope` setting — `active` on both is the default here,
`all` puts the terminated people back.

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

Each load runs inside one transaction.

## The tabs

| Tab | Shows |
|-----|-------|
| **Employees** | Everyone, searchable, filterable by status, source, business unit, problem and whether they are picked; click a row to edit |
| **305 / 350 / 360** | The records exactly as the extract will write them, under the template's own headings |
| **Hierarchy** | The org tree, filterable by status and selection; one person's complete chain up and down, and where the chain breaks |
| **Maps** | The six lookup tables, editable in place — every edit re-derives; the Supervisor Map is validated row by row |
| **Exceptions** | **Fixes** — everything wrong, grouped by what would fix it, with the fix on the card — or every error and warning as a raw list |
| **Extract** | The Login ID rule, the file shape, preview, write, and clearing employees by source |

Clicking an employee opens the editor: the exceptions against that person, the
include flags, the ADP fields, the derived values read-only beside them, and
the three records the person produces rendered as they will appear in the file.

## Still open

* **The Login ID shape.** Set once SAP confirms it. Everything else is ready.
* **`207199` and `TBDTUN`.** Two supervisor IDs nobody in this load matches —
  see [Where it breaks today](#where-it-breaks-today). One is probably a non-US
  record; the other is a placeholder in the Supervisor Map. Both are one click
  each on the Fixes tab once somebody confirms the right answer.
* **Brandon Hibbitts' Supervisor ID reads `000207199N`** — ten characters, and
  not the three-letter-prefix shape every other row uses. `MID(H,4,6)` happens
  to pull `207199` out of it correctly, but by luck rather than by rule. Worth
  asking ADP whether that column is reliably formatted.
* **13 roots.** Ten people have no supervisor and no reports at all. Real tops,
  or ADP simply having no manager on the record?
* **The 34 people with no work email.** A `sources` fallback brings them in.
* **Non-US employees.** The Country Map holds only `USA → US` and the Language
  Map only resolves through the BU default, so the rest of the world will come
  out `Legal Country Not Mapped`. Both are map rows, not code.
* **Custom 3 Expense Profile** pointing at the Travel map — see above.
