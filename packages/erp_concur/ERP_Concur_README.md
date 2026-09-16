# ERP_Concur

Cut the Concur import files down to a vendor, or to a set of purchase orders, and write
the slice back out unchanged.

Drop the three files `ConcurExtracts/ConcurExtracts/PurchaseOrder.dtsx` produces — the
purchase order file, the vendor file and the PO receipt file — on the page. Pick a vendor
and you get every order that names it and every receipt against those orders. Pick orders
and you get the vendors behind them. Press **Write extract** and three smaller files land
in `output/`, every line byte for byte the line that came in.

Drop the **import results** report Concur sends back on the same page and the orders it
rejected are picked for you — one button, and the next extract is the reprocessing file.
See **What Concur said** below.

```
python3 ERP_Concur_App.py            # http://127.0.0.1:5063
python3 ERP_Concur_Export.py --vendor 008820044 --verify     # same thing from the CLI
```

Port **5063**. 5057–5059 are `m3_security`, `adp_concur` and `mig_sync`, 5060/5061 are the
SIP ports Chrome blocks, 5062 is `sheet_security`.

State lives in `~/sqlite/doppio.db` alongside the other tools' tables, under the
`ERP_Concur_` prefix. `--db` or `ERP_CONCUR_DB` moves it.

---

## The one promise this tool makes

**The subset is the source.** Every record written out is its original line, replayed from
the `raw` column, in the original order. Nothing is re-serialised from the parsed columns,
nothing is trimmed, no date is reformatted, no field is re-quoted, and the line endings of
each source file are reproduced — including whether the file ended with one.

That is the point. The question these cuts answer is *what does Concur do with the real
file, reduced to two purchase orders* — and an output that had quietly fixed the CHAR
padding or the date format would answer a different question. Ryerson's postal code goes
out as `44149␣␣␣␣␣` because that is what the extract contains.

Every write is checked. `verify()` looks up each emitted line in the raw text the parser
stored and reports anything it does not recognise; the Write extract button runs it and
says so in the toast, and `--verify` makes it the CLI's exit code. With both sample orders
picked, the purchase order cut is byte-identical to `input/po.txt`.

A record can be corrected before the cut - see **Editing records before the cut** below -
and for one that has been, "the source" means the database's `raw`, not the original
file's line anymore: the edit rebuilds `raw` from the edited fields and `verify()` checks
against that. Everything untouched is still replayed byte for byte from the file that was
dropped.

The files are deliberately **not** named the way the dtsx names its own. Each carries a
label and a timestamp, and the manifest beside them carries `_Selection_`:

```
POInvoice_<label>_<stamp>.txt
vendor_<label>_<stamp>.txt
purch_receipt_<label>_<stamp>.txt
FMG_Concur_Selection_<label>_<stamp>_manifest.txt
```

A partial file called `purchase_order_import_t1209783x4lw_.txt` sitting in a Concur pickup
folder gets loaded as the nightly one. `adp_concur` sidesteps the same trap the same way.
The manifest beside them lists the vendors and orders that were in the cut, how each got
picked, and the findings that cover it — because these get mailed to somebody who loads
them a week later, and "which orders were in the one that failed" is otherwise a question
only this database can answer.

---

## Files

| File | What it is |
|---|---|
| `ERP_Concur_App.py` | Flask routes; every one returns JSON. |
| `templates/ERP_Concur_Index.html` | The whole page — one file, house style. |
| `ERP_Concur_Spec.py` | **Generated.** The seven record layouts, field by field. |
| `ERP_Concur_GenSpec.py` | The generator that emits it from the spec workbook. |
| `ERP_Concur_Parse.py` | Reads the three files; decides which is which. |
| `ERP_Concur_Findings.py` | Every check. Rebuilt from scratch on each load. |
| `ERP_Concur_Results.py` | Reads the import results report and resolves it back to orders. |
| `ERP_Concur_Db.py` | Schema, and the vendor ⇄ order selection rules. |
| `ERP_Concur_Export.py` | The writer and the CLI. |
| `input/` | Where files are dropped; `Load from input/` reads it. Holds the two sample orders. |
| `output/` | Where extracts land. |

`ERP_Concur_Spec.py` is generated from
`~/Doppio/FMG/Concur_Record_Type_Specifications.xlsx` — the mapping workbook built from
the SAP Help Portal's Concur Invoice Professional Edition Administration Guides, version
2026_08. Re-run `python3 ERP_Concur_GenSpec.py` if that workbook gains a record type or a
field; it asserts the seven widths (4 / 62 / 61 / 20 / 20 / 52 / 29) and fails rather than
emit a layout that disagrees with the guide. Do not hand-edit the generated file.

---

## Three things about these files

**A record type does not identify a layout.** Record type `200` is the vendor record (62
fields), the purchase order header (61 fields) *and* the receipt record (29 fields) — the
same literal in three files, because the receipt writer reuses the header's constant. So a
file is identified by the signature of what is inside it, not by its name and not by the
first record type it happens to start with. Every record votes; the kind with the most
wins. A file with no recognisable record at all is refused with a message saying what it
did contain.

**Records belonging to one order are grouped by position, not by key.** The dtsx writes,
per order: the `300` lines, then the `210` bill-to, then the `220` ship-to, then the `200`
header. *Header last.* So the records belonging to a header are the ones standing in front
of it, and that is the link used here. Parsing the order number back out of a line's
External ID would look more explicit and be strictly worse: the charge branch of
`p_concurinvoicepo_get_poline` emits two lines with the **same** External ID, and the
`210`/`220` rows carry the vendor code in their External ID rather than anything resembling
an order number. Position is the only link that holds for every record in the file. When a
line's External ID disagrees with its group's entity + order number, the grouping stands
and a `line_key_prefix` finding says so.

**Quoting is the dtsx's `Escape()`, which does not trim.** A field is quoted only when it
contains a comma, a double quote, CR or LF. `csv.reader` with the default dialect reads
exactly that — so padding from the ERP's CHAR columns arrives *inside* the value, and is
preserved all the way back out.

---

## Picking

The selection is two sets in one table, `ERP_Concur_Selection`: vendors by vendor code,
orders by order number. It lives in the database, not the browser — it survives a reload,
every list filters on it with a join instead of a query string carrying a hundred keys, and
the extract and the screen read the same rows so they cannot disagree.

`propagate_selection()` closes it over the link, both ways:

- picking a **vendor** by name pulls in every order whose header names it;
- picking an **order** pulls in the vendor its header names, but only as a passenger — that
  vendor does not, in turn, pull in the rest of its orders. Only a vendor picked by name
  (from the Vendors tab, or `--vendor`) does that. Without this, picking one order from a
  large vendor would quietly select every other order that vendor has, which is not what
  picking one order means.

Each row carries `direct` — 1 when you asked for that thing by name, 0 when it came along
for the ride — and that is what makes an un-pick the exact inverse of the pick it mirrors:

- dropping a vendor drops its orders, *unless* an order was picked by hand or belongs to
  another picked vendor;
- dropping an order drops its vendor *only* if that vendor was never picked by hand and has
  no other picked order left.

The toast says what came along, every time. A silent extra order in an extract is a
surprise three days later.

Lines, bill-to/ship-to rows and receipts are **not** picked separately — they follow their
order. Receipts are taken by order number, not by line: a receipt whose Line Item External
ID resolves to nothing is exactly the error-1001 case worth reproducing, so it stays in the
cut and gets reported instead of filtered out. Receipts for an order that is not in the
purchase order file at all are left out; there is nothing for them to attach to.

Matching an order to a vendor for *selection* is done on vendor code alone, deliberately —
not on the `(vendor_code, vendor_address_code)` pair Concur matches on. A header whose pair
does not resolve is the error-2000 case, and it has to end up in the extract *together with
its vendor row* so the mismatch can be looked at. The pair is checked as a finding instead,
and shown in the **Pair** column on the Purchase orders tab.

Re-dropping a corrected file replaces that file and everything parsed from it, but leaves
the selection alone. Anything picked that is no longer in the file is listed on the Extract
tab and in the manifest — an extract short of what was asked for should say so, not round
down quietly.

---

## Editing records before the cut

Sometimes the fix for a finding is not "drop this order," it is "the vendor address code is
wrong, correct it and re-cut." Every record - vendor, PO header, line, bill-to/ship-to
address and receipt - can be edited from its detail panel (the "Field by field" table opens
an input per field; a header's fields are also reachable straight from the purchase order
panel). Position 1, Record Type, is the one field that never opens for editing - it is what
tells this tool which table the record belongs to and how to group it, not something worth
letting a click corrupt.

Saving an edit (`ERP_Concur_Parse.edit_record`) re-joins the record's fields into one line
with the same quoting rule `split_record()` reads with - a field is quoted only when it
holds a comma, a double quote, CR or LF, with an embedded quote doubled - and writes that
back into `raw`. It also refreshes whichever named columns the loader for that record type
would have filled from the same positions (`VENDOR_MAP`, `PO_HEADER_MAP`, `PO_LINE_MAP`,
`PO_ADDRESS_MAP`, `RECEIPT_MAP`, `SETTINGS_MAP` in `ERP_Concur_Parse.py`), so a search, a
sort or the vendor/order pairing check sees the new value exactly like one that had always
been there - and a PO line's `is_charge` flag is recomputed too, since it depends on the
account code and description that may have just changed. The findings pass reruns
immediately after, so a correction that fixes error 2000 shows the error gone without a
separate rebuild.

The one thing an edit does not touch is which file the record came from or where it sits in
the file's order - grouping is still positional (see below), and an edit only ever changes
the fields the loader reads out of one existing line.

---

## Findings

Two kinds, and nothing here blocks an extract.

**Per record, driven by the spec** — one loop over `ERP_Concur_Spec`, so it picks up a new
field the day the workbook gains one: field count against the guide's width, a required
field left empty, a value carrying padding, a date that is not `yyyy-mm-dd`.

**Across records, the three matches Concur actually performs:**

| Finding | Concur's answer |
|---|---|
| `vendor_not_found` / `vendor_pair_not_found` | Error 2000, "no vendor found for the supplied Vendor Code and Vendor Address Code" |
| `duplicate_line_external_id` → `receipt_line_ambiguous` | Error 1001 on the receipt, "The External ID is missing or invalid" |
| `receipt_line_not_found` | Error 1001 — a *lookup* failure, not a validation complaint |
| `no_expense_or_account` / `both_expense_and_account` | Error 5001, Field Code AccountCode |

**And what Concur replied**, when a results report is loaded: every row of it that is not
Info becomes a `concur_<code>` finding (`concur_5001`, `concur_4002`, `concur_sequence`,
`concur_not_imported`) against the order it resolves to, so the Purchase orders tab's error
count is what actually happened rather than what was predicted. Reading the two together is
the point: a `5001` this tool found *and* a `5001` that came back is the check confirmed,
and a rejection with no finding in front of it is a check missing here. Run 73 makes that
concrete — 22 lines were flagged `no_expense_or_account`, and Concur rejected 18 orders
with 5001.

Plus `address_gap`, `char_padding`, `date_format`, `po_without_lines`,
`record_without_header`, `duplicate_vendor_pair`, `zero_received_quantity`,
`no_goods_receipt_number`, `line_key_prefix`, `address_external_id`, `file_missing`,
`concur_report_mismatch`, `concur_unresolved_row`, `concur_count_mismatch`.

Severity is advisory: **error** means Concur will reject the record, **warning** means it
will load something wrong, **info** means it looks odd and is worth an eye. The table is
rebuilt from scratch on every load, so it always describes the files as they stand. Click
any finding to open the record it is about, field by field, with padding highlighted —
that highlight is the only way to see a trailing space in a postal code.

### What the two sample orders produce

`input/` holds the reference output captured 2026-09-11: M3 PO **2100003074** and JBA PO
**7710284**. Every record's field count matches the spec exactly; the problems are all in
content. Loading all three files gives 15 findings — 4 errors, 4 warnings, 7 info — and
they are the known defects, which makes this the regression case for any proc change:

- **2 × `duplicate_line_external_id`** — TOTAL TAX (line 3) and TOTAL FREIGHT (line 4) both
  ship `0088210000307400`. `line_number` was made unique with the `99…99` scheme;
  `external_id` was not.
- **2 × `receipt_line_not_found`** and **2 × `receipt_po_not_found`** — four receipt rows
  for PO 2100000215, which is not in this purchase order file. The receipt proc has no
  extract_id or changed-since filter, so it re-emits every receipt every run.
- **1 × `address_gap`** — ULINE ships Address Line 1 empty with the street in Line 2
  (`CIDADR.SAADR1` blank, street in `SAADR2`). The spec does not permit the gap.
- **1 × `char_padding`** — Ryerson's postal code is `44149` with five trailing spaces.
- **3 × `zero_received_quantity`** — JBA receipts with a received date and nothing received.
- **4 × `address_external_id`** — both the 210 and the 220 on both orders carry the *vendor
  code* where the spec wants an External ID. Reported as info, not a defect: it is how the
  procs are written rather than a one-off, and it wants confirming against the import spec.

One known defect is **not** detectable from the files alone: JBA's Goods Receipt Numbers
(`76771028404`, `76771028401`, …) are the `ROW_NUMBER()` fallback in
`p_concurinvoicepo_get_poreceipt`, which is not stable across runs. Nothing in a single
file says so — you need two runs to see it move. `no_goods_receipt_number` catches only the
case where even the fallback missed.

---

## What Concur said

A run comes back as a spreadsheet — `Purchase_Order_Import_<entity> Run-<n>.xls` — with
three columns: Level, Record Identifier, Message. Drop it on the same page as the extract
files. It is never written out and never part of a cut; it is read to answer one question,
**which orders have to go again**, and the answer is a selection.

**The Record Identifier is a line number.** It is the 1-based position of the record in the
purchase order file that was sent — the Info row that closes the report carries the line
count of the whole file, which is the quickest way to confirm the pairing. So the report is
resolved against the file it answers and nothing else: identifier 272 is the `200` header
at line 272, and the order it names is the order at that line.

That matters because **most messages name no order**. Of the 43 error rows in Run 73, 19
carry a purchase order number in their text (`Error Code: 5001`, `Error Code: 4002`); the
other 24 are twelve pairs of

```
The sequence of the record types for a purchase order is invalid.        ← the 210
Due to an error importing a purchase order record, the purchase          ← its 200 header
order was not imported.
```

with nothing in them but a line number. Both halves resolve through the positional grouping
the parser already uses — the `210` at 3235 belongs to the header at 3238 — so the pair
collapses onto one order. Nineteen named orders plus twelve resolved by line is 31, which is
exactly what the report's own summary line says failed. That agreement is the check worth
having: if the two readings ever disagree, the Results tab says so rather than rounding.

**Both readings are kept, and compared.** `stated_po` comes from the message, `resolved_po`
from the line number, and `aligned` is whether they agree where both are known. A report read
against the *wrong* run still resolves — to real orders that happen to sit at those lines —
and nothing else in either file would say so. The banner on the tab is red when they
disagree, and `concur_report_mismatch` says it again in the findings.

**Pick every failed order** takes all of them in one go, with the reason recorded against
each row as `failed in Run 73`, so the manifest beside the extract says why those
thirty-one orders were in a file together. Their vendors, lines and receipts come along
exactly as they do for an order picked by hand. Then **Write extract** as usual.

The report is replaced, not accumulated: one run is in play at a time. Re-dropping the
purchase order file re-resolves it, because every identifier in it means a different line
then.

Reading a `.xls` needs `xlrd`; `.xlsx` and a report pasted into a `.csv` are read with what
is already here. Which file is which is decided by content — a spreadsheet is always a
report, a text file only if it carries the report's headings — so there is one drop zone
and nothing to get wrong.

---

## The CLI

```bash
# everything for one vendor, checked, into output/
python3 ERP_Concur_Export.py --vendor 008820044 --label ULINE --verify

# two orders, ignoring whatever is picked in the app
python3 ERP_Concur_Export.py --replace --po 7710284 --po 2100003074

# whatever is picked in the app right now, vendor file only
python3 ERP_Concur_Export.py --kind vendor
```

`--vendor` and `--po` are repeatable and propagate exactly as the page does. Without either,
the selection is used as it stands. `--verify` exits non-zero if any written line is not
found in the source.

---

## Open questions

- **The 210/220 External ID.** Both carry the vendor code. Confirm against the current PO
  import spec whether Concur ties the address to the request through that field; if it
  does, both procs need the order number there instead.
- **Charge-line External IDs.** Giving charges their own namespace (`'99' + charge_sequence`)
  fixes the collision, but re-keys charge lines Concur has already accepted — they would
  load as new line items, not updates. Worth deciding before changing the proc.
- **Receipts for orders outside the PO extract's window.** Silently dropped from the cut
  here, and reported. Whether the receipt proc should filter them is a proc question.
