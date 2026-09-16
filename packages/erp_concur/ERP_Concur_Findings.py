"""
ERP_Concur_Findings - what is wrong with the files that are loaded.

Two kinds of check live here.

**Per record, driven by the spec.** Field count against the guide's width, a
required field left empty, a value carrying CHAR padding, a date that is not
yyyy-mm-dd. These need no knowledge of Concur beyond
Concur_Record_Type_Specifications.xlsx, so they are one loop over
ERP_Concur_Spec rather than a list of hand-written rules, and they pick up a
new field type the day the workbook gains one.

**Across records, the three matches Concur actually performs.** A purchase
order header is matched to a vendor on the exact pair (Vendor Code, Vendor
Address Code) - error 2000 when it misses. A receipt is matched to a line on
Line Item External ID, byte for byte - error 1001 when it misses, including
when it misses because two lines share one External ID. And a 300 line must
carry exactly one of Expense Type or Account Code - error 5001 when it carries
neither.

**What Concur said.** When an import results report is loaded, every row of
it that is not Info is added as a `concur_*` finding against the order it
resolves to - see ERP_Concur_Results for how a Record Identifier becomes an
order number. These are not predictions; they are the answer.

Nothing here blocks an extract. The point of the tool is to take a slice of
the real files to Concur and see what happens; a finding is what to look at
when it comes back, and the count on the toolbar is so you know what you are
shipping before you ship it. Severity is therefore advisory: `error` means
Concur will reject the record, `warning` means it will load something wrong,
`info` means it looks odd and is worth an eye.

The table is rebuilt from scratch every time, so it always describes the files
as they stand.
"""
from __future__ import annotations

import logging
import re
import sqlite3

import ERP_Concur_Results as results
import ERP_Concur_Spec as spec
from ERP_Concur_Parse import split_record

log = logging.getLogger("ERP_Concur_Findings")

DATE_RE = re.compile(r"^\d{4}-\d{2}-\d{2}$")

# Tables the per-record pass walks: table, primary key, spec key (or None when
# the record type decides it), and how to label a row in the findings list.
RECORD_SOURCES = [
    ("ERP_Concur_Settings",    "settings_key", "settings", "vendor_100",
     "Import Settings"),
    ("ERP_Concur_Vendors",     "vendor_key",   "vendor",   "vendor_200", None),
    ("ERP_Concur_PoHeaders",   "po_key",       "po",       "po_200",     None),
    ("ERP_Concur_PoLines",     "line_key",     "line",     "po_300",     None),
    ("ERP_Concur_PoLineAllocations", "alloc_key", "allocation", "po_400", None),
    ("ERP_Concur_PoAddresses", "addr_key",     "address",  None,         None),
    ("ERP_Concur_Receipts",    "receipt_key",  "receipt",  "receipt_200", None),
]


class Findings:
    """Collects findings and writes them in one go."""

    def __init__(self, conn: sqlite3.Connection):
        self.conn = conn
        self.rows: list[tuple] = []

    def add(self, severity: str, kind: str, scope: str, ref_key, message: str,
            label: str = "", field: str = "", po_number: str = "",
            vendor_code: str = "") -> None:
        self.rows.append((severity, kind, scope, ref_key, po_number,
                          vendor_code, label, field, message))

    def flush(self, commit: bool = True) -> int:
        self.conn.execute("DELETE FROM ERP_Concur_Findings")
        self.conn.executemany(
            "INSERT INTO ERP_Concur_Findings (severity, kind, scope, ref_key, "
            "po_number, vendor_code, label, field, message) "
            "VALUES (?,?,?,?,?,?,?,?,?)", self.rows)
        if commit:
            self.conn.commit()
        return len(self.rows)


def _label(scope: str, row: dict) -> str:
    if scope == "vendor":
        return f"{row.get('vendor_code')} {row.get('vendor_name') or ''}".strip()
    if scope == "po":
        return f"PO {row.get('po_number')}"
    if scope == "line":
        return f"line {row.get('line_number')} of {row.get('external_id')}"
    if scope == "allocation":
        return f"allocation {row.get('amount')} (line {row.get('line_key')})"
    if scope == "address":
        return f"{row.get('record_type')} {row.get('external_id')}"
    if scope == "receipt":
        return f"receipt {row.get('goods_receipt_number') or row.get('line_no')}"
    return "Import Settings"


# --------------------------------------------------------- the spec-driven pass


def _per_record(f: Findings) -> None:
    for table, pk, scope, spec_key, fixed_label in RECORD_SOURCES:
        for r in f.conn.execute(f"SELECT * FROM {table}"):
            row = dict(r)
            key = row[pk]
            sk = spec_key or ("po_210" if row.get("record_type") == "210"
                              else "po_220")
            fields = spec.fields(sk)
            values = split_record(row["raw"])
            label = fixed_label or _label(scope, row)
            po = str(row.get("po_number") or "")
            ven = str(row.get("vendor_code") or "")
            common = {"label": label, "po_number": po, "vendor_code": ven}

            width = len(fields)
            if len(values) != width:
                f.add("error", "field_count", scope, key,
                      f"{len(values)} fields, the {spec.label(sk)} record type "
                      f"takes {width}. Everything after the first missing or "
                      f"extra comma is in the wrong position.", **common)

            padded = []
            for i, value in enumerate(values):
                if not value or value == value.strip():
                    continue
                name = fields[i][0] if i < width else f"field {i + 1}"
                padded.append(f"{name} ({value!r})")
            if padded:
                f.add("warning", "char_padding", scope, key,
                      "Value carries padding the ERP's CHAR columns left on "
                      "it; Escape() only quotes on comma, quote, CR and LF, so "
                      "it reaches Concur as part of the value: "
                      + ", ".join(padded), **common)

            for i, (name, required, _proc) in enumerate(fields):
                value = values[i] if i < len(values) else ""
                if required == "Y" and not value.strip():
                    f.add("error", "required_empty", scope, key,
                          f"{name} is required by the {spec.label(sk)} spec "
                          "and is empty.", field=name, **common)
                if value.strip() and name.endswith("Date") \
                        and not DATE_RE.match(value.strip()):
                    # The dtsx's Escape() ends in a bare ToString(), so a SQL
                    # date materialises in the server's own culture -
                    # '6/16/2026 12:00:00 AM' where Concur wants 2026-06-16.
                    # Only order_date is formatted in the C#; the rest are
                    # latent until they go non-null.
                    f.add("error", "date_format", scope, key,
                          f"{name} is {value!r}, not yyyy-mm-dd. Fix it in the "
                          "proc with convert(char(10), <col>, 23) - Escape() "
                          "formats nothing.", field=name, **common)


# ------------------------------------------------------------- the link passes


def _vendor_match(f: Findings) -> None:
    """The error-2000 check: a header's (code, address code) pair must exist."""
    if not f.conn.execute("SELECT 1 FROM ERP_Concur_Vendors LIMIT 1").fetchone():
        return                       # no vendor file loaded; nothing to match
    for r in f.conn.execute(
            "SELECT h.*, "
            " (SELECT COUNT(*) FROM ERP_Concur_Vendors v "
            "   WHERE v.vendor_code = h.vendor_code) AS by_code, "
            " (SELECT COUNT(*) FROM ERP_Concur_Vendors v "
            "   WHERE v.vendor_code = h.vendor_code "
            "     AND v.vendor_address_code = h.vendor_address_code) AS by_pair "
            "FROM ERP_Concur_PoHeaders h"):
        row = dict(r)
        common = {"label": f"PO {row['po_number']}",
                  "po_number": row["po_number"],
                  "vendor_code": row["vendor_code"]}
        if row["by_pair"]:
            continue
        if row["by_code"]:
            have = [x[0] for x in f.conn.execute(
                "SELECT vendor_address_code FROM ERP_Concur_Vendors "
                "WHERE vendor_code = ?", (row["vendor_code"],))]
            f.add("error", "vendor_pair_not_found", "po", row["po_key"],
                  f"Vendor {row['vendor_code']} is in the vendor file but not "
                  f"with address code {row['vendor_address_code']!r} "
                  f"(it has {', '.join(repr(h) for h in have)}). Concur matches "
                  "a purchase order on the exact pair: Error Code 2000, "
                  "'There was no vendor found for the supplied Vendor Code and "
                  "Vendor Address Code'.", field="Vendor Address Code", **common)
        else:
            f.add("error", "vendor_not_found", "po", row["po_key"],
                  f"Vendor {row['vendor_code']} is not in the vendor file at "
                  "all. Either the vendor proc filtered it out or the vendor "
                  "file has not been loaded into Concur yet - the dtsx writes "
                  "the purchase order file FIRST, so a vendor first seen on a "
                  "PO can be extracted before it has ever been sent.",
                  field="Vendor Code", **common)


def _line_keys(f: Findings) -> None:
    """The error-1001 groundwork: External ID has to be unique per line."""
    dupes = f.conn.execute(
        "SELECT external_id, COUNT(*) n FROM ERP_Concur_PoLines "
        "GROUP BY external_id HAVING COUNT(*) > 1").fetchall()
    for external_id, n in dupes:
        rows = [dict(x) for x in f.conn.execute(
            "SELECT l.*, h.po_number FROM ERP_Concur_PoLines l "
            "LEFT JOIN ERP_Concur_PoHeaders h ON h.po_key = l.po_key "
            "WHERE l.external_id = ?", (external_id,))]
        charges = sum(1 for x in rows if x["is_charge"])
        for row in rows:
            f.add("error", "duplicate_line_external_id", "line", row["line_key"],
                  f"{n} line records share External ID {external_id!r} "
                  f"(lines {', '.join(str(x['line_number']) for x in rows)})."
                  + (" These are the charge rows from the second leg of the "
                     "UNION in p_concurinvoicepo_get_poline: line_number was "
                     "made unique with the 99…99 scheme but external_id was "
                     "not, so charges reuse the real line's key. External ID is "
                     "Concur's line identity, and this is the usual cause of "
                     "Error Code 1001 on a later receipt."
                     if charges else
                     " External ID is Concur's line identity; two lines with "
                     "one key make every receipt against it ambiguous."),
                  label=f"line {row['line_number']} of {external_id}",
                  field="External ID", po_number=row.get("po_number") or "")

    for row in f.conn.execute(
            "SELECT l.line_key, l.external_id, l.line_number, l.entity_id, "
            "h.po_number FROM ERP_Concur_PoLines l "
            "JOIN ERP_Concur_PoHeaders h ON h.po_key = l.po_key"):
        row = dict(row)
        prefix = f"{row['entity_id'] or ''}{row['po_number']}"
        if prefix and not row["external_id"].startswith(prefix):
            f.add("warning", "line_key_prefix", "line", row["line_key"],
                  f"External ID {row['external_id']!r} does not start with this "
                  f"order's entity and number ({prefix!r}). The line is grouped "
                  "with this header by its position in the file - header last, "
                  "the way the dtsx writes it - so the grouping stands, but the "
                  "key does not agree with it.",
                  label=f"line {row['line_number']} of {row['external_id']}",
                  field="External ID", po_number=row["po_number"])


def _expense_or_account(f: Findings) -> None:
    """Exactly one of Expense Type / Account Code - the error-5001 rule."""
    for row in f.conn.execute(
            "SELECT l.line_key, l.line_number, l.external_id, l.expense_type, "
            "l.account_code, l.is_charge, l.description, h.po_number "
            "FROM ERP_Concur_PoLines l "
            "LEFT JOIN ERP_Concur_PoHeaders h ON h.po_key = l.po_key"):
        row = dict(row)
        et = (row["expense_type"] or "").strip()
        ac = (row["account_code"] or "").strip()
        common = {"label": f"line {row['line_number']} of {row['external_id']}",
                  "po_number": row.get("po_number") or ""}
        if not et and not ac:
            f.add("error", "no_expense_or_account", "line", row["line_key"],
                  "Neither Expense Type nor Account Code is populated - the one "
                  "combination Concur rejects: Error Code 5001, Field Code "
                  "AccountCode."
                  + (" This is a charge row, where account_code is a CASE over "
                     "the description with no default - any description that is "
                     "not TAX, FREIGHT or OTHER lands here."
                     if row["is_charge"] else ""),
                  field="Account Code", **common)
        elif et and ac:
            f.add("error", "both_expense_and_account", "line", row["line_key"],
                  f"Both Expense Type ({et!r}) and Account Code ({ac!r}) are "
                  "populated. Concur takes one or the other, never both.",
                  field="Expense Type", **common)


def _allocations(f: Findings) -> None:
    """
    400 checks. The spec workbook's own note on the PO - 400 tab says no proc
    in the deployed script emits this record type - but real production files
    carry plenty of them, so that note is stale (or this file did not come
    through the path it describes). Either way, a 400 here is real and worth
    checking, not a hypothetical.
    """
    for row in f.conn.execute(
            "SELECT a.*, h.po_number FROM ERP_Concur_PoLineAllocations a "
            "LEFT JOIN ERP_Concur_PoHeaders h ON h.po_key = a.po_key "
            "WHERE a.line_key IS NULL"):
        row = dict(row)
        f.add("error", "allocation_without_line", "allocation", row["alloc_key"],
              "This 400 record does not immediately precede a 300 line record, "
              "so there is no line for it to allocate. A 400 carries no "
              "External ID or line number of its own; the line right after it "
              "in the file is the only thing that says which line it splits.",
              label=f"allocation at line {row['line_no']}",
              po_number=row.get("po_number") or "")

    for row in f.conn.execute(
            "SELECT a.*, h.po_number FROM ERP_Concur_PoLineAllocations a "
            "LEFT JOIN ERP_Concur_PoHeaders h ON h.po_key = a.po_key "
            "WHERE a.synthetic = 1"):
        row = dict(row)
        f.add("info", "synthetic_allocation", "allocation", row["alloc_key"],
              f"This 400 (Amount {row['amount']}) was added by 'Fill missing "
              "allocations', not read from the source file. It is a trivial "
              "Quantity x Unit Price placeholder, not a real account split - "
              "this line had no real 400 to begin with - so it will go out in "
              "an extract exactly like this if the line is picked.",
              label=f"allocation at line {row['line_no']}",
              po_number=row.get("po_number") or "")

    for row in f.conn.execute(
            "SELECT l.line_key, l.line_number, l.external_id, l.quantity, "
            "l.unit_price, h.po_number, "
            "(SELECT COUNT(*) FROM ERP_Concur_PoLineAllocations a "
            "  WHERE a.line_key = l.line_key) n_alloc, "
            "(SELECT SUM(CAST(a.amount AS REAL)) FROM ERP_Concur_PoLineAllocations a "
            "  WHERE a.line_key = l.line_key AND a.amount <> '') alloc_total "
            "FROM ERP_Concur_PoLines l "
            "LEFT JOIN ERP_Concur_PoHeaders h ON h.po_key = l.po_key"):
        row = dict(row)
        if not row["n_alloc"]:
            continue
        qty = float_or_none(row["quantity"]) or 1.0
        price = float_or_none(row["unit_price"])
        extended = qty * price if price is not None else None
        total = row["alloc_total"]
        label = f"line {row['line_number']} of {row['external_id']}"
        common = {"label": label, "po_number": row.get("po_number") or ""}
        if extended is not None and total is not None \
                and round(extended - total, 2) != 0:
            f.add("warning", "allocation_amount_mismatch", "line", row["line_key"],
                  f"{row['n_alloc']} allocation(s) total {total}, which does not "
                  f"match Quantity x Unit Price ({extended}). In every other "
                  "instance of this in production data the two agree exactly, "
                  "so a mismatch here is worth checking against the source "
                  "file directly rather than assumed to be this tool's doing.",
                  field="Amount", **common)


def _receipt_match(f: Findings) -> None:
    """Receipts against headers and lines, the other half of error 1001."""
    have_po = bool(f.conn.execute(
        "SELECT 1 FROM ERP_Concur_PoHeaders LIMIT 1").fetchone())
    for r in f.conn.execute("SELECT * FROM ERP_Concur_Receipts"):
        row = dict(r)
        key = row["receipt_key"]
        label = (f"receipt {row['goods_receipt_number']}"
                 if row["goods_receipt_number"] else f"line {row['line_no']}")
        common = {"label": label, "po_number": row["po_number"]}
        if have_po:
            n_po = f.conn.execute(
                "SELECT COUNT(*) FROM ERP_Concur_PoHeaders WHERE po_number = ?",
                (row["po_number"],)).fetchone()[0]
            if not n_po:
                f.add("warning", "receipt_po_not_found", "receipt", key,
                      f"Purchase Order Number {row['po_number']!r} has no header "
                      "in the purchase order file. The receipt proc has no "
                      "extract_id or changed-since filter, so it re-emits every "
                      "receipt every run, including ones whose order is long "
                      "out of the PO extract's window.", **common)
            n_line = f.conn.execute(
                "SELECT COUNT(*) FROM ERP_Concur_PoLines WHERE external_id = ?",
                (row["line_item_external_id"],)).fetchone()[0]
            if not n_line:
                f.add("error", "receipt_line_not_found", "receipt", key,
                      f"Line Item External ID {row['line_item_external_id']!r} "
                      "matches no 300 line record. It has to equal the line's "
                      "External ID byte for byte: Error Code 1001, 'The External "
                      "ID is missing or invalid' - a lookup failure, not a "
                      "validation complaint.", field="Line Item External ID",
                      **common)
            elif n_line > 1:
                f.add("error", "receipt_line_ambiguous", "receipt", key,
                      f"Line Item External ID {row['line_item_external_id']!r} "
                      f"matches {n_line} line records. Concur cannot tell which "
                      "line this receipt is against; see the duplicate External "
                      "ID finding on those lines.",
                      field="Line Item External ID", **common)
        qty = (row["received_quantity"] or "").strip()
        if qty and float_or_none(qty) == 0:
            f.add("info", "zero_received_quantity", "receipt", key,
                  "Received Quantity is 0 with a received date of "
                  f"{row['received_date'] or '(blank)'} - an unreceived line "
                  "being emitted as a receipt.", field="Received Quantity",
                  **common)
        if not row["goods_receipt_number"].strip():
            f.add("warning", "no_goods_receipt_number", "receipt", key,
                  "Goods Receipt Number is empty. The proc falls back to "
                  "system_entity_id + po_num + po_seq + ROW_NUMBER(), which is "
                  "not stable across runs, so an empty one means even that "
                  "missed.", field="Goods Receipt Number", **common)


def float_or_none(value: str):
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def _addresses(f: Findings) -> None:
    """Address gaps."""
    for r in f.conn.execute(
            "SELECT a.*, h.po_number FROM ERP_Concur_PoAddresses a "
            "LEFT JOIN ERP_Concur_PoHeaders h ON h.po_key = a.po_key"):
        row = dict(r)
        kind = "Bill-to" if row["record_type"] == "210" else "Ship-to"
        common = {"label": f"{row['record_type']} {row['external_id']}",
                  "po_number": row.get("po_number") or ""}
        if not (row["address1"] or "").strip() and (
                (row["address2"] or "").strip() or (row["address3"] or "").strip()):
            f.add("warning", "address_gap", "address", row["addr_key"],
                  f"{kind} Address 1 is empty while a later address line is "
                  "filled. The spec is explicit that Address 1 comes first and "
                  "gaps are not permitted.", field="Address 1", **common)

    for r in f.conn.execute("SELECT * FROM ERP_Concur_Vendors"):
        row = dict(r)
        if not (row["address1"] or "").strip() and (
                (row["address2"] or "").strip() or (row["address3"] or "").strip()):
            f.add("warning", "address_gap", "vendor", row["vendor_key"],
                  "Address Line 1 is empty and the street is in Address Line 2 "
                  f"({row['address2']!r}). The vendor spec does not permit a "
                  "gap; in M3 this is CIDADR.SAADR1 blank with the street in "
                  "SAADR2.", field="Address Line 1",
                  label=f"{row['vendor_code']} {row['vendor_name'] or ''}".strip(),
                  vendor_code=row["vendor_code"])


def _structure(f: Findings) -> None:
    """Files that are missing, headers with no lines, records with no header."""
    loaded = {r[0] for r in f.conn.execute(
        "SELECT DISTINCT kind FROM ERP_Concur_Files")}
    for kind, label in (("po", "purchase order"), ("vendor", "vendor"),
                        ("receipt", "PO receipt")):
        if kind not in loaded:
            f.add("info", "file_missing", "file", None,
                  f"No {label} file is loaded, so the checks that need it are "
                  "not running.", label=f"{label} file")

    for r in f.conn.execute(
            "SELECT h.po_key, h.po_number, h.vendor_code, "
            " (SELECT COUNT(*) FROM ERP_Concur_PoLines l WHERE l.po_key = h.po_key) n "
            "FROM ERP_Concur_PoHeaders h"):
        row = dict(r)
        if not row["n"]:
            f.add("error", "po_without_lines", "po", row["po_key"],
                  "No 300 line records precede this header. A purchase order "
                  "with no lines has nothing to receive against.",
                  label=f"PO {row['po_number']}", po_number=row["po_number"],
                  vendor_code=row["vendor_code"])

    for table, pk, scope in (("ERP_Concur_PoLines", "line_key", "line"),
                             ("ERP_Concur_PoLineAllocations", "alloc_key",
                              "allocation"),
                             ("ERP_Concur_PoAddresses", "addr_key", "address")):
        for r in f.conn.execute(
                f"SELECT {pk}, line_no FROM {table} WHERE po_key IS NULL"):
            f.add("error", "record_without_header", scope, r[0],
                  f"This record is at line {r[1]} of the purchase order file "
                  "with no 200 header after it, so there is no order to attach "
                  "it to. The file looks truncated.", label=f"line {r[1]}")

    dupes = f.conn.execute(
        "SELECT vendor_code, vendor_address_code, COUNT(*) n "
        "FROM ERP_Concur_Vendors GROUP BY vendor_code, vendor_address_code "
        "HAVING COUNT(*) > 1").fetchall()
    for code, addr, n in dupes:
        for r in f.conn.execute(
                "SELECT vendor_key, vendor_name FROM ERP_Concur_Vendors "
                "WHERE vendor_code = ? AND vendor_address_code = ?", (code, addr)):
            f.add("warning", "duplicate_vendor_pair", "vendor", r[0],
                  f"{n} vendor records share the pair ({code}, {addr}). That "
                  "pair is the identity a purchase order is matched on, so the "
                  "later row wins and the earlier one is silently replaced.",
                  label=f"{code} {r[1] or ''}".strip(), vendor_code=code)


# --------------------------------------------------------------------- entry


def rebuild(conn: sqlite3.Connection, commit: bool = True) -> dict:
    """Re-run every check over the files that are loaded. Returns a summary."""
    f = Findings(conn)
    _per_record(f)
    _structure(f)
    _vendor_match(f)
    _line_keys(f)
    _expense_or_account(f)
    _allocations(f)
    _receipt_match(f)
    _addresses(f)
    # What Concur actually replied, if a results report is loaded. It goes in
    # the same table as the checks above on purpose: a 5001 this tool
    # predicted and a 5001 that came back are the same fault, once from the
    # file and once from the system, and a rejection with no finding in front
    # of it is a check that is missing here.
    results.add_findings(f)
    total = f.flush(commit=commit)
    summary = {"total": total}
    for sev in ("error", "warning", "info"):
        summary[sev] = sum(1 for r in f.rows if r[0] == sev)
    log.info("findings: %s", summary)
    return summary


def by_kind(conn: sqlite3.Connection) -> list[dict]:
    """The findings list's own summary - one row per kind, worst first."""
    return [dict(r) for r in conn.execute(
        "SELECT kind, severity, COUNT(*) n, MIN(message) sample "
        "FROM ERP_Concur_Findings GROUP BY kind, severity "
        "ORDER BY CASE severity WHEN 'error' THEN 0 WHEN 'warning' THEN 1 "
        "ELSE 2 END, n DESC")]
