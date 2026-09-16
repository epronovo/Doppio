"""
ERP_Concur_Parse - read the three Concur extract files into the database.

Three things are worth knowing about these files before reading this module.

**They are positional, and the record type alone does not identify a layout.**
Record type 200 is the vendor record (62 fields), the purchase order header
(61 fields) and the receipt record (29 fields) - the same literal in three
different files, because the receipt writer reuses the header's constant. So a
file is identified by the *signature* of what is in it, not by its name (the
dtsx names them purchase_order_import_<entity>_.txt and friends, but nobody
keeps those names by the time they reach a laptop) and not by the first record
type it happens to start with.

**Records belonging to one purchase order are grouped by position, not by key.**
PurchaseOrder.dtsx writes, per PO: the 300 lines, then the 210 bill-to, then
the 220 ship-to, then the 200 header. Header last. So the records belonging to
a header are the ones standing in front of it, and that is the link used here.
Parsing the PO number back out of a line's External ID would seem more
explicit but is strictly worse: the charge branch of
p_concurinvoicepo_get_poline emits two lines with the *same* External ID, and
the 210/220 rows carry the vendor code in their External ID rather than
anything resembling a PO number. Position is the only link that is true of
every record in the file.

A 400 (line item allocation) is positional too, and the wrong way round from
every other child record here: a 400 (or a run of them) comes *before* the 300
it allocates, not after - confirmed against real production data (PO
2100000017: ten 400s, each one Amount matching the Quantity x Unit Price of
the very next 300, not the one before it, to the cent). The spec gives a 400
no External ID or line number of its own, so the 300 immediately following a
run of them is the only thing that says which line they belong to.

**Every field is kept twice.** The raw line goes into `raw` byte for byte, and
the handful of fields the page searches and the selection needs are parsed out
beside it. The extract replays `raw`. Parsing is for looking; it is never the
source of what gets written.

Quoting follows the dtsx's own Escape(): a field is quoted only when it
contains a comma, a double quote, CR or LF, and an embedded quote is doubled.
csv.reader with the default dialect reads exactly that, so it is what is used
here - with the one caveat that Escape() does not trim, so CHAR padding from
the ERP arrives inside the value and is preserved.
"""
from __future__ import annotations

import csv
import hashlib
import io
import json
import logging
import sqlite3
from pathlib import Path

import ERP_Concur_Spec as spec
from ERP_Concur_Db import FILE_KINDS, replace_file

log = logging.getLogger("ERP_Concur_Parse")

# Which (record type, field count) pairs vote for which kind of file. The
# field count is what separates the three 200s. A count that matches nothing
# still votes on record type alone, so a file with a field added to every row
# is recognised and then reported field by field, rather than rejected.
SIGNATURES: dict[tuple[str, int], str] = {
    ("100", 4): "vendor",
    ("200", 62): "vendor",
    ("200", 61): "po",
    ("210", 20): "po",
    ("220", 20): "po",
    ("300", 52): "po",
    ("400", 42): "po",
    ("200", 29): "receipt",
}

# The fall-back vote, on record type alone.
BY_RECORD_TYPE: dict[str, str] = {
    "100": "vendor", "210": "po", "220": "po", "300": "po", "400": "po",
}

# Charge lines - the second leg of the UNION in p_concurinvoicepo_get_poline.
# They are recognised by the account codes that leg hardcodes, because that is
# the only thing on the record that says so: nothing in the file marks a
# charge as a charge. Used for the External ID collision finding and for the
# Lines tab's "charges only" filter, never to decide what gets extracted.
CHARGE_ACCOUNT_CODES = {"5555", "9999", "4444"}
CHARGE_DESCRIPTIONS = ("TOTAL TAX", "TOTAL FREIGHT", "TOTAL OTHER")

# 1-based raw CSV position of every column a loader pulls out of a record,
# named the same as the column it fills. The loaders below and edit_record()
# both read positions from these maps rather than from literals scattered
# twice, so the two can never drift apart.
VENDOR_MAP = {
    "vendor_code": 2, "vendor_address_id": 15, "vendor_address_code": 16,
    "vendor_name": 3, "currency": 17, "payment_term_days": 18,
    "address1": 28, "address2": 29, "address3": 30, "city": 31, "state": 32,
    "postal_code": 33, "country_code": 34, "contact_email": 53,
}
SETTINGS_MAP = {
    "error_threshold": 2, "default_country_code": 3, "pay_method_type": 4,
}
PO_HEADER_MAP = {
    "po_number": 2, "policy_external_id": 3, "currency_code": 4,
    "vendor_code": 5, "vendor_address_code": 6, "order_date": 7,
    "payment_terms": 12, "tax": 15, "shipping": 16, "ledger_code": 32,
    "entity_id": 38,
}
PO_LINE_MAP = {
    "external_id": 2, "line_number": 3, "supplier_part_id": 4,
    "expense_type": 7, "account_code": 8, "description": 9, "quantity": 10,
    "unit_price": 11, "uom": 13, "entity_id": 33,
}
PO_ADDRESS_MAP = {
    "external_id": 2, "name": 3, "address1": 4, "address2": 5,
    "address3": 6, "city": 7, "state": 8, "postal_code": 9,
    "country_code": 10,
}
PO_ALLOC_MAP = {
    "amount": 2,
}
RECEIPT_MAP = {
    "po_number": 2, "line_item_external_id": 3, "goods_receipt_number": 4,
    "delivery_slip_number": 5, "uom": 6, "received_quantity": 7,
    "received_date": 8, "is_deleted": 9,
}


class ParseError(Exception):
    """A file that cannot be read as a Concur extract at all."""


def read_lines(path: Path) -> tuple[list[str], str, bool]:
    """
    Split a file into lines without losing what its line endings were.

    Returned as (lines, newline, trailing_newline) so the extract can write
    the same endings back - a CRLF file that comes out LF is a different file,
    and the point of this tool is that the subset is the source.
    """
    data = path.read_bytes()
    text = data.decode("utf-8-sig")
    newline = "\r\n" if "\r\n" in text else ("\r" if "\r" in text else "\n")
    trailing = text.endswith(("\n", "\r"))
    lines = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    if lines and lines[-1] == "":
        lines.pop()
    return lines, newline, trailing


def split_record(line: str) -> list[str]:
    """One record's fields, honouring the dtsx's quoting and nothing else."""
    return next(csv.reader(io.StringIO(line)))


def identify(lines: list[str]) -> tuple[str, dict[str, int]]:
    """
    Decide which of the three files this is, by what is in it.

    Every record casts a vote; the kind with the most wins. A 300 record is
    decisive on its own (only the purchase order file has one), and so is a
    100 (only the vendor file does) - but they are counted rather than
    special-cased, so a file that is some of each is reported as ambiguous
    instead of being half-loaded.
    """
    votes: dict[str, int] = {}
    counts: dict[str, int] = {}
    for line in lines:
        if not line.strip():
            continue
        fields = split_record(line)
        rt = fields[0].strip()
        counts[rt] = counts.get(rt, 0) + 1
        kind = SIGNATURES.get((rt, len(fields))) or BY_RECORD_TYPE.get(rt)
        if kind:
            votes[kind] = votes.get(kind, 0) + 1
    if not votes:
        raise ParseError(
            "No Concur records recognised. Expected record types 100/200 "
            "(vendor), 200/210/220/300 (purchase order) or 200 (receipt); "
            f"found {', '.join(sorted(counts)) or 'nothing'}.")
    kind = max(votes, key=lambda k: votes[k])
    return kind, counts


def _spec_key(kind: str, record_type: str) -> str | None:
    return FILE_KINDS.get(kind, {}).get("records", {}).get(record_type)


def _at(fields: list[str], position: int) -> str:
    """Field by 1-based position in the record, empty when the record is short."""
    return fields[position - 1] if len(fields) >= position else ""


# --------------------------------------------------------------- the loaders


def _load_vendor_file(conn, file_id, lines):
    n = {"settings": 0, "vendors": 0, "skipped": 0}
    for i, line in enumerate(lines, start=1):
        if not line.strip():
            continue
        f = split_record(line)
        rt = f[0].strip()
        if rt == "100":
            m = SETTINGS_MAP
            conn.execute(
                "INSERT INTO ERP_Concur_Settings (file_id, line_no, "
                "error_threshold, default_country_code, pay_method_type, raw) "
                "VALUES (?,?,?,?,?,?)",
                (file_id, i, _at(f, m["error_threshold"]),
                 _at(f, m["default_country_code"]),
                 _at(f, m["pay_method_type"]), line))
            n["settings"] += 1
        elif rt == "200":
            m = VENDOR_MAP
            conn.execute(
                "INSERT INTO ERP_Concur_Vendors (file_id, line_no, vendor_code, "
                "vendor_address_id, vendor_address_code, vendor_name, currency, "
                "payment_term_days, address1, address2, address3, city, state, "
                "postal_code, country_code, contact_email, field_count, raw) "
                "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                (file_id, i, _at(f, m["vendor_code"]), _at(f, m["vendor_address_id"]),
                 _at(f, m["vendor_address_code"]), _at(f, m["vendor_name"]),
                 _at(f, m["currency"]), _at(f, m["payment_term_days"]),
                 _at(f, m["address1"]), _at(f, m["address2"]), _at(f, m["address3"]),
                 _at(f, m["city"]), _at(f, m["state"]), _at(f, m["postal_code"]),
                 _at(f, m["country_code"]), _at(f, m["contact_email"]),
                 len(f), line))
            n["vendors"] += 1
        else:
            n["skipped"] += 1
    return n


def _load_po_file(conn, file_id, lines):
    """
    Walk the file accumulating 400 / 300 / 210 / 220 records until a 200
    header closes the group, then write the group out pointing at that
    header.

    A 400 carries no External ID or line number of its own - Amount and
    twenty-two future/custom fields are the whole record - and, unlike every
    other child record here, it comes *before* the 300 it allocates rather
    than after: confirmed against real production data, where a run of 400s
    is immediately followed by the one 300 whose Quantity x Unit Price they
    match, to the cent, every time. So a run of 400s is held in a small buffer
    until the next record settles what it belongs to: a 300 claims the whole
    buffer, anything else (210/220/200, or the file ending) means the buffer
    had nothing to attach to and is written with line_key NULL - the findings
    pass reports it, the same as any other orphan here.

    Anything left over at the end - lines with no header after them - is
    written with po_key NULL rather than discarded, and the findings pass
    reports it. A truncated file is a thing you want to see, not a thing you
    want silently rounded down.
    """
    n = {"headers": 0, "line_items": 0, "allocations": 0,
         "unlinked_allocations": 0, "addresses": 0, "orphans": 0,
         "skipped": 0}
    group = 0
    # Each entry: {"line_no", "rt", "fields", "raw"}, plus "allocs" (a list of
    # (line_no, fields, raw) triples) on every "300" entry - the 400s that ran
    # immediately before it. The raw line is carried here rather than read
    # from the enclosing loop variable: by the time flush() runs, that
    # variable holds the 200 header, and every pending record would be
    # stored with the header's text.
    pending: list[dict] = []
    # 400s seen since the last 300 (or the last group boundary), waiting to
    # find out whether a 300 is coming to claim them.
    alloc_buffer: list[tuple[int, list[str], str]] = []

    def _insert_alloc(po_key, line_key, line_no, raw_fields, raw) -> None:
        m = PO_ALLOC_MAP
        conn.execute(
            "INSERT INTO ERP_Concur_PoLineAllocations (file_id, po_key, "
            "line_key, line_no, po_group, amount, field_count, raw) "
            "VALUES (?,?,?,?,?,?,?,?)",
            (file_id, po_key, line_key, line_no, group,
             _at(raw_fields, m["amount"]), len(raw_fields), raw))
        n["allocations"] += 1
        if line_key is None:
            n["unlinked_allocations"] += 1
        if po_key is None:
            n["orphans"] += 1

    def _orphan_alloc_buffer() -> None:
        for a_line_no, a_f, a_raw in alloc_buffer:
            pending.append({"line_no": a_line_no, "rt": "400", "fields": a_f,
                            "raw": a_raw})
        alloc_buffer.clear()

    def flush(po_key: int | None, po_group: int) -> None:
        for item in pending:
            line_no, rt, f, raw = (item["line_no"], item["rt"], item["fields"],
                                   item["raw"])
            if rt == "300":
                m = PO_LINE_MAP
                desc = _at(f, m["description"])
                account = _at(f, m["account_code"]).strip()
                is_charge = (account in CHARGE_ACCOUNT_CODES
                             or desc.strip().upper() in CHARGE_DESCRIPTIONS)
                cur = conn.execute(
                    "INSERT INTO ERP_Concur_PoLines (file_id, po_key, line_no, "
                    "po_group, external_id, line_number, supplier_part_id, "
                    "expense_type, account_code, description, quantity, "
                    "unit_price, uom, entity_id, is_charge, field_count, raw) "
                    "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                    (file_id, po_key, line_no, po_group, _at(f, m["external_id"]),
                     _at(f, m["line_number"]), _at(f, m["supplier_part_id"]),
                     _at(f, m["expense_type"]), _at(f, m["account_code"]), desc,
                     _at(f, m["quantity"]), _at(f, m["unit_price"]),
                     _at(f, m["uom"]), _at(f, m["entity_id"]),
                     1 if is_charge else 0, len(f), raw))
                n["line_items"] += 1
                line_key = cur.lastrowid
                if po_key is None:
                    n["orphans"] += 1
                for a_line_no, a_f, a_raw in item["allocs"]:
                    _insert_alloc(po_key, line_key, a_line_no, a_f, a_raw)
            elif rt == "400":
                _insert_alloc(po_key, None, line_no, f, raw)
            else:
                m = PO_ADDRESS_MAP
                conn.execute(
                    "INSERT INTO ERP_Concur_PoAddresses (file_id, po_key, "
                    "line_no, po_group, record_type, external_id, name, "
                    "address1, address2, address3, city, state, postal_code, "
                    "country_code, field_count, raw) "
                    "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                    (file_id, po_key, line_no, po_group, rt, _at(f, m["external_id"]),
                     _at(f, m["name"]), _at(f, m["address1"]), _at(f, m["address2"]),
                     _at(f, m["address3"]), _at(f, m["city"]), _at(f, m["state"]),
                     _at(f, m["postal_code"]), _at(f, m["country_code"]),
                     len(f), raw))
                n["addresses"] += 1
                if po_key is None:
                    n["orphans"] += 1
        pending.clear()

    for i, line in enumerate(lines, start=1):
        if not line.strip():
            continue
        f = split_record(line)
        rt = f[0].strip()
        if rt == "400":
            alloc_buffer.append((i, f, line))
        elif rt == "300":
            pending.append({"line_no": i, "rt": rt, "fields": f, "raw": line,
                            "allocs": list(alloc_buffer)})
            alloc_buffer.clear()
        elif rt in ("210", "220"):
            _orphan_alloc_buffer()
            pending.append({"line_no": i, "rt": rt, "fields": f, "raw": line})
        elif rt == "200":
            _orphan_alloc_buffer()
            group += 1
            m = PO_HEADER_MAP
            cur = conn.execute(
                "INSERT INTO ERP_Concur_PoHeaders (file_id, line_no, po_group, "
                "po_number, policy_external_id, currency_code, vendor_code, "
                "vendor_address_code, order_date, payment_terms, tax, shipping, "
                "ledger_code, entity_id, field_count, raw) "
                "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                (file_id, i, group, _at(f, m["po_number"]),
                 _at(f, m["policy_external_id"]), _at(f, m["currency_code"]),
                 _at(f, m["vendor_code"]), _at(f, m["vendor_address_code"]),
                 _at(f, m["order_date"]), _at(f, m["payment_terms"]),
                 _at(f, m["tax"]), _at(f, m["shipping"]), _at(f, m["ledger_code"]),
                 _at(f, m["entity_id"]), len(f), line))
            n["headers"] += 1
            flush(cur.lastrowid, group)
        else:
            n["skipped"] += 1
    _orphan_alloc_buffer()
    flush(None, group + 1)
    return n


def _load_receipt_file(conn, file_id, lines):
    n = {"receipts": 0, "skipped": 0}
    for i, line in enumerate(lines, start=1):
        if not line.strip():
            continue
        f = split_record(line)
        if f[0].strip() != "200":
            n["skipped"] += 1
            continue
        m = RECEIPT_MAP
        conn.execute(
            "INSERT INTO ERP_Concur_Receipts (file_id, line_no, po_number, "
            "line_item_external_id, goods_receipt_number, delivery_slip_number, "
            "uom, received_quantity, received_date, is_deleted, field_count, raw) "
            "VALUES (?,?,?,?,?,?,?,?,?,?,?,?)",
            (file_id, i, _at(f, m["po_number"]), _at(f, m["line_item_external_id"]),
             _at(f, m["goods_receipt_number"]), _at(f, m["delivery_slip_number"]),
             _at(f, m["uom"]), _at(f, m["received_quantity"]),
             _at(f, m["received_date"]), _at(f, m["is_deleted"]), len(f), line))
        n["receipts"] += 1
    return n


LOADERS = {"vendor": _load_vendor_file, "po": _load_po_file,
           "receipt": _load_receipt_file}


def load_file(path: Path, conn: sqlite3.Connection, kind: str | None = None,
              commit: bool = True) -> dict:
    """
    Read one file into the database, replacing whatever file of its kind was
    loaded before.

    `kind` forces the decision when a file is genuinely ambiguous; normally it
    is left out and identify() decides from the content.
    """
    path = Path(path)
    lines, newline, trailing = read_lines(path)
    detected, record_counts = identify(lines)
    kind = kind or detected
    if kind not in FILE_KINDS:
        raise ParseError(f"{kind!r} is not a Concur extract file kind.")

    replace_file(conn, kind)
    data = path.read_bytes()
    cur = conn.execute(
        "INSERT INTO ERP_Concur_Files (kind, file_name, file_path, sha256, "
        "byte_size, line_count, record_counts, newline, trailing_nl) "
        "VALUES (?,?,?,?,?,?,?,?,?)",
        (kind, path.name, str(path), hashlib.sha256(data).hexdigest(),
         len(data), len(lines), json.dumps(record_counts),
         {"\n": "LF", "\r\n": "CRLF", "\r": "CR"}[newline],
         1 if trailing else 0))
    file_id = cur.lastrowid

    counts = LOADERS[kind](conn, file_id, lines)
    if commit:
        conn.commit()
    result = {"kind": kind, "label": FILE_KINDS[kind]["label"],
              "file": path.name, "file_id": file_id, "detected": detected,
              "forced": kind != detected, "lines": len(lines),
              "record_counts": record_counts, "newline": newline, **counts}
    log.info("%s: %s", path.name, result)
    return result


def load_many(paths: list[Path], conn: sqlite3.Connection) -> tuple[list, list]:
    """
    Load a drop of files. Order matters only in one respect: a purchase order
    file has to be in before the selection can be propagated, so the caller
    re-propagates afterwards rather than this function guessing.

    A file that cannot be read does not stop the others - its error comes back
    in the second list so the page can say which one it was.
    """
    loaded, errors = [], []
    for p in paths:
        try:
            loaded.append(load_file(Path(p), conn))
        except Exception as exc:                              # noqa: BLE001
            log.error("%s: %s", p, exc)
            errors.append({"file": Path(p).name, "message": str(exc)})
    return loaded, errors


# scope -> (table, primary key column, spec key | None when it depends on
# the row's own record_type, as an address's does). Shared by record_detail
# and edit_record so the two can never name a table differently.
RECORD_SCOPES: dict[str, tuple[str, str, str | None]] = {
    "vendor":     ("ERP_Concur_Vendors",           "vendor_key",   "vendor_200"),
    "settings":   ("ERP_Concur_Settings",          "settings_key", "vendor_100"),
    "po":         ("ERP_Concur_PoHeaders",         "po_key",       "po_200"),
    "line":       ("ERP_Concur_PoLines",           "line_key",     "po_300"),
    "allocation": ("ERP_Concur_PoLineAllocations", "alloc_key",    "po_400"),
    "receipt":    ("ERP_Concur_Receipts",          "receipt_key",  "receipt_200"),
    "address":    ("ERP_Concur_PoAddresses",       "addr_key",     None),
}

# scope -> the position map its loader fills from. No entry for "address"
# because its columns are the same regardless of whether the row is a 210 or
# a 220, so PO_ADDRESS_MAP covers both.
RECORD_MAPS: dict[str, dict[str, int]] = {
    "vendor": VENDOR_MAP, "settings": SETTINGS_MAP, "po": PO_HEADER_MAP,
    "line": PO_LINE_MAP, "allocation": PO_ALLOC_MAP, "receipt": RECEIPT_MAP,
    "address": PO_ADDRESS_MAP,
}


def _float_or_none(value: str) -> float | None:
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def synthesize_allocations(conn: sqlite3.Connection, commit: bool = True) -> dict:
    """
    Add one 400 record to every loaded 300 line that has none, Amount equal to
    Quantity x Unit Price - the trivial allocation that satisfies the one
    thing Concur checks (the 400s under a line sum to its extended amount)
    without claiming to know a real account split.

    Most lines in a real file arrive with a genuine 400 already (see
    _load_po_file); this only ever touches the ones that do not. For those,
    there is still no real split to write - whatever produced the others gave
    this particular line none - so a synthesized row is a placeholder, not a
    derived fact, and is marked `synthetic = 1` for exactly that reason. It is
    written here, on request, rather than silently at export the way the
    vendor file's default Import Settings record is: that record is one
    Concur requires and the dtsx already hardcodes, where a 400 is optional
    and this app has no source for the number it is putting in Amount, so
    whoever adds one should see it in the review before it goes out, not
    discover it in the file afterwards.

    A line with no Unit Price has nothing to compute an amount from and is
    left alone; a blank Quantity is treated as 1, the same assumption the
    reconciliation finding makes.

    Custom_1 is carried over from the 300 line's own Custom_1 (position 33 on
    that record - stored here as PoLines.entity_id, the same column the
    header and address External IDs are built from) rather than left blank:
    an allocation belongs to the same entity as the line it splits, and
    Concur's own custom-field validation runs against whatever is in this
    position regardless of record type.
    """
    n = {"lines_seen": 0, "added": 0, "skipped_no_price": 0}
    width = len(spec.fields("po_400"))
    custom_1_pos = spec.names("po_400").index("Custom_1")
    rows = conn.execute(
        "SELECT line_key, file_id, po_key, line_no, po_group, quantity, "
        "unit_price, entity_id FROM ERP_Concur_PoLines l WHERE NOT EXISTS "
        "(SELECT 1 FROM ERP_Concur_PoLineAllocations a "
        " WHERE a.line_key = l.line_key)").fetchall()
    for row in rows:
        row = dict(row)
        n["lines_seen"] += 1
        price = _float_or_none(row["unit_price"])
        if price is None:
            n["skipped_no_price"] += 1
            continue
        qty = _float_or_none(row["quantity"]) or 1.0
        amount = f"{round(qty * price, 2):.2f}"
        values = [""] * width
        values[0], values[1] = "400", amount
        values[custom_1_pos] = row["entity_id"] or ""
        raw = rebuild_raw(values)
        conn.execute(
            "INSERT INTO ERP_Concur_PoLineAllocations (file_id, po_key, "
            "line_key, line_no, po_group, amount, synthetic, field_count, raw) "
            "VALUES (?,?,?,?,?,?,1,?,?)",
            (row["file_id"], row["po_key"], row["line_key"], row["line_no"],
             row["po_group"], amount, width, raw))
        n["added"] += 1
    if commit:
        conn.commit()
    return n


def record_detail(conn: sqlite3.Connection, scope: str, key: int) -> dict:
    """
    One record, field by field, named from the spec - what the detail panel
    shows. The raw line is re-split here rather than reconstructed from the
    parsed columns, so the panel shows what is in the file.
    """
    table, pk, kind_of = RECORD_SCOPES[scope]
    row = conn.execute(f"SELECT * FROM {table} WHERE {pk} = ?", (key,)).fetchone()
    if row is None:
        raise FileNotFoundError(f"No {scope} record {key}.")
    row = dict(row)
    spec_key = kind_of or ("po_210" if row.get("record_type") == "210" else "po_220")
    values = split_record(row["raw"])
    return {"scope": scope, "key": key, "row": row, "spec_key": spec_key,
            "spec_label": spec.label(spec_key),
            "expected_width": spec.WIDTHS.get(spec_key, 0),
            "fields": spec.describe(spec_key, values)}


def rebuild_raw(values: list[str]) -> str:
    """
    Re-join a record's fields into one line the way the dtsx's own Escape()
    would write it: quoted only when a field holds a comma, a double quote,
    CR or LF, with an embedded quote doubled. csv.writer's default dialect
    does exactly that - the write side of the same dialect split_record()
    reads with - so a field that did not need editing round-trips unchanged.
    """
    buf = io.StringIO()
    csv.writer(buf, lineterminator="").writerow(values)
    return buf.getvalue()


def edit_record(conn: sqlite3.Connection, scope: str, key: int,
                edits: dict, commit: bool = True) -> dict:
    """
    Change fields on one record by 1-based raw CSV position, then rebuild
    `raw` from the edited fields and refresh whichever named columns mirror
    them - the same positions the loader for this scope filled them from, so
    a search or a selection join sees the edit exactly like a value that had
    always been there.

    `edits` maps position -> new value (positions and values may arrive as
    strings, since they come off JSON). A position past the end of the
    record extends it with empty fields first, the same way a short record
    reads as empty there. Position 1 (Record Type) cannot be edited - it is
    what makes this tool group and classify the record in the first place,
    not something Concur reads back on its own.
    """
    if scope not in RECORD_SCOPES:
        raise ValueError(f"Unknown record scope {scope!r}.")
    table, pk, _ = RECORD_SCOPES[scope]
    row = conn.execute(f"SELECT * FROM {table} WHERE {pk} = ?", (key,)).fetchone()
    if row is None:
        raise FileNotFoundError(f"No {scope} record {key}.")
    values = split_record(dict(row)["raw"])

    for pos, new_value in edits.items():
        pos = int(pos)
        if pos < 2:
            raise ValueError("The record type (field 1) cannot be edited.")
        while len(values) < pos:
            values.append("")
        values[pos - 1] = "" if new_value is None else str(new_value)
    raw = rebuild_raw(values)

    set_cols, args = ["raw = ?", "field_count = ?"], [raw, len(values)]
    for name, pos in RECORD_MAPS[scope].items():
        set_cols.append(f"{name} = ?")
        args.append(_at(values, pos))
    if scope == "line":
        m = PO_LINE_MAP
        account = _at(values, m["account_code"]).strip()
        desc = _at(values, m["description"])
        is_charge = (account in CHARGE_ACCOUNT_CODES
                     or desc.strip().upper() in CHARGE_DESCRIPTIONS)
        set_cols.append("is_charge = ?")
        args.append(1 if is_charge else 0)
    args.append(key)
    conn.execute(f"UPDATE {table} SET {', '.join(set_cols)} WHERE {pk} = ?", args)
    if commit:
        conn.commit()
    return record_detail(conn, scope, key)
