"""
ERP_Concur_App - Flask front end for cutting Concur extracts down to a vendor
or a purchase order.

Drop the three files ConcurExtracts/PurchaseOrder.dtsx writes - the purchase
order file, the vendor file and the PO receipt file - on the page, or press
"Load from input/" to take whatever is sitting in the package's input folder.
They are parsed into the database, the tabs fill, and the findings pass says
what is wrong with them.

Then pick. Picking a vendor takes every purchase order that names it, and every
receipt against those orders. Picking purchase orders takes the vendors behind
them. The two are the same selection read from either end, which is the whole
point of the tool: "everything for this vendor" and "the vendors behind these
orders" are one question.

Write extract drops a subset of each file in output/, every line byte for byte
the line that came in - or, for a record whose fields were corrected first in
its detail panel, byte for byte the line edit_record() rebuilt for it. See
ERP_Concur_Export for why replaying `raw` unchanged matters, and
ERP_Concur_Parse.edit_record for how an edit updates it.

The import results report Concur sends back after a run is dropped on the same
page. It is not an extract and is never written out; it is read to find out
which orders were rejected, and "Pick every failed order" turns the answer
into the next selection. See ERP_Concur_Results.

Every route returns JSON; the page itself is templates/ERP_Concur_Index.html.
"""
from __future__ import annotations

import argparse
import logging
import sqlite3
import sys
from pathlib import Path

from flask import Flask, g, jsonify, render_template, request, send_file

import ERP_Concur_Results as results
import ERP_Concur_Spec as spec
from ERP_Concur_Db import (
    ALL_KINDS,
    DEFAULT_DB_PATH,
    FILE_KINDS,
    clear_all,
    connect,
    counts as db_counts,
    files as db_files,
    picked_clause,
    propagate_selection,
    resolve_db_path,
    selection_add,
    selection_clear,
    selection_ids,
    selection_remove,
    selection_summary,
)
from ERP_Concur_Export import (
    DEFAULT_OUTPUT_DIR,
    verify as verify_extract,
    write_extract,
)
from ERP_Concur_Findings import by_kind, rebuild as rebuild_findings
from ERP_Concur_Parse import (
    ParseError,
    edit_record,
    load_many,
    record_detail,
    synthesize_allocations,
)

BASE_DIR = Path(__file__).parent.resolve()
INPUT_DIR = BASE_DIR / "input"
UPLOAD_DIR = INPUT_DIR / "dropped"

# What "Load from input/" will consider. The extract files are text; the
# results report is a spreadsheet, and which of the two a file is gets decided
# by its content in _load_paths, not by the suffix.
LOADABLE = {".txt", ".csv", ".dat", ".tsv", ".xls", ".xlsx", ".xlsm"}

log = logging.getLogger("ERP_Concur_App")

app = Flask(__name__, template_folder=str(BASE_DIR / "templates"))
app.config["MAX_CONTENT_LENGTH"] = 256 * 1024 * 1024
app.config["ERP_CONCUR_DB"] = None
app.config["ERP_CONCUR_OUT"] = DEFAULT_OUTPUT_DIR

# Columns each list may sort on. Anything else falls back to the default order
# rather than reaching the query.
SORTS = {
    "vendors": {"vendor_code", "vendor_name", "vendor_address_code", "currency",
                "payment_term_days", "city", "state", "country_code", "line_no"},
    "pos": {"po_number", "vendor_code", "vendor_address_code", "order_date",
            "currency_code", "entity_id", "ledger_code", "line_no"},
    "lines": {"external_id", "line_number", "supplier_part_id", "description",
              "quantity", "unit_price", "uom", "account_code", "line_no"},
    "receipts": {"po_number", "line_item_external_id", "goods_receipt_number",
                 "delivery_slip_number", "received_quantity", "received_date",
                 "line_no"},
}


# ---------------------------------------------------------------- plumbing


def db() -> sqlite3.Connection:
    if "db" not in g:
        g.db = connect(app.config["ERP_CONCUR_DB"])
    return g.db


@app.teardown_appcontext
def _close_db(_exc):
    conn = g.pop("db", None)
    if conn is not None:
        conn.close()


@app.errorhandler(ParseError)
def _handle_parse(exc):
    return jsonify(status="error", message=str(exc)), 400


@app.errorhandler(Exception)
def _handle(exc):
    if isinstance(exc, FileNotFoundError):
        return jsonify(status="error", message=str(exc)), 404
    log.exception("request failed")
    return jsonify(status="error", message=str(exc)), 500


def _order_by(name: str, default: str) -> str:
    sort = request.args.get("sort") or ""
    if sort not in SORTS[name]:
        return default
    direction = "DESC" if request.args.get("dir") == "desc" else "ASC"
    return f"{sort} COLLATE NOCASE {direction}"


def _paging() -> tuple[int, int]:
    try:
        size = max(10, min(2000, int(request.args.get("size", 200))))
    except ValueError:
        size = 200
    try:
        page = max(1, int(request.args.get("page", 1)))
    except ValueError:
        page = 1
    return page, size


def _page(conn, sql: str, args: list, count_sql: str) -> dict:
    page, size = _paging()
    total = conn.execute(count_sql, args).fetchone()[0]
    rows = [dict(r) for r in conn.execute(
        sql + " LIMIT ? OFFSET ?", args + [size, (page - 1) * size])]
    return {"rows": rows, "total": total, "page": page, "size": size}


def _body() -> dict:
    return request.get_json(force=True, silent=True) or {}


def _out_dir() -> Path:
    return Path(app.config["ERP_CONCUR_OUT"])


# ------------------------------------------------------------------- pages


@app.route("/")
def index():
    return render_template("ERP_Concur_Index.html")


@app.route("/api/status")
def api_status():
    conn = db()
    return jsonify(status="success",
                   db=resolve_db_path(app.config["ERP_CONCUR_DB"]),
                   input_dir=str(INPUT_DIR), out_dir=str(_out_dir()),
                   counts=db_counts(conn), files=db_files(conn),
                   kinds=ALL_KINDS,
                   selection=selection_summary(conn),
                   results=results.summary(conn),
                   extracts=[dict(r) for r in conn.execute(
                       "SELECT extract_key, stamp, label, scope, n_vendor, n_po, "
                       "n_line, n_receipt, written_at FROM ERP_Concur_Extracts "
                       "ORDER BY extract_key DESC LIMIT 8")])


# ------------------------------------------------------------------ loading


def _after_load(conn) -> dict:
    """
    Re-close the selection and re-run the findings after any file changes.

    Both have to happen together: a new purchase order file can bring a picked
    vendor new orders, and every finding is about the files as they now stand.
    """
    propagate_selection(conn)
    # A re-dropped purchase order file changes what every Record Identifier in
    # the loaded results report means, so the report is resolved again before
    # the findings that quote it are rebuilt.
    resolved = results.resolve(conn, commit=False)
    conn.commit()
    return {"findings": rebuild_findings(conn),
            "selection": selection_summary(conn),
            "results": results.summary(conn) if resolved["failed_pos"] or
                       results.file_row(conn) else {"loaded": False}}


def _load_paths(paths: list[Path], conn) -> tuple[list, list]:
    """
    Load a drop of files, whatever is in it.

    A results report and the three extract files arrive through the same drop
    zone and are told apart by content - a spreadsheet is always a report, a
    text file only if it carries the report's headings - so nobody has to
    remember which control takes which file.
    """
    reports = [p for p in paths if results.is_results_file(p)]
    extracts = [p for p in paths if p not in reports]
    loaded, errors = load_many(extracts, conn) if extracts else ([], [])
    for p in reports:
        try:
            loaded.append(results.load_results(conn, p))
        except Exception as exc:                                  # noqa: BLE001
            log.error("%s: %s", p, exc)
            errors.append({"file": Path(p).name, "message": str(exc)})
    return loaded, errors


@app.route("/api/upload", methods=["POST"])
def api_upload():
    uploads = request.files.getlist("files")
    if not uploads:
        return jsonify(status="error", message="No files received."), 400
    UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
    paths = []
    for fs in uploads:
        name = Path(fs.filename or "").name
        if not name:
            continue
        dest = UPLOAD_DIR / name
        fs.save(dest)
        paths.append(dest)
    conn = db()
    loaded, errors = _load_paths(paths, conn)
    if not loaded:
        return jsonify(status="error", loaded=loaded, errors=errors,
                       message="; ".join(f"{e['file']}: {e['message']}"
                                         for e in errors)
                               or "Nothing could be read from those files."), 400
    return jsonify(status="success", loaded=loaded, errors=errors,
                   **_after_load(conn))


@app.route("/api/input")
def api_input_list():
    """What is sitting in input/, so the page can offer it before loading it."""
    INPUT_DIR.mkdir(parents=True, exist_ok=True)
    rows = []
    for p in sorted(INPUT_DIR.rglob("*")):
        if p.is_file() and p.suffix.lower() in LOADABLE:
            rows.append({"name": str(p.relative_to(INPUT_DIR)),
                         "bytes": p.stat().st_size,
                         "results": results.is_results_file(p)})
    return jsonify(status="success", dir=str(INPUT_DIR), rows=rows)


@app.route("/api/input/load", methods=["POST"])
def api_input_load():
    body = _body()
    names = body.get("names")
    INPUT_DIR.mkdir(parents=True, exist_ok=True)
    if names:
        paths = []
        for n in names:
            p = (INPUT_DIR / n).resolve()
            # Keep the route inside input/ - it takes a name from the browser.
            if not str(p).startswith(str(INPUT_DIR.resolve())) or not p.is_file():
                return jsonify(status="error",
                               message=f"{n} is not a file in input/."), 400
            paths.append(p)
    else:
        paths = [p for p in sorted(INPUT_DIR.rglob("*"))
                 if p.is_file() and p.suffix.lower() in LOADABLE]
    if not paths:
        return jsonify(status="error",
                       message=f"Nothing to load in {INPUT_DIR}."), 400
    conn = db()
    loaded, errors = _load_paths(paths, conn)
    if not loaded:
        return jsonify(status="error", loaded=loaded, errors=errors,
                       message="; ".join(f"{e['file']}: {e['message']}"
                                         for e in errors)), 400
    return jsonify(status="success", loaded=loaded, errors=errors,
                   **_after_load(conn))


@app.route("/api/findings/rebuild", methods=["POST"])
def api_findings_rebuild():
    conn = db()
    return jsonify(status="success", findings=rebuild_findings(conn),
                   counts=db_counts(conn))


@app.route("/api/allocations/synthesize", methods=["POST"])
def api_allocations_synthesize():
    """
    Add a trivial 400 (Quantity x Unit Price) to every 300 line that has none.

    Explicit and on request, never automatic: see
    ERP_Concur_Parse.synthesize_allocations for why a 400 is not the same
    kind of fabricated default as the vendor file's Import Settings record.
    """
    conn = db()
    added = synthesize_allocations(conn, commit=False)
    findings = rebuild_findings(conn)
    conn.commit()
    return jsonify(status="success", added=added, findings=findings,
                   counts=db_counts(conn))


# ------------------------------------------------------------------ vendors


@app.route("/api/vendors")
def api_vendors():
    conn = db()
    where, args = ["1=1"], []
    q = (request.args.get("q") or "").strip()
    if q:
        like = f"%{q}%"
        where.append("(v.vendor_code LIKE ? OR v.vendor_name LIKE ? "
                     "OR v.vendor_address_code LIKE ? OR v.city LIKE ? "
                     "OR v.contact_email LIKE ?)")
        args += [like] * 5
    if request.args.get("problems") == "any":
        where.append("EXISTS (SELECT 1 FROM ERP_Concur_Findings f "
                     "WHERE f.scope = 'vendor' AND f.ref_key = v.vendor_key)")
    if request.args.get("orphans") == "1":
        # A vendor no purchase order in this file refers to. Not a fault - the
        # vendor proc has no changed-since filter, so it sends everything - but
        # it is the quick way to see which vendors are actually in play.
        where.append("v.vendor_code NOT IN "
                     "(SELECT vendor_code FROM ERP_Concur_PoHeaders)")
    clause = " AND ".join(where) + picked_clause(
        request.args.get("picked") or "", "v.vendor_code", "vendor")
    sql = ("SELECT v.*, "
           " (v.vendor_code IN (SELECT id FROM ERP_Concur_Selection "
           "   WHERE kind = 'vendor')) AS picked, "
           " (SELECT COUNT(*) FROM ERP_Concur_PoHeaders h "
           "   WHERE h.vendor_code = v.vendor_code) AS n_po, "
           " (SELECT COUNT(*) FROM ERP_Concur_Findings f "
           "   WHERE f.scope = 'vendor' AND f.ref_key = v.vendor_key) AS n_finding "
           f"FROM ERP_Concur_Vendors v WHERE {clause} "
           f"ORDER BY {_order_by('vendors', 'v.vendor_code COLLATE NOCASE')}")
    return jsonify(status="success", **_page(
        conn, sql, args,
        f"SELECT COUNT(*) FROM ERP_Concur_Vendors v WHERE {clause}"))


@app.route("/api/vendors/keys")
def api_vendor_keys():
    """Every vendor code the current filter matches - for "select all in view"."""
    conn = db()
    q = (request.args.get("q") or "").strip()
    where, args = ["1=1"], []
    if q:
        like = f"%{q}%"
        where.append("(vendor_code LIKE ? OR vendor_name LIKE ? "
                     "OR vendor_address_code LIKE ? OR city LIKE ? "
                     "OR contact_email LIKE ?)")
        args += [like] * 5
    if request.args.get("orphans") == "1":
        where.append("vendor_code NOT IN "
                     "(SELECT vendor_code FROM ERP_Concur_PoHeaders)")
    clause = " AND ".join(where) + picked_clause(
        request.args.get("picked") or "", "vendor_code", "vendor")
    return jsonify(status="success", ids=[r[0] for r in conn.execute(
        f"SELECT DISTINCT vendor_code FROM ERP_Concur_Vendors WHERE {clause}",
        args)])


# ---------------------------------------------------------- purchase orders


def _po_filter() -> tuple[str, list]:
    where, args = ["1=1"], []
    q = (request.args.get("q") or "").strip()
    if q:
        like = f"%{q}%"
        where.append("(h.po_number LIKE ? OR h.vendor_code LIKE ? "
                     "OR h.vendor_address_code LIKE ? OR h.entity_id LIKE ? "
                     "OR EXISTS (SELECT 1 FROM ERP_Concur_Vendors v "
                     "  WHERE v.vendor_code = h.vendor_code AND v.vendor_name LIKE ?) "
                     "OR EXISTS (SELECT 1 FROM ERP_Concur_PoLines l "
                     "  WHERE l.po_key = h.po_key AND (l.description LIKE ? "
                     "        OR l.supplier_part_id LIKE ?)))")
        args += [like] * 7
    entity = (request.args.get("entity") or "").strip()
    if entity:
        where.append("h.entity_id = ?")
        args.append(entity)
    vendor = (request.args.get("vendor") or "").strip()
    if vendor:
        where.append("h.vendor_code = ?")
        args.append(vendor)
    if request.args.get("problems") == "any":
        where.append("EXISTS (SELECT 1 FROM ERP_Concur_Findings f "
                     "WHERE f.po_number = h.po_number)")
    if request.args.get("problems") == "errors":
        where.append("EXISTS (SELECT 1 FROM ERP_Concur_Findings f "
                     "WHERE f.po_number = h.po_number AND f.severity = 'error')")
    clause = " AND ".join(where) + picked_clause(
        request.args.get("picked") or "", "h.po_number", "po")
    return clause, args


@app.route("/api/pos")
def api_pos():
    conn = db()
    clause, args = _po_filter()
    sql = ("SELECT h.*, "
           " (h.po_number IN (SELECT id FROM ERP_Concur_Selection "
           "   WHERE kind = 'po')) AS picked, "
           " (SELECT vendor_name FROM ERP_Concur_Vendors v "
           "   WHERE v.vendor_code = h.vendor_code LIMIT 1) AS vendor_name, "
           " (SELECT COUNT(*) FROM ERP_Concur_Vendors v "
           "   WHERE v.vendor_code = h.vendor_code "
           "     AND v.vendor_address_code = h.vendor_address_code) AS vendor_pair_ok, "
           " (SELECT COUNT(*) FROM ERP_Concur_PoLines l "
           "   WHERE l.po_key = h.po_key) AS n_line, "
           " (SELECT COUNT(*) FROM ERP_Concur_PoLines l "
           "   WHERE l.po_key = h.po_key AND l.is_charge = 1) AS n_charge, "
           " (SELECT COUNT(*) FROM ERP_Concur_Receipts r "
           "   WHERE r.po_number = h.po_number) AS n_receipt, "
           " (SELECT COUNT(*) FROM ERP_Concur_Findings f "
           "   WHERE f.po_number = h.po_number) AS n_finding, "
           " (SELECT COUNT(*) FROM ERP_Concur_Findings f "
           "   WHERE f.po_number = h.po_number AND f.severity = 'error') AS n_error "
           f"FROM ERP_Concur_PoHeaders h WHERE {clause} "
           f"ORDER BY {_order_by('pos', 'h.po_group')}")
    out = _page(conn, sql, args,
                f"SELECT COUNT(*) FROM ERP_Concur_PoHeaders h WHERE {clause}")
    return jsonify(status="success", **out)


@app.route("/api/pos/keys")
def api_po_keys():
    conn = db()
    clause, args = _po_filter()
    return jsonify(status="success", ids=[r[0] for r in conn.execute(
        f"SELECT h.po_number FROM ERP_Concur_PoHeaders h WHERE {clause}", args)])


@app.route("/api/pos/entities")
def api_po_entities():
    conn = db()
    return jsonify(status="success", rows=[r[0] for r in conn.execute(
        "SELECT DISTINCT entity_id FROM ERP_Concur_PoHeaders "
        "WHERE entity_id <> '' ORDER BY entity_id")])


@app.route("/api/pos/<int:po_key>")
def api_po_detail(po_key: int):
    """One order with everything hanging off it - what the detail panel shows."""
    conn = db()
    header = conn.execute("SELECT * FROM ERP_Concur_PoHeaders WHERE po_key = ?",
                          (po_key,)).fetchone()
    if header is None:
        raise FileNotFoundError(f"No purchase order {po_key}.")
    header = dict(header)
    lines = [dict(r) for r in conn.execute(
        "SELECT l.*, "
        " (SELECT COUNT(*) FROM ERP_Concur_Receipts r "
        "   WHERE r.line_item_external_id = l.external_id) AS n_receipt, "
        " (SELECT COUNT(*) FROM ERP_Concur_PoLines x "
        "   WHERE x.external_id = l.external_id) AS n_share_key, "
        " (SELECT COUNT(*) FROM ERP_Concur_PoLineAllocations a "
        "   WHERE a.line_key = l.line_key) AS n_allocation "
        "FROM ERP_Concur_PoLines l WHERE l.po_key = ? ORDER BY l.line_no",
        (po_key,))]
    return jsonify(
        status="success", header=header, lines=lines,
        allocations=[dict(r) for r in conn.execute(
            "SELECT * FROM ERP_Concur_PoLineAllocations WHERE po_key = ? "
            "ORDER BY line_no", (po_key,))],
        addresses=[dict(r) for r in conn.execute(
            "SELECT * FROM ERP_Concur_PoAddresses WHERE po_key = ? "
            "ORDER BY record_type", (po_key,))],
        receipts=[dict(r) for r in conn.execute(
            "SELECT * FROM ERP_Concur_Receipts WHERE po_number = ? "
            "ORDER BY line_no", (header["po_number"],))],
        vendors=[dict(r) for r in conn.execute(
            "SELECT * FROM ERP_Concur_Vendors WHERE vendor_code = ? "
            "ORDER BY line_no", (header["vendor_code"],))],
        findings=[dict(r) for r in conn.execute(
            "SELECT * FROM ERP_Concur_Findings WHERE po_number = ? "
            "ORDER BY severity, kind", (header["po_number"],))],
        picked=header["po_number"] in selection_ids(conn, "po"))


# -------------------------------------------------------------------- lines


@app.route("/api/lines")
def api_lines():
    conn = db()
    where, args = ["1=1"], []
    q = (request.args.get("q") or "").strip()
    if q:
        like = f"%{q}%"
        where.append("(l.external_id LIKE ? OR l.description LIKE ? "
                     "OR l.supplier_part_id LIKE ? OR l.account_code LIKE ? "
                     "OR h.po_number LIKE ?)")
        args += [like] * 5
    charges = request.args.get("charges") or ""
    if charges == "only":
        where.append("l.is_charge = 1")
    elif charges == "not":
        where.append("l.is_charge = 0")
    if request.args.get("dupes") == "1":
        where.append("l.external_id IN (SELECT external_id FROM ERP_Concur_PoLines "
                     "GROUP BY external_id HAVING COUNT(*) > 1)")
    clause = " AND ".join(where) + picked_clause(
        request.args.get("picked") or "", "h.po_number", "po")
    base = ("FROM ERP_Concur_PoLines l "
            "LEFT JOIN ERP_Concur_PoHeaders h ON h.po_key = l.po_key "
            f"WHERE {clause}")
    sql = ("SELECT l.*, h.po_number, h.vendor_code, "
           " (h.po_number IN (SELECT id FROM ERP_Concur_Selection "
           "   WHERE kind = 'po')) AS picked, "
           " (SELECT COUNT(*) FROM ERP_Concur_Receipts r "
           "   WHERE r.line_item_external_id = l.external_id) AS n_receipt, "
           " (SELECT COUNT(*) FROM ERP_Concur_PoLines x "
           "   WHERE x.external_id = l.external_id) AS n_share_key, "
           " (SELECT COUNT(*) FROM ERP_Concur_PoLineAllocations a "
           "   WHERE a.line_key = l.line_key) AS n_allocation, "
           " (SELECT COUNT(*) FROM ERP_Concur_Findings f "
           "   WHERE f.scope = 'line' AND f.ref_key = l.line_key) AS n_finding "
           + base + f" ORDER BY {_order_by('lines', 'l.line_no')}")
    return jsonify(status="success", **_page(
        conn, sql, args, "SELECT COUNT(*) " + base))


# ----------------------------------------------------------------- receipts


@app.route("/api/receipts")
def api_receipts():
    conn = db()
    where, args = ["1=1"], []
    q = (request.args.get("q") or "").strip()
    if q:
        like = f"%{q}%"
        where.append("(r.po_number LIKE ? OR r.line_item_external_id LIKE ? "
                     "OR r.goods_receipt_number LIKE ? "
                     "OR r.delivery_slip_number LIKE ?)")
        args += [like] * 4
    match = request.args.get("match") or ""
    if match == "orphan":
        where.append("r.line_item_external_id NOT IN "
                     "(SELECT external_id FROM ERP_Concur_PoLines)")
    elif match == "matched":
        where.append("r.line_item_external_id IN "
                     "(SELECT external_id FROM ERP_Concur_PoLines)")
    elif match == "ambiguous":
        where.append("r.line_item_external_id IN "
                     "(SELECT external_id FROM ERP_Concur_PoLines "
                     " GROUP BY external_id HAVING COUNT(*) > 1)")
    if request.args.get("zero") == "1":
        where.append("CAST(r.received_quantity AS REAL) = 0")
    clause = " AND ".join(where) + picked_clause(
        request.args.get("picked") or "", "r.po_number", "po")
    sql = ("SELECT r.*, "
           " (r.po_number IN (SELECT id FROM ERP_Concur_Selection "
           "   WHERE kind = 'po')) AS picked, "
           " (SELECT COUNT(*) FROM ERP_Concur_PoLines l "
           "   WHERE l.external_id = r.line_item_external_id) AS n_line, "
           " (SELECT COUNT(*) FROM ERP_Concur_PoHeaders h "
           "   WHERE h.po_number = r.po_number) AS n_header, "
           " (SELECT COUNT(*) FROM ERP_Concur_Findings f "
           "   WHERE f.scope = 'receipt' AND f.ref_key = r.receipt_key) AS n_finding "
           f"FROM ERP_Concur_Receipts r WHERE {clause} "
           f"ORDER BY {_order_by('receipts', 'r.line_no')}")
    return jsonify(status="success", **_page(
        conn, sql, args, f"SELECT COUNT(*) FROM ERP_Concur_Receipts r WHERE {clause}"))


# ----------------------------------------------------------------- findings


@app.route("/api/findings")
def api_findings():
    conn = db()
    where, args = ["1=1"], []
    for arg, column in (("severity", "severity"), ("kind", "kind"),
                        ("scope", "scope")):
        value = (request.args.get(arg) or "").strip()
        if value:
            where.append(f"{column} = ?")
            args.append(value)
    q = (request.args.get("q") or "").strip()
    if q:
        like = f"%{q}%"
        where.append("(message LIKE ? OR label LIKE ? OR field LIKE ? "
                     "OR po_number LIKE ? OR vendor_code LIKE ?)")
        args += [like] * 5
    if request.args.get("picked") == "only":
        # A finding is "in the selection" if it names a picked order or a
        # picked vendor, or names neither and so is about the files as a whole.
        where.append("((po_number = '' AND vendor_code = '') "
                     "OR po_number IN (SELECT id FROM ERP_Concur_Selection "
                     "                 WHERE kind = 'po') "
                     "OR vendor_code IN (SELECT id FROM ERP_Concur_Selection "
                     "                   WHERE kind = 'vendor'))")
    clause = " AND ".join(where)
    sql = (f"SELECT * FROM ERP_Concur_Findings WHERE {clause} "
           "ORDER BY CASE severity WHEN 'error' THEN 0 WHEN 'warning' THEN 1 "
           "ELSE 2 END, kind, finding_key")
    out = _page(conn, sql, args,
                f"SELECT COUNT(*) FROM ERP_Concur_Findings WHERE {clause}")
    return jsonify(status="success", summary=by_kind(conn), **out)


# ------------------------------------------------------------------ results
#
# What came back from a run: the report itself, the orders it says failed, and
# the one action that turns the second into a selection.


@app.route("/api/results")
def api_results():
    """The report row by row, filtered the way the Results tab filters it."""
    conn = db()
    where, args = ["1=1"], []
    q = (request.args.get("q") or "").strip()
    if q:
        like = f"%{q}%"
        where.append("(r.text LIKE ? OR r.stated_po LIKE ? OR r.resolved_po LIKE ? "
                     "OR r.error_code LIKE ? OR r.line_item_external_id LIKE ?)")
        args += [like] * 5
    level = (request.args.get("level") or "").strip()
    if level:
        where.append("r.level = ?")
        args.append(level)
    code = (request.args.get("code") or "").strip()
    if code:
        # The chips are labelled the way summary() groups them, so the filter
        # has to speak the same language rather than only matching a code.
        if code == "Record sequence":
            where.append("r.text LIKE '%sequence of the record types%'")
        elif code == "Order not imported":
            where.append("r.text LIKE '%was not imported%'")
        elif code.startswith("Error "):
            where.append("r.error_code = ?")
            args.append(code.split(" ", 1)[1])
        else:
            where.append("r.level = ?")
            args.append(code)
    if request.args.get("resolved") == "not":
        where.append(f"{results.EFFECTIVE_PO} = ''")
    clause = " AND ".join(where) + picked_clause(
        request.args.get("picked") or "", results.EFFECTIVE_PO, "po")
    sql = (f"SELECT r.*, {results.EFFECTIVE_PO} AS po_number, "
           f" ({results.EFFECTIVE_PO} IN (SELECT id FROM ERP_Concur_Selection "
           "   WHERE kind = 'po')) AS picked, "
           " (SELECT h.po_key FROM ERP_Concur_PoHeaders h "
           f"   WHERE h.po_number = {results.EFFECTIVE_PO} LIMIT 1) AS po_key "
           f"FROM ERP_Concur_Results r WHERE {clause} ORDER BY r.row_no")
    out = _page(conn, sql, args,
                f"SELECT COUNT(*) FROM ERP_Concur_Results r WHERE {clause}")
    return jsonify(status="success", summary=results.summary(conn), **out)


@app.route("/api/results/orders")
def api_results_orders():
    """One row per order the report rejected."""
    conn = db()
    return jsonify(status="success", rows=results.failed_orders(conn),
                   summary=results.summary(conn))


@app.route("/api/results/pick", methods=["POST"])
def api_results_pick():
    """
    Pick (or drop) every order the report says failed, in one go.

    This is the point of reading the report at all, so it is one button and
    not a filter the user has to get right. The vendors behind those orders
    come along through the usual propagation, and the reason recorded against
    each row names the run - so the manifest beside the extract says "failed
    in Run 73" rather than "picked", a week later when somebody asks why
    these thirty-one orders were in a file together.
    """
    conn = db()
    pos = results.failed_po_numbers(conn)
    if not pos:
        return jsonify(status="error",
                       message="No failed order resolves to a purchase order. "
                               "Load the purchase order file this report "
                               "answers, then try again."), 400
    on = bool(_body().get("on", True))
    report = results.file_row(conn) or {}
    reason = "failed in " + (report.get("run") or report.get("file_name")
                             or "the results report")
    if on:
        out = selection_add(conn, pos=pos, reason=reason)
    else:
        out = selection_remove(conn, pos=pos)
    return jsonify(status="success", orders=len(pos), reason=reason, **out)


# ------------------------------------------------------------------- detail


@app.route("/api/record/<scope>/<int:key>")
def api_record(scope: str, key: int):
    if scope not in ("vendor", "settings", "po", "line", "allocation", "receipt", "address"):
        return jsonify(status="error", message=f"Unknown record {scope!r}."), 400
    conn = db()
    detail = record_detail(conn, scope, key)
    detail["findings"] = [dict(r) for r in conn.execute(
        "SELECT * FROM ERP_Concur_Findings WHERE scope = ? AND ref_key = ? "
        "ORDER BY severity", (scope, key))]
    return jsonify(status="success", **detail)


@app.route("/api/record/<scope>/<int:key>/edit", methods=["POST"])
def api_record_edit(scope: str, key: int):
    if scope not in ("vendor", "settings", "po", "line", "allocation", "receipt", "address"):
        return jsonify(status="error", message=f"Unknown record {scope!r}."), 400
    edits = _body().get("edits") or {}
    if not edits:
        return jsonify(status="error", message="Nothing to change."), 400
    conn = db()
    try:
        detail = edit_record(conn, scope, key, edits, commit=False)
    except (ValueError, FileNotFoundError) as exc:
        conn.rollback()
        return jsonify(status="error", message=str(exc)), 400
    # A record's field can be the difference between a pairing error and a
    # clean one, so the findings pass has to see the edit before it counts.
    findings = rebuild_findings(conn)
    conn.commit()
    detail["findings"] = [dict(r) for r in conn.execute(
        "SELECT * FROM ERP_Concur_Findings WHERE scope = ? AND ref_key = ? "
        "ORDER BY severity", (scope, key))]
    return jsonify(status="success", **detail, findings_summary=findings,
                   counts=db_counts(conn))


@app.route("/api/spec")
def api_spec():
    """The layouts themselves, for the Spec reference on the Findings tab."""
    return jsonify(status="success", widths=spec.WIDTHS, spec={
        k: {"record_type": v[0], "label": v[1], "source": v[2],
            "fields": [{"position": i + 1, "name": n, "required": r, "proc": p}
                       for i, (n, r, p) in enumerate(v[3])]}
        for k, v in spec.SPEC.items()})


# ---------------------------------------------------------------- selection


@app.route("/api/selection")
def api_selection():
    conn = db()
    return jsonify(status="success", summary=selection_summary(conn),
                   vendors=[dict(r) for r in conn.execute(
                       "SELECT s.id, s.direct, s.reason, "
                       " (SELECT vendor_name FROM ERP_Concur_Vendors v "
                       "   WHERE v.vendor_code = s.id LIMIT 1) AS name "
                       "FROM ERP_Concur_Selection s WHERE s.kind = 'vendor' "
                       "ORDER BY s.id")],
                   pos=[dict(r) for r in conn.execute(
                       "SELECT s.id, s.direct, s.reason, h.vendor_code, "
                       " h.order_date, h.po_key, "
                       " (SELECT COUNT(*) FROM ERP_Concur_PoLines l "
                       "   WHERE l.po_key = h.po_key) AS n_line, "
                       " (SELECT COUNT(*) FROM ERP_Concur_Receipts r "
                       "   WHERE r.po_number = s.id) AS n_receipt "
                       "FROM ERP_Concur_Selection s "
                       "LEFT JOIN ERP_Concur_PoHeaders h ON h.po_number = s.id "
                       "WHERE s.kind = 'po' ORDER BY s.id")])


@app.route("/api/selection/add", methods=["POST"])
def api_selection_add():
    body = _body()
    conn = db()
    return jsonify(status="success", **selection_add(
        conn, vendors=body.get("vendors"), pos=body.get("pos"),
        reason=body.get("reason") or "picked"))


@app.route("/api/selection/remove", methods=["POST"])
def api_selection_remove():
    body = _body()
    conn = db()
    return jsonify(status="success", **selection_remove(
        conn, vendors=body.get("vendors"), pos=body.get("pos")))


@app.route("/api/selection/clear", methods=["POST"])
def api_selection_clear():
    conn = db()
    n = selection_clear(conn)
    return jsonify(status="success", cleared=n, **selection_summary(conn))


# ----------------------------------------------------------------- extracts


@app.route("/api/extract", methods=["POST"])
def api_extract():
    body = _body()
    conn = db()
    if not (selection_ids(conn, "vendor") or selection_ids(conn, "po")):
        return jsonify(status="error",
                       message="Nothing is picked, so there is nothing to "
                               "write. Pick a vendor or a purchase order "
                               "first."), 400
    result = write_extract(conn, out_dir=_out_dir(),
                           label=body.get("label") or "",
                           kinds=body.get("kinds") or None)
    if not result["files"]:
        return jsonify(status="error", message="; ".join(
            f"no {s['kind']} file - {s['why']}" for s in result["skipped"]),
            **result), 400
    if body.get("verify", True):
        result["verify"] = verify_extract(conn, result["files"])
    return jsonify(status="success", **result)


@app.route("/api/extract/preview")
def api_extract_preview():
    """
    What an extract would contain, without writing anything - the numbers the
    Extract tab shows before you press the button.
    """
    conn = db()
    from ERP_Concur_Export import CUTS
    preview = {}
    for kind in CUTS:
        src = conn.execute("SELECT file_name FROM ERP_Concur_Files WHERE kind = ?",
                           (kind,)).fetchone()
        if not src:
            preview[kind] = {"loaded": False}
            continue
        lines, counts = CUTS[kind](conn)
        preview[kind] = {"loaded": True, "source": src[0], "rows": len(lines),
                         "counts": counts,
                         "head": lines[:6], "tail": lines[-3:] if len(lines) > 9 else []}
    return jsonify(status="success", preview=preview,
                   selection=selection_summary(conn))


@app.route("/api/download/<path:name>")
def api_download(name: str):
    path = (_out_dir() / name).resolve()
    if not str(path).startswith(str(_out_dir().resolve())) or not path.is_file():
        raise FileNotFoundError(f"No extract called {name}.")
    return send_file(path, as_attachment=True, download_name=path.name,
                     mimetype="text/plain")


@app.route("/api/extracts/<int:extract_key>")
def api_extract_row(extract_key: int):
    conn = db()
    r = conn.execute("SELECT * FROM ERP_Concur_Extracts WHERE extract_key = ?",
                     (extract_key,)).fetchone()
    if r is None:
        raise FileNotFoundError(f"No extract {extract_key}.")
    import json as _json
    row = dict(r)
    row["files"] = _json.loads(row["files"] or "[]")
    return jsonify(status="success", **row)


# -------------------------------------------------------------------- clear


@app.route("/api/clear", methods=["POST"])
def api_clear():
    body = _body()
    conn = db()
    done = clear_all(conn, selection=bool(body.get("selection", True)))
    return jsonify(status="success", cleared=done, counts=db_counts(conn),
                   selection=selection_summary(conn))


# --------------------------------------------------------------------- main


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(
        description="ERP -> Concur extract picker front end.")
    ap.add_argument("--host", default="127.0.0.1")
    # 5057-5059 are m3_security, adp_concur and mig_sync; 5060/5061 are the
    # SIP ports Chrome blocks; 5062 is sheet_security; 5064 is
    # deploy_packages; 5065 is launcher.
    ap.add_argument("--port", type=int, default=5063)
    ap.add_argument("--db", default=None,
                    help=f"SQLite path (default {DEFAULT_DB_PATH})")
    ap.add_argument("--out", default=str(DEFAULT_OUTPUT_DIR),
                    help="where extracts are written")
    ap.add_argument("--debug", action="store_true")
    args = ap.parse_args(argv)

    logging.basicConfig(
        level=logging.DEBUG if args.debug else logging.INFO,
        format="%(asctime)s  %(levelname)-8s  %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
        handlers=[logging.StreamHandler(sys.stdout)])

    app.config["ERP_CONCUR_DB"] = args.db
    app.config["ERP_CONCUR_OUT"] = Path(args.out)
    INPUT_DIR.mkdir(parents=True, exist_ok=True)
    Path(args.out).mkdir(parents=True, exist_ok=True)

    connect(args.db).close()          # fail early if it cannot be opened
    log.info("Database : %s", resolve_db_path(args.db))
    log.info("Input    : %s", INPUT_DIR)
    log.info("Output   : %s", args.out)
    log.info("Open     : http://%s:%s", args.host, args.port)
    app.run(host=args.host, port=args.port, debug=args.debug, threaded=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
