"""
ADP_Concur_App - Flask front end for the ADP -> Concur employee load.

Drop the workbook on the page. It is saved to input/adp_concur/, parsed into
the database, and the tabs fill: the employees, the six maps they are derived
through, the 305 / 350 / 360 records that come out, and everything that went
wrong on the way. When it looks right, Write extract drops the flat file in the
outbound folder for SAP Concur to collect.

Every route returns JSON; the page itself is templates/ADP_Concur_Index.html.
"""
from __future__ import annotations

import argparse
import json
import logging
import sqlite3
import sys
from pathlib import Path

from flask import Flask, g, jsonify, render_template, request, send_file

from ADP_Concur_Db import (
    ADP_COLUMNS,
    DEFAULT_DB_PATH,
    DERIVED_COLUMNS,
    EMPLOYEE_EDITABLE_FIELDS,
    clear_employees,
    clear_maps,
    clear_preview,
    connect,
    counts as db_counts,
    picked_clause,
    resolve_db_path,
    selection_add,
    selection_clear,
    selection_keys,
    selection_remove,
    selection_summary,
)
from ADP_Concur_Export import (
    ADP_Concur_export,
    DEFAULT_OUTPUT_DIR,
    held_back,
    outbound_dir,
)
from ADP_Concur_Fix import ADP_Concur_apply_fix, ADP_Concur_fix_plan
from ADP_Concur_Hierarchy import (
    ADP_Concur_chain_up,
    ADP_Concur_direct_reports,
    ADP_Concur_forest,
    ADP_Concur_hierarchy_problems,
    ADP_Concur_hierarchy_stats,
    ADP_Concur_subtree,
    ADP_Concur_subtree_keys,
    ADP_Concur_validate_supervisor_map,
)
from ADP_Concur_Import import ADP_Concur_import_workbook
from ADP_Concur_Map import (
    ADP_Concur_derive,
    CONFIG_PATH,
    FIELD_MAP,
    build_record,
    layout_width,
    load_config,
    save_config,
    selected_employees,
)

BASE_DIR = Path(__file__).parent.resolve()
UPLOAD_DIR = BASE_DIR / "input" / "adp_concur"

log = logging.getLogger("ADP_Concur_App")

app = Flask(__name__, template_folder=str(BASE_DIR / "templates"))
app.config["MAX_CONTENT_LENGTH"] = 256 * 1024 * 1024
app.config["ADP_CONCUR_DB"] = None

# The six lookup tables, as the Maps tab shows them: label, table, key column,
# and the columns that are editable in the grid.
MAP_TABLES = {
    "org": ("Org Map", "ADP_Concur_OrgMap", "map_key",
            ["business_unit_desc", "home_department_desc", "home_department_code",
             "org_unit_1", "org_unit_2", "default_language", "currency"]),
    "status": ("Status Map", "ADP_Concur_StatusMap", "map_key",
               ["position_status", "concur_status"]),
    "country": ("Country Map", "ADP_Concur_CountryMap", "map_key",
                ["adp_country", "concur_country"]),
    "language": ("Language Map", "ADP_Concur_LanguageMap", "map_key",
                 ["language_desc", "language_code", "adp_language"]),
    "salary": ("Salary Map", "ADP_Concur_SalaryMap", "map_key",
               ["pay_grade_code", "pay_grade_desc", "expense_map", "travel_map"]),
    "supervisor": ("Supervisor Map", "ADP_Concur_SupervisorMap", "map_key",
                   ["file_number", "employee_name", "supervisor_name",
                    "supervisor_id", "note"]),
}

# Columns the Employees list may sort on. Anything else falls back to the
# default order rather than reaching the query.
EMPLOYEE_SORTS = {c for _, c in ADP_COLUMNS} | {c for _, c in DERIVED_COLUMNS} | {
    "employee_key", "source", "row_state", "duplicate_rows"}


# ---------------------------------------------------------------- plumbing


def db() -> sqlite3.Connection:
    # One connection per request context; Flask tears it down for us.
    if "db" not in g:
        g.db = connect(app.config["ADP_CONCUR_DB"])
    return g.db


@app.teardown_appcontext
def _close_db(_exc):
    conn = g.pop("db", None)
    if conn is not None:
        conn.close()


@app.errorhandler(Exception)
def _handle(exc):
    if isinstance(exc, FileNotFoundError):
        return jsonify(status="error", message=str(exc)), 404
    log.exception("request failed")
    return jsonify(status="error", message=str(exc)), 500


def _order_by(sort: str, direction: str, allowed: set, default: str) -> str:
    """Sorting happens in SQL so it orders the whole result set, not the page."""
    if sort not in allowed:
        return default
    return f"{sort} COLLATE NOCASE {'DESC' if direction == 'desc' else 'ASC'}"


def _paging() -> tuple[int, int]:
    try:
        size = max(10, min(1000, int(request.args.get("size", 100))))
    except ValueError:
        size = 100
    try:
        page = max(1, int(request.args.get("page", 1)))
    except ValueError:
        page = 1
    return page, size


# ------------------------------------------------------------------- pages


@app.route("/")
def index():
    return render_template("ADP_Concur_Index.html")


@app.route("/api/status")
def api_status():
    conn = db()
    cfg = load_config()
    imports = [dict(r) for r in conn.execute(
        "SELECT import_id, file_name, row_count, imported_at "
        "FROM ADP_Concur_Imports ORDER BY import_id DESC LIMIT 6")]
    extracts = [dict(r) for r in conn.execute(
        "SELECT * FROM ADP_Concur_Extracts ORDER BY extract_key DESC LIMIT 6")]
    return jsonify(status="success",
                   db=resolve_db_path(app.config["ADP_CONCUR_DB"]),
                   counts=db_counts(conn),
                   config_path=str(CONFIG_PATH),
                   outbound=str(outbound_dir(cfg)),
                   imports=imports, extracts=extracts)


# ------------------------------------------------------------------ upload


@app.route("/api/upload", methods=["POST"])
def api_upload():
    files = request.files.getlist("files")
    if not files:
        return jsonify(status="error", message="No files received."), 400

    UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
    conn = db()
    results, errors = [], []

    for fs in files:
        name = Path(fs.filename or "").name
        if not name:
            continue
        if not name.lower().endswith((".xlsx", ".xlsm")):
            errors.append({"file": name, "message": "Not an .xlsx workbook."})
            continue
        dest = UPLOAD_DIR / name
        fs.save(dest)
        try:
            results.append(ADP_Concur_import_workbook(dest, conn=conn))
        except Exception as exc:                       # noqa: BLE001
            log.error("import %s: %s", name, exc)
            errors.append({"file": name, "message": str(exc)})

    if not results:
        return jsonify(
            status="error", results=results, errors=errors,
            message="; ".join(f"{e['file']}: {e['message']}" for e in errors)
                    or "Nothing could be loaded from those files."), 400
    return jsonify(status="success", results=results, errors=errors)


# --------------------------------------------------------------- employees


def _employee_filter() -> tuple[str, list]:
    """Search box, status, source and problem filters, as one WHERE clause."""
    where = ["e.row_state <> 'deleted'"]
    args: list = []

    q = (request.args.get("q") or "").strip()
    if q:
        like = f"%{q}%"
        where.append("(e.file_number LIKE ? OR e.legal_first_name LIKE ? "
                     "OR e.legal_last_name LIKE ? OR e.work_email LIKE ? "
                     "OR e.job_title LIKE ? OR e.business_unit_desc LIKE ? "
                     "OR e.login_id LIKE ?)")
        args += [like] * 7

    status = request.args.get("status") or ""
    if status == "active":
        where.append("e.concur_status = 'Y'")
    elif status == "inactive":
        where.append("e.concur_status = 'N'")
    elif status == "unmapped":
        where.append("e.concur_status NOT IN ('Y','N')")

    source = request.args.get("source") or ""
    if source in ("adp", "manual"):
        where.append("e.source = ?")
        args.append(source)

    problems = request.args.get("problems") or ""
    if problems == "errors":
        where.append("EXISTS (SELECT 1 FROM ADP_Concur_Exceptions x "
                     "WHERE x.employee_key = e.employee_key AND x.severity = 'error')")
    elif problems == "any":
        where.append("EXISTS (SELECT 1 FROM ADP_Concur_Exceptions x "
                     "WHERE x.employee_key = e.employee_key)")
    elif problems == "clean":
        where.append("NOT EXISTS (SELECT 1 FROM ADP_Concur_Exceptions x "
                     "WHERE x.employee_key = e.employee_key)")
    elif problems == "duplicates":
        where.append("e.duplicate_rows > 1")

    bu = request.args.get("bu") or ""
    if bu:
        where.append("e.business_unit_desc = ?")
        args.append(bu)

    clause = " AND ".join(where)
    # 'only picked' is a filter like any other, so it composes with the rest:
    # picked + errors answers "what is wrong inside my pilot".
    clause += picked_clause(request.args.get("picked") or "")
    return clause, args


@app.route("/api/employees")
def api_employees():
    conn = db()
    where, args = _employee_filter()
    order = _order_by(request.args.get("sort", ""), request.args.get("dir", "asc"),
                      EMPLOYEE_SORTS, "legal_last_name COLLATE NOCASE, "
                                      "legal_first_name COLLATE NOCASE")
    page, size = _paging()
    total = conn.execute(
        f"SELECT COUNT(*) FROM ADP_Concur_Employees e WHERE {where}", args).fetchone()[0]
    rows = [dict(r) for r in conn.execute(
        f"""
        SELECT e.*,
               (SELECT COUNT(*) FROM ADP_Concur_Exceptions x
                 WHERE x.employee_key = e.employee_key AND x.severity = 'error') AS n_errors,
               (SELECT COUNT(*) FROM ADP_Concur_Exceptions x
                 WHERE x.employee_key = e.employee_key AND x.severity = 'warning') AS n_warnings
          FROM ADP_Concur_Employees e
         WHERE {where} ORDER BY {order} LIMIT ? OFFSET ?
        """, args + [size, (page - 1) * size])]
    return jsonify(status="success", rows=rows, total=total, page=page, size=size)


@app.route("/api/employees/business-units")
def api_business_units():
    conn = db()
    return jsonify(status="success", rows=[r[0] for r in conn.execute(
        "SELECT DISTINCT business_unit_desc FROM ADP_Concur_Employees "
        "WHERE business_unit_desc <> '' ORDER BY 1")])


@app.route("/api/employees/keys")
def api_employee_keys():
    """
    Every employee key matching the current filter, not just the page on
    screen. What 'Select all in view' runs, so a selection can be built from a
    filter without paging through it.
    """
    conn = db()
    where, args = _employee_filter()
    return jsonify(status="success", keys=[r[0] for r in conn.execute(
        f"SELECT e.employee_key FROM ADP_Concur_Employees e WHERE {where}", args)])


@app.route("/api/employees/<int:employee_key>")
def api_employee(employee_key: int):
    conn = db()
    row = conn.execute("SELECT * FROM ADP_Concur_Employees WHERE employee_key = ?",
                       (employee_key,)).fetchone()
    if row is None:
        return jsonify(status="error", message="No such employee."), 404
    exceptions = [dict(r) for r in conn.execute(
        "SELECT severity, field, message FROM ADP_Concur_Exceptions "
        "WHERE employee_key = ? ORDER BY severity, field", (employee_key,))]
    cfg = load_config()
    records = {rt: build_record(dict(row), rt, layout_width(conn, rt), cfg)
               for rt in ("305", "350", "360")}
    # The chain the approver fields are built from, so it is visible on the
    # record it affects rather than only on the Hierarchy tab.
    chain = ADP_Concur_chain_up(conn, row["file_number"])
    # The editor is about this one person, so it never filters - it shows the
    # chain and the reports as they really are.
    reports, _hidden = ADP_Concur_direct_reports(conn, row["file_number"])
    return jsonify(status="success", row=dict(row), exceptions=exceptions,
                   records=records, chain=chain, reports=reports,
                   n_subtree=len(ADP_Concur_subtree(conn, row["file_number"])),
                   adp_columns=ADP_COLUMNS, derived_columns=DERIVED_COLUMNS)


# -------------------------------------------------------------- selection


@app.route("/api/selection")
def api_selection():
    """
    Who is picked, how many, and how they got there.

    The count is what the header and both selection bars show, so it comes
    from the database rather than from anything the page is holding - reload
    the browser and the pilot selection is still there.
    """
    conn = db()
    return jsonify(status="success", keys=selection_keys(conn),
                   **selection_summary(conn))


@app.route("/api/selection/add", methods=["POST"])
def api_selection_add():
    """
    Add to the selection.

    Three ways in, all landing in the same set: explicit `keys`, a whole
    organisation via `file_numbers` with `subtree`, or `filter` to take
    everything matching the current Employees filter. `reason` is stored
    against each person so the selection can be explained later.
    """
    conn = db()
    body = request.get_json(force=True, silent=True) or {}
    reason = str(body.get("reason") or "").strip()

    keys = [int(k) for k in (body.get("keys") or [])]
    if body.get("file_numbers"):
        file_numbers = [str(f) for f in body["file_numbers"]]
        if body.get("subtree"):
            keys += ADP_Concur_subtree_keys(conn, file_numbers)
        else:
            marks = ",".join("?" * len(file_numbers))
            keys += [r[0] for r in conn.execute(
                f"SELECT employee_key FROM ADP_Concur_Employees "
                f"WHERE file_number IN ({marks}) AND row_state <> 'deleted'",
                file_numbers)]
    if not keys:
        return jsonify(status="error", message="Nobody named."), 400

    added = selection_add(conn, keys, reason)
    return jsonify(status="success", added=added, offered=len(set(keys)),
                   keys=selection_keys(conn), **selection_summary(conn))


@app.route("/api/selection/remove", methods=["POST"])
def api_selection_remove():
    conn = db()
    body = request.get_json(force=True, silent=True) or {}
    keys = [int(k) for k in (body.get("keys") or [])]
    if body.get("file_numbers") and body.get("subtree"):
        keys += ADP_Concur_subtree_keys(conn, [str(f) for f in body["file_numbers"]])
    removed = selection_remove(conn, keys)
    return jsonify(status="success", removed=removed,
                   keys=selection_keys(conn), **selection_summary(conn))


@app.route("/api/selection/clear", methods=["POST"])
def api_selection_clear():
    conn = db()
    return jsonify(status="success", removed=selection_clear(conn),
                   keys=[], **selection_summary(conn))


# -------------------------------------------------------------- hierarchy


@app.route("/api/hierarchy/tree")
def api_hierarchy_tree():
    """
    The whole forest, nested. 167 people is small enough to send at once.

    Two filters, and they combine: `picked=only` prunes to the selection, and
    `status=active|inactive` prunes to who is still employed. Either way the
    managers holding a surviving branch up are kept, marked `wanted: false`.

    The stats are deliberately of the *whole* tree, not the filtered one -
    "13 roots, 4 levels deep" is a fact about the hierarchy, and it should not
    change because somebody hid the leavers.
    """
    conn = db()
    picked = (request.args.get("picked") or "") == "only"
    status = request.args.get("status") or ""
    if status not in ("", "active", "inactive"):
        status = ""
    roots = ADP_Concur_forest(conn, picked_only=picked, status=status)

    def count(node):
        return 1 + sum(count(c) for c in node["children"])

    shown = sum(count(r) for r in roots)
    return jsonify(status="success", roots=roots, picked_only=picked,
                   status_filter=status, shown=shown,
                   wanted=sum(_count_wanted(r) for r in roots),
                   stats=ADP_Concur_hierarchy_stats(conn))


def _count_wanted(node: dict) -> int:
    """How many of the nodes on screen actually matched the filter."""
    return (1 if node.get("wanted") else 0) + sum(
        _count_wanted(c) for c in node["children"])


@app.route("/api/hierarchy/employee/<file_number>")
def api_hierarchy_employee(file_number: str):
    """
    One person's complete chain: everything above them, everything below.

    The chain itself is never filtered - it is the truth about who approves
    this person, and a filtered chain would be a lie. The direct reports do
    follow the tree's filter so the two agree, but the number hidden comes
    back with them so the panel can say so out loud.
    """
    conn = db()
    chain = ADP_Concur_chain_up(conn, file_number)
    if not chain["chain"]:
        return jsonify(status="error", message=chain["detail"]), 404
    picked = (request.args.get("picked") or "") == "only"
    status = request.args.get("status") or ""
    reports, hidden = ADP_Concur_direct_reports(conn, file_number, picked, status)
    subtree = ADP_Concur_subtree(conn, file_number, include_self=False)
    return jsonify(status="success", file_number=file_number, chain=chain,
                   reports=reports, reports_hidden=hidden,
                   subtree=subtree, n_subtree=len(subtree))


@app.route("/api/hierarchy/problems")
def api_hierarchy_problems():
    return jsonify(status="success", **ADP_Concur_hierarchy_problems(db()))


@app.route("/api/hierarchy/subtree-keys", methods=["POST"])
def api_hierarchy_subtree_keys():
    """
    Resolve 'this person and everyone under them' into employee keys.

    Done on the server because the client's copy of the tree is a snapshot -
    a map edit between rendering the tree and pressing the button would
    otherwise scope the extract to a hierarchy that no longer exists.
    """
    body = request.get_json(force=True, silent=True) or {}
    file_numbers = [str(f) for f in (body.get("file_numbers") or []) if str(f).strip()]
    if not file_numbers:
        return jsonify(status="error", message="Nobody named."), 400
    return jsonify(status="success",
                   keys=ADP_Concur_subtree_keys(db(), file_numbers))


@app.route("/api/employees/<int:employee_key>", methods=["POST"])
def api_employee_save(employee_key: int):
    conn = db()
    body = request.get_json(force=True, silent=True) or {}
    fields = {k: v for k, v in body.items() if k in EMPLOYEE_EDITABLE_FIELDS}
    flags = {k: (1 if body[k] else 0) for k in
             ("include_305", "include_350", "include_360") if k in body}
    if not fields and not flags:
        return jsonify(status="error", message="Nothing to save."), 400

    sets = ", ".join(f"{k} = ?" for k in list(fields) + list(flags))
    conn.execute(
        f"UPDATE ADP_Concur_Employees SET {sets}, "
        "row_state = CASE WHEN row_state = 'new' THEN 'new' ELSE 'modified' END, "
        "modified_at = datetime('now') WHERE employee_key = ?",
        list(fields.values()) + list(flags.values()) + [employee_key])
    conn.commit()
    # A field edit can change what every lookup returns, so re-derive.
    ADP_Concur_derive(conn, load_config())
    return api_employee(employee_key)


@app.route("/api/employees/new", methods=["POST"])
def api_employee_new():
    """
    Add someone ADP does not have.

    This is how a branch that has not gone live on ADP yet gets its people into
    the load. The record is marked 'manual', which keeps it out of the way of
    the ADP merge and lets a later ADP cut take it over cleanly - the File
    Number is the key either way.
    """
    conn = db()
    body = request.get_json(force=True, silent=True) or {}
    file_number = str(body.get("file_number") or "").strip()
    if not file_number:
        return jsonify(status="error",
                       message="A File Number is required - it is the Employee "
                               "ID every Concur record is keyed on."), 400
    if conn.execute("SELECT 1 FROM ADP_Concur_Employees WHERE file_number = ?",
                    (file_number,)).fetchone():
        return jsonify(status="error",
                       message=f"File Number {file_number} is already held."), 409

    fields = {k: v for k, v in body.items() if k in EMPLOYEE_EDITABLE_FIELDS}
    fields["file_number"] = file_number
    conn.execute(
        f"INSERT INTO ADP_Concur_Employees ({', '.join(fields)}, source, row_state) "
        f"VALUES ({', '.join('?' for _ in fields)}, 'manual', 'new')",
        list(fields.values()))
    conn.commit()
    ADP_Concur_derive(conn, load_config())
    row = conn.execute("SELECT employee_key FROM ADP_Concur_Employees "
                       "WHERE file_number = ?", (file_number,)).fetchone()
    return api_employee(row["employee_key"])


@app.route("/api/employees/<int:employee_key>", methods=["DELETE"])
def api_employee_delete(employee_key: int):
    """
    Soft delete - the row stays and is simply left out of the extract, so it
    can be restored. A manual record can be removed outright with ?hard=1.
    """
    conn = db()
    if request.args.get("hard") == "1":
        conn.execute("DELETE FROM ADP_Concur_Employees "
                     "WHERE employee_key = ? AND source = 'manual'", (employee_key,))
    else:
        conn.execute("UPDATE ADP_Concur_Employees SET row_state = 'deleted', "
                     "modified_at = datetime('now') WHERE employee_key = ?",
                     (employee_key,))
    conn.commit()
    ADP_Concur_derive(conn, load_config())
    return jsonify(status="success")


@app.route("/api/employees/<int:employee_key>/restore", methods=["POST"])
def api_employee_restore(employee_key: int):
    conn = db()
    conn.execute("UPDATE ADP_Concur_Employees SET row_state = 'modified' "
                 "WHERE employee_key = ?", (employee_key,))
    conn.commit()
    ADP_Concur_derive(conn, load_config())
    return jsonify(status="success")


# -------------------------------------------------------------------- maps


@app.route("/api/maps/<name>")
def api_map(name: str):
    if name not in MAP_TABLES:
        return jsonify(status="error", message="No such map."), 404
    label, table, key, columns = MAP_TABLES[name]
    conn = db()
    q = (request.args.get("q") or "").strip()
    where, args = "1=1", []
    if q:
        where = " OR ".join(f"{c} LIKE ?" for c in columns)
        args = [f"%{q}%"] * len(columns)
    rows = [dict(r) for r in conn.execute(
        f"SELECT * FROM {table} WHERE {where} ORDER BY {columns[0]} COLLATE NOCASE", args)]

    # The Supervisor Map is the one map whose rows can be wrong in a way you
    # cannot see by reading them - both ends have to resolve to an employee -
    # so it carries its verdicts into the grid.
    validation = None
    if name == "supervisor":
        checked = ADP_Concur_validate_supervisor_map(conn)
        verdicts = {r["file_number"]: r for r in checked["rows"]}
        for row in rows:
            v = verdicts.get(row["file_number"])
            row["verdict"] = v["verdict"] if v else ""
            row["severity"] = v["severity"] if v else ""
            row["detail"] = v["detail"] if v else ""
        validation = checked["totals"]

    return jsonify(status="success", label=label, key=key,
                   columns=columns, rows=rows, total=len(rows),
                   validation=validation)


@app.route("/api/maps/supervisor/validate")
def api_validate_supervisor_map():
    """
    Every Supervisor Map row checked against the employees.

    Worth having on its own route as well as inline in the grid, because it is
    the check somebody wants to run deliberately before a load rather than
    only notice while scrolling.
    """
    return jsonify(status="success", **ADP_Concur_validate_supervisor_map(db()))


@app.route("/api/maps/<name>/row", methods=["POST"])
def api_map_save(name: str):
    """Add or change one mapping row, then re-derive so the effect is visible."""
    if name not in MAP_TABLES:
        return jsonify(status="error", message="No such map."), 404
    _, table, key, columns = MAP_TABLES[name]
    conn = db()
    body = request.get_json(force=True, silent=True) or {}
    fields = {c: str(body.get(c) or "").strip() for c in columns if c in body}
    if not fields:
        return jsonify(status="error", message="Nothing to save."), 400

    map_key = body.get(key)
    if map_key:
        conn.execute(
            f"UPDATE {table} SET {', '.join(f'{c} = ?' for c in fields)}, "
            f"row_state = 'modified' WHERE {key} = ?",
            list(fields.values()) + [map_key])
    else:
        conn.execute(
            f"INSERT OR REPLACE INTO {table} ({', '.join(fields)}, row_state) "
            f"VALUES ({', '.join('?' for _ in fields)}, 'new')", list(fields.values()))
    conn.commit()
    result = ADP_Concur_derive(conn, load_config())
    return jsonify(status="success", derive=result)


@app.route("/api/maps/<name>/row/<int:map_key>", methods=["DELETE"])
def api_map_delete(name: str, map_key: int):
    if name not in MAP_TABLES:
        return jsonify(status="error", message="No such map."), 404
    _, table, key, _cols = MAP_TABLES[name]
    conn = db()
    conn.execute(f"DELETE FROM {table} WHERE {key} = ?", (map_key,))
    conn.commit()
    return jsonify(status="success", derive=ADP_Concur_derive(conn, load_config()))


# ----------------------------------------------------------------- records


@app.route("/api/records/<record_type>")
def api_records(record_type: str):
    """
    The 305 / 350 / 360 tabs: the records exactly as the extract will write
    them, with the template's own headings above the columns.

    Only the positions the field map fills are shown by default - a 137-column
    table of mostly-empty cells is not readable, and the empty ones are empty
    by design. ?full=1 shows every position.
    """
    if record_type not in FIELD_MAP:
        return jsonify(status="error", message="No such record type."), 404
    conn = db()
    cfg = load_config()
    width = layout_width(conn, record_type)
    # 'only picked' here shows the records a selection-scoped extract would
    # actually write - the same list, built the same way, before it is a file.
    picked = request.args.get("picked") or ""
    people = selected_employees(conn, record_type, cfg,
                               selection_keys(conn) if picked == "only" else None)

    layout = {r["position"]: dict(r) for r in conn.execute(
        "SELECT position, column_ref, heading, max_width FROM ADP_Concur_Layouts "
        "WHERE record_type = ?", (record_type,))}

    from ADP_Concur_Map import column_ref_to_index
    used = sorted(column_ref_to_index(ref) for ref in FIELD_MAP[record_type])
    positions = list(range(1, width + 1)) if request.args.get("full") == "1" else used

    page, size = _paging()
    window = people[(page - 1) * size: (page - 1) * size + size]
    rows = []
    for emp in window:
        record = build_record(emp, record_type, width, cfg)
        rows.append({"file_number": emp["file_number"],
                     "employee_key": emp["employee_key"],
                     "values": [record[p - 1] for p in positions]})

    columns = [{"position": p,
                "column_ref": layout.get(p, {}).get("column_ref", ""),
                "heading": layout.get(p, {}).get("heading", ""),
                "max_width": layout.get(p, {}).get("max_width", "")}
               for p in positions]
    return jsonify(status="success", record_type=record_type, width=width,
                   columns=columns, rows=rows, total=len(people),
                   page=page, size=size, picked=picked,
                   full=request.args.get("full") == "1")


# -------------------------------------------------------------- exceptions


@app.route("/api/exceptions")
def api_exceptions():
    conn = db()
    severity = request.args.get("severity") or ""
    field = request.args.get("field") or ""
    where, args = ["1=1"], []
    if severity in ("error", "warning"):
        where.append("x.severity = ?")
        args.append(severity)
    if field:
        where.append("x.field = ?")
        args.append(field)
    clause = " AND ".join(where) + picked_clause(request.args.get("picked") or "", "x")
    rows = [dict(r) for r in conn.execute(
        f"""
        SELECT x.*, e.position_status, e.concur_status, e.business_unit_desc,
               e.source
          FROM ADP_Concur_Exceptions x
          LEFT JOIN ADP_Concur_Employees e ON e.employee_key = x.employee_key
         WHERE {clause}
         ORDER BY CASE x.severity WHEN 'error' THEN 0 ELSE 1 END,
                  x.field, x.employee_name
        """, args)]
    # The summary counts the same scope as the list, so 'only picked' gives
    # the shape of what is wrong inside the pilot rather than overall.
    summary = [dict(r) for r in conn.execute(
        f"SELECT x.severity, x.field, COUNT(*) AS n FROM ADP_Concur_Exceptions x "
        f"WHERE 1=1{picked_clause(request.args.get('picked') or '', 'x')} "
        f"GROUP BY x.severity, x.field ORDER BY x.severity, n DESC")]
    return jsonify(status="success", rows=rows, total=len(rows), summary=summary)


# ------------------------------------------------------------------- fixes


@app.route("/api/fixes")
def api_fixes():
    """
    The exception list regrouped by what would fix it.

    Two people blocked by the same absent supervisor are one missing person,
    not two problems - this is the list you can actually work down.
    """
    return jsonify(status="success", **ADP_Concur_fix_plan(db()))


@app.route("/api/fixes/apply", methods=["POST"])
def api_fixes_apply():
    """
    Apply one remedy and re-derive.

    The reply carries the error and warning counts either side of the fix, so
    the page can say what it actually achieved rather than claiming success -
    creating a person from guesses can resolve two problems and introduce one,
    and that shows up here as it happens.
    """
    body = request.get_json(force=True, silent=True) or {}
    action = str(body.get("action") or "")
    fields = body.get("fields") or {}
    try:
        result = ADP_Concur_apply_fix(db(), action, fields)
    except ValueError as exc:
        return jsonify(status="error", message=str(exc)), 400
    return jsonify(status="success", counts=db_counts(db()), **result)


@app.route("/api/employees/picker")
def api_employee_picker():
    """
    Every employee as 'file number — name', for the supervisor picker.

    Small enough to send whole, which keeps the picker instant and means it
    cannot offer somebody who is not in the load.
    """
    return jsonify(status="success", rows=[dict(r) for r in db().execute(
        """
        SELECT file_number,
               legal_last_name || ', ' || legal_first_name AS name, job_title
          FROM ADP_Concur_Employees WHERE row_state <> 'deleted'
         ORDER BY legal_last_name COLLATE NOCASE, legal_first_name COLLATE NOCASE
        """)])


# ------------------------------------------------------------------ config


@app.route("/api/config")
def api_config():
    return jsonify(status="success", config=load_config(), path=str(CONFIG_PATH))


@app.route("/api/config", methods=["POST"])
def api_config_save():
    """
    Save the config and re-derive.

    The Login ID rule lives here, so this is the route that answers the
    question in Kelly's email: change the rule, save, and every login ID in
    the extract follows it.
    """
    body = request.get_json(force=True, silent=True) or {}
    cfg = load_config()
    for key, value in body.items():
        if isinstance(value, dict) and isinstance(cfg.get(key), dict):
            cfg[key].update(value)
        else:
            cfg[key] = value
    save_config(cfg)
    conn = db()
    return jsonify(status="success", config=cfg,
                   derive=ADP_Concur_derive(conn, cfg))


@app.route("/api/derive", methods=["POST"])
def api_derive():
    conn = db()
    return jsonify(status="success", derive=ADP_Concur_derive(conn, load_config()),
                   counts=db_counts(conn))


# ------------------------------------------------------------------ export


@app.route("/api/export", methods=["GET", "POST"])
def api_export():
    """
    Build the extract, or write it.

    POST with `keys` scopes it to a selection - the people ticked on the
    Employees tab, or a whole organisation picked off the Hierarchy tab. The
    resulting file is named and recorded as a selection so a pilot load sitting
    in the pickup folder is never mistaken for the full company.
    """
    conn = db()
    cfg = load_config()
    body = request.get_json(force=True, silent=True) or {}
    dry = (request.args.get("write") != "1") and not body.get("write")

    # The selection lives in the database, so the file is scoped by asking for
    # it rather than by the page sending a list of keys - the extract and the
    # 'only picked' view are then reading the same thing by construction.
    keys, label = None, ""
    if body.get("selection") or request.args.get("selection") == "1":
        keys = selection_keys(conn)
        label = str(body.get("label") or "") or selection_summary(conn)["label"]
    elif body.get("keys") is not None:
        keys = [int(k) for k in body["keys"]]
        label = str(body.get("label") or "")

    res = ADP_Concur_export(conn, cfg, keys=keys, selection_label=label,
                            dry_run=dry)
    res["held_back"] = held_back(conn, keys)
    return jsonify(status="success", **res)


@app.route("/api/download/<path:name>")
def api_download(name: str):
    """Serve a written extract back through the browser."""
    cfg = load_config()
    target = (outbound_dir(cfg) / Path(name).name).resolve()
    root = outbound_dir(cfg).resolve()
    if root not in target.parents and target.parent != root:
        return jsonify(status="error", message="Outside the outbound folder."), 400
    if not target.exists():
        return jsonify(status="error", message="No such file."), 404
    return send_file(str(target), as_attachment=True, download_name=target.name)


@app.route("/api/held-back")
def api_held_back():
    return jsonify(status="success", rows=held_back(db()))


# ------------------------------------------------------------------- clear


@app.route("/api/clear", methods=["POST"])
def api_clear():
    """
    Start again.

    'employees' takes a list of sources - 'adp', 'manual', or both - because
    those are two different decisions. 'maps' empties the six lookup tables and
    the captured layouts, which the next workbook refills.

    A clear is irreversible for the hand-keyed people, so it needs `confirm`
    to be the literal word CLEAR. That is the same gate the M3_Security tool
    puts on its own clear, and for the same reason.
    """
    body = request.get_json(force=True, silent=True) or {}
    what = body.get("what")
    conn = db()

    if str(body.get("confirm") or "").strip().upper() != "CLEAR":
        return jsonify(status="error",
                       message="Type CLEAR to confirm - this cannot be undone."), 400

    if what == "employees":
        # An absent key means "the usual thing"; an empty list means the caller
        # deliberately named nothing, and must not fall through to a default
        # that deletes 166 people.
        sources = body.get("sources")
        if sources is None:
            sources = ["adp"]
        try:
            out = clear_employees(conn, sources)
        except ValueError as exc:
            return jsonify(status="error", message=str(exc)), 400
    elif what == "maps":
        out = clear_maps(conn)
    else:
        return jsonify(status="error", message="Nothing named to clear."), 400
    ADP_Concur_derive(conn, load_config())
    return jsonify(status="success", cleared=out, counts=db_counts(conn))


@app.route("/api/clear/preview")
def api_clear_preview():
    """
    What a clear would take, before it takes it.

    Per source, with the counts that matter: how many are picked (the selection
    goes with them) and how many are hand-keyed and therefore gone for good.
    """
    sources = [s for s in (request.args.get("sources") or "adp").split(",") if s]
    return jsonify(status="success", **clear_preview(db(), sources))


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="ADP -> Concur employee load front end.")
    ap.add_argument("--host", default="127.0.0.1")
    ap.add_argument("--port", type=int, default=5058)
    ap.add_argument("--db", default=None, help=f"SQLite path (default {DEFAULT_DB_PATH})")
    ap.add_argument("--debug", action="store_true")
    args = ap.parse_args(argv)

    logging.basicConfig(
        level=logging.DEBUG if args.debug else logging.INFO,
        format="%(asctime)s  %(levelname)-8s  %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
        handlers=[logging.StreamHandler(sys.stdout)])

    app.config["ADP_CONCUR_DB"] = args.db
    UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
    cfg = load_config()
    outbound_dir(cfg).mkdir(parents=True, exist_ok=True)

    # Fail early if the database cannot be opened or migrated.
    connect(args.db).close()
    log.info("Database : %s", resolve_db_path(args.db))
    log.info("Uploads  : %s", UPLOAD_DIR)
    log.info("Outbound : %s", outbound_dir(cfg))
    log.info("Config   : %s", CONFIG_PATH)
    log.info("Open     : http://%s:%s", args.host, args.port)

    app.run(host=args.host, port=args.port, debug=args.debug, threaded=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
