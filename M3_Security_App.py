"""
M3_Security_App - Flask front end for the M3 / Infor OS security capture.

Six tabs over one SQLite capture: users, security roles, functional security
roles, SES400 function authorisations, MNS405 role definitions and MNS410
role-per-user rows. Everything the page needs arrives as JSON from the ~50
routes below; the page itself is templates/M3_Security_Index.html.

Anything that talks to M3 goes through M3_Security_M3Api and, when it is slow,
through start_job() so the browser can poll /api/job/<id> instead of holding a
request open.

Reconstructed from M3_Security_App.cpython-312.pyc after the source was lost.
Logic matches the bytecode; comments and formatting are not original.
"""

from __future__ import annotations

import argparse
import logging
import sqlite3
import sys
import threading
import time
import traceback
import uuid
from pathlib import Path

from flask import (
    Flask,
    g,
    jsonify,
    render_template,
    request,
    send_file,
)

from M3_Security_Bod import (
    CONFIG_PATH as BOD_CONFIG_PATH,
    BodConfigError,
    tenant_bod_config,
    DEFAULT_BOD_DIR,
    load_bod_config,
    save_bod_config,
)
from M3_Security_Db import (
    DEFAULT_DB_PATH,
    USER_EDITABLE_FIELDS,
    add_roles_from_m3,
    connect,
    resolve_db_path,
    roles_missing_from_m3,
    strip_role_members,
    strip_user_roles,
    tenants as db_tenants,
)
from M3_Security_Export import (
    changed_counts,
    export_functional,
    export_roles,
    export_users,
    mark_pushed,
)
from M3_Security_M3Api import (
    DEFAULT_BATCH_SIZE,
    FUNCTION_STATUS,
    DEFAULT_COMPANY,
    DEFAULT_DIVISION,
    DEFAULT_IONAPI_DIR,
    CONFIG_PATH as M3_CONFIG_PATH,
    M3ApiError,
    M3Client,
    add_role_members,
    discover_companies,
    load_m3_config,
    save_tenant_company,
    tenant_company,
    create_missing_function_roles,
    create_roles,
    delete_function_roles,
    list_ionapi_files,
    mns405_add,
    mns405_delete,
    mns405_sync,
    mns405_update,
    mns410_add,
    mns410_delete,
    mns410_sync,
    mns410_update,
    remove_role_members,
    sync_function_roles,
    sync_role_members,
    sync_roles,
    sync_users,
    update_function_role_status,
)
from M3_Security_Import import import_file, parse_file_name

BASE_DIR = Path(__file__).parent.resolve()
UPLOAD_DIR = BASE_DIR / "input" / "m3_security"
OUTPUT_DIR = BASE_DIR / "output" / "m3_security"
BOD_DIR = DEFAULT_BOD_DIR

log = logging.getLogger("M3_Security_App")

app = Flask(__name__, template_folder=str(BASE_DIR / "templates"))
app.config["MAX_CONTENT_LENGTH"] = 256 * 1024 * 1024
app.config["M3_SECURITY_DB"] = None


# ---------------------------------------------------------------- plumbing


def db() -> sqlite3.Connection:
    # One connection per request context; Flask tears it down for us.
    if "db" not in g:
        g.db = connect(app.config["M3_SECURITY_DB"])
    return g.db


@app.teardown_appcontext
def _close_db(_exc):
    conn = g.pop("db", None)
    if conn is not None:
        conn.close()


@app.errorhandler(Exception)
def _handle(exc):
    # Flask's own HTTP errors carry a numeric .code; pass those straight
    # through rather than dressing them up as a 500 with a traceback.
    if hasattr(exc, "code") and isinstance(getattr(exc, "code"), int):
        return jsonify(status="error", message=str(exc)), exc.code
    log.error("%s", traceback.format_exc())
    return jsonify(status="error", message=str(exc),
                   trace=traceback.format_exc()), 500


def _touch_user(conn, user_key, state="modified"):
    # A row that is still 'new' stays 'new' - it has never been exported.
    conn.execute(
        "UPDATE M3_Security_Users SET row_state = CASE WHEN row_state = 'new' "
        "THEN 'new' ELSE ? END, modified_at = datetime('now') "
        "WHERE user_key = ?",
        (state, user_key),
    )


@app.route("/")
def index():
    return render_template("M3_Security_Index.html")


# ------------------------------------------------------------------ status


@app.route("/api/status")
def api_status():
    conn = db()
    out = []
    for tenant in db_tenants(conn):
        row = conn.execute(
            """
            SELECT
              (SELECT COUNT(*) FROM M3_Security_Users
                WHERE tenant = ?1 AND row_state <> 'deleted')            AS users,
              (SELECT COUNT(*) FROM M3_Security_Users
                WHERE tenant = ?1 AND row_state IN ('modified','new'))   AS users_changed,
              (SELECT COUNT(*) FROM M3_Security_Users
                WHERE tenant = ?1 AND row_state = 'deleted')             AS users_deleted,
              (SELECT COUNT(*) FROM M3_Security_Roles
                WHERE tenant = ?1 AND row_state <> 'deleted')            AS roles,
              (SELECT COUNT(*) FROM M3_Security_Roles
                WHERE tenant = ?1 AND row_state IN ('modified','new'))   AS roles_changed,
              (SELECT COUNT(*) FROM M3_Security_RoleAssignments
                WHERE tenant = ?1 AND row_state <> 'deleted')            AS assignments
            """,
            (tenant,),
        ).fetchone()
        imports = conn.execute(
            "SELECT file_kind, file_name, row_count, imported_at "
            "FROM M3_Security_Imports WHERE tenant = ? "
            "ORDER BY import_id DESC LIMIT 6",
            (tenant,),
        ).fetchall()
        out.append({"tenant": tenant,
                    **{k: row[k] for k in row.keys()},
                    "imports": [dict(i) for i in imports]})
    return jsonify(status="success",
                   db=resolve_db_path(app.config["M3_SECURITY_DB"]),
                   tenants=out)


# ------------------------------------------------------------------ upload


@app.route("/api/upload", methods=["POST"])
def api_upload():
    files = request.files.getlist("files")
    if not files:
        return jsonify(status="error", message="No files received."), 400

    # tenant_override pins every file; tenant is only a fallback for files
    # whose name does not carry one.
    tenant_fallback = (request.form.get("tenant") or "").strip() or None
    tenant_override = (request.form.get("tenant_override") or "").strip() or None

    UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
    conn = db()
    results, errors = [], []

    saved = []
    for fs in files:
        name = Path(fs.filename or "").name
        if not name:
            continue
        if not name.lower().endswith(".csv"):
            errors.append({"file": name, "message": "Not a .csv file."})
            continue
        dest = UPLOAD_DIR / name
        fs.save(dest)
        saved.append(dest)

    # Files whose name names a tenant go first, so the ones that do not can
    # inherit the tenant those established.
    saved.sort(key=lambda p: parse_file_name(p.name)["tenant"] is None)
    discovered = None

    for dest in saved:
        try:
            res = import_file(dest, conn=conn, tenant=tenant_override,
                              tenant_fallback=tenant_fallback or discovered)
            discovered = discovered or res["tenant"]
            results.append(res)
        except Exception as exc:
            log.error("import %s: %s", dest.name, exc)
            errors.append({"file": dest.name, "message": str(exc)})

    if not results:
        return jsonify(
            status="error", results=results, errors=errors,
            message="; ".join(f"{e['file']}: {e['message']}"
                              for e in errors)
                    or "Nothing could be loaded from those files.")
    return jsonify(status="success", results=results, errors=errors)


# ------------------------------------------------------------------- users


def _order_by(sorts: dict, default: str) -> str:
    key = (request.args.get("sort") or "").strip()
    down = (request.args.get("dir") or "asc").lower() == "desc"
    expr = sorts.get(key)
    if not expr:
        return default
    # A sort key may map to several columns, applied in order.
    terms = [expr] if isinstance(expr, str) else list(expr)
    parts = [f"{e} {'DESC' if down else 'ASC'}" for e in terms]
    # The default trails every explicit sort so paging stays stable.
    return ", ".join(parts) + ", " + default


USER_SORTS = {
    "last_name": "LOWER(COALESCE(u.last_name,''))",
    "first_name": "LOWER(COALESCE(u.first_name,''))",
    "email_id": "LOWER(COALESCE(u.email_id,''))",
    "title": "LOWER(COALESCE(u.title,''))",
    "status": "LOWER(COALESCE(u.status,''))",
    "upn": "LOWER(COALESCE(u.upn,''))",
    "n_security": "n_security",
    "n_functional": "n_functional",
    "last_login_date": "COALESCE(u.last_login_date,'')",
    "row_state": "u.row_state",
}
USER_DEFAULT_SORT = ("LOWER(COALESCE(u.last_name, '')), "
                     "LOWER(COALESCE(u.first_name, '')), u.user_key")


def _user_filter() -> tuple[str, list]:
    """WHERE clause + args shared by the user list and the select-all helper."""
    tenant = request.args.get("tenant", "")
    q = (request.args.get("q") or "").strip()
    status = (request.args.get("status") or "").strip()
    role = (request.args.get("role") or "").strip()
    func_role = (request.args.get("func_role") or "").strip()
    state = (request.args.get("state") or "").strip()

    where = ["u.tenant = ?"]
    args = [tenant]
    if not request.args.get("include_deleted"):
        where.append("u.row_state <> 'deleted'")
    if q:
        where.append("(u.first_name LIKE ? OR u.last_name LIKE ? OR u.email_id LIKE ? "
                     " OR u.user_name LIKE ? OR u.upn LIKE ? OR u.person_id LIKE ?)")
        args += [f"%{q}%"] * 6
    if status:
        where.append("u.status = ?")
        args.append(status)
    if state:
        where.append("u.row_state = ?")
        args.append(state)
    if role:
        where.append("EXISTS (SELECT 1 FROM M3_Security_UserRoles ur "
                     "WHERE ur.user_key = u.user_key AND ur.row_state <> 'deleted' "
                     "AND ur.role_name = ?)")
        args.append(role)
    if func_role:
        where.append("EXISTS (SELECT 1 FROM M3_Security_UserRoles ur "
                     "WHERE ur.user_key = u.user_key AND ur.row_state <> 'deleted' "
                     "AND ur.role_type = 'Functional' AND ur.role_name = ?)")
        args.append(func_role)
    if request.args.get("has_functional") == "1":
        where.append("EXISTS (SELECT 1 FROM M3_Security_UserRoles ur "
                     "WHERE ur.user_key = u.user_key AND ur.row_state <> 'deleted' "
                     "AND ur.role_type = 'Functional')")
    return " AND ".join(where), args


@app.route("/api/users")
def api_users():
    conn = db()
    page = max(1, int(request.args.get("page", 1)))
    size = min(500, max(10, int(request.args.get("size", 50))))
    clause, args = _user_filter()

    total = conn.execute(
        f"SELECT COUNT(*) FROM M3_Security_Users u WHERE {clause}",
        args).fetchone()[0]

    rows = conn.execute(
        f"""
        SELECT u.user_key, u.person_id, u.first_name, u.last_name, u.email_id,
               u.title, u.status, u.user_name, u.upn, u.user_alias,
               u.last_login_date, u.created_date, u.row_state,
               (SELECT COUNT(*) FROM M3_Security_UserRoles ur
                 WHERE ur.user_key = u.user_key AND ur.role_type = 'Security'
                   AND ur.row_state <> 'deleted')  AS n_security,
               (SELECT COUNT(*) FROM M3_Security_UserRoles ur
                 WHERE ur.user_key = u.user_key AND ur.role_type = 'Functional'
                   AND ur.row_state <> 'deleted')  AS n_functional,
               -- when the user holds exactly one role of a type the list shows
               -- its name instead of the count; MIN over that single row is it
               (SELECT MIN(ur.role_name) FROM M3_Security_UserRoles ur
                 WHERE ur.user_key = u.user_key AND ur.role_type = 'Security'
                   AND ur.row_state <> 'deleted')  AS security_role,
               (SELECT MIN(ur.role_name) FROM M3_Security_UserRoles ur
                 WHERE ur.user_key = u.user_key AND ur.role_type = 'Functional'
                   AND ur.row_state <> 'deleted')  AS functional_role
        FROM M3_Security_Users u
        WHERE {clause}
        ORDER BY {_order_by(USER_SORTS, USER_DEFAULT_SORT)}
        LIMIT ? OFFSET ?
        """,
        args + [size, (page - 1) * size]).fetchall()

    return jsonify(status="success", total=total, page=page, size=size,
                   rows=[dict(r) for r in rows])


@app.route("/api/users/keys")
def api_user_keys():
    """Every user key matching the current filter, unpaged, for select-all."""
    conn = db()
    clause, args = _user_filter()
    rows = conn.execute(
        f"""
        SELECT u.user_key, u.email_id,
               (SELECT COUNT(*) FROM M3_Security_UserRoles ur
                 WHERE ur.user_key = u.user_key AND ur.role_type = 'Functional'
                   AND ur.row_state <> 'deleted')  AS n_functional
        FROM M3_Security_Users u WHERE {clause}
        ORDER BY LOWER(COALESCE(u.last_name, '')), u.user_key
        """,
        args).fetchall()
    return jsonify(status="success", total=len(rows),
                   rows=[dict(r) for r in rows])


@app.route("/api/users/remove-roles", methods=["POST"])
def api_users_remove_roles():
    """
    Take roles off a set of users, locally.

    role_type 'Functional' is the FunctionalSecurityRoleN block, 'Security' the
    SecurityRoleN block. roles limits it to named roles; empty means every role
    of that type the selected users hold.

    Nothing is sent to M3 - the users are flagged as changed so the next users
    export carries them without those roles, and importing that into IFS is
    what actually removes them. Preview unless confirm == 'REMOVE'.
    """
    conn = db()
    data = request.get_json(force=True) or {}
    tenant = (data.get("tenant") or "").strip()
    keys = [int(k) for k in (data.get("user_keys") or [])]
    role_type = (data.get("role_type") or "Functional").strip()
    names = [str(r) for r in (data.get("roles") or [])]
    confirm = data.get("confirm")
    if not tenant:
        return jsonify(status="error", message="Tenant is required."), 400
    if not keys:
        return jsonify(status="error", message="No users selected."), 400
    if role_type not in ("Functional", "Security"):
        return jsonify(status="error",
                       message="role_type must be Functional or Security."), 400
    if confirm and confirm != "REMOVE":
        return jsonify(status="error", message="Type REMOVE to confirm."), 400

    marks = ",".join("?" * len(keys))
    args = [tenant, role_type] + keys
    name_clause = ""
    if names:
        name_clause = " AND ur.role_name IN (%s)" % ",".join("?" * len(names))
        args += names

    # What the removal would touch, one row per role.
    breakdown = [dict(r) for r in conn.execute(
        f"""
        SELECT ur.role_name AS role, COUNT(DISTINCT ur.user_key) AS users
        FROM M3_Security_UserRoles ur
        WHERE ur.tenant = ? AND ur.role_type = ? AND ur.row_state <> 'deleted'
          AND ur.user_key IN ({marks}){name_clause}
        GROUP BY ur.role_name ORDER BY LOWER(ur.role_name)
        """,
        args).fetchall()]
    totals = {
        "selected_users": len(keys),
        "roles": len(breakdown),
        "role_links": sum(b["users"] for b in breakdown),
        "users": conn.execute(
            f"""
            SELECT COUNT(DISTINCT ur.user_key) FROM M3_Security_UserRoles ur
            WHERE ur.tenant = ? AND ur.role_type = ? AND ur.row_state <> 'deleted'
              AND ur.user_key IN ({marks}){name_clause}
            """,
            args).fetchone()[0],
    }
    label = "functional" if role_type == "Functional" else "security"

    if not confirm:
        return jsonify(
            status="success", dry_run=True, plan=breakdown, totals=totals,
            role_type=role_type,
            message=f"{totals['role_links']} {label} role assignment(s) would be removed from "
                    f"{totals['users']} of the "
                    f"{totals['selected_users']} selected user(s), across "
                    f"{totals['roles']} role(s).")

    res = strip_user_roles(conn, tenant, keys, role_type, names or None)
    return jsonify(
        status="success", dry_run=False, plan=breakdown, totals=totals,
        role_type=role_type, result=res,
        message=f"{res['role_links']} {label} role assignment(s) removed from "
                f"{res['users']} user(s). Export the users and import them to apply it in IFS.")


@app.route("/api/users/<int:user_key>")
def api_user(user_key: int):
    conn = db()
    u = conn.execute(
        "SELECT * FROM M3_Security_Users WHERE user_key = ?",
        (user_key,)).fetchone()
    if not u:
        return jsonify(status="error", message="User not found."), 404
    roles = conn.execute(
        "SELECT role_type, role_name FROM M3_Security_UserRoles "
        "WHERE user_key = ? AND row_state <> 'deleted' "
        "ORDER BY role_type, seq, user_role_key",
        (user_key,)).fetchall()
    return jsonify(
        status="success",
        user=dict(u),
        security_roles=[r["role_name"] for r in roles if r["role_type"] == "Security"],
        functional_roles=[r["role_name"] for r in roles if r["role_type"] == "Functional"],
        editable=USER_EDITABLE_FIELDS)


@app.route("/api/users/<int:user_key>", methods=["POST"])
def api_user_save(user_key: int):
    conn = db()
    data = request.get_json(force=True) or {}
    fields = {k: (v or "") for k, v in (data.get("fields") or {}).items()
              if k in USER_EDITABLE_FIELDS}
    if fields:
        sets = ", ".join(f"{k} = ?" for k in fields)
        conn.execute(
            f"UPDATE M3_Security_Users SET {sets} WHERE user_key = ?",
            list(fields.values()) + [user_key])

    # Roles are replaced wholesale per block, keeping the order given.
    for key, role_type in (("security_roles", "Security"),
                           ("functional_roles", "Functional")):
        if key not in data:
            continue
        names = [str(n).strip() for n in (data[key] or []) if str(n).strip()]
        seen, ordered = set(), []
        for n in names:
            if n not in seen:
                seen.add(n)
                ordered.append(n)
        row = conn.execute(
            "SELECT tenant, person_id FROM M3_Security_Users WHERE user_key = ?",
            (user_key,)).fetchone()
        conn.execute(
            "DELETE FROM M3_Security_UserRoles WHERE user_key = ? AND role_type = ?",
            (user_key, role_type))
        conn.executemany(
            "INSERT INTO M3_Security_UserRoles (user_key, tenant, person_id, "
            "role_type, seq, role_name, row_state) "
            "VALUES (?, ?, ?, ?, ?, ?, 'modified')",
            [(user_key, row["tenant"], row["person_id"], role_type, i, n)
             for i, n in enumerate(ordered, start=1)])

    _touch_user(conn, user_key)
    conn.commit()
    return jsonify(status="success")


@app.route("/api/users/new", methods=["POST"])
def api_user_new():
    conn = db()
    data = request.get_json(force=True) or {}
    tenant = (data.get("tenant") or "").strip()
    if not tenant:
        return jsonify(status="error", message="Tenant is required."), 400
    person_id = (data.get("person_id") or str(uuid.uuid4())).strip()
    fields = {k: (v or "") for k, v in (data.get("fields") or {}).items()
              if k in USER_EDITABLE_FIELDS}
    cols = ["tenant", "person_id", "user_guid", "row_state"] + list(fields)
    vals = [tenant, person_id, person_id, "new"] + list(fields.values())
    cur = conn.execute(
        f"INSERT INTO M3_Security_Users ({', '.join(cols)}) "
        f"VALUES ({', '.join('?' * len(cols))})",
        vals)
    conn.commit()
    return jsonify(status="success", user_key=cur.lastrowid, person_id=person_id)


@app.route("/api/users/<int:user_key>/state", methods=["POST"])
def api_user_state(user_key: int):
    conn = db()
    state = (request.get_json(force=True) or {}).get("row_state", "unchanged")
    if state not in ("unchanged", "modified", "new", "deleted"):
        return jsonify(status="error", message="Bad row_state."), 400
    conn.execute(
        "UPDATE M3_Security_Users SET row_state = ?, "
        "modified_at = datetime('now') WHERE user_key = ?",
        (state, user_key))
    conn.commit()
    return jsonify(status="success")


@app.route("/api/users/<int:user_key>", methods=["DELETE"])
def api_user_delete(user_key: int):
    conn = db()
    # hard=1 really removes the row; otherwise it is only flagged, so the
    # next export can carry the deletion into IFS.
    hard = request.args.get("hard") == "1"
    if hard:
        conn.execute("DELETE FROM M3_Security_Users WHERE user_key = ?",
                     (user_key,))
    else:
        conn.execute(
            "UPDATE M3_Security_Users SET row_state = 'deleted', "
            "modified_at = datetime('now') WHERE user_key = ?",
            (user_key,))
    conn.commit()
    return jsonify(status="success")


# ------------------------------------------------------------------- roles


ROLE_SORTS = {
    "name": "LOWER(r.name)",
    "in_m3": "COALESCE(r.in_m3, -1)",
    "description": "LOWER(COALESCE(r.description,''))",
    "m3_description": "LOWER(COALESCE(r.m3_description,''))",
    "n_members": "n_members",
    "m3_member_count": "COALESCE(r.m3_member_count, -1)",
    "n_user_links": "n_user_links",
    "row_state": "r.row_state",
}


def _role_filter() -> tuple[str, list]:
    """WHERE clause + args shared by the role list and the select-all helper."""
    tenant = request.args.get("tenant", "")
    q = (request.args.get("q") or "").strip()
    m3 = (request.args.get("m3") or "").strip()
    where = ["r.tenant = ?", "r.row_state <> 'deleted'"]
    args = [tenant]
    if q:
        where.append("(r.name LIKE ? OR r.description LIKE ?)")
        args += [f"%{q}%", f"%{q}%"]
    # in_m3 is NULL until the role has been checked against M3 at least once.
    if m3 == "yes":
        where.append("r.in_m3 = 1")
    elif m3 == "no":
        where.append("COALESCE(r.in_m3, 0) = 0")
    elif m3 == "unchecked":
        where.append("r.in_m3 IS NULL")
    return " AND ".join(where), args


@app.route("/api/roles")
def api_roles():
    conn = db()
    page = max(1, int(request.args.get("page", 1)))
    size = min(500, max(10, int(request.args.get("size", 50))))
    clause, args = _role_filter()

    total = conn.execute(
        f"SELECT COUNT(*) FROM M3_Security_Roles r WHERE {clause}",
        args).fetchone()[0]

    rows = conn.execute(
        f"""
        SELECT r.role_key, r.name, r.description, r.row_state,
               r.in_m3, r.m3_description, r.m3_role_type, r.m3_member_count,
               r.m3_checked_at,
               (SELECT COUNT(*) FROM M3_Security_RoleAssignments a
                 WHERE a.role_key = r.role_key AND a.row_state <> 'deleted'
                   AND a.email_id <> '')                     AS n_members,
               (SELECT COUNT(*) FROM M3_Security_UserRoles ur
                 WHERE ur.tenant = r.tenant AND ur.role_name = r.name
                   AND ur.row_state <> 'deleted')            AS n_user_links
        FROM M3_Security_Roles r
        WHERE {clause}
        ORDER BY {_order_by(ROLE_SORTS, 'LOWER(r.name)')}
        LIMIT ? OFFSET ?
        """,
        args + [size, (page - 1) * size]).fetchall()

    return jsonify(status="success", total=total, page=page, size=size,
                   rows=[dict(r) for r in rows])


@app.route("/api/roles/keys")
def api_role_keys():
    """
    Every role key matching the current filter, unpaged - this is what the
    'select all not in M3' control uses, so a selection is not limited to the
    page on screen.
    """
    conn = db()
    clause, args = _role_filter()
    rows = conn.execute(
        f"SELECT r.role_key, r.name, r.in_m3 FROM M3_Security_Roles r "
        f"WHERE {clause} ORDER BY LOWER(r.name)",
        args).fetchall()
    return jsonify(status="success", total=len(rows),
                   rows=[dict(r) for r in rows])


@app.route("/api/roles/remove-members-bulk", methods=["POST"])
def api_roles_remove_members_bulk():
    """
    Strip every member from a set of roles, locally.

    This never calls M3. Roles that do not exist in M3 have nothing to delete
    there; roles that do are handled one at a time through the M3 panel, where
    the MNS410MI/Dlt run is previewed and confirmed on its own.

    Without confirm this is a preview. A real run needs confirm == 'REMOVE'.
    """
    conn = db()
    data = request.get_json(force=True) or {}
    tenant = (data.get("tenant") or "").strip()
    keys = [int(k) for k in (data.get("role_keys") or [])]
    confirm = (data.get("confirm") or "").strip()
    if not tenant:
        return jsonify(status="error", message="Tenant is required."), 400
    if not keys:
        return jsonify(status="error", message="No roles selected."), 400

    marks = ",".join("?" * len(keys))
    rows = conn.execute(
        f"""
        SELECT r.role_key, r.name, r.in_m3,
               (SELECT COUNT(*) FROM M3_Security_RoleAssignments a
                 WHERE a.role_key = r.role_key AND a.email_id <> ''
                   AND a.row_state <> 'deleted')            AS members,
               (SELECT COUNT(*) FROM M3_Security_UserRoles ur
                 WHERE ur.tenant = r.tenant AND ur.role_name = r.name
                   AND ur.row_state <> 'deleted')           AS user_links
        FROM M3_Security_Roles r
        WHERE r.tenant = ? AND r.row_state <> 'deleted' AND r.role_key IN ({marks})
        ORDER BY LOWER(r.name)
        """,
        [tenant] + keys).fetchall()
    plan = [dict(r) for r in rows]

    totals = {
        "roles": len(plan),
        "roles_with_members": sum(1 for p in plan if p["members"]),
        "members": sum(p["members"] for p in plan),
        "user_links": sum(p["user_links"] for p in plan),
        "in_m3": sum(1 for p in plan if p["in_m3"] == 1),
        "not_in_m3": sum(1 for p in plan if p["in_m3"] != 1),
    }

    if confirm != "REMOVE":
        if confirm:
            return jsonify(status="error",
                           message="Type REMOVE to confirm."), 400
        return jsonify(status="success", dry_run=True, plan=plan, totals=totals,
                       message=f"{totals['members']} member(s) would be removed from "
                               f"{totals['roles_with_members']} of "
                               f"{totals['roles']} selected role(s).")

    done = []
    for p in plan:
        res = strip_role_members(conn, tenant, p["name"], commit=False)
        done.append({**p, "removed": res["assignments"],
                     "user_links_removed": res["user_links"]})
    conn.commit()
    return jsonify(
        status="success", dry_run=False, plan=done, totals=totals,
        message=f"{sum(d['removed'] for d in done)} member(s) removed from "
                f"{totals['roles']} role(s). The roles stay in the export with no members.")


@app.route("/api/roles/<int:role_key>")
def api_role(role_key: int):
    conn = db()
    r = conn.execute(
        "SELECT * FROM M3_Security_Roles WHERE role_key = ?",
        (role_key,)).fetchone()
    if not r:
        return jsonify(status="error", message="Role not found."), 404
    members = conn.execute(
        "SELECT assignment_key, email_id, row_state "
        "FROM M3_Security_RoleAssignments WHERE role_key = ? "
        "AND row_state <> 'deleted' "
        "ORDER BY COALESCE(seq, 1000000000), assignment_key",
        (role_key,)).fetchall()
    return jsonify(status="success", role=dict(r),
                   members=[dict(m) for m in members])


@app.route("/api/roles/<int:role_key>", methods=["POST"])
def api_role_save(role_key: int):
    conn = db()
    data = request.get_json(force=True) or {}
    r = conn.execute(
        "SELECT * FROM M3_Security_Roles WHERE role_key = ?",
        (role_key,)).fetchone()
    if not r:
        return jsonify(status="error", message="Role not found."), 404

    name = (data.get("name") or r["name"]).strip()
    desc = data.get("description")
    desc = r["description"] if desc is None else str(desc).strip()
    conn.execute(
        "UPDATE M3_Security_Roles SET name = ?, description = ?, "
        "row_state = CASE WHEN row_state = 'new' THEN 'new' ELSE 'modified' END, "
        "modified_at = datetime('now') WHERE role_key = ?",
        (name, desc, role_key))

    # A rename has to follow the role into every place it is referenced by name.
    if name != r["name"]:
        conn.execute(
            "UPDATE M3_Security_RoleAssignments SET role_name = ? "
            "WHERE role_key = ?",
            (name, role_key))
        conn.execute(
            "UPDATE M3_Security_UserRoles SET role_name = ?, "
            "row_state = 'modified' WHERE tenant = ? AND role_name = ?",
            (name, r["tenant"], r["name"]))

    if "members" in data:
        emails, seen = [], set()
        for e in (data["members"] or []):
            e = str(e).strip()
            if not e:
                continue
            if e.lower() not in seen:
                seen.add(e.lower())
                emails.append(e)
        keep = conn.execute(
            "SELECT email_id, seq FROM M3_Security_RoleAssignments "
            "WHERE role_key = ?",
            (role_key,)).fetchall()
        # Members that were already there keep their slot and stay unchanged.
        old_seq = {k["email_id"]: k["seq"] for k in keep}
        conn.execute("DELETE FROM M3_Security_RoleAssignments WHERE role_key = ?",
                     (role_key,))
        payload = [(role_key, r["tenant"], name, e, old_seq.get(e),
                    "unchanged" if e in old_seq else "new")
                   for e in (emails or [""])]
        conn.executemany(
            "INSERT OR IGNORE INTO M3_Security_RoleAssignments "
            "(role_key, tenant, role_name, email_id, seq, row_state) "
            "VALUES (?, ?, ?, ?, ?, ?)",
            payload)

    conn.commit()
    return jsonify(status="success")


@app.route("/api/roles/new", methods=["POST"])
def api_role_new():
    conn = db()
    data = request.get_json(force=True) or {}
    tenant = (data.get("tenant") or "").strip()
    name = (data.get("name") or "").strip()
    if not tenant or not name:
        return jsonify(status="error",
                       message="Tenant and name are required."), 400
    try:
        cur = conn.execute(
            "INSERT INTO M3_Security_Roles (tenant, name, description, row_state) "
            "VALUES (?, ?, ?, 'new')",
            (tenant, name, (data.get("description") or name).strip()))
    except sqlite3.IntegrityError:
        return jsonify(status="error",
                       message=f"Role '{name}' already exists."), 400
    role_key = cur.lastrowid
    # An empty member row keeps the role in the export even with no members.
    conn.execute(
        "INSERT INTO M3_Security_RoleAssignments "
        "(role_key, tenant, role_name, email_id, row_state) "
        "VALUES (?, ?, ?, '', 'new')",
        (role_key, tenant, name))
    conn.commit()
    return jsonify(status="success", role_key=role_key)


@app.route("/api/roles/add-from-mns405", methods=["POST"])
def api_roles_add_from_mns405():
    """
    Copy the MNS405 roles the capture is missing into IFS Security Roles.

    Without a confirm this is a preview: the plan comes back and nothing is
    written. With confirm='CREATE' the roles are added as 'new', which is what
    puts them in the next Security Role export. Nothing is sent to M3 - the
    roles are already there.
    """
    conn = db()
    data = request.get_json(force=True) or {}
    tenant = (data.get("tenant") or "").strip()
    confirm = data.get("confirm")
    if not tenant:
        return jsonify(status="error", message="A tenant is required."), 400

    # An explicit list limits it; without one, every missing role goes.
    names = data.get("names")

    plan = roles_missing_from_m3(conn, tenant)
    if names is not None:
        wanted = {str(n).strip() for n in names if str(n).strip()}
        plan = [p for p in plan if p["roll"] in wanted]

    if not confirm:
        return jsonify(
            status="success", dry_run=True, plan=plan, added=[], skipped=[],
            totals={"added": 0, "skipped": 0, "considered": len(plan)},
            message=f"{len(plan)} role(s) in MNS405 are not in the IFS "
                    f"Security Roles capture.")

    if confirm != "CREATE":
        return jsonify(status="error", message="Type CREATE to confirm."), 400
    if not plan:
        return jsonify(status="error", message="Nothing to add."), 400

    res = add_roles_from_m3(conn, tenant, [p["roll"] for p in plan])
    n = res["totals"]["added"]
    return jsonify(
        status="success", dry_run=False, plan=plan, **res,
        message=f"{n} role(s) added to IFS Security Roles. They are flagged "
                f"new - export the changes to create them in IFS.")


@app.route("/api/roles/<int:role_key>", methods=["DELETE"])
def api_role_delete(role_key: int):
    conn = db()
    if request.args.get("hard") == "1":
        conn.execute("DELETE FROM M3_Security_Roles WHERE role_key = ?",
                     (role_key,))
    else:
        conn.execute(
            "UPDATE M3_Security_Roles SET row_state = 'deleted', "
            "modified_at = datetime('now') WHERE role_key = ?",
            (role_key,))
    conn.commit()
    return jsonify(status="success")


@app.route("/api/role-names")
def api_role_names():
    """
    Pick-list of role names known for a tenant (both sources).

    type 'Functional' narrows it to the functional security roles, 'Security'
    to the plain ones; without it the list is everything, as before.
    """
    conn = db()
    tenant = request.args.get("tenant", "")
    role_type = (request.args.get("type") or "").strip()
    if role_type == "Functional":
        rows = conn.execute(
            """
            SELECT name FROM M3_Security_FunctionalRoles
             WHERE tenant = ? AND row_state <> 'deleted'
            UNION
            SELECT DISTINCT role_name FROM M3_Security_UserRoles
             WHERE tenant = ? AND row_state <> 'deleted'
               AND role_type = 'Functional'
            ORDER BY 1
            """,
            (tenant, tenant)).fetchall()
    elif role_type == "Security":
        rows = conn.execute(
            """
            SELECT name FROM M3_Security_Roles
             WHERE tenant = ? AND row_state <> 'deleted'
            UNION
            SELECT DISTINCT role_name FROM M3_Security_UserRoles
             WHERE tenant = ? AND row_state <> 'deleted'
               AND role_type = 'Security'
            ORDER BY 1
            """,
            (tenant, tenant)).fetchall()
    else:
        rows = conn.execute(
            """
            SELECT name FROM M3_Security_Roles
             WHERE tenant = ? AND row_state <> 'deleted'
            UNION
            SELECT DISTINCT role_name FROM M3_Security_UserRoles
             WHERE tenant = ? AND row_state <> 'deleted'
            ORDER BY 1
            """,
            (tenant, tenant)).fetchall()
    return jsonify(status="success", names=[r[0] for r in rows])


# ----------------------------------------------------------------- exports


@app.route("/api/changes")
def api_changes():
    """How many records a changes-only export would carry."""
    conn = db()
    tenant = request.args.get("tenant", "")
    if not tenant:
        return jsonify(status="success", users=0, roles=0, functional=0,
                       missing_from_mns405=0)
    # Also how many MNS405 roles the capture has never seen, so the roles
    # toolbar can offer to pull them in without a round trip of its own.
    missing = conn.execute(
        """
        SELECT COUNT(*) FROM M3_Security_M3RoleDefs d
         WHERE d.tenant = ?
           AND NOT EXISTS (SELECT 1 FROM M3_Security_Roles r
                            WHERE r.tenant = d.tenant AND r.name = d.roll)
        """,
        (tenant,)).fetchone()[0]
    return jsonify(status="success", missing_from_mns405=missing,
                   **changed_counts(conn, tenant))


@app.route("/api/export/<kind>")
def api_export(kind: str):
    conn = db()
    tenant = request.args.get("tenant", "")
    new_name = request.args.get("new_name") == "1"
    scope = request.args.get("scope", "changes")
    if kind not in ("users", "roles", "functional"):
        return jsonify(status="error",
                       message="kind must be users, roles or functional."), 400
    if scope not in ("changes", "full"):
        return jsonify(status="error",
                       message="scope must be changes or full."), 400
    fn = {"users": export_users, "roles": export_roles,
          "functional": export_functional}[kind]
    try:
        path = fn(conn, tenant, OUTPUT_DIR / scope,
                  reuse_source_name=not new_name, scope=scope)
    except ValueError as exc:
        return jsonify(status="error", message=str(exc)), 400
    return send_file(str(path), as_attachment=True,
                     download_name=path.name, mimetype="text/csv")


@app.route("/api/mark-pushed", methods=["POST"])
def api_mark_pushed():
    """Clear the change flags once a delta has been pushed into M3."""
    conn = db()
    data = request.get_json(force=True) or {}
    tenant = (data.get("tenant") or "").strip()
    kind = data.get("kind", "both")
    if not tenant:
        return jsonify(status="error", message="Tenant is required."), 400
    if kind not in ("users", "roles", "both"):
        return jsonify(status="error",
                       message="kind must be users, roles or both."), 400
    return jsonify(status="success", **mark_pushed(conn, tenant, kind))


@app.route("/api/output-dir")
def api_output_dir():
    return jsonify(status="success", output_dir=str(OUTPUT_DIR),
                   upload_dir=str(UPLOAD_DIR))


# ---------------------------------------------------------------- M3 client


def _company_for(tenant: str) -> tuple:
    """
    (cono, divi) to send for a tenant, and where the answer came from.

    A --company / --division given on the command line pins every tenant;
    otherwise the tenant's entry in M3_Security_M3.json decides; with neither,
    both are left off the request and M3 uses the service account's own default
    company and division. Guessing here is what produced 403s on tenants that
    do not number their companies 001/100.
    """
    cono = app.config.get("M3_COMPANY", DEFAULT_COMPANY)
    divi = app.config.get("M3_DIVISION", DEFAULT_DIVISION)
    if cono or divi:
        return cono, divi, "command line"
    cono, divi = tenant_company(tenant, M3_CONFIG_PATH)
    if cono or divi:
        return cono, divi, "M3_Security_M3.json"
    return None, None, "service account default"


def _client(tenant: str) -> M3Client:
    """Build an M3 client for a tenant using the app's connection settings."""
    if not tenant:
        raise M3ApiError("Tenant is required.")
    cono, divi, _src = _company_for(tenant)
    return M3Client(
        tenant,
        ionapi_dir=app.config.get("M3_IONAPI_DIR", DEFAULT_IONAPI_DIR),
        company=cono,
        division=divi,
        m3user=app.config.get("M3_USER"),
        batch_size=app.config.get("M3_BATCH_SIZE", DEFAULT_BATCH_SIZE))


# ------------------------------------------------------------ job plumbing


# Long M3 runs happen on a worker thread; the browser polls /api/job/<id>.
_JOBS: dict[str, dict] = {}
_JOBS_LOCK = threading.Lock()
_JOB_KEEP = 40


def _job_set(job_id: str, **fields) -> None:
    with _JOBS_LOCK:
        job = _JOBS.get(job_id)
        if job is not None:
            job.update(fields)


def start_job(label: str, fn) -> str:
    """
    Run fn(conn, progress) on a worker thread.

    progress(done, total, message) updates what /api/job/<id> reports.
    """
    job_id = str(uuid.uuid4())
    with _JOBS_LOCK:
        # Keep the newest _JOB_KEEP finished jobs; a running one is never cut.
        for old in sorted(_JOBS, key=lambda k: _JOBS[k]["started_at"])[:-_JOB_KEEP]:
            if _JOBS[old]["status"] != "running":
                _JOBS.pop(old, None)
        _JOBS[job_id] = {"id": job_id, "label": label, "status": "running",
                         "done": 0, "total": 0, "phase": "", "result": None,
                         "error": None, "started_at": time.time(),
                         "finished_at": None}

    # The worker gets its own connection - sqlite3 objects are not shareable.
    db_path = app.config["M3_SECURITY_DB"]

    def progress(done: int, total: int, message: str = "") -> None:
        _job_set(job_id, done=done, total=total, phase=message or "")

    def run() -> None:
        conn = connect(db_path)
        try:
            result = fn(conn, progress)
            _job_set(job_id, status="done", result=result,
                     finished_at=time.time())
        except Exception as exc:
            log.error("job %s failed: %s", label, traceback.format_exc())
            _job_set(job_id, status="error", error=str(exc),
                     finished_at=time.time())
        finally:
            conn.close()

    threading.Thread(target=run, name=f"m3sec-{label}", daemon=True).start()
    return job_id


@app.route("/api/job/<job_id>")
def api_job(job_id: str):
    with _JOBS_LOCK:
        job = _JOBS.get(job_id)
        snapshot = dict(job) if job else None
    if snapshot is None:
        return jsonify(status="error", message="Unknown job."), 404
    snapshot.pop("started_at", None)
    snapshot.pop("finished_at", None)
    return jsonify(status="success", job=snapshot)


@app.errorhandler(M3ApiError)
def _handle_m3(exc):
    return jsonify(status="error", message=str(exc)), 400


# ------------------------------------------------------------- M3: general


@app.route("/api/m3/config")
def api_m3_config():
    """Connection settings and which tenants have an .ionapi file."""
    files = list_ionapi_files(app.config.get("M3_IONAPI_DIR", DEFAULT_IONAPI_DIR))
    tenant = (request.args.get("tenant") or "").strip()
    cono, divi, source = _company_for(tenant) if tenant else (None, None, "")
    return jsonify(
        status="success",
        ionapi_dir=str(app.config.get("M3_IONAPI_DIR", DEFAULT_IONAPI_DIR)),
        config_file=str(M3_CONFIG_PATH),
        tenant=tenant,
        company=cono or "",
        division=divi or "",
        source=source,
        pinned=bool(cono or divi),
        m3user=app.config.get("M3_USER"),
        files=[{"file": f["file"], "tenant": f.get("tenant")} for f in files])


@app.route("/api/m3/company", methods=["POST"])
def api_m3_company():
    """Pin a tenant's CONO / DIVI, or clear both by sending blanks."""
    data = request.get_json(force=True) or {}
    tenant = (data.get("tenant") or "").strip()
    if not tenant:
        return jsonify(status="error", message="Tenant is required."), 400
    entry = save_tenant_company(tenant, data.get("company"),
                                data.get("division"), M3_CONFIG_PATH)
    cono, divi, source = _company_for(tenant)
    return jsonify(
        status="success", file=str(M3_CONFIG_PATH), tenant=tenant,
        company=cono or "", division=divi or "", source=source,
        saved=entry,
        message=f"{tenant} pinned to company {entry['cono'] or '—'}"
                f" division {entry['divi'] or '—'}"
                if entry["cono"] or entry["divi"]
                else f"{tenant} unpinned — M3 will use the service account's own company and division.")


@app.route("/api/m3/companies", methods=["POST"])
def api_m3_companies():
    """Read the companies and divisions this tenant actually has."""
    data = request.get_json(force=True) or {}
    tenant = (data.get("tenant") or "").strip()
    client = _client(tenant)
    rows = discover_companies(client)
    return jsonify(status="success", tenant=tenant, companies=rows,
                   message=f"{len(rows)} company/companies in {tenant}"
                           if rows
                           else f"M3 returned no companies for {tenant}.")


@app.route("/api/m3/sync-roles", methods=["POST"])
def api_m3_sync_roles():
    """Flag which captured roles exist in M3 (MNS405MI/Lst)."""
    conn = db()
    data = request.get_json(force=True) or {}
    tenant = (data.get("tenant") or "").strip()
    client = _client(tenant)
    with_members = bool(data.get("with_members"))
    with_users = bool(data.get("with_users"))

    def work(conn, progress):
        if with_users:
            progress(0, 1, "Reading M3 users")
            users = sync_users(conn, client)
        result = sync_roles(conn, client, with_members=with_members,
                            progress=progress)
        if with_users:
            result["m3_users"] = users
        result["ionapi"] = client.ionapi_path.name
        return result

    return jsonify(status="success", job_id=start_job("Check M3", work))


@app.route("/api/m3/role")
def api_m3_role():
    """
    M3 membership for one role, side by side with what the CSV capture holds.
    Pass refresh=1 to re-read it from M3 rather than use the stored copy.
    """
    conn = db()
    tenant = request.args.get("tenant", "")
    role = (request.args.get("role") or "").strip()
    if not role:
        return jsonify(status="error", message="role is required."), 400

    if request.args.get("refresh") == "1":
        sync_role_members(conn, _client(tenant), role)

    m3 = conn.execute(
        """
        SELECT m.usid, m.valid_from, m.valid_to,
               u.email AS email, u.name AS name, u.status AS status
        FROM M3_Security_M3Members m
        LEFT JOIN M3_Security_M3Users u
               ON u.tenant = m.tenant AND u.usid = m.usid
        WHERE m.tenant = ? AND m.role_name = ?
        ORDER BY m.usid
        """,
        (tenant, role)).fetchall()
    csv_members = [r[0] for r in conn.execute(
        "SELECT email_id FROM M3_Security_RoleAssignments WHERE tenant = ? "
        "AND role_name = ? AND email_id <> '' AND row_state <> 'deleted' "
        "ORDER BY email_id",
        (tenant, role)).fetchall()]

    # Matching is by email, case-insensitively - M3 stores a USID, the capture
    # stores the address the role was granted to.
    m3_emails = {(r["email"] or "").strip().lower() for r in m3 if r["email"]}
    csv_lower = {e.lower() for e in csv_members}
    role_row = conn.execute(
        "SELECT in_m3, m3_description, m3_role_type, m3_checked_at "
        "FROM M3_Security_Roles WHERE tenant = ? AND name = ?",
        (tenant, role)).fetchone()

    return jsonify(
        status="success",
        role=role,
        in_m3=role_row["in_m3"] if role_row else None,
        m3_description=role_row["m3_description"] if role_row else None,
        m3_role_type=role_row["m3_role_type"] if role_row else None,
        checked_at=role_row["m3_checked_at"] if role_row else None,
        m3_members=[dict(r) for r in m3],
        csv_members=csv_members,
        only_in_m3=sorted(
            (r["email"] or r["usid"])
            for r in m3
            if not r["email"] or r["email"].strip().lower() not in csv_lower),
        only_in_csv=sorted(e for e in csv_members if e.lower() not in m3_emails))


@app.route("/api/m3/remove-members", methods=["POST"])
def api_m3_remove_members():
    """
    Remove every member of a role in M3.

    Without confirm this is a dry run: it re-reads the membership from M3 and
    reports what a real run would delete. A real run needs confirm == role.
    """
    conn = db()
    data = request.get_json(force=True) or {}
    tenant = (data.get("tenant") or "").strip()
    role = (data.get("role") or "").strip()
    confirm = data.get("confirm")
    if not role:
        return jsonify(status="error", message="role is required."), 400

    client = _client(tenant)
    clear_local = bool(data.get("clear_local", True))
    if confirm and str(confirm).strip() != role:
        return jsonify(status="error",
                       message=f"Confirmation does not match. To remove every member of '{role}'"
                               f" the confirm value must be exactly '{role}'"
                               f"."), 400
    if not confirm:
        return jsonify(status="success",
                       **remove_role_members(conn, client, role, dry_run=True,
                                             clear_local=clear_local))
    return jsonify(status="success", job_id=start_job(
        f"Remove members of {role}",
        lambda c, p: remove_role_members(c, client, role, dry_run=False,
                                         confirm=confirm,
                                         clear_local=clear_local, progress=p)))


@app.route("/api/m3/create-roles", methods=["POST"])
def api_m3_create_roles():
    """
    Create the selected captured roles in M3 with MNS405MI/Add.

    The roles can be named three ways: 'roles' outright, 'role_keys' from the
    security roles tab, or 'fsr_keys' from the functional security role tab -
    which stands for every security role inside those functional roles, so one
    functional role creates the whole set it is made of.

    Without confirm this is a preview: it returns what would be created and
    what would be skipped, with a reason per skip. A real run needs
    confirm == 'CREATE'.
    """
    conn = db()
    data = request.get_json(force=True) or {}
    tenant = (data.get("tenant") or "").strip()
    keys = [int(k) for k in (data.get("role_keys") or [])]
    fsr_keys = [int(k) for k in (data.get("fsr_keys") or [])]
    names = [str(n) for n in (data.get("roles") or [])]
    confirm = data.get("confirm")
    if not tenant:
        return jsonify(status="error", message="Tenant is required."), 400

    if keys and not names:
        marks = ",".join("?" * len(keys))
        names = [r[0] for r in conn.execute(
            f"SELECT name FROM M3_Security_Roles WHERE tenant = ? "
            f"AND role_key IN ({marks}) ORDER BY LOWER(name)",
            [tenant] + keys).fetchall()]
    sources = []
    if fsr_keys and not names:
        marks = ",".join("?" * len(fsr_keys))
        rows = conn.execute(
            f"""
            SELECT m.security_role AS role,
                   GROUP_CONCAT(DISTINCT f.name) AS functional_roles
            FROM M3_Security_FunctionalRoleMembers m
            JOIN M3_Security_FunctionalRoles f ON f.fsr_key = m.fsr_key
            WHERE m.tenant = ? AND m.fsr_key IN ({marks})
              AND m.row_state <> 'deleted' AND m.security_role <> ''
            GROUP BY m.security_role
            ORDER BY LOWER(m.security_role)
            """,
            [tenant] + fsr_keys).fetchall()
        # sources says which functional role each security role came from.
        names = [r["role"] for r in rows]
        sources = [dict(r) for r in rows]
        if not names:
            return jsonify(status="error",
                           message="The selected functional role(s) hold no security roles."), 400
    if not names:
        return jsonify(status="error", message="No roles selected."), 400

    client = _client(tenant)
    role_type = data.get("role_type") or None
    if confirm and confirm != "CREATE":
        return jsonify(status="error", message="Type CREATE to confirm."), 400
    if not confirm:
        return jsonify(status="success", sources=sources,
                       **create_roles(conn, client, names, dry_run=True,
                                      role_type=role_type))
    return jsonify(status="success", job_id=start_job(
        "Create roles in M3",
        lambda c, p: create_roles(c, client, names, dry_run=False,
                                  confirm=confirm, role_type=role_type,
                                  progress=p)))


@app.route("/api/m3/add-members", methods=["POST"])
def api_m3_add_members():
    """
    Give the captured members of the selected roles that role in M3
    (MNS410MI/Add). USID comes from matching the member's email address to the
    M3 user list. Preview unless confirm == 'ADD'.
    """
    conn = db()
    data = request.get_json(force=True) or {}
    tenant = (data.get("tenant") or "").strip()
    keys = [int(k) for k in (data.get("role_keys") or [])]
    names = [str(n) for n in (data.get("roles") or [])]
    confirm = data.get("confirm")
    if not tenant:
        return jsonify(status="error", message="Tenant is required."), 400

    if keys and not names:
        marks = ",".join("?" * len(keys))
        names = [r[0] for r in conn.execute(
            f"SELECT name FROM M3_Security_Roles WHERE tenant = ? "
            f"AND role_key IN ({marks}) ORDER BY LOWER(name)",
            [tenant] + keys).fetchall()]
    if not names:
        return jsonify(status="error", message="No roles selected."), 400

    known = conn.execute(
        "SELECT COUNT(*) FROM M3_Security_M3Users WHERE tenant = ?",
        (tenant,)).fetchone()[0]
    if not known:
        return jsonify(status="error",
                       message="No M3 user list held for this tenant, so emails "
                               "cannot be resolved to a USID. Run Check M3 first."
                       ), 400

    client = _client(tenant)
    if confirm and confirm != "ADD":
        return jsonify(status="error", message="Type ADD to confirm."), 400
    if not confirm:
        return jsonify(status="success",
                       **add_role_members(conn, client, names, dry_run=True))
    return jsonify(status="success", job_id=start_job(
        "Add members in M3",
        lambda c, p: add_role_members(c, client, names, dry_run=False,
                                      confirm=confirm, progress=p)))


# --------------------------------------------------------------- functions


@app.route("/api/functions/sync", methods=["POST"])
def api_functions_sync():
    """Pull SES400MI/Lst and flag which of its roles MNS405 is missing."""
    data = request.get_json(force=True) or {}
    tenant = (data.get("tenant") or "").strip()
    client = _client(tenant)
    return jsonify(status="success", job_id=start_job(
        "Load functions from SES400",
        lambda c, p: {**sync_function_roles(c, client, progress=p),
                      "ionapi": client.ionapi_path.name}))


@app.route("/api/functions/status")
def api_functions_status():
    """What is held for the Functions tab right now."""
    conn = db()
    tenant = request.args.get("tenant", "")
    row = conn.execute(
        """
        SELECT
          (SELECT COUNT(*) FROM M3_Security_FunctionRoles WHERE tenant = ?1)          AS authorisations,
          (SELECT COUNT(DISTINCT fnid) FROM M3_Security_FunctionRoles WHERE tenant = ?1) AS functions,
          (SELECT COUNT(*) FROM M3_Security_FunctionRoleStatus WHERE tenant = ?1)     AS roles,
          (SELECT COUNT(*) FROM M3_Security_FunctionRoleStatus
            WHERE tenant = ?1 AND COALESCE(in_mns405, 0) = 0)                         AS missing,
          (SELECT MAX(checked_at) FROM M3_Security_FunctionRoles WHERE tenant = ?1)   AS checked_at
        """,
        (tenant,)).fetchone()
    divisions = [dict(r) for r in conn.execute(
        "SELECT DISTINCT cono, divi FROM M3_Security_FunctionRoles "
        "WHERE tenant = ? ORDER BY cono, divi",
        (tenant,)).fetchall()]
    statuses = [dict(r) for r in conn.execute(
        "SELECT COALESCE(stat, '') AS stat, COUNT(*) AS n "
        "FROM M3_Security_FunctionRoles WHERE tenant = ? "
        "GROUP BY COALESCE(stat, '') ORDER BY stat",
        (tenant,)).fetchall()]
    return jsonify(status="success", **dict(row), divisions=divisions,
                   statuses=statuses,
                   status_values=[{"code": k, "label": v}
                                  for k, v in FUNCTION_STATUS.items()])


FN_ROLE_SORTS = {
    "name": "LOWER(f.roll)",
    "in_mns405": "COALESCE(in_mns405, 0)",
    "description": "LOWER(COALESCE(description,''))",
    "n_functions": "n_functions",
    "n_rows": "n_rows",
}
FN_FUNCTION_SORTS = {
    "name": "LOWER(f.fnid)",
    "n_roles": "n_roles",
    "n_rows": "n_rows",
    "n_missing": "n_missing",
}


def _fn_filter(alias: str = "f") -> tuple[str, list]:
    tenant = request.args.get("tenant", "")
    cono = (request.args.get("cono") or "").strip()
    divi = (request.args.get("divi") or "").strip()
    stat = (request.args.get("stat") or "").strip()
    where = [f"{alias}.tenant = ?"]
    args = [tenant]
    if cono:
        where.append(f"{alias}.cono = ?")
        args.append(cono)
    if divi:
        where.append(f"{alias}.divi = ?")
        args.append(divi)
    if stat:
        where.append(f"COALESCE({alias}.stat, '') = ?")
        # The page sends "(blank)" for rows SES400 left without a status.
        args.append("" if stat == "(blank)" else stat)
    return " AND ".join(where), args


@app.route("/api/functions")
def api_functions():
    """
    The Functions tab list.

    by=role      one row per role, with how many functions it is authorised to
    by=function  one row per function, with how many roles hold it
    """
    conn = db()
    by = (request.args.get("by") or "role").strip()
    if by not in ("role", "function"):
        return jsonify(status="error", message="by must be role or function."), 400
    q = (request.args.get("q") or "").strip()
    missing_only = request.args.get("missing") == "1"
    page = max(1, int(request.args.get("page", 1)))
    size = min(500, max(10, int(request.args.get("size", 50))))

    clause, args = _fn_filter()
    if by == "role":
        if q:
            clause += " AND f.roll LIKE ?"
            args.append(f"%{q}%")
        if missing_only:
            clause += " AND COALESCE((SELECT s.in_mns405 FROM M3_Security_FunctionRoleStatus s " \
                      " WHERE s.tenant = f.tenant AND s.roll = f.roll), 0) = 0"
        total = conn.execute(
            f"SELECT COUNT(DISTINCT f.roll) FROM M3_Security_FunctionRoles f WHERE "
            f"{clause}", args).fetchone()[0]
        rows = conn.execute(
            f"""
            SELECT f.roll AS name,
                   COUNT(DISTINCT f.fnid) AS n_functions,
                   COUNT(*)               AS n_rows,
                   (SELECT s.in_mns405 FROM M3_Security_FunctionRoleStatus s
                     WHERE s.tenant = f.tenant AND s.roll = f.roll) AS in_mns405,
                   (SELECT r.description FROM M3_Security_Roles r
                     WHERE r.tenant = f.tenant AND r.name = f.roll)  AS description
            FROM M3_Security_FunctionRoles f
            WHERE {clause}
            GROUP BY f.roll
            ORDER BY {_order_by(FN_ROLE_SORTS, 'LOWER(f.roll)')}
            LIMIT ? OFFSET ?
            """,
            args + [size, (page - 1) * size]).fetchall()
    else:
        if q:
            clause += " AND f.fnid LIKE ?"
            args.append(f"%{q}%")
        total = conn.execute(
            f"SELECT COUNT(DISTINCT f.fnid) FROM M3_Security_FunctionRoles f WHERE "
            f"{clause}", args).fetchone()[0]
        rows = conn.execute(
            f"""
            SELECT f.fnid AS name,
                   COUNT(DISTINCT f.roll) AS n_roles,
                   COUNT(*)               AS n_rows,
                   SUM(CASE WHEN COALESCE((SELECT s.in_mns405
                                             FROM M3_Security_FunctionRoleStatus s
                                            WHERE s.tenant = f.tenant AND s.roll = f.roll), 0) = 0
                            THEN 1 ELSE 0 END) AS n_missing
            FROM M3_Security_FunctionRoles f
            WHERE {clause}
            GROUP BY f.fnid
            ORDER BY {_order_by(FN_FUNCTION_SORTS, 'LOWER(f.fnid)')}
            LIMIT ? OFFSET ?
            """,
            args + [size, (page - 1) * size]).fetchall()

    return jsonify(status="success", by=by, total=total, page=page, size=size,
                   rows=[dict(r) for r in rows])


@app.route("/api/functions/detail")
def api_functions_detail():
    """
    The other side of the pairing.

    by=role     name is a role     -> the functions it is authorised to
    by=function name is a function -> the roles authorised to it
    """
    conn = db()
    by = (request.args.get("by") or "role").strip()
    name = (request.args.get("name") or "").strip()
    if not name:
        return jsonify(status="error", message="name is required."), 400
    clause, args = _fn_filter()

    if by == "role":
        rows = conn.execute(
            f"""
            SELECT f.function_role_key, f.fnid AS name, f.cono, f.divi, f.stat
            FROM M3_Security_FunctionRoles f
            WHERE {clause} AND f.roll = ?
            ORDER BY f.fnid, f.cono, f.divi
            """,
            args + [name]).fetchall()
        status = conn.execute(
            "SELECT in_mns405 FROM M3_Security_FunctionRoleStatus "
            "WHERE tenant = ? AND roll = ?",
            (request.args.get("tenant", ""), name)).fetchone()
        return jsonify(status="success", by=by, name=name,
                       in_mns405=status["in_mns405"] if status else None,
                       rows=[dict(r) for r in rows])

    rows = conn.execute(
        f"""
        SELECT f.function_role_key, f.roll AS name, f.cono, f.divi, f.stat,
               (SELECT s.in_mns405 FROM M3_Security_FunctionRoleStatus s
                 WHERE s.tenant = f.tenant AND s.roll = f.roll) AS in_mns405
        FROM M3_Security_FunctionRoles f
        WHERE {clause} AND f.fnid = ?
        ORDER BY f.roll, f.cono, f.divi
        """,
        args + [name]).fetchall()
    return jsonify(status="success", by=by, name=name,
                   rows=[dict(r) for r in rows])


@app.route("/api/functions/keys")
def api_function_keys():
    """
    Authorisation keys behind the current view, unpaged.

    With `name` (and `by`) it expands one aggregate row - every authorisation
    for that role, or for that function. Without it, everything matching the
    current filter. This is what lets a tick on an aggregate row select the
    records underneath it.
    """
    conn = db()
    by = (request.args.get("by") or "role").strip()
    if by not in ("role", "function"):
        return jsonify(status="error", message="by must be role or function."), 400
    name = (request.args.get("name") or "").strip()
    q = (request.args.get("q") or "").strip()
    missing_only = request.args.get("missing") == "1"

    clause, args = _fn_filter()
    column = "roll" if by == "role" else "fnid"
    if name:
        clause += f" AND f.{column} = ?"
        args.append(name)
    elif q:
        clause += f" AND f.{column} LIKE ?"
        args.append(f"%{q}%")
    if by == "role" and missing_only:
        clause += " AND COALESCE((SELECT s.in_mns405 FROM M3_Security_FunctionRoleStatus s " \
                  " WHERE s.tenant = f.tenant AND s.roll = f.roll), 0) = 0"

    rows = conn.execute(
        f"""
        SELECT f.function_role_key, f.fnid, f.roll, f.cono, f.divi
        FROM M3_Security_FunctionRoles f
        WHERE {clause}
        ORDER BY f.fnid, f.roll, f.cono, f.divi
        """,
        args).fetchall()
    return jsonify(status="success", total=len(rows),
                   rows=[dict(r) for r in rows])


@app.route("/api/functions/update-status", methods=["POST"])
def api_functions_update_status():
    """
    Set the status on the selected SES400 authorisations (SES400MI/Upd).

    Only the key fields and STAT are sent, so the option flags on each
    authorisation are left alone. Preview unless confirm == 'UPDATE'.
    """
    conn = db()
    data = request.get_json(force=True) or {}
    tenant = (data.get("tenant") or "").strip()
    keys = [int(k) for k in (data.get("keys") or [])]
    new_status = (data.get("status") or "").strip()
    confirm = data.get("confirm")
    if not tenant:
        return jsonify(status="error", message="Tenant is required."), 400
    if not keys:
        return jsonify(status="error", message="No authorisations selected."), 400
    if not new_status:
        return jsonify(status="error", message="A status is required."), 400
    client = _client(tenant)
    if confirm and confirm != "UPDATE":
        return jsonify(status="error", message="Type UPDATE to confirm."), 400

    if not confirm:
        return jsonify(status="success", **update_function_role_status(
            conn, client, keys, new_status, dry_run=True))
    return jsonify(status="success", job_id=start_job(
        f"Set status {new_status}",
        lambda c, p: update_function_role_status(
            c, client, keys, new_status, dry_run=False,
            confirm=confirm, progress=p)))


@app.route("/api/functions/delete", methods=["POST"])
def api_functions_delete():
    """
    Delete the selected SES400 authorisations (SES400MI/Dlt).

    Preview unless confirm == 'DELETE'.
    """
    conn = db()
    data = request.get_json(force=True) or {}
    tenant = (data.get("tenant") or "").strip()
    keys = [int(k) for k in (data.get("keys") or [])]
    confirm = data.get("confirm")
    if not tenant:
        return jsonify(status="error", message="Tenant is required."), 400
    if not keys:
        return jsonify(status="error", message="No authorisations selected."), 400
    client = _client(tenant)
    if confirm and confirm != "DELETE":
        return jsonify(status="error", message="Type DELETE to confirm."), 400

    if not confirm:
        return jsonify(status="success",
                       **delete_function_roles(conn, client, keys, dry_run=True))
    return jsonify(status="success", job_id=start_job(
        "Delete SES400 authorisations",
        lambda c, p: delete_function_roles(c, client, keys, dry_run=False,
                                           confirm=confirm, progress=p)))


@app.route("/api/functions/add-missing-roles", methods=["POST"])
def api_functions_add_missing():
    """
    Add the SES400 roles MNS405 does not have. Name and description are both
    the role itself, since SES400 carries neither. Preview unless
    confirm == 'CREATE'.
    """
    conn = db()
    data = request.get_json(force=True) or {}
    tenant = (data.get("tenant") or "").strip()
    roles = [str(r) for r in (data.get("roles") or [])] or None
    confirm = data.get("confirm")
    client = _client(tenant)
    if confirm and confirm != "CREATE":
        return jsonify(status="error", message="Type CREATE to confirm."), 400

    if not confirm:
        return jsonify(status="success", **create_missing_function_roles(
            conn, client, roles, dry_run=True))
    return jsonify(status="success", job_id=start_job(
        "Add missing roles to MNS405",
        lambda c, p: create_missing_function_roles(
            c, client, roles, dry_run=False, confirm=confirm,
            role_type=data.get("role_type") or None, progress=p)))


# ------------------------------------------------------------------ MNS405


@app.route("/api/mns405/sync", methods=["POST"])
def api_mns405_sync():
    data = request.get_json(force=True) or {}
    client = _client((data.get("tenant") or "").strip())
    return jsonify(status="success", job_id=start_job(
        "Load MNS405 roles",
        lambda c, p: mns405_sync(c, client, progress=p)))


MNS405_SORTS = {
    "roll": "LOWER(d.roll)",
    "tx40": "LOWER(COALESCE(d.tx40,''))",
    "tx15": "LOWER(COALESCE(d.tx15,''))",
    "rolt": "LOWER(COALESCE(d.rolt,''))",
    "n_users": "n_users",
    "n_functions": "n_functions",
}


@app.route("/api/mns405")
def api_mns405():
    conn = db()
    tenant = request.args.get("tenant", "")
    q = (request.args.get("q") or "").strip()
    page = max(1, int(request.args.get("page", 1)))
    size = min(500, max(10, int(request.args.get("size", 50))))

    where, args = ["d.tenant = ?"], [tenant]
    if q:
        where.append("(d.roll LIKE ? OR d.tx40 LIKE ? OR d.tx15 LIKE ?)")
        args += [f"%{q}%"] * 3
    clause = " AND ".join(where)
    total = conn.execute(
        f"SELECT COUNT(*) FROM M3_Security_M3RoleDefs d WHERE {clause}",
        args).fetchone()[0]
    rows = conn.execute(
        f"""
        SELECT d.role_def_key, d.roll, d.tx40, d.tx15, d.rolt, d.txid,
               (SELECT COUNT(*) FROM M3_Security_M3Members m
                 WHERE m.tenant = d.tenant AND m.role_name = d.roll)   AS n_users,
               (SELECT COUNT(DISTINCT f.fnid) FROM M3_Security_FunctionRoles f
                 WHERE f.tenant = d.tenant AND f.roll = d.roll)        AS n_functions
        FROM M3_Security_M3RoleDefs d
        WHERE {clause}
        ORDER BY {_order_by(MNS405_SORTS, 'LOWER(d.roll)')} LIMIT ? OFFSET ?
        """,
        args + [size, (page - 1) * size]).fetchall()
    checked = conn.execute(
        "SELECT MAX(checked_at) FROM M3_Security_M3RoleDefs WHERE tenant = ?",
        (tenant,)).fetchone()[0]
    return jsonify(status="success", total=total, page=page, size=size,
                   checked_at=checked, rows=[dict(r) for r in rows])


@app.route("/api/mns405/keys")
def api_mns405_keys():
    conn = db()
    tenant = request.args.get("tenant", "")
    q = (request.args.get("q") or "").strip()
    where, args = ["tenant = ?"], [tenant]
    if q:
        where.append("(roll LIKE ? OR tx40 LIKE ? OR tx15 LIKE ?)")
        args += [f"%{q}%"] * 3
    rows = conn.execute(
        f"SELECT role_def_key, roll FROM M3_Security_M3RoleDefs "
        f"WHERE {' AND '.join(where)} ORDER BY LOWER(roll)",
        args).fetchall()
    return jsonify(status="success", total=len(rows),
                   rows=[dict(r) for r in rows])


@app.route("/api/mns405/write", methods=["POST"])
def api_mns405_write():
    """Add, update or delete MNS405 role definitions. Preview unless confirmed."""
    conn = db()
    data = request.get_json(force=True) or {}
    tenant = (data.get("tenant") or "").strip()
    action = (data.get("action") or "").strip()
    confirm = data.get("confirm")
    client = _client(tenant)
    expect = {"add": "CREATE", "update": "UPDATE", "delete": "DELETE"}.get(action)
    if not expect:
        return jsonify(status="error",
                       message="action must be add, update or delete."), 400
    if confirm and confirm != expect:
        return jsonify(status="error", message=f"Type {expect} to confirm."), 400

    if action == "delete":
        keys = [int(k) for k in (data.get("keys") or [])]
        # Deleting a role definition can also raise a BOD for the downstream
        # systems that mirror MNS405.
        emit_bods = data.get("emit_bods", True) is not False
        fn = lambda c, p: mns405_delete(c, client, keys, dry_run=False,
                                        confirm=confirm, emit_bods=emit_bods,
                                        bod_dir=BOD_DIR, progress=p)
        preview = lambda: mns405_delete(conn, client, keys, dry_run=True,
                                        emit_bods=emit_bods)
    else:
        rows = data.get("rows") or []
        write = mns405_add if action == "add" else mns405_update
        fn = lambda c, p: write(c, client, rows, dry_run=False,
                                confirm=confirm, progress=p)
        preview = lambda: write(conn, client, rows, dry_run=True)

    if not confirm:
        return jsonify(status="success", **preview())
    return jsonify(status="success",
                   job_id=start_job(f"MNS405 {action}", fn))


# ------------------------------------------------------------------ MNS410


@app.route("/api/mns410/sync", methods=["POST"])
def api_mns410_sync():
    data = request.get_json(force=True) or {}
    client = _client((data.get("tenant") or "").strip())
    role = (data.get("role") or "").strip()
    user = (data.get("user") or "").strip()
    return jsonify(status="success", job_id=start_job(
        "Load MNS410 roles per user",
        lambda c, p: mns410_sync(c, client, role, user, progress=p)))


MNS410_SORTS = {
    "role_name": "LOWER(m.role_name)",
    "in_mns405": "COALESCE(in_mns405, 0)",
    "usid": "LOWER(m.usid)",
    "email": "LOWER(COALESCE(u.email, u.name, ''))",
    "valid_from": "COALESCE(m.valid_from,'')",
    "valid_to": "COALESCE(m.valid_to,'')",
}


@app.route("/api/mns410")
def api_mns410():
    conn = db()
    tenant = request.args.get("tenant", "")
    q = (request.args.get("q") or "").strip()
    role = (request.args.get("role") or "").strip()
    page = max(1, int(request.args.get("page", 1)))
    size = min(500, max(10, int(request.args.get("size", 50))))

    where, args = ["m.tenant = ?"], [tenant]
    if q:
        where.append("(m.role_name LIKE ? OR m.usid LIKE ?)")
        args += [f"%{q}%"] * 2
    if role:
        where.append("m.role_name = ?")
        args.append(role)
    clause = " AND ".join(where)
    total = conn.execute(
        f"SELECT COUNT(*) FROM M3_Security_M3Members m WHERE {clause}",
        args).fetchone()[0]
    rows = conn.execute(
        f"""
        SELECT m.m3_member_key, m.role_name, m.usid, m.valid_from, m.valid_to,
               u.email AS email, u.name AS user_name,
               (SELECT 1 FROM M3_Security_M3RoleDefs d
                 WHERE d.tenant = m.tenant AND d.roll = m.role_name) AS in_mns405
        FROM M3_Security_M3Members m
        LEFT JOIN M3_Security_M3Users u
               ON u.tenant = m.tenant AND u.usid = m.usid
        WHERE {clause}
        ORDER BY {_order_by(MNS410_SORTS, 'LOWER(m.role_name), LOWER(m.usid)')}
        LIMIT ? OFFSET ?
        """,
        args + [size, (page - 1) * size]).fetchall()
    summary = conn.execute(
        "SELECT COUNT(*) AS rows, COUNT(DISTINCT role_name) AS roles, "
        "COUNT(DISTINCT usid) AS users, MAX(checked_at) AS checked_at "
        "FROM M3_Security_M3Members WHERE tenant = ?",
        (tenant,)).fetchone()
    return jsonify(status="success", total=total, page=page, size=size,
                   summary=dict(summary), rows=[dict(r) for r in rows])


@app.route("/api/mns410/keys")
def api_mns410_keys():
    conn = db()
    tenant = request.args.get("tenant", "")
    q = (request.args.get("q") or "").strip()
    role = (request.args.get("role") or "").strip()
    where, args = ["tenant = ?"], [tenant]
    if q:
        where.append("(role_name LIKE ? OR usid LIKE ?)")
        args += [f"%{q}%"] * 2
    if role:
        where.append("role_name = ?")
        args.append(role)
    rows = conn.execute(
        f"SELECT m3_member_key, role_name, usid FROM M3_Security_M3Members "
        f"WHERE {' AND '.join(where)} ORDER BY role_name, usid",
        args).fetchall()
    return jsonify(status="success", total=len(rows),
                   rows=[dict(r) for r in rows])


@app.route("/api/mns410/write", methods=["POST"])
def api_mns410_write():
    """Add, update or delete role-per-user rows. Preview unless confirmed."""
    conn = db()
    data = request.get_json(force=True) or {}
    tenant = (data.get("tenant") or "").strip()
    action = (data.get("action") or "").strip()
    confirm = data.get("confirm")
    client = _client(tenant)
    expect = {"add": "CREATE", "update": "UPDATE", "delete": "DELETE"}.get(action)
    if not expect:
        return jsonify(status="error",
                       message="action must be add, update or delete."), 400
    if confirm and confirm != expect:
        return jsonify(status="error", message=f"Type {expect} to confirm."), 400

    if action == "delete":
        keys = [int(k) for k in (data.get("keys") or [])]
        fn = lambda c, p: mns410_delete(c, client, keys, dry_run=False,
                                        confirm=confirm, progress=p)
        preview = lambda: mns410_delete(conn, client, keys, dry_run=True)
    else:
        rows = data.get("rows") or []
        write = mns410_add if action == "add" else mns410_update
        fn = lambda c, p: write(c, client, rows, dry_run=False,
                                confirm=confirm, progress=p)
        preview = lambda: write(conn, client, rows, dry_run=True)

    if not confirm:
        return jsonify(status="success", **preview())
    return jsonify(status="success",
                   job_id=start_job(f"MNS410 {action}", fn))


# --------------------------------------------------------------------- BODs


@app.route("/api/bod/config")
def api_bod_config():
    """The BOD identity per tenant, and whether the current tenant has one."""
    tenant = (request.args.get("tenant") or "").strip()
    mine = tenant_bod_config(tenant) if tenant else {}
    return jsonify(status="success", file=str(BOD_CONFIG_PATH), tenant=tenant,
                   owner=mine.get("owner", ""),
                   sequence=mine.get("sequence", 0),
                   version=mine.get("version", ""),
                   application=mine.get("application", ""),
                   environment=mine.get("environment", ""),
                   bod_dir=str(BOD_DIR))


@app.route("/api/bod/config", methods=["POST"])
def api_bod_config_save():
    data = request.get_json(force=True) or {}
    tenant = (data.get("tenant") or "").strip()
    if not tenant:
        return jsonify(status="error", message="Tenant is required."), 400
    cfg = load_bod_config()
    entry = cfg.setdefault(tenant, {})
    if "owner" in data:
        entry["owner"] = (data.get("owner") or "").strip()
    if data.get("sequence") not in (None, ""):
        try:
            entry["sequence"] = int(data["sequence"])
        except (TypeError, ValueError):
            return jsonify(status="error",
                           message="sequence must be a whole number."), 400
    for key in ("version", "application", "environment"):
        if data.get(key):
            entry[key] = str(data[key]).strip()
    save_bod_config(cfg)
    return jsonify(status="success", file=str(BOD_CONFIG_PATH), **entry)


@app.route("/api/bod/file/<path:name>")
def api_bod_file(name: str):
    """Download one generated BOD."""
    # Only the basename is honoured, and the result has to stay inside BOD_DIR.
    path = (BOD_DIR / Path(name).name).resolve()
    if not str(path).startswith(str(BOD_DIR.resolve())) or not path.exists():
        return jsonify(status="error", message="Not found."), 404
    return send_file(str(path), as_attachment=True, download_name=path.name,
                     mimetype="application/xml")


# ------------------------------------------- functional security roles (FSR)


FSR_SORTS = {
    "name": "LOWER(f.name)",
    "description": "LOWER(COALESCE(f.description,''))",
    "n_roles": "n_roles",
    "n_users": "n_users",
    "row_state": "f.row_state",
}


def _fsr_filter() -> tuple[str, list]:
    tenant = request.args.get("tenant", "")
    q = (request.args.get("q") or "").strip()
    role = (request.args.get("role") or "").strip()
    where = ["f.tenant = ?", "f.row_state <> 'deleted'"]
    args = [tenant]
    if q:
        where.append("(f.name LIKE ? OR f.description LIKE ?)")
        args += [f"%{q}%", f"%{q}%"]
    if role:
        where.append("EXISTS (SELECT 1 FROM M3_Security_FunctionalRoleMembers m "
                     "WHERE m.fsr_key = f.fsr_key AND m.row_state <> 'deleted' "
                     "AND m.security_role = ?)")
        args.append(role)
    return " AND ".join(where), args


@app.route("/api/fsr")
def api_fsr():
    conn = db()
    page = max(1, int(request.args.get("page", 1)))
    size = min(500, max(10, int(request.args.get("size", 50))))
    clause, args = _fsr_filter()
    total = conn.execute(
        f"SELECT COUNT(*) FROM M3_Security_FunctionalRoles f WHERE {clause}",
        args).fetchone()[0]
    rows = conn.execute(
        f"""
        SELECT f.fsr_key, f.name, f.description, f.row_state,
               (SELECT COUNT(*) FROM M3_Security_FunctionalRoleMembers m
                 WHERE m.fsr_key = f.fsr_key AND m.row_state <> 'deleted'
                   AND m.security_role <> '')                      AS n_roles,
               (SELECT COUNT(*) FROM M3_Security_UserRoles ur
                 WHERE ur.tenant = f.tenant AND ur.role_name = f.name
                   AND ur.role_type = 'Functional'
                   AND ur.row_state <> 'deleted')                  AS n_users
        FROM M3_Security_FunctionalRoles f
        WHERE {clause}
        ORDER BY {_order_by(FSR_SORTS, 'LOWER(f.name)')} LIMIT ? OFFSET ?
        """,
        args + [size, (page - 1) * size]).fetchall()
    summary = conn.execute(
        """
        SELECT
          (SELECT COUNT(*) FROM M3_Security_FunctionalRoles
            WHERE tenant = ?1 AND row_state <> 'deleted')          AS roles,
          (SELECT COUNT(*) FROM M3_Security_FunctionalRoleMembers
            WHERE tenant = ?1 AND row_state <> 'deleted'
              AND security_role <> '')                             AS members,
          (SELECT COUNT(*) FROM M3_Security_FunctionalRoles
            WHERE tenant = ?1 AND row_state IN ('modified','new'))  AS changed,
          (SELECT COUNT(*) FROM M3_Security_Imports
            WHERE tenant = ?1 AND file_kind = 'functional')        AS files
        """,
        (request.args.get("tenant", ""),)).fetchone()
    return jsonify(status="success", total=total, page=page, size=size,
                   summary=dict(summary), rows=[dict(r) for r in rows])


@app.route("/api/fsr/keys")
def api_fsr_keys():
    conn = db()
    clause, args = _fsr_filter()
    rows = conn.execute(
        f"SELECT f.fsr_key, f.name FROM M3_Security_FunctionalRoles f "
        f"WHERE {clause} ORDER BY LOWER(f.name)",
        args).fetchall()
    return jsonify(status="success", total=len(rows),
                   rows=[dict(r) for r in rows])


@app.route("/api/fsr/roles")
def api_fsr_role_names():
    """Every security role name available to put inside a functional role."""
    conn = db()
    tenant = request.args.get("tenant", "")
    rows = conn.execute(
        """
        SELECT name FROM M3_Security_Roles
         WHERE tenant = ? AND row_state <> 'deleted'
        UNION
        SELECT DISTINCT security_role FROM M3_Security_FunctionalRoleMembers
         WHERE tenant = ? AND security_role <> ''
        UNION
        SELECT DISTINCT roll FROM M3_Security_M3RoleDefs WHERE tenant = ?
        ORDER BY 1
        """,
        (tenant, tenant, tenant)).fetchall()
    return jsonify(status="success", names=[r[0] for r in rows])


@app.route("/api/fsr/<int:fsr_key>")
def api_fsr_detail(fsr_key: int):
    conn = db()
    f = conn.execute(
        "SELECT * FROM M3_Security_FunctionalRoles WHERE fsr_key = ?",
        (fsr_key,)).fetchone()
    if not f:
        return jsonify(status="error", message="Not found."), 404
    members = [r[0] for r in conn.execute(
        "SELECT security_role FROM M3_Security_FunctionalRoleMembers "
        "WHERE fsr_key = ? AND row_state <> 'deleted' AND security_role <> '' "
        "ORDER BY COALESCE(seq, 1000000000), fsr_member_key",
        (fsr_key,))]
    users = [r[0] for r in conn.execute(
        "SELECT DISTINCT u.email_id FROM M3_Security_UserRoles ur "
        "JOIN M3_Security_Users u ON u.user_key = ur.user_key "
        "WHERE ur.tenant = ? AND ur.role_name = ? "
        "AND ur.role_type = 'Functional' AND ur.row_state <> 'deleted' "
        "ORDER BY LOWER(u.email_id)",
        (f["tenant"], f["name"]))]
    return jsonify(status="success", fsr=dict(f), roles=members, users=users)


@app.route("/api/fsr/<int:fsr_key>", methods=["POST"])
def api_fsr_save(fsr_key: int):
    """Save the description and the security roles inside a functional role."""
    conn = db()
    data = request.get_json(force=True) or {}
    f = conn.execute(
        "SELECT * FROM M3_Security_FunctionalRoles WHERE fsr_key = ?",
        (fsr_key,)).fetchone()
    if not f:
        return jsonify(status="error", message="Not found."), 404

    name = (data.get("name") or f["name"]).strip()
    desc = data.get("description")
    desc = f["description"] if desc is None else str(desc).strip()
    conn.execute(
        "UPDATE M3_Security_FunctionalRoles SET name = ?, description = ?, "
        "row_state = CASE WHEN row_state = 'new' THEN 'new' ELSE 'modified' END, "
        "modified_at = datetime('now') WHERE fsr_key = ?",
        (name, desc, fsr_key))
    if name != f["name"]:
        conn.execute(
            "UPDATE M3_Security_FunctionalRoleMembers SET fsr_name = ? "
            "WHERE fsr_key = ?",
            (name, fsr_key))

    if "roles" in data:
        # Security role names are compared case-insensitively but stored as given.
        seen, ordered = set(), []
        for r in (data["roles"] or []):
            r = str(r).strip()
            if not r:
                continue
            if r.upper() not in seen:
                seen.add(r.upper())
                ordered.append(r)
        old = {r["security_role"]: r["seq"] for r in conn.execute(
            "SELECT security_role, seq FROM M3_Security_FunctionalRoleMembers "
            "WHERE fsr_key = ?",
            (fsr_key,)).fetchall()}
        conn.execute(
            "DELETE FROM M3_Security_FunctionalRoleMembers WHERE fsr_key = ?",
            (fsr_key,))
        conn.executemany(
            "INSERT OR IGNORE INTO M3_Security_FunctionalRoleMembers "
            "(fsr_key, tenant, fsr_name, security_role, email_id, seq, row_state) "
            "VALUES (?, ?, ?, ?, '', ?, ?)",
            [(fsr_key, f["tenant"], name, r, old.get(r),
              "unchanged" if r in old else "new") for r in ordered])

    conn.commit()
    return jsonify(status="success")


@app.route("/api/fsr/new", methods=["POST"])
def api_fsr_new():
    conn = db()
    data = request.get_json(force=True) or {}
    tenant = (data.get("tenant") or "").strip()
    name = (data.get("name") or "").strip()
    if not tenant or not name:
        return jsonify(status="error",
                       message="Tenant and name are required."), 400
    try:
        cur = conn.execute(
            "INSERT INTO M3_Security_FunctionalRoles "
            "(tenant, name, description, row_state, seq) "
            "VALUES (?, ?, ?, 'new', (SELECT COALESCE(MAX(seq), 0) + 1 "
            "FROM M3_Security_FunctionalRoles  WHERE tenant = ?))",
            (tenant, name, (data.get("description") or name).strip(), tenant))
    except sqlite3.IntegrityError:
        return jsonify(status="error",
                       message=f"'{name}' already exists."), 400
    conn.commit()
    return jsonify(status="success", fsr_key=cur.lastrowid)


@app.route("/api/fsr/<int:fsr_key>", methods=["DELETE"])
def api_fsr_delete(fsr_key: int):
    conn = db()
    if request.args.get("hard") == "1":
        conn.execute(
            "DELETE FROM M3_Security_FunctionalRoles WHERE fsr_key = ?",
            (fsr_key,))
    else:
        conn.execute(
            "UPDATE M3_Security_FunctionalRoles SET row_state = 'deleted', "
            "modified_at = datetime('now') WHERE fsr_key = ?",
            (fsr_key,))
    conn.commit()
    return jsonify(status="success")


# -------------------------------------------------------------- clearing


# Per tab: what to say, which tables to empty (in dependency order) and which
# table's row_state decides whether unexported changes would be lost.
CLEAR_KINDS: dict[str, dict] = {
    "users": {
        "label": "user",
        "parts": [("user(s)", "M3_Security_Users", ""),
                  ("role assignment(s)", "M3_Security_UserRoles", "")],
        "changed": ("M3_Security_Users",
                    "row_state IN ('modified','new','deleted')"),
    },
    "roles": {
        "label": "security role",
        "parts": [("security role(s)", "M3_Security_Roles", ""),
                  ("member assignment(s)", "M3_Security_RoleAssignments", "")],
        "changed": ("M3_Security_Roles",
                    "row_state IN ('modified','new','deleted')"),
    },
    "functional": {
        "label": "functional security role",
        "parts": [("functional role(s)", "M3_Security_FunctionalRoles", ""),
                  ("security role assignment(s)",
                   "M3_Security_FunctionalRoleMembers",
                   "security_role <> ''")],
        "changed": ("M3_Security_FunctionalRoles",
                    "row_state IN ('modified','new')"),
    },
}


@app.route("/api/clear/<kind>", methods=["POST"])
def api_clear(kind: str):
    """
    Empty one IFS capture for one tenant. Preview unless confirm == 'CLEAR'.

    Deletes the tab's own tables plus its rows in M3_Security_Imports, so the
    next drop starts from scratch rather than merging into, or being compared
    against, what is already held.
    """
    spec = CLEAR_KINDS.get(kind)
    if not spec:
        return jsonify(status="error",
                       message="kind must be users, roles or functional."), 400
    conn = db()
    data = request.get_json(force=True) or {}
    tenant = (data.get("tenant") or "").strip()
    confirm = data.get("confirm")
    if not tenant:
        return jsonify(status="error", message="Tenant is required."), 400
    if confirm and confirm != "CLEAR":
        return jsonify(status="error", message="Type CLEAR to confirm."), 400

    def count(table: str, extra: str = "") -> int:
        sql = f"SELECT COUNT(*) FROM {table} WHERE tenant = ?"
        if extra:
            sql += f" AND {extra}"
        return conn.execute(sql, (tenant,)).fetchone()[0]

    plan = [{"label": label, "count": count(table, extra)}
            for label, table, extra in spec["parts"]]
    totals = {
        "rows": plan[0]["count"],
        "files": conn.execute(
            "SELECT COUNT(*) FROM M3_Security_Imports "
            "WHERE tenant = ? AND file_kind = ?",
            (tenant, kind)).fetchone()[0],
        "changed": count(*spec["changed"]),
    }
    summary = ", ".join(f"{p['count']:,} {p['label']}" for p in plan)

    if not confirm:
        return jsonify(
            status="success", dry_run=True, kind=kind, plan=plan, totals=totals,
            message=f"{summary} from {totals['files']} file(s) would be cleared for '"
                    f"{tenant}'."
                    if totals["rows"]
                    else f"No {spec['label']} data held for '{tenant}'.")

    cur = conn.cursor()
    # Children first, so a foreign key never blocks the parent delete.
    for _, table, _extra in reversed(spec["parts"]):
        cur.execute(f"DELETE FROM {table} WHERE tenant = ?", (tenant,))
    cur.execute("DELETE FROM M3_Security_Imports WHERE tenant = ? AND file_kind = ?",
                (tenant, kind))
    conn.commit()
    return jsonify(
        status="success", dry_run=False, kind=kind, plan=plan,
        totals=totals,
        message=f"Cleared {summary} for '{tenant}"
                f"'. Drop the export file(s) again to start over.")


@app.route("/api/fsr/clear", methods=["POST"])
def api_fsr_clear():
    """Kept so an older page still works; /api/clear/functional is the route."""
    return api_clear("functional")


@app.route("/api/fsr/remove-roles", methods=["POST"])
def api_fsr_remove_roles():
    """
    Take security roles out of a set of functional roles.

    roles empty means every security role they contain. Preview unless
    confirm == 'REMOVE'.
    """
    conn = db()
    data = request.get_json(force=True) or {}
    tenant = (data.get("tenant") or "").strip()
    keys = [int(k) for k in (data.get("keys") or [])]
    names = [str(r) for r in (data.get("roles") or [])]
    confirm = data.get("confirm")
    if not tenant:
        return jsonify(status="error", message="Tenant is required."), 400
    if not keys:
        return jsonify(status="error",
                       message="No functional roles selected."), 400
    if confirm and confirm != "REMOVE":
        return jsonify(status="error", message="Type REMOVE to confirm."), 400

    marks = ",".join("?" * len(keys))
    args = [tenant] + keys
    name_clause = ""
    if names:
        name_clause = " AND m.security_role IN (%s)" % ",".join("?" * len(names))
        args += names

    breakdown = [dict(r) for r in conn.execute(
        f"""
        SELECT m.security_role AS role, COUNT(DISTINCT m.fsr_key) AS roles
        FROM M3_Security_FunctionalRoleMembers m
        WHERE m.tenant = ? AND m.row_state <> 'deleted' AND m.security_role <> ''
          AND m.fsr_key IN ({marks}){name_clause}
        GROUP BY m.security_role ORDER BY LOWER(m.security_role)
        """,
        args).fetchall()]
    totals = {"selected": len(keys), "roles": len(breakdown),
              "links": sum(b["roles"] for b in breakdown)}

    if not confirm:
        return jsonify(
            status="success", dry_run=True, plan=breakdown,
            totals=totals,
            message=f"{totals['links']} security role assignment(s) would be removed from "
                    f"{totals['selected']} functional role(s), across "
                    f"{totals['roles']} security role(s).")

    cur = conn.cursor()
    cur.execute(
        f"""
        UPDATE M3_Security_FunctionalRoleMembers SET row_state = 'deleted'
         WHERE tenant = ? AND fsr_key IN ({marks}) AND row_state <> 'deleted'
           AND security_role <> ''
           {name_clause.replace('m.security_role', 'security_role')}
        """,
        args)
    removed = cur.rowcount
    cur.execute(
        f"""
        UPDATE M3_Security_FunctionalRoles SET row_state =
               CASE WHEN row_state = 'new' THEN 'new' ELSE 'modified' END,
               modified_at = datetime('now')
         WHERE tenant = ? AND fsr_key IN ({marks})
        """,
        [tenant] + keys)
    conn.commit()
    return jsonify(
        status="success", dry_run=False, plan=breakdown, totals=totals,
        removed=removed,
        message=f"{removed} security role assignment(s) removed from "
                f"{len(keys)} functional role(s). Export and import to apply it in IFS.")


@app.route("/api/m3/log")
def api_m3_log():
    conn = db()
    tenant = request.args.get("tenant", "")
    role = request.args.get("role")
    sql = "SELECT * FROM M3_Security_M3Log WHERE tenant = ?"
    args = [tenant]
    if role:
        sql += " AND role_name = ?"
        args.append(role)
    sql += " ORDER BY m3_log_key DESC LIMIT 200"
    return jsonify(status="success",
                   rows=[dict(r) for r in conn.execute(sql, args).fetchall()])


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="M3 security capture front end.")
    ap.add_argument("--host", default="127.0.0.1")
    ap.add_argument("--port", type=int, default=5057)
    ap.add_argument("--db", default=None, help=f"SQLite path (default {DEFAULT_DB_PATH})")
    ap.add_argument("--ionapi-dir", default=str(DEFAULT_IONAPI_DIR),
                    help="Folder holding the .ionapi files")
    ap.add_argument("--company", default=DEFAULT_COMPANY,
                    help="CONO for every tenant. Left off, each tenant's entry "
                         "in M3_Security_M3.json is used, and failing that the "
                         "request omits it so M3 uses the service account's own "
                         "default company")
    ap.add_argument("--division", default=DEFAULT_DIVISION,
                    help="DIVI for every tenant, same rules as --company")
    ap.add_argument("--m3user", default=None,
                    help="Optional m3user to run the API calls as")
    ap.add_argument("--batch-size", type=int, default=DEFAULT_BATCH_SIZE,
                    help=f"Records per M3 API call (default {DEFAULT_BATCH_SIZE})")
    ap.add_argument("--debug", action="store_true")
    args = ap.parse_args(argv)

    logging.basicConfig(
        level=logging.DEBUG if args.debug else logging.INFO,
        format="%(asctime)s  %(levelname)-8s  %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
        handlers=[logging.StreamHandler(sys.stdout)])

    app.config["M3_SECURITY_DB"] = args.db
    app.config["M3_IONAPI_DIR"] = Path(args.ionapi_dir)
    app.config["M3_COMPANY"] = args.company
    app.config["M3_DIVISION"] = args.division
    app.config["M3_USER"] = args.m3user
    app.config["M3_BATCH_SIZE"] = args.batch_size
    UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    # Fail early if the database cannot be opened or migrated.
    conn = connect(args.db)
    conn.close()
    log.info("Database : %s", resolve_db_path(args.db))
    log.info("Uploads  : %s", UPLOAD_DIR)
    log.info("Exports  : %s", OUTPUT_DIR)
    log.info("BODs     : %s", BOD_DIR)
    log.info("ION API  : %s (cono %s, divi %s, %s per call)",
             args.ionapi_dir, args.company, args.division or "(blank)",
             args.batch_size)
    log.info("Open     : http://%s:%s", args.host, args.port)

    app.run(host=args.host, port=args.port, debug=args.debug, threaded=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
