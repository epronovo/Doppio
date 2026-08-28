"""
M3_Security_Export - write the capture back out in the inbound IFS format.

Three files come out of here, each matching the shape of the export it came
from so it can be sent straight back in:

  * the users export (ExportFile_...), whose column set is rebuilt from the
    file that was loaded, including how many SecurityRoleN / FunctionalSecurityRoleN
    slots it carried;
  * the Security Role export;
  * the Functional Security Role export.

Two scopes: 'changes' writes only what was edited (each changed record with its
COMPLETE role list or membership, so a delta is never read as a revocation),
'full' writes everything held for the tenant.

Records flagged for removal are never written, in either scope.

Reconstructed from M3_Security_Export.cpython-312.pyc after the source was lost.
Logic matches the bytecode; comments and formatting are not original.
"""
from __future__ import annotations

import argparse
import csv
import json
import logging
import sqlite3
import sys
from datetime import datetime, timezone
from pathlib import Path

from M3_Security_Db import (
    DEFAULT_DB_PATH,
    ROLE_BLOCKS,
    USER_BASE_COLUMNS,
    USER_TAIL_COLUMNS,
    connect,
    latest_import,
    resolve_db_path,
)

log = logging.getLogger("M3_Security_Export")

DEFAULT_OUT_DIR = Path(__file__).parent / "output"

# .NET counts ticks from 0001-01-01.
_NET_EPOCH = datetime(1, 1, 1, tzinfo=timezone.utc)

SCOPES = ("changes", "full")

# A user is "changed" when the record was edited or added, or when any role
# they hold was touched - including a role that was removed, which only shows
# up as a deleted UserRoles row.
_USERS_CHANGED_WHERE = """
    u.row_state <> 'deleted'
AND (
        u.row_state IN ('modified', 'new')
     OR EXISTS (SELECT 1 FROM M3_Security_UserRoles ur
                 WHERE ur.user_key = u.user_key
                   AND ur.row_state IN ('modified', 'new', 'deleted'))
    )
"""

# Same idea for a role: its own fields, or its membership.
_ROLES_CHANGED_WHERE = """
    r.row_state <> 'deleted'
AND (
        r.row_state IN ('modified', 'new')
     OR EXISTS (SELECT 1 FROM M3_Security_RoleAssignments a2
                 WHERE a2.role_key = r.role_key
                   AND a2.row_state IN ('modified', 'new', 'deleted'))
    )
"""


def _check_scope(scope: str) -> str:
    if scope not in SCOPES:
        raise ValueError(f"scope must be one of {SCOPES}, got '{scope}'.")
    return scope


def changed_counts(conn: sqlite3.Connection, tenant: str) -> dict:
    """How many users / roles a changes-only export would carry."""
    users = conn.execute(
        f"SELECT COUNT(*) FROM M3_Security_Users u "
        f"WHERE u.tenant = ? AND {_USERS_CHANGED_WHERE}",
        (tenant,),
    ).fetchone()[0]
    roles = conn.execute(
        f"SELECT COUNT(*) FROM M3_Security_Roles r "
        f"WHERE r.tenant = ? AND {_ROLES_CHANGED_WHERE}",
        (tenant,),
    ).fetchone()[0]
    functional = conn.execute(
        """
        SELECT COUNT(*) FROM M3_Security_FunctionalRoles f
         WHERE f.tenant = ? AND f.row_state <> 'deleted'
           AND (f.row_state IN ('modified', 'new')
                OR EXISTS (SELECT 1 FROM M3_Security_FunctionalRoleMembers m2
                            WHERE m2.fsr_key = f.fsr_key
                              AND m2.row_state IN ('modified', 'new', 'deleted')))
        """,
        (tenant,),
    ).fetchone()[0]
    return {"users": users, "roles": roles, "functional": functional}


def mark_pushed(conn: sqlite3.Connection, tenant: str, kind: str = "both") -> dict:
    """
    Reset the change flags after a delta has been pushed into M3, so the next
    export starts from a clean baseline.  Rows flagged for removal are dropped
    at this point - they were never exported and are not coming back.
    """
    done = {"users": 0, "roles": 0, "assignments": 0, "purged": 0}
    cur = conn.cursor()

    if kind in ("users", "both"):
        cur.execute(
            "DELETE FROM M3_Security_Users WHERE tenant = ? AND row_state = 'deleted'",
            (tenant,),
        )
        done["purged"] += cur.rowcount
        cur.execute(
            "UPDATE M3_Security_UserRoles SET row_state = 'unchanged' "
            "WHERE tenant = ? AND row_state <> 'unchanged'",
            (tenant,),
        )
        cur.execute(
            "UPDATE M3_Security_Users SET row_state = 'unchanged' "
            "WHERE tenant = ? AND row_state <> 'unchanged'",
            (tenant,),
        )
        done["users"] = cur.rowcount

    if kind in ("roles", "both"):
        cur.execute(
            "DELETE FROM M3_Security_Roles WHERE tenant = ? AND row_state = 'deleted'",
            (tenant,),
        )
        done["purged"] += cur.rowcount
        cur.execute(
            "UPDATE M3_Security_RoleAssignments SET row_state = 'unchanged' "
            "WHERE tenant = ? AND row_state <> 'unchanged'",
            (tenant,),
        )
        done["assignments"] = cur.rowcount
        cur.execute(
            "UPDATE M3_Security_Roles SET row_state = 'unchanged' "
            "WHERE tenant = ? AND row_state <> 'unchanged'",
            (tenant,),
        )
        done["roles"] = cur.rowcount

    conn.commit()
    return done


def dotnet_ticks(when: datetime | None = None) -> str:
    """.NET DateTime.Ticks - 100ns intervals since 0001-01-01."""
    when = when or datetime.now(timezone.utc)
    if when.tzinfo is None:
        when = when.replace(tzinfo=timezone.utc)
    delta = when - _NET_EPOCH
    return str(delta.days * 864000000000 + delta.seconds * 10000000 + delta.microseconds * 10)


def users_file_name(tenant: str, guid: str | None = None, ticks: str | None = None) -> str:
    """
    Rebuild the users export name, keeping the guid exactly as it arrived.

    IFS writes the guid dashed (0ea904a0-a76b-...); some copies have the dashes
    stripped. Whichever form came in is what goes back out, so the file looks
    like the one it replaces. Only when there is no guid at all do we invent
    one, in the dashed form IFS uses.
    """
    guid = (guid or "").strip() or "-".join(
        ["00000000", "0000", "0000", "0000", "000000000000"])
    ticks = ticks or dotnet_ticks()
    return f"ExportFile_{tenant}_{guid}_{ticks}.csv"


def functional_file_name(tenant: str, when: datetime | None = None) -> str:
    when = when or datetime.now()
    stamp = when.strftime("%m_%d_%Y_%H_%M_%S.") + f"{when.microsecond:06d}0"
    return f"Functional Security Role_{tenant}_{stamp}.csv"


def roles_file_name(tenant: str, when: datetime | None = None) -> str:
    when = when or datetime.now()
    stamp = when.strftime("%m_%d_%Y_%H_%M_%S.") + f"{when.microsecond:06d}0"
    return f"Security Role_{tenant}_{stamp}.csv"


def reusable_source_name(src, tenant: str) -> str | None:
    """
    The imported file's name, but only if it identifies the tenant.

    A Security Role or Functional Security Role export saved without the
    tenant in its name can still be loaded - the tenant selected in the front
    end supplies it - but reusing that name on the way out would produce a file
    that needs the same manual step to get back in. In that case we generate a
    proper name instead, so an export is always self-describing.
    """
    name = src["file_name"] if src and src["file_name"] else None
    return name if name and tenant in name else None


def export_users(
    conn: sqlite3.Connection,
    tenant: str,
    out_dir: str | Path = DEFAULT_OUT_DIR,
    file_name: str | None = None,
    reuse_source_name: bool = True,
    scope: str = "changes",
) -> Path:
    """
    Rebuild the IFS user export for one tenant. Returns the file path.

    scope='changes' (default) writes only added/edited users;
    scope='full' writes every user held for the tenant.
    """
    _check_scope(scope)

    where = _USERS_CHANGED_WHERE if scope == "changes" else "u.row_state <> 'deleted'"
    users = conn.execute(
        f"""
        SELECT u.* FROM M3_Security_Users u
        WHERE u.tenant = ? AND {where}
        ORDER BY COALESCE(u.source_row, 1000000000), u.user_key
        """,
        (tenant,),
    ).fetchall()
    if not users:
        raise ValueError(
            f"No changed users to export for tenant '{tenant}' - edit or add a "
            f"user first, or run a full export."
            if scope == "changes"
            else f"No users held for tenant '{tenant}'."
        )

    # The roles each user holds, in export order.
    role_rows = conn.execute(
        """
        SELECT user_key, role_type, role_name
        FROM M3_Security_UserRoles
        WHERE tenant = ? AND row_state <> 'deleted'
        ORDER BY user_key, role_type, seq, user_role_key
        """,
        (tenant,),
    ).fetchall()

    by_user = {}
    for r in role_rows:
        by_user.setdefault(r["user_key"], {}).setdefault(r["role_type"], []).append(
            r["role_name"])

    # How many role columns to write: at least as many as the source file had,
    # and never fewer than the roles someone actually holds.
    src = latest_import(conn, "users", tenant)
    widths = {}
    for prefix, role_type in ROLE_BLOCKS:
        needed = max(
            (len(v.get(role_type, [])) for v in by_user.values()),
            default=0,
        )
        base = 0
        if src:
            base = (src["security_slots"] if prefix == "SecurityRole"
                    else src["functional_slots"]) or 0
        widths[prefix] = max(base, needed)

    # Columns the source file carried. A file that was missing some of them is
    # written back the same way rather than gaining columns it never had.
    base_cols = list(USER_BASE_COLUMNS)
    tail_cols = list(USER_TAIL_COLUMNS)
    if src is not None and "header_json" in src.keys() and src["header_json"]:
        try:
            want = json.loads(src["header_json"])
            keep_base = set(want.get("base") or [])
            keep_tail = set(want.get("tail") or [])
            if keep_base:
                base_cols = [c for c in USER_BASE_COLUMNS if c[0] in keep_base]
                tail_cols = [c for c in USER_TAIL_COLUMNS if c[0] in keep_tail]
                if want.get("missing"):
                    log.info(
                        "Users export: %s column(s) the source file did not "
                        "carry are left out again: %s",
                        len(want["missing"]), ", ".join(want["missing"]))
        except (json.JSONDecodeError, TypeError):
            log.warning(
                "import %s has an unreadable header_json - writing the full "
                "column set", src["import_id"])

    header = [h for h, _ in base_cols]
    for prefix, _ in ROLE_BLOCKS:
        header += [f"{prefix}{i}" for i in range(1, widths[prefix] + 1)]
    header += [h for h, _ in tail_cols]

    if file_name is None:
        file_name = reusable_source_name(src, tenant) if reuse_source_name else None
        if file_name is None:
            file_name = users_file_name(
                tenant,
                src["file_guid"] if src else None,
                None,
            )

    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = out_dir / file_name

    with open(out_path, "w", encoding="utf-8", newline="") as fh:
        w = csv.writer(fh, lineterminator="\r\n", quoting=csv.QUOTE_MINIMAL)
        w.writerow(header)
        for u in users:
            row = [u[col] or "" for _, col in base_cols]
            held = by_user.get(u["user_key"], {})
            for prefix, role_type in ROLE_BLOCKS:
                names = held.get(role_type, [])
                row += names + [""] * (widths[prefix] - len(names))
            row += [u[col] or "" for _, col in tail_cols]
            w.writerow(row)

    log.info("Users export (%s): %s (%s rows, %s columns)",
             scope, out_path, len(users), len(header))
    return out_path


def export_roles(
    conn: sqlite3.Connection,
    tenant: str,
    out_dir: str | Path = DEFAULT_OUT_DIR,
    file_name: str | None = None,
    reuse_source_name: bool = True,
    scope: str = "changes",
) -> Path:
    """
    Rebuild the Security Role export for one tenant. Returns the path.

    scope='changes' (default) writes only roles whose name, description or
    membership changed - each one with its complete current member list, so a
    delta never looks like a revocation of the untouched members.
    scope='full' writes every role held for the tenant.
    """
    _check_scope(scope)

    where = _ROLES_CHANGED_WHERE if scope == "changes" else "r.row_state <> 'deleted'"
    rows = conn.execute(
        f"""
        SELECT r.name, r.description, a.email_id
        FROM M3_Security_Roles r
        JOIN M3_Security_RoleAssignments a ON a.role_key = r.role_key
        WHERE r.tenant = ?
          AND a.row_state <> 'deleted'
          AND {where}
        ORDER BY COALESCE(r.seq, 1000000000), r.role_key,
                 COALESCE(a.seq, 1000000000), a.assignment_key
        """,
        (tenant,),
    ).fetchall()
    if not rows:
        raise ValueError(
            f"No changed security roles to export for tenant '{tenant}' - edit "
            f"a role or its members first, or run a full export."
            if scope == "changes"
            else f"No security roles held for tenant '{tenant}'."
        )

    src = latest_import(conn, "roles", tenant)
    if file_name is None:
        file_name = (reusable_source_name(src, tenant) if reuse_source_name
                     else None) or roles_file_name(tenant)

    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = out_dir / file_name

    with open(out_path, "w", encoding="utf-8-sig", newline="") as fh:
        w = csv.writer(fh, lineterminator="\r\n", quoting=csv.QUOTE_MINIMAL)
        w.writerow(["Name", "Description", "EmailId"])
        for r in rows:
            w.writerow([r["name"], r["description"] or "", r["email_id"] or "", ""])

    log.info("Roles export (%s): %s (%s rows)", scope, out_path, len(rows))
    return out_path


def export_functional(
    conn: sqlite3.Connection,
    tenant: str,
    out_dir: str | Path = DEFAULT_OUT_DIR,
    file_name: str | None = None,
    reuse_source_name: bool = True,
    scope: str = "changes",
) -> Path:
    """
    Write the functional security roles back out in the inbound format.

    Two details of that format matter:

      * UTF-8 BOM, CRLF, and a trailing comma on rows that carry a member,
        because EmailId is the fourth column and IFS leaves it empty.
      * a functional role with NO members is written as two fields with no
        trailing commas at all - "FSR_OVEX,OVEX" - which is exactly how IFS
        exports it. Writing four fields there would not match.

    scope='changes' writes only the functional roles that were edited.
    """
    _check_scope(scope)

    where = ["f.tenant = ?", "f.row_state <> 'deleted'"]
    if scope == "changes":
        where.append(
            "(f.row_state IN ('modified', 'new') OR EXISTS (SELECT 1 FROM "
            "M3_Security_FunctionalRoleMembers m2             WHERE m2.fsr_key "
            "= f.fsr_key               AND m2.row_state IN ('modified', 'new', "
            "'deleted')))")
    clause = " AND ".join(where)

    roles = conn.execute(
        f"""
        SELECT f.fsr_key, f.name, f.description
        FROM M3_Security_FunctionalRoles f
        WHERE {clause}
        ORDER BY COALESCE(f.seq, 1000000000), f.fsr_key
        """,
        (tenant,),
    ).fetchall()
    if not roles:
        raise ValueError(
            f"No changed functional security roles to export for tenant "
            f"'{tenant}' - edit one first, or run a full export."
            if scope == "changes"
            else f"No functional security roles held for tenant '{tenant}'."
        )

    members = {}
    for m in conn.execute(
        """
        SELECT fsr_key, security_role, email_id
        FROM M3_Security_FunctionalRoleMembers
        WHERE tenant = ? AND row_state <> 'deleted'
          AND (security_role <> '' OR email_id <> '')
        ORDER BY COALESCE(seq, 1000000000), fsr_member_key
        """,
        (tenant,),
    ).fetchall():
        members.setdefault(m["fsr_key"], []).append(m)

    src = latest_import(conn, "functional", tenant)
    if file_name is None:
        file_name = (reusable_source_name(src, tenant) if reuse_source_name
                     else None) or functional_file_name(tenant)

    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = out_dir / file_name

    n_rows = 0
    with open(out_path, "w", encoding="utf-8-sig", newline="") as fh:
        w = csv.writer(fh, lineterminator="\r\n", quoting=csv.QUOTE_MINIMAL)
        w.writerow(["Functional SecurityRole Name", "Description",
                    "SecurityRole", "EmailId"])
        for r in roles:
            rows = members.get(r["fsr_key"], [])
            if not rows:
                w.writerow([r["name"], r["description"] or ""])
                n_rows += 1
                continue
            for m in rows:
                w.writerow([r["name"], r["description"] or "",
                            m["security_role"], m["email_id"]])
                n_rows += 1

    log.info("Functional security role export (%s): %s (%s role(s), %s rows)",
             scope, out_path, len(roles), n_rows)
    return out_path


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(
        description="Export M3 security data from doppio.db back to CSV.")
    ap.add_argument("--tenant", required=True)
    ap.add_argument("--kind",
                    choices=["users", "roles", "functional", "both", "all"],
                    default="both")
    ap.add_argument("--scope", choices=list(SCOPES), default="changes",
                    help="'changes' (default) exports only modified/added "
                         "records; 'full' exports everything")
    ap.add_argument("--out-dir", default=str(DEFAULT_OUT_DIR))
    ap.add_argument("--db", default=None,
                    help=f"SQLite path (default {DEFAULT_DB_PATH})")
    ap.add_argument("--new-name", action="store_true",
                    help="Generate a fresh time-stamped file name instead of "
                         "reusing the imported file name")
    ap.add_argument("--mark-pushed", action="store_true",
                    help="After a successful export, clear the change flags so "
                         "the next delta starts from a clean baseline")
    args = ap.parse_args(argv)

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s  %(levelname)-8s  %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
        handlers=[logging.StreamHandler(sys.stdout)],
    )

    conn = connect(args.db)
    log.info("Database: %s", resolve_db_path(args.db))
    out_dir = Path(args.out_dir) / args.scope
    rc = 0
    written = 0
    try:
        if args.scope == "changes":
            c = changed_counts(conn, args.tenant)
            log.info("Pending changes: %s user(s), %s role(s)",
                     c["users"], c["roles"])

        for kind, fn in (("users", export_users), ("roles", export_roles),
                         ("functional", export_functional)):
            wanted = (args.kind == kind or args.kind == "all"
                      or (args.kind == "both" and kind in ("users", "roles")))
            if not wanted:
                continue
            try:
                fn(conn, args.tenant, out_dir,
                   reuse_source_name=not args.new_name, scope=args.scope)
                written += 1
            except Exception as exc:
                log.error("%s: %s", kind, exc)
                rc = 1

        if args.mark_pushed:
            if written and rc == 0:
                done = mark_pushed(conn, args.tenant, args.kind)
                log.info("Baseline reset: %s user(s), %s role(s), %s "
                         "assignment(s) cleared, %s removal(s) purged",
                         done["users"], done["roles"], done["assignments"],
                         done["purged"])
            else:
                log.warning("Baseline NOT reset - the export did not complete "
                            "cleanly.")
    finally:
        conn.close()
    return rc


if __name__ == "__main__":
    raise SystemExit(main())
