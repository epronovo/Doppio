"""
M3_Security_Import - load the Infor OS / M3 security exports into doppio.db.

Three file kinds, decided from the header row rather than the file name:

  * the users export (ExportFile_...), which MERGES on PersonId - a re-drop
    refreshes the fields but keeps the roles held here, because role work done
    locally has not reached IFS yet and a file load must not undo it;
  * the Security Role export, which is a full refresh per tenant;
  * the Functional Security Role export, which MERGES, because IFS caps that
    export at 75 rows and a real capture arrives as several files.

The tenant comes from the file name where it is there, and from the tenant
selected in the front end where it is not.

Reconstructed from M3_Security_Import.cpython-312.pyc after the source was lost.
Logic matches the bytecode; comments and formatting are not original.
"""
from __future__ import annotations

import argparse
import csv
import json
import logging
import re
import sqlite3
import sys
from pathlib import Path

from M3_Security_Db import (
    DEFAULT_DB_PATH,
    ROLE_BLOCKS,
    USER_BASE_COLUMNS,
    USER_TAIL_COLUMNS,
    connect,
    resolve_db_path,
)

# SecurityAccessProfiles can be enormous on a real export.
csv.field_size_limit(1000000000)

log = logging.getLogger("M3_Security_Import")

# IFS writes the guid dashed; some copies have the dashes stripped. Both are
# accepted, and whichever came in is what goes back out.
GUID = (r"(?:[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}"
        r"-[0-9a-fA-F]{12}|[0-9a-fA-F]{32})")

RE_USERS_NAME = re.compile(
    r"^ExportFile_(?P<tenant>.+?)_(?P<guid>" + GUID + r")_(?P<ticks>\d+)$")

# The label IFS puts in front of a role export. Dropping it leaves the tenant.
RE_EXPORT_LABEL = re.compile(
    r"^(?:functional[_\s-]*security[_\s-]*role|security[_\s-]*role"
    r"|export[_\s-]*file)[_\s-]+",
    re.IGNORECASE,
)

# A tenant is written as <id>_<ENV>. The id has letters and digits; the
# environment is a short upper-case token (TST, TRN, PRD, DEM).
RE_TENANT_TOKEN = re.compile(r"^[A-Za-z0-9]{2,32}$")
RE_TENANT_ID = re.compile(r"^(?=.*\d)(?=.*[A-Za-z])[A-Za-z0-9]{6,24}$")
RE_TENANT_ENV = re.compile(r"^[A-Z]{2,4}$")

# 08_17_2026_15_07_05.6295484
RE_ROLE_STAMP = re.compile(r"(\d{2}_\d{2}_\d{4}_\d{2}_\d{2}_\d{2}\.\d+)")

USERS_HEADER_KEY = "personid"

USER_KEY_COLUMN = "PersonId"
ROLES_HEADER = ["name", "description", "emailid"]
FSR_HEADER = ["functional securityrole name", "description", "securityrole"]


def parse_file_name(file_name: str) -> dict:
    """
    Pull whatever identifying pieces are folded into the file name.

    Returns keys: tenant, guid, ticks, stamp  (any may be None).
    """
    stem = Path(file_name).stem
    out = {"tenant": None, "guid": None, "ticks": None, "stamp": None}

    # The users export carries all three.
    m = RE_USERS_NAME.match(stem)
    if m:
        out.update(tenant=m.group("tenant"),
                   guid=m.group("guid"),
                   ticks=m.group("ticks"))
        return out

    m = RE_ROLE_STAMP.search(stem)
    if m:
        out["stamp"] = m.group(1)
        stem_wo_stamp = stem.replace(m.group(1), " ")
    else:
        stem_wo_stamp = stem

    # Drop the label IFS writes in front; whatever is left should be the
    # tenant, however it is spelled.
    label = RE_EXPORT_LABEL.match(stem_wo_stamp)
    if label:
        rest = [t for t in re.split(r"[_\s]+", stem_wo_stamp[label.end():]) if t]
        plausible = (1 <= len(rest) <= 2
                     and all(RE_TENANT_TOKEN.match(t) for t in rest))
        if plausible:
            if (out["stamp"] or len(rest) == 2) and RE_TENANT_ENV.match(rest[1]):
                out["tenant"] = "_".join(rest)
                return out

    # No usable label: look for an <id>_<ENV> pair anywhere in the name.
    tokens = [t for t in re.split(r"[_\s]+", stem_wo_stamp) if t]
    for i in range(len(tokens) - 1):
        if RE_TENANT_ID.match(tokens[i]) and RE_TENANT_ENV.match(tokens[i + 1]):
            out["tenant"] = f"{tokens[i]}_{tokens[i + 1]}"
            return out

    for tok in tokens:
        if len(tok) >= 10 and RE_TENANT_ID.match(tok):
            out["tenant"] = tok
            return out

    return out


def detect_kind(header: list[str]) -> str:
    """'users' or 'roles', decided from the header row."""
    norm = [h.strip().lower().lstrip("﻿") for h in header]
    if norm and norm[0] == USERS_HEADER_KEY:
        return "users"
    if norm[:3] == FSR_HEADER:
        return "functional"
    if norm[:3] == ROLES_HEADER:
        return "roles"
    raise ValueError(
        "Unrecognised CSV header. Expected an IFS user export (starts with "
        "PersonId) or a Security Role export (Name,Description,EmailId). "
        f"Got: {header[:5]}")


def read_header(path: Path) -> list[str]:
    with open(path, "r", encoding="utf-8-sig", newline="") as fh:
        return next(csv.reader(fh))


def _role_slot_map(header: list[str]) -> dict[str, list[tuple[int, int]]]:
    """
    prefix -> [(column index, slot number), ...] sorted by slot number.
    """
    slots = {p: [] for p, _ in ROLE_BLOCKS}
    # Longest prefix first, so FunctionalSecurityRole1 is not read as
    # SecurityRole-anything.
    prefixes = sorted((p for p, _ in ROLE_BLOCKS), key=len, reverse=True)
    for idx, col in enumerate(header):
        col = col.strip()
        for prefix in prefixes:
            if col.startswith(prefix) and col[len(prefix):].isdigit():
                slots[prefix].append((idx, int(col[len(prefix):])))
                break
    for prefix in slots:
        slots[prefix].sort(key=lambda x: x[1])
    return slots


def resolve_tenant(file_name: str, meta: dict, tenant: str | None = None,
                   tenant_fallback: str | None = None,
                   hint: str = "") -> tuple[str, str]:
    """
    Decide which tenant a file belongs to, and say where that came from.

    Precedence: an explicit ``tenant`` (``--tenant``, or the front end asking
    for an override) wins, then the tenant folded into the file name, then
    ``tenant_fallback`` - the tenant selected in the front end, which is what
    rescues a Security Role or Functional Security Role export that was saved
    without the tenant in its name.

    Returns (tenant, source) where source is 'override', 'file name' or
    'selected'.
    """
    tenant = (tenant or "").strip()
    if tenant:
        return (tenant, "override")
    if meta.get("tenant"):
        return (meta["tenant"], "file name")
    fallback = (tenant_fallback or "").strip()
    if fallback:
        return (fallback, "selected")
    raise ValueError(
        f"Could not determine the tenant from '{file_name}'. "
        + (hint + " " if hint else "")
        + "Select a tenant in the front end first, or pass --tenant.")


def import_users(path: Path, conn: sqlite3.Connection, tenant: str | None = None,
                 tenant_fallback: str | None = None) -> dict:
    meta = parse_file_name(path.name)
    tenant, tenant_source = resolve_tenant(
        path.name, meta, tenant, tenant_fallback,
        "Expected ExportFile_<TENANT>_<guid>_<ticks>.csv.")

    with open(path, "r", encoding="utf-8-sig", newline="") as fh:
        reader = csv.reader(fh)
        header = next(reader)
        header = [h.strip() for h in header]
        col_ix = {h: i for i, h in enumerate(header)}

        # Real exports differ in which columns they carry. A column this file
        # does not have is stored as null rather than treated as an error, and
        # is remembered so the export can leave it out again.
        missing = [h for h, _ in USER_BASE_COLUMNS + USER_TAIL_COLUMNS
                   if h not in col_ix]
        if USER_KEY_COLUMN not in col_ix:
            raise ValueError(
                f"Users export has no '{USER_KEY_COLUMN}' column - header "
                f"starts {header[:4]}")
        if missing:
            log.warning("%s: %s column(s) not in this export, stored as null: %s",
                        path.name, len(missing), ", ".join(missing))

        slots = _role_slot_map(header)
        prefix_to_type = dict(ROLE_BLOCKS)
        base_ix = {h: col_ix.get(h) for h, _ in USER_BASE_COLUMNS}
        tail_ix = {h: col_ix.get(h) for h, _ in USER_TAIL_COLUMNS}

        # Stored on the import row so the export can rebuild this shape.
        header_json = json.dumps({
            "base": [h for h, _ in USER_BASE_COLUMNS if h in col_ix],
            "tail": [h for h, _ in USER_TAIL_COLUMNS if h in col_ix],
            "missing": missing,
        })

        db_cols = ([c for _, c in USER_BASE_COLUMNS]
                   + [c for _, c in USER_TAIL_COLUMNS])
        insert_sql = (
            f"INSERT INTO M3_Security_Users (tenant, import_id, source_row, "
            f"{', '.join(db_cols)}) VALUES (?, ?, ?, "
            f"{', '.join('?' * len(db_cols))})")

        # An update only touches the columns this file actually carried, so a
        # partial export does not blank what it left out. PersonId is the key
        # and is never written.
        upd_cols = [c for h, c in USER_BASE_COLUMNS + USER_TAIL_COLUMNS
                    if h in col_ix and c != "person_id"]
        update_sql = (
            "UPDATE M3_Security_Users SET "
            + ", ".join(f"{c} = COALESCE(NULLIF(?, ''), {c})" for c in upd_cols)
            + ", import_id = ?, source_row = ?, modified_at = datetime('now') "
              "WHERE user_key = ?")
        upd_pos = [i for i, (h, c)
                   in enumerate(USER_BASE_COLUMNS + USER_TAIL_COLUMNS)
                   if h in col_ix and c != "person_id"]

        cur = conn.cursor()
        cur.execute(
            """
            INSERT INTO M3_Security_Imports
                (file_kind, file_name, tenant, file_guid, file_ticks,
                 security_slots, functional_slots, header_json, row_count)
            VALUES ('users', ?, ?, ?, ?, ?, ?, ?, 0)
            """,
            (path.name, tenant, meta["guid"], meta["ticks"],
             len(slots["SecurityRole"]), len(slots["FunctionalSecurityRole"]),
             header_json),
        )
        import_id = cur.lastrowid

        # Users already held keep their key, their row_state and their roles.
        existing = {r[0]: r[1] for r in conn.execute(
            "SELECT person_id, user_key FROM M3_Security_Users WHERE tenant = ?",
            (tenant,)).fetchall()}
        held_before = set(existing)
        n_kept = conn.execute(
            "SELECT COUNT(*) FROM M3_Security_UserRoles WHERE tenant = ?",
            (tenant,)).fetchone()[0]

        n_users = 0
        n_roles = 0
        n_new = 0
        n_updated = 0
        n_skipped = 0
        seen_ids = set()
        for row_no, row in enumerate(reader, start=1):
            if not any(v.strip() for v in row):
                continue
            if len(row) < len(header):
                row = row + [""] * (len(header) - len(row))

            values = []
            for h, _ in USER_BASE_COLUMNS:
                i = base_ix[h]
                values.append(row[i].strip() if i is not None else None)
            for h, _ in USER_TAIL_COLUMNS:
                i = tail_ix[h]
                values.append(row[i].strip() if i is not None else None)

            person_id = values[0]
            if not person_id:
                n_skipped += 1
                log.warning("%s: row %d has no %s, skipped",
                            path.name, row_no, USER_KEY_COLUMN)
                continue

            n_users += 1
            seen_ids.add(person_id)

            user_key = existing.get(person_id)
            if user_key is not None:
                cur.execute(update_sql,
                            (*[values[i] for i in upd_pos],
                             import_id, row_no, user_key))
                n_updated += 1
                continue

            cur.execute(insert_sql, (tenant, import_id, row_no, *values))
            user_key = cur.lastrowid
            existing[person_id] = user_key
            n_new += 1

            role_rows = []
            for prefix, role_type in prefix_to_type.items():
                seen = set()
                seq = 0
                for idx, _slot in slots[prefix]:
                    name = row[idx].strip()
                    if not name or name in seen:
                        continue
                    seen.add(name)
                    seq += 1
                    role_rows.append(
                        (user_key, tenant, person_id, role_type, seq, name))
            if not role_rows:
                continue
            cur.executemany(
                "INSERT OR IGNORE INTO M3_Security_UserRoles (user_key, tenant, "
                "person_id, role_type, seq, role_name) VALUES (?, ?, ?, ?, ?, ?)",
                role_rows)
            n_roles += len(role_rows)

        n_absent = len(held_before - seen_ids)

        cur.execute(
            "UPDATE M3_Security_Imports SET row_count = ? WHERE import_id = ?",
            (n_users, import_id))

        conn.commit()

    return {
        "kind": "users",
        "tenant": tenant,
        "tenant_source": tenant_source,
        "import_id": import_id,
        "file_name": path.name,
        "users": n_users,
        "new_users": n_new,
        "updated_users": n_updated,
        "users_not_in_file": n_absent,
        "skipped_rows": n_skipped,
        "role_links": n_roles,
        "role_links_kept": n_kept,
        "security_slots": len(slots["SecurityRole"]),
        "functional_slots": len(slots["FunctionalSecurityRole"]),
        "missing_columns": missing,
        "guid": meta["guid"],
        "ticks": meta["ticks"],
    }


def import_roles(path: Path, conn: sqlite3.Connection, tenant: str | None = None,
                 tenant_fallback: str | None = None) -> dict:
    meta = parse_file_name(path.name)
    tenant, tenant_source = resolve_tenant(
        path.name, meta, tenant, tenant_fallback,
        "Add the tenant to the file name (e.g. 'Security "
        "Role_ZFQP353QZYV89ZHG_TST_08_17_2026....csv').")

    with open(path, "r", encoding="utf-8-sig", newline="") as fh:
        reader = csv.reader(fh)
        next(reader)

        cur = conn.cursor()
        cur.execute(
            """
            INSERT INTO M3_Security_Imports
                (file_kind, file_name, tenant, file_stamp, row_count)
            VALUES ('roles', ?, ?, ?, 0)
            """,
            (path.name, tenant, meta["stamp"]),
        )
        import_id = cur.lastrowid

        # This file is a full refresh for the tenant.
        cur.execute("DELETE FROM M3_Security_RoleAssignments WHERE tenant = ?",
                    (tenant,))
        cur.execute("DELETE FROM M3_Security_Roles WHERE tenant = ?", (tenant,))

        role_keys = {}
        n_rows = 0
        n_assign = 0
        for seq, row in enumerate(reader, start=1):
            if not row or not any(v.strip() for v in row):
                continue
            name = row[0].strip()
            desc = row[1].strip() if len(row) > 1 else ""
            email = row[2].strip() if len(row) > 2 else ""
            if not name:
                continue
            n_rows += 1

            key = role_keys.get(name)
            if key is None:
                cur.execute(
                    "INSERT INTO M3_Security_Roles (tenant, name, description, "
                    "import_id, seq) VALUES (?, ?, ?, ?, ?)",
                    (tenant, name, desc, import_id, seq))
                key = cur.lastrowid
                role_keys[name] = key
            elif desc:
                cur.execute(
                    "UPDATE M3_Security_Roles SET description = ? WHERE "
                    "role_key = ? AND COALESCE(description, '') = ''",
                    (desc, key))

            cur.execute(
                "INSERT OR IGNORE INTO M3_Security_RoleAssignments (role_key, "
                "tenant, role_name, email_id, seq) VALUES (?, ?, ?, ?, ?)",
                (key, tenant, name, email, seq))
            n_assign += cur.rowcount

        cur.execute(
            "UPDATE M3_Security_Imports SET row_count = ? WHERE import_id = ?",
            (n_rows, import_id))
        conn.commit()

    return {
        "kind": "roles",
        "tenant": tenant,
        "tenant_source": tenant_source,
        "import_id": import_id,
        "file_name": path.name,
        "roles": len(role_keys),
        "assignments": n_assign,
        "rows": n_rows,
        "stamp": meta["stamp"],
    }


def import_functional(path: Path, conn: sqlite3.Connection,
                      tenant: str | None = None,
                      tenant_fallback: str | None = None) -> dict:
    """
    Load a "Functional Security Role" export.

    IFS caps this export at 75 roles per file, so several files make up the
    whole picture. This one MERGES rather than replacing: existing functional
    roles keep their members and anything already present is skipped, so
    loading file 2 does not wipe file 1 and re-loading the same file twice
    changes nothing.

    A functional role with no security roles comes through as a two-field row
    (FSR_OVEX,OVEX) rather than one with trailing commas - that is a role with
    an empty membership, not a malformed line.
    """
    meta = parse_file_name(path.name)
    tenant, tenant_source = resolve_tenant(
        path.name, meta, tenant, tenant_fallback,
        "Add the tenant to the file name (e.g. 'Functional Security "
        "Role_ZFQP353QZYV89ZHG_TST_08_18_2026....csv').")

    with open(path, "r", encoding="utf-8-sig", newline="") as fh:
        reader = csv.reader(fh)
        next(reader)

        cur = conn.cursor()
        cur.execute(
            """
            INSERT INTO M3_Security_Imports
                (file_kind, file_name, tenant, file_stamp, row_count)
            VALUES ('functional', ?, ?, ?, 0)
            """,
            (path.name, tenant, meta["stamp"]),
        )
        import_id = cur.lastrowid

        existing_fsr = {r[0]: r[1] for r in conn.execute(
            "SELECT name, fsr_key FROM M3_Security_FunctionalRoles "
            "WHERE tenant = ?", (tenant,)).fetchall()}
        existing_members = {(r[0], r[1], r[2]) for r in conn.execute(
            "SELECT fsr_name, security_role, email_id FROM "
            "M3_Security_FunctionalRoleMembers WHERE tenant = ?",
            (tenant,)).fetchall()}
        next_seq = conn.execute(
            "SELECT COALESCE(MAX(seq), 0) FROM M3_Security_FunctionalRoles "
            "WHERE tenant = ?", (tenant,)).fetchone()[0] + 1

        n_rows = 0
        new_fsr, seen_fsr = [], set()
        added_members, dup_members, empty_fsr = 0, 0, 0

        for seq, row in enumerate(reader, start=1):
            if not row or not any(v.strip() for v in row):
                continue
            name = row[0].strip()
            if not name:
                continue
            desc = row[1].strip() if len(row) > 1 else ""
            role = row[2].strip() if len(row) > 2 else ""
            email = row[3].strip() if len(row) > 3 else ""
            n_rows += 1
            seen_fsr.add(name)

            key = existing_fsr.get(name)
            if key is None:
                cur.execute(
                    "INSERT INTO M3_Security_FunctionalRoles (tenant, name, "
                    "description, seq, import_id, row_state) "
                    "VALUES (?, ?, ?, ?, ?, 'unchanged')",
                    (tenant, name, desc, next_seq, import_id))
                key = cur.lastrowid
                existing_fsr[name] = key
                next_seq += 1
                new_fsr.append(name)
            elif desc:
                cur.execute(
                    "UPDATE M3_Security_FunctionalRoles SET description = ? "
                    "WHERE fsr_key = ? AND COALESCE(description, '') = ''",
                    (desc, key))

            # A two-field row is a role with no members, not a bad line.
            if not role and not email:
                empty_fsr += 1
                continue

            if (name, role, email) in existing_members:
                dup_members += 1
                continue

            cur.execute(
                "INSERT OR IGNORE INTO M3_Security_FunctionalRoleMembers "
                "(fsr_key, tenant, fsr_name, security_role, email_id, seq) "
                "VALUES (?, ?, ?, ?, ?, ?)",
                (key, tenant, name, role, email, seq))
            if cur.rowcount:
                added_members += 1
                existing_members.add((name, role, email))
            else:
                dup_members += 1

        cur.execute(
            "UPDATE M3_Security_Imports SET row_count = ? WHERE import_id = ?",
            (n_rows, import_id))
        conn.commit()

    return {
        "kind": "functional",
        "tenant": tenant,
        "tenant_source": tenant_source,
        "import_id": import_id,
        "file_name": path.name,
        "rows": n_rows,
        "functional_roles": len(seen_fsr),
        "new_functional_roles": len(new_fsr),
        "existing_functional_roles": len(seen_fsr) - len(new_fsr),
        "members_added": added_members,
        "members_skipped": dup_members,
        "empty_roles": empty_fsr,
        "stamp": meta["stamp"],
    }


def import_file(file_path: str | Path, conn: sqlite3.Connection | None = None,
                tenant: str | None = None, db_path: str | None = None,
                tenant_fallback: str | None = None) -> dict:
    """
    Import one CSV, choosing the loader from the header row.

    ``tenant`` overrides whatever the file name says; ``tenant_fallback`` is
    only used when the file name carries no tenant of its own - that is the
    tenant selected in the front end.
    """
    path = Path(file_path)
    if not path.exists():
        raise FileNotFoundError(path)

    own_conn = conn is None
    conn = conn or connect(db_path)
    try:
        kind = detect_kind(read_header(path))
        if kind == "users":
            return import_users(path, conn, tenant, tenant_fallback)
        if kind == "functional":
            return import_functional(path, conn, tenant, tenant_fallback)
        return import_roles(path, conn, tenant, tenant_fallback)
    finally:
        if own_conn:
            conn.close()


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(
        description="Load Infor OS / M3 security exports into doppio.db.")
    ap.add_argument("files", nargs="+", help="CSV export file(s)")
    ap.add_argument("--tenant", default=None,
                    help="Override the tenant instead of reading it from the "
                         "file name")
    ap.add_argument("--tenant-fallback", default=None, dest="tenant_fallback",
                    help="Tenant to use only for files whose name does not "
                         "carry one (the front end passes the selected tenant "
                         "here)")
    ap.add_argument("--db", default=None,
                    help=f"SQLite path (default {DEFAULT_DB_PATH})")
    ap.add_argument("--verbose", action="store_true")
    args = ap.parse_args(argv)

    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(asctime)s  %(levelname)-8s  %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
        handlers=[logging.StreamHandler(sys.stdout)],
    )

    conn = connect(args.db)
    log.info("Database: %s", resolve_db_path(args.db))
    rc = 0
    try:
        for f in args.files:
            try:
                res = import_file(f, conn=conn, tenant=args.tenant,
                                  tenant_fallback=args.tenant_fallback)
                if res.get("tenant_source") == "selected":
                    log.info("%s has no tenant in its name - loaded into %s",
                             res["file_name"], res["tenant"])
                if res["kind"] == "users":
                    log.info("%s -> tenant %s: %s users, %s role links "
                             "(%s security slots, %s functional slots)",
                             res["file_name"], res["tenant"], res["users"],
                             res["role_links"], res["security_slots"],
                             res["functional_slots"])
                elif res["kind"] == "functional":
                    log.info("%s -> tenant %s: %s functional role(s) (%s new, "
                             "%s already held), %s member(s) added, %s "
                             "duplicate(s) skipped",
                             res["file_name"], res["tenant"],
                             res["functional_roles"],
                             res["new_functional_roles"],
                             res["existing_functional_roles"],
                             res["members_added"], res["members_skipped"])
                else:
                    log.info("%s -> tenant %s: %s roles, %s assignments "
                             "(%s rows)",
                             res["file_name"], res["tenant"], res["roles"],
                             res["assignments"], res["rows"])
            except Exception as exc:
                log.error("%s: %s", f, exc)
                rc = 1
    finally:
        conn.close()
    return rc


if __name__ == "__main__":
    raise SystemExit(main())
