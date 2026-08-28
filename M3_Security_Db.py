"""
M3_Security_Db - schema and connection helpers for the M3 security capture.

Everything the capture holds lives in one SQLite file, by default
~/sqlite/doppio.db, matching Sheet2Db.py and config.py. An explicit --db
argument wins, then the M3_SECURITY_DB environment variable, then the default.

Reconstructed from M3_Security_Db.cpython-312.pyc after the source was lost.
Logic matches the bytecode; comments and formatting are not original.
"""
from __future__ import annotations

import os
import sqlite3
from pathlib import Path

# The database lives beside the other Doppio SQLite files.
SQLITE_DIR = Path.home() / "sqlite"
DEFAULT_DB_PATH = str(SQLITE_DIR / "doppio.db")

# Overrides the default without a command-line argument.
DB_ENV_VAR = "M3_SECURITY_DB"

# The IFS users export, column header -> database column. Order matters: it is
# the order the export is written back out in.
USER_BASE_COLUMNS: list[tuple[str, str]] = [
    ("PersonId", "person_id"),
    ("FirstName", "first_name"),
    ("LastName", "last_name"),
    ("EmailId", "email_id"),
    ("Title", "title"),
    ("Status", "status"),
    ("FederatedId", "federated_id"),
    ("UserGuid", "user_guid"),
    ("User Name", "user_name"),
    ("Created Date", "created_date"),
    ("LastLogin Date", "last_login_date"),
    ("UPN", "upn"),
    ("User Alias", "user_alias"),
    ("GenericProperty_Language", "gp_language"),
    ("GenericProperty_LanguageOrigin", "gp_language_origin"),
    ("GenericProperty_Locale", "gp_locale"),
    ("GenericProperty_LocaleOrigin", "gp_locale_origin"),
    ("GenericProperty_ApplicationTimezone", "gp_application_timezone"),
    ("GenericProperty_Timezone", "gp_timezone"),
    ("GenericProperty_TimezoneOrigin", "gp_timezone_origin"),
]

# Written after the role columns rather than before them.
USER_TAIL_COLUMNS: list[tuple[str, str]] = [
    ("SecurityAccessProfiles", "security_access_profiles"),
]

# The two repeating column blocks in the users export, and the role_type each
# one is stored under.
ROLE_BLOCKS: list[tuple[str, str]] = [
    ("SecurityRole", "Security"),
    ("FunctionalSecurityRole", "Functional"),
]

# What the user editor is allowed to write. PersonId, UserGuid and the dates
# are identity or IFS-owned and stay out.
USER_EDITABLE_FIELDS = [
    "first_name",
    "last_name",
    "email_id",
    "title",
    "status",
    "federated_id",
    "user_name",
    "upn",
    "user_alias",
    "gp_language",
    "gp_language_origin",
    "gp_locale",
    "gp_locale_origin",
    "gp_application_timezone",
    "gp_timezone",
    "gp_timezone_origin",
    "security_access_profiles",
]

SCHEMA = """
PRAGMA foreign_keys = ON;

-- ---------------------------------------------------------------- users
CREATE TABLE IF NOT EXISTS M3_Security_Users (
    user_key                 INTEGER PRIMARY KEY AUTOINCREMENT,
    tenant                   TEXT    NOT NULL,
    person_id                TEXT    NOT NULL,
    first_name               TEXT,
    last_name                TEXT,
    email_id                 TEXT,
    title                    TEXT,
    status                   TEXT,
    federated_id             TEXT,
    user_guid                TEXT,
    user_name                TEXT,
    created_date             TEXT,
    last_login_date          TEXT,
    upn                      TEXT,
    user_alias               TEXT,
    gp_language              TEXT,
    gp_language_origin       TEXT,
    gp_locale                TEXT,
    gp_locale_origin         TEXT,
    gp_application_timezone  TEXT,
    gp_timezone              TEXT,
    gp_timezone_origin       TEXT,
    security_access_profiles TEXT,
    row_state                TEXT    NOT NULL DEFAULT 'unchanged',
    import_id                INTEGER,
    source_row               INTEGER,
    created_at               TEXT    NOT NULL DEFAULT (datetime('now')),
    modified_at              TEXT,
    UNIQUE (tenant, person_id)
);
CREATE INDEX IF NOT EXISTS ix_m3sec_users_tenant  ON M3_Security_Users (tenant);
CREATE INDEX IF NOT EXISTS ix_m3sec_users_email   ON M3_Security_Users (tenant, email_id);
CREATE INDEX IF NOT EXISTS ix_m3sec_users_state   ON M3_Security_Users (row_state);

-- ------------------------------------------------------------ user roles
CREATE TABLE IF NOT EXISTS M3_Security_UserRoles (
    user_role_key INTEGER PRIMARY KEY AUTOINCREMENT,
    user_key      INTEGER NOT NULL
                  REFERENCES M3_Security_Users (user_key) ON DELETE CASCADE,
    tenant        TEXT    NOT NULL,
    person_id     TEXT    NOT NULL,
    role_type     TEXT    NOT NULL,          -- 'Security' | 'Functional'
    seq           INTEGER NOT NULL,          -- 1-based position in the export
    role_name     TEXT    NOT NULL,
    row_state     TEXT    NOT NULL DEFAULT 'unchanged',
    UNIQUE (user_key, role_type, role_name)
);
CREATE INDEX IF NOT EXISTS ix_m3sec_userroles_user ON M3_Security_UserRoles (user_key, role_type, seq);
CREATE INDEX IF NOT EXISTS ix_m3sec_userroles_role ON M3_Security_UserRoles (tenant, role_name);

-- -------------------------------------------------------- security roles
CREATE TABLE IF NOT EXISTS M3_Security_Roles (
    role_key    INTEGER PRIMARY KEY AUTOINCREMENT,
    tenant      TEXT    NOT NULL,
    name        TEXT    NOT NULL,
    description TEXT,
    row_state   TEXT    NOT NULL DEFAULT 'unchanged',
    import_id   INTEGER,
    seq         INTEGER,                     -- first appearance in the export
    created_at  TEXT    NOT NULL DEFAULT (datetime('now')),
    modified_at TEXT,
    UNIQUE (tenant, name)
);
CREATE INDEX IF NOT EXISTS ix_m3sec_roles_tenant ON M3_Security_Roles (tenant);

-- ---------------------------------------------------- role -> user links
CREATE TABLE IF NOT EXISTS M3_Security_RoleAssignments (
    assignment_key INTEGER PRIMARY KEY AUTOINCREMENT,
    role_key       INTEGER NOT NULL
                   REFERENCES M3_Security_Roles (role_key) ON DELETE CASCADE,
    tenant         TEXT    NOT NULL,
    role_name      TEXT    NOT NULL,
    email_id       TEXT    NOT NULL DEFAULT '',
    row_state      TEXT    NOT NULL DEFAULT 'unchanged',
    seq            INTEGER,
    created_at     TEXT    NOT NULL DEFAULT (datetime('now')),
    UNIQUE (role_key, email_id)
);
CREATE INDEX IF NOT EXISTS ix_m3sec_assign_role  ON M3_Security_RoleAssignments (role_key, seq);
CREATE INDEX IF NOT EXISTS ix_m3sec_assign_email ON M3_Security_RoleAssignments (tenant, email_id);

-- --------------------------------------------------------- import audit
CREATE TABLE IF NOT EXISTS M3_Security_Imports (
    import_id       INTEGER PRIMARY KEY AUTOINCREMENT,
    file_kind       TEXT    NOT NULL,        -- 'users' | 'roles'
    file_name       TEXT    NOT NULL,
    tenant          TEXT    NOT NULL,
    file_guid       TEXT,                    -- users export: 32-char guid
    file_ticks      TEXT,                    -- users export: .NET ticks
    file_stamp      TEXT,                    -- roles export: MM_DD_YYYY_HH_MM_SS.fffffff
    security_slots  INTEGER,                 -- SecurityRoleN column count
    functional_slots INTEGER,                -- FunctionalSecurityRoleN column count
    header_json     TEXT,                    -- the base/tail columns this file had
    row_count       INTEGER,
    imported_at     TEXT    NOT NULL DEFAULT (datetime('now'))
);
CREATE INDEX IF NOT EXISTS ix_m3sec_imports_tenant ON M3_Security_Imports (tenant, file_kind);

-- --------------------------------------------- role members live from M3
CREATE TABLE IF NOT EXISTS M3_Security_M3Members (
    m3_member_key INTEGER PRIMARY KEY AUTOINCREMENT,
    tenant        TEXT    NOT NULL,
    role_name     TEXT    NOT NULL,          -- ROLL
    usid          TEXT    NOT NULL,          -- USID
    valid_from    TEXT,                      -- FVDT
    valid_to      TEXT,                      -- VTDT
    checked_at    TEXT    NOT NULL DEFAULT (datetime('now')),
    UNIQUE (tenant, role_name, usid)
);
CREATE INDEX IF NOT EXISTS ix_m3sec_m3members_role ON M3_Security_M3Members (tenant, role_name);
CREATE INDEX IF NOT EXISTS ix_m3sec_m3members_usid ON M3_Security_M3Members (tenant, usid);

-- ------------------------------------------------ USID -> email from M3
CREATE TABLE IF NOT EXISTS M3_Security_M3Users (
    m3_user_key INTEGER PRIMARY KEY AUTOINCREMENT,
    tenant      TEXT    NOT NULL,
    usid        TEXT    NOT NULL,            -- USID
    name        TEXT,                        -- TX40
    email       TEXT,                        -- EMAL
    status      TEXT,                        -- STAT
    checked_at  TEXT    NOT NULL DEFAULT (datetime('now')),
    UNIQUE (tenant, usid)
);
CREATE INDEX IF NOT EXISTS ix_m3sec_m3users_email ON M3_Security_M3Users (tenant, LOWER(email));

-- ---------------------------------------------------- writes sent to M3
CREATE TABLE IF NOT EXISTS M3_Security_M3Log (
    m3_log_key  INTEGER PRIMARY KEY AUTOINCREMENT,
    tenant      TEXT    NOT NULL,
    program     TEXT    NOT NULL,
    transaction_name TEXT NOT NULL,
    role_name   TEXT,
    usid        TEXT,
    dry_run     INTEGER NOT NULL DEFAULT 0,
    outcome     TEXT    NOT NULL,            -- 'ok' | 'error' | 'preview'
    message     TEXT,
    payload     TEXT,
    response    TEXT,
    run_id      TEXT,
    logged_at   TEXT    NOT NULL DEFAULT (datetime('now'))
);
CREATE INDEX IF NOT EXISTS ix_m3sec_m3log_run  ON M3_Security_M3Log (run_id);
CREATE INDEX IF NOT EXISTS ix_m3sec_m3log_role ON M3_Security_M3Log (tenant, role_name);

-- ------------------------------------------ function authorisations (SES400)
CREATE TABLE IF NOT EXISTS M3_Security_FunctionRoles (
    function_role_key INTEGER PRIMARY KEY AUTOINCREMENT,
    tenant     TEXT NOT NULL,
    fnid       TEXT NOT NULL,          -- FNID, the function
    roll       TEXT NOT NULL,          -- ROLL, the role authorised to it
    cono       TEXT,                   -- CONO
    divi       TEXT,                   -- DIVI
    stat       TEXT,                   -- STAT
    txid       TEXT,                   -- TXID
    checked_at TEXT NOT NULL DEFAULT (datetime('now')),
    UNIQUE (tenant, fnid, roll, cono, divi)
);
CREATE INDEX IF NOT EXISTS ix_m3sec_fnrole_fn   ON M3_Security_FunctionRoles (tenant, fnid);
CREATE INDEX IF NOT EXISTS ix_m3sec_fnrole_roll ON M3_Security_FunctionRoles (tenant, roll);

-- ------------------------------- IFS functional security roles (FSR)
-- A functional security role contains security roles. Loaded from the IFS
-- "Functional Security Role" export, which caps at 75 roles per file, so
-- several files are merged together rather than replacing each other.
CREATE TABLE IF NOT EXISTS M3_Security_FunctionalRoles (
    fsr_key     INTEGER PRIMARY KEY AUTOINCREMENT,
    tenant      TEXT NOT NULL,
    name        TEXT NOT NULL,          -- Functional SecurityRole Name
    description TEXT,                   -- Description
    row_state   TEXT NOT NULL DEFAULT 'unchanged',
    seq         INTEGER,                -- first appearance across imports
    import_id   INTEGER,
    created_at  TEXT NOT NULL DEFAULT (datetime('now')),
    modified_at TEXT,
    UNIQUE (tenant, name)
);
CREATE INDEX IF NOT EXISTS ix_m3sec_fsr ON M3_Security_FunctionalRoles (tenant, name);

-- The security roles inside each functional role. EmailId is carried because
-- the export has the column, though IFS leaves it empty on this file.
CREATE TABLE IF NOT EXISTS M3_Security_FunctionalRoleMembers (
    fsr_member_key INTEGER PRIMARY KEY AUTOINCREMENT,
    fsr_key        INTEGER NOT NULL
                   REFERENCES M3_Security_FunctionalRoles (fsr_key) ON DELETE CASCADE,
    tenant         TEXT NOT NULL,
    fsr_name       TEXT NOT NULL,
    security_role  TEXT NOT NULL DEFAULT '',   -- SecurityRole
    email_id       TEXT NOT NULL DEFAULT '',   -- EmailId
    row_state      TEXT NOT NULL DEFAULT 'unchanged',
    seq            INTEGER,
    created_at     TEXT NOT NULL DEFAULT (datetime('now')),
    UNIQUE (fsr_key, security_role, email_id)
);
CREATE INDEX IF NOT EXISTS ix_m3sec_fsrmem ON M3_Security_FunctionalRoleMembers (tenant, fsr_name);
CREATE INDEX IF NOT EXISTS ix_m3sec_fsrmem_role ON M3_Security_FunctionalRoleMembers (tenant, security_role);

-- ------------------------------------------- MNS405 role master from M3
-- The role definitions as they are in M3, independent of the CSV capture in
-- M3_Security_Roles. This is what the MNS405 Roles tab reads and writes.
CREATE TABLE IF NOT EXISTS M3_Security_M3RoleDefs (
    role_def_key INTEGER PRIMARY KEY AUTOINCREMENT,
    tenant     TEXT NOT NULL,
    roll       TEXT NOT NULL,          -- ROLL
    tx40       TEXT,                   -- TX40 description
    tx15       TEXT,                   -- TX15 name
    rolt       TEXT,                   -- ROLT role type
    txid       TEXT,                   -- TXID
    checked_at TEXT NOT NULL DEFAULT (datetime('now')),
    UNIQUE (tenant, roll)
);
CREATE INDEX IF NOT EXISTS ix_m3sec_roledefs ON M3_Security_M3RoleDefs (tenant, roll);

-- Roles seen in SES400 and whether MNS405 has them. Kept separate from
-- M3_Security_Roles because SES400 roles need not appear in the CSV capture.
CREATE TABLE IF NOT EXISTS M3_Security_FunctionRoleStatus (
    tenant     TEXT NOT NULL,
    roll       TEXT NOT NULL,
    in_mns405  INTEGER,                -- 1 yes, 0 no, NULL not checked
    checked_at TEXT NOT NULL DEFAULT (datetime('now')),
    PRIMARY KEY (tenant, roll)
);
"""

# Added to M3_Security_Roles after the fact, so _migrate() puts them on
# databases created before the live-M3 columns existed.
_ROLE_M3_COLUMNS: list[tuple[str, str]] = [
    ("in_m3", "INTEGER"),
    ("m3_description", "TEXT"),
    ("m3_name", "TEXT"),
    ("m3_role_type", "TEXT"),
    ("m3_text_id", "TEXT"),
    ("m3_member_count", "INTEGER"),
    ("m3_checked_at", "TEXT"),
]


def resolve_db_path(db_path: str | os.PathLike | None = None) -> str:
    """Pick the database file: explicit arg > env var > default."""
    if db_path:
        return str(db_path)
    return os.environ.get(DB_ENV_VAR) or DEFAULT_DB_PATH


def connect(db_path: str | os.PathLike | None = None, init: bool = True) -> sqlite3.Connection:
    """Open (and by default initialise) the doppio database."""
    path = Path(resolve_db_path(db_path))
    path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(path))
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    if init:
        init_db(conn)
    return conn


def init_db(conn: sqlite3.Connection) -> None:
    """Create every M3_Security_* object if it does not already exist."""
    conn.executescript(SCHEMA)
    _migrate(conn)
    conn.commit()


def _migrate(conn: sqlite3.Connection) -> None:
    """Bring an existing database up to the current column set."""
    have = {r[1] for r in conn.execute("PRAGMA table_info(M3_Security_Roles)")}
    for col, decl in _ROLE_M3_COLUMNS:
        if col not in have:
            conn.execute(f"ALTER TABLE M3_Security_Roles ADD COLUMN {col} {decl}")

    have = {r[1] for r in conn.execute("PRAGMA table_info(M3_Security_Imports)")}
    if "header_json" not in have:
        conn.execute("ALTER TABLE M3_Security_Imports ADD COLUMN header_json TEXT")


def tenants(conn: sqlite3.Connection) -> list[str]:
    """
    Distinct tenants currently held in the database.

    Every table that carries a tenant is consulted, not just Users and Roles.
    Clearing one capture must not make a tenant vanish from the picker while
    its other captures - functional roles, the M3 side, the import history -
    are still held, and a tenant read straight from M3 should be selectable
    before any CSV has been loaded.
    """
    rows = conn.execute(
        """
        SELECT tenant FROM M3_Security_Users
        UNION SELECT tenant FROM M3_Security_Roles
        UNION SELECT tenant FROM M3_Security_FunctionalRoles
        UNION SELECT tenant FROM M3_Security_Imports
        UNION SELECT tenant FROM M3_Security_M3RoleDefs
        UNION SELECT tenant FROM M3_Security_FunctionRoles
        UNION SELECT tenant FROM M3_Security_M3Members
        ORDER BY 1
        """
    ).fetchall()
    return [r[0] for r in rows if r[0]]


def strip_role_members(conn: sqlite3.Connection, tenant: str, role_name: str, commit: bool = True) -> dict:
    """
    Take every member off a role locally.

    The role itself stays. A blank-EmailId row is left behind so the role still
    appears in the roles export with no members - which is exactly how the
    inbound file represents a member-less role - rather than vanishing from it.

    Returns the counts touched. Nothing is sent to M3 here.
    """
    role = conn.execute(
        "SELECT role_key FROM M3_Security_Roles WHERE tenant = ? AND name = ?",
        (tenant, role_name),
    ).fetchone()

    cur = conn.cursor()
    assignments = 0

    if role:
        role_key = role["role_key"]
        cur.execute(
            "UPDATE M3_Security_RoleAssignments SET row_state = 'deleted' "
            "WHERE role_key = ? AND email_id <> '' AND row_state <> 'deleted'",
            (role_key,),
        )
        assignments = cur.rowcount

        blank = conn.execute(
            "SELECT assignment_key, row_state FROM M3_Security_RoleAssignments "
            "WHERE role_key = ? AND email_id = ''",
            (role_key,),
        ).fetchone()
        if blank is None:
            cur.execute(
                "INSERT INTO M3_Security_RoleAssignments "
                "(role_key, tenant, role_name, email_id, row_state) "
                "VALUES (?, ?, ?, '', 'new')",
                (role_key, tenant, role_name),
            )
        elif blank["row_state"] == "deleted":
            cur.execute(
                "UPDATE M3_Security_RoleAssignments SET row_state = 'modified' "
                "WHERE assignment_key = ?",
                (blank["assignment_key"],),
            )

        cur.execute(
            "UPDATE M3_Security_Roles SET row_state = CASE WHEN row_state = 'new' "
            "THEN 'new' ELSE 'modified' END, modified_at = datetime('now') "
            "WHERE role_key = ?",
            (role_key,),
        )

    cur.execute(
        "UPDATE M3_Security_UserRoles SET row_state = 'deleted' "
        "WHERE tenant = ? AND role_name = ? AND row_state <> 'deleted'",
        (tenant, role_name),
    )
    user_links = cur.rowcount

    cur.execute(
        """
        UPDATE M3_Security_Users
           SET row_state = CASE WHEN row_state IN ('new', 'deleted')
                                THEN row_state ELSE 'modified' END,
               modified_at = datetime('now')
         WHERE tenant = ?
           AND user_key IN (SELECT user_key FROM M3_Security_UserRoles
                             WHERE tenant = ? AND role_name = ?)
        """,
        (tenant, tenant, role_name),
    )
    users_touched = cur.rowcount

    if commit:
        conn.commit()

    return {
        "role": role_name,
        "assignments": assignments,
        "user_links": user_links,
        "users": users_touched,
        "missing": role is None,
    }


def strip_user_roles(
    conn: sqlite3.Connection,
    tenant: str,
    user_keys: list[int],
    role_type: str = "Functional",
    role_names: list[str] | None = None,
    commit: bool = True,
) -> dict:
    """
    Take roles off a set of users.

    role_type picks the block: 'Functional' for the FunctionalSecurityRoleN
    columns, 'Security' for SecurityRoleN. role_names limits it to named roles;
    None means every role of that type the users hold.

    The users are flagged as changed so they reach the next users export, which
    is what actually removes the roles in IFS. Nothing is sent to M3.
    """
    if not user_keys:
        return {"users": 0, "role_links": 0, "roles": []}

    marks = ",".join("?" * len(user_keys))
    args = [tenant, role_type] + list(user_keys)

    name_clause = ""
    if role_names:
        name_clause = " AND role_name IN (%s)" % ",".join("?" * len(role_names))
        args += list(role_names)

    affected = conn.execute(
        f"""
        SELECT DISTINCT role_name FROM M3_Security_UserRoles
         WHERE tenant = ? AND role_type = ? AND row_state <> 'deleted'
           AND user_key IN ({marks}){name_clause}
         ORDER BY role_name
        """,
        args,
    ).fetchall()

    cur = conn.cursor()
    cur.execute(
        f"""
        UPDATE M3_Security_UserRoles SET row_state = 'deleted'
         WHERE tenant = ? AND role_type = ? AND row_state <> 'deleted'
           AND user_key IN ({marks}){name_clause}
        """,
        args,
    )
    role_links = cur.rowcount

    cur.execute(
        f"""
        UPDATE M3_Security_Users
           SET row_state = CASE WHEN row_state IN ('new', 'deleted')
                                THEN row_state ELSE 'modified' END,
               modified_at = datetime('now')
         WHERE tenant = ? AND user_key IN ({marks})
           AND user_key IN (SELECT user_key FROM M3_Security_UserRoles
                             WHERE tenant = ? AND role_type = ?
                               AND row_state = 'deleted')
        """,
        [tenant] + list(user_keys) + [tenant, role_type],
    )
    users = cur.rowcount

    if commit:
        conn.commit()

    return {
        "users": users,
        "role_links": role_links,
        "roles": [r[0] for r in affected],
    }


def roles_missing_from_m3(conn: sqlite3.Connection, tenant: str) -> list[dict]:
    """
    MNS405 role definitions the IFS Security Roles capture does not hold.

    Matching is by name: M3_Security_M3RoleDefs.roll against
    M3_Security_Roles.name, which is the same pairing 'Check M3' uses to set
    in_m3. The counts come along so the preview can show what each role is
    carrying in M3 before anything is written.
    """
    return [dict(r) for r in conn.execute(
        """
        SELECT d.roll, d.tx40, d.tx15, d.rolt, d.txid,
               (SELECT COUNT(*) FROM M3_Security_M3Members m
                 WHERE m.tenant = d.tenant AND m.role_name = d.roll)  AS n_users,
               (SELECT COUNT(DISTINCT f.fnid) FROM M3_Security_FunctionRoles f
                 WHERE f.tenant = d.tenant AND f.roll = d.roll)       AS n_functions
        FROM M3_Security_M3RoleDefs d
        WHERE d.tenant = ?
          AND NOT EXISTS (SELECT 1 FROM M3_Security_Roles r
                           WHERE r.tenant = d.tenant AND r.name = d.roll)
        ORDER BY d.roll
        """,
        (tenant,)).fetchall()]


def add_roles_from_m3(conn: sqlite3.Connection, tenant: str,
                      names: list[str] | None = None,
                      commit: bool = True) -> dict:
    """
    Copy MNS405 role definitions into the IFS Security Roles capture.

    The roles are added as 'new', which is what puts them in the next Security
    Role export - and that export is what actually creates them in IFS. Nothing
    is sent to M3 here; they are already there, which is the whole point.

    TX40 becomes the description (the role name if TX40 is blank) and each role
    gets the blank-EmailId member row, so a role with no members still appears
    in the export rather than vanishing from it - the same shape '+ Role' uses.
    The M3 columns are filled in at the same time, because a role copied from
    MNS405 is by definition in M3 and should not need a 'Check M3' to say so.

    names limits it to the roles given; None takes every missing one. A name
    already in the capture is skipped, so running this twice is harmless.
    """
    missing = roles_missing_from_m3(conn, tenant)
    if names is not None:
        wanted = {n.strip() for n in names if n and n.strip()}
        missing = [m for m in missing if m["roll"] in wanted]

    cur = conn.cursor()
    added = []
    skipped = []
    for m in missing:
        name = m["roll"]
        description = (m["tx40"] or "").strip() or name
        try:
            cur.execute(
                "INSERT INTO M3_Security_Roles "
                "(tenant, name, description, row_state, in_m3, m3_description, "
                " m3_name, m3_role_type, m3_text_id, m3_member_count, "
                " m3_checked_at) "
                "VALUES (?, ?, ?, 'new', 1, ?, ?, ?, ?, ?, datetime('now'))",
                (tenant, name, description, m["tx40"], m["tx15"], m["rolt"],
                 m["txid"], m["n_users"]))
        except sqlite3.IntegrityError:
            # Already there - something added it between the plan and the write.
            skipped.append(name)
            continue
        role_key = cur.lastrowid
        cur.execute(
            "INSERT INTO M3_Security_RoleAssignments "
            "(role_key, tenant, role_name, email_id, row_state) "
            "VALUES (?, ?, ?, '', 'new')",
            (role_key, tenant, name))
        added.append({"name": name, "description": description,
                      "role_key": role_key, "n_users": m["n_users"],
                      "n_functions": m["n_functions"]})

    if commit:
        conn.commit()

    return {"tenant": tenant, "added": added, "skipped": skipped,
            "totals": {"added": len(added), "skipped": len(skipped),
                       "considered": len(missing)}}


def latest_import(conn: sqlite3.Connection, file_kind: str, tenant: str | None = None) -> sqlite3.Row | None:
    """Most recent import row for a file kind (optionally per tenant)."""
    sql = "SELECT * FROM M3_Security_Imports WHERE file_kind = ?"
    args = [file_kind]
    if tenant:
        sql += " AND tenant = ?"
        args.append(tenant)
    sql += " ORDER BY import_id DESC LIMIT 1"
    return conn.execute(sql, args).fetchone()


if __name__ == "__main__":
    import argparse

    ap = argparse.ArgumentParser(description="Create the M3_Security_* schema.")
    ap.add_argument("--db", default=None, help=f"SQLite path (default {DEFAULT_DB_PATH})")
    args = ap.parse_args()

    c = connect(args.db)
    print(f"Schema ready in {resolve_db_path(args.db)}")
    for t in c.execute(
        "SELECT name FROM sqlite_master WHERE type='table' "
        "AND name LIKE 'M3_Security_%' ORDER BY name"
    ):
        print("  ", t[0])
    c.close()
