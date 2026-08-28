"""
M3_Security_M3Api - live M3 calls over the ION REST API.

The rest of the M3 Security tool works on a captured copy of Infor OS security:
roles and role assignments read out of an IFS export.  This module is the half
that talks to M3 itself, so the capture can be compared with - and pushed into -
what M3 actually holds.

Three areas of M3 are covered:

  * MNS405  role definitions        (Lst / Get / Add / Upd / Dlt)
  * MNS410  role per user           (Lst / Add / Upd / Dlt)
  * SES400  function authorisations (Lst / Upd / Dlt)

plus MNS150MI/LstUserData for the USID -> email map that lets M3 members line up
with the email addresses in the capture, and MNS095MI / MNS100MI for the company
and division discovery behind the company picker.

Connection details come from an .ionapi file: every file in the ionapi directory
is read and the one whose "ti" equals the requested tenant wins, so nothing has
to be selected by hand.  Authentication is the service-account password grant.

Company and division are deliberately not assumed.  Left unset they are dropped
from the request and M3 falls back to the service account's own defaults, which
is right on most tenants; a tenant that needs a specific company pins it in
M3_Security_M3.json.  Precedence is the explicit argument, then the file, then
nothing at all.

Every bulk operation posts up to batch_size records in one m3api-rest call
instead of one call per record, and reports per record: the response results
array comes back in the same order as the transactions that went in, so
result[j] belongs to record[j].  A short array means M3 dropped the tail, and
those records are reported as failures rather than counted as successes.

Every write goes through the same shape:

    dry_run=True (the default) sends nothing and returns the plan
    a real run needs dry_run=False AND a matching confirm string

and lands one audit row per record in M3_Security_M3Log, so a bulk call can be
read back afterwards record by record.  Nothing here writes to M3 by accident.

Deleting an MNS405 role additionally publishes a SyncSecurityRoleMaster BOD
carrying Operation DELETE, the same document M3 publishes from the CMNROL smart
rule, and strips the role off the captured IFS users.

The module is importable (the Flask app calls these functions directly) and also
runs as a command line tool - see main() at the bottom for the flags.

Reconstructed from M3_Security_M3Api.cpython-312.pyc after the source was lost.
Logic matches the bytecode; comments and formatting are not original.
"""

from __future__ import annotations

import argparse
import json
import logging
import sqlite3
import sys
import time
import uuid
from pathlib import Path

import requests

from M3_Security_Bod import (
    BodConfigError,
    build_and_write,
    tenant_bod_config,
)
from M3_Security_Db import (
    DEFAULT_DB_PATH,
    connect,
    resolve_db_path,
    strip_role_members,
)

log = logging.getLogger("M3_Security_M3Api")

BASE_DIR = Path(__file__).parent.resolve()
DEFAULT_IONAPI_DIR = BASE_DIR / "ionapi"

# Company and division are not assumed.  Hard-coding them (they used to be the
# 001 / 100 of the original curl example) breaks on any tenant that numbers its
# companies differently: M3 answers 403 because the service account has no
# authority for a company that is not its own.
#
# Left as None they are omitted from the request entirely and M3 uses the
# service account's own default company and division from its MNS150 record,
# which is the right answer most of the time.
#
# Order of precedence: the explicit argument, then M3_Security_M3.json, then
# nothing.
DEFAULT_COMPANY: str | None = None
DEFAULT_DIVISION: str | None = None

# Per-tenant company / division overrides - see load_m3_config().
CONFIG_PATH = BASE_DIR / "M3_Security_M3.json"

DEFAULT_TIMEOUT = 120

# m3api-rest takes an array of transactions in one body, so a bulk call is one
# HTTP round trip per batch_size records rather than one per record.
DEFAULT_BATCH_SIZE = 500

# SES400 status codes.  M3 only defines these two.
FUNCTION_STATUS = {
    "10": "Preliminary",
    "20": "Active",
}


def function_status_label(stat: str | None) -> str:
    """'20' -> '20 Active'. Unknown codes are passed through as they are."""
    stat = (stat or "").strip()
    if not stat:
        return ""
    name = FUNCTION_STATUS.get(stat)
    return f"{stat} {name}" if name else stat


def _noop_progress(done: int, total: int, message: str = "") -> None:
    """Default progress sink - see the 'progress' argument on the bulk calls."""


class M3ApiError(RuntimeError):
    """Anything that goes wrong talking to M3."""


# .ionapi discovery, and the per-tenant company / division overrides.
# ---------------------------------------------------------------------------

def list_ionapi_files(ionapi_dir: str | Path = DEFAULT_IONAPI_DIR) -> list[dict]:
    """Every readable .ionapi file with the tenant it points at."""
    out = []
    d = Path(ionapi_dir)
    if not d.is_dir():
        return out
    for f in sorted(d.glob("*.ionapi"), key=lambda p: p.name.lower()):
        try:
            cfg = json.loads(f.read_text())
        except Exception as exc:
            out.append({"file": f.name, "tenant": None, "error": str(exc)})
            continue
        out.append({"file": f.name, "tenant": cfg.get("ti"),
                    "url": cfg.get("iu"), "path": str(f)})
    return out


def load_m3_config(path: str | Path = CONFIG_PATH) -> dict:
    """
    Per-tenant company / division overrides. Missing file is fine.

    {
      "DOPPIO_DEM": {"cono": "100", "divi": ""},
      "_default":   {"cono": "",    "divi": ""}
    }

    A blank or absent value means "do not send it", which lets M3 fall back to
    the service account's own default company and division.
    """
    path = Path(path)
    if not path.exists():
        return {}
    try:
        return json.loads(path.read_text() or "{}")
    except json.JSONDecodeError as exc:
        raise M3ApiError(f"{path.name} is not valid JSON: {exc}") from exc


def tenant_company(tenant: str, path: str | Path = CONFIG_PATH) -> tuple:
    """(cono, divi) for a tenant - None for either when it is not pinned."""
    cfg = load_m3_config(path)
    entry = {**cfg.get("_default", {}), **cfg.get(tenant, {})}
    cono = str(entry.get("cono") or "").strip() or None
    divi = str(entry.get("divi") or "").strip() or None
    return cono, divi


def save_tenant_company(tenant: str, cono: str | None, divi: str | None,
                        path: str | Path = CONFIG_PATH) -> dict:
    """Pin (or unpin, with blanks) a tenant's company and division."""
    path = Path(path)
    cfg = load_m3_config(path)
    cfg[tenant] = {"cono": (cono or "").strip(), "divi": (divi or "").strip()}
    path.write_text(json.dumps(cfg, indent=2) + "\n")
    return cfg[tenant]


def discover_companies(client: "M3Client") -> list[dict]:
    """
    Ask M3 which companies and divisions this tenant actually has.

    MNS095MI/Lst gives the companies, MNS100MI/LstDivisions the divisions
    under each. Run with no cono/divi on the request, so it works before we
    know what is valid - which is the whole point of it.
    """
    probe = M3Client(client.tenant, client.ionapi_path.parent,
                     company=None, division=None, m3user=client.m3user,
                     timeout=client.timeout, batch_size=client.batch_size)
    out = []
    body = probe.execute("MNS095MI", [
        {"transaction": "Lst", "record": {},
         "selectedColumns": ["CONO", "TX40"]}])
    for r in probe._records(body, "Lst"):
        cono = (r.get("CONO") or "").strip()
        if not cono:
            continue
        entry = {"cono": cono, "name": (r.get("TX40") or "").strip(),
                 "divisions": []}
        try:
            dbody = probe.execute("MNS100MI", [
                {"transaction": "LstDivisions", "record": {"CONO": cono},
                 "selectedColumns": ["CONO", "DIVI", "TX15"]}])
            for d in probe._records(dbody, "LstDivisions"):
                divi = (d.get("DIVI") or "").strip()
                if not divi: continue
                entry["divisions"].append(
                    {"divi": divi, "name": (d.get("TX15") or "").strip()})
        except M3ApiError as exc:
            log.warning("divisions for company %s: %s", cono, exc)
        out.append(entry)
    return out


def resolve_ionapi(tenant: str,
                   ionapi_dir: str | Path = DEFAULT_IONAPI_DIR) -> Path:
    """Find the .ionapi file whose 'ti' matches the tenant."""
    matches = [e for e in list_ionapi_files(ionapi_dir)
               if (e.get("tenant") or "").strip().upper() == tenant.strip().upper()]
    if not matches:
        known = ", ".join(sorted(
            e["tenant"] for e in list_ionapi_files(ionapi_dir) if e.get("tenant")
        )) or "none"
        raise M3ApiError(
            f"No .ionapi file in {ionapi_dir} has ti='{tenant}'. Tenants available: "
            f"{known}")

    if len(matches) > 1:
        log.warning("Several .ionapi files match %s (%s) - using %s",
                    tenant, ", ".join(m["file"] for m in matches),
                    matches[0]["file"])
    return Path(matches[0]["path"])


# The client itself.
# ---------------------------------------------------------------------------
class M3Client:
    """Minimal, non-interactive M3 REST client for the security routines."""

    def __init__(self,
                 tenant: str,
                 ionapi_dir: str | Path = DEFAULT_IONAPI_DIR,
                 company: str | None = DEFAULT_COMPANY,
                 division: str | None = DEFAULT_DIVISION,
                 m3user: str | None = None,
                 timeout: int = DEFAULT_TIMEOUT,
                 batch_size: int = DEFAULT_BATCH_SIZE):
        self.tenant = tenant
        # Blank counts as unset: both mean "leave it off the request".
        self.company = str(company).strip() if company not in (None, "") else None
        self.division = str(division).strip() if division not in (None, "") else None
        self.m3user = m3user
        self.timeout = timeout
        self.batch_size = max(1, int(batch_size or DEFAULT_BATCH_SIZE))

        self.ionapi_path = resolve_ionapi(tenant, ionapi_dir)
        self._cfg = json.loads(self.ionapi_path.read_text())
        for key in ("ti", "iu", "pu", "ot", "ci", "cs", "saak", "sask"):
            if self._cfg.get(key): continue
            raise M3ApiError(
                f"{self.ionapi_path.name} is missing '{key}' - it does not look like a service-account .ionapi file.")


        self._token = None
        self._token_at = 0.0
        self.session = requests.Session()

    # The URL carries the company / division, so it is rebuilt per call.
    @property
    def api_url(self) -> str:
        iu = self._cfg["iu"].rstrip("/")
        url = f"{iu}/{self._cfg['ti']}/M3/m3api-rest/v2/execute?maxrecs=0&extendedresult=true&righttrim=true"
        # Unset company / division are left off the query string entirely, so
        # M3 falls back to the service account's own defaults rather than to a
        # guessed 001 / 100 that it would answer 403 for.
        if self.company:
            url += f"&cono={self.company}"
        if self.division:
            url += f"&divi={self.division}"
        if self.m3user:
            url += f"&m3user={self.m3user}"
        return url

    def _fetch_token(self) -> None:
        cfg = self._cfg
        token_url = cfg["pu"].rstrip("/") + "/" + cfg["ot"].lstrip("/")
        payload = {
            "client_id": cfg["ci"],
            "client_secret": cfg["cs"],
            "grant_type": "password",
            "username": cfg["saak"],
            "password": cfg["sask"],
        }
        try:
            r = self.session.post(
                token_url, data=payload, timeout=self.timeout,
                headers={"Content-Type": "application/x-www-form-urlencoded"})

            r.raise_for_status()
            self._token = r.json()["access_token"]
            self._token_at = time.time()
        except requests.RequestException as exc:
            raise M3ApiError(
                f"Token request failed for {self.tenant} ("
                f"{self.ionapi_path.name}): {exc}"
            ) from exc
        except (KeyError, ValueError) as exc:
            raise M3ApiError(f"Token response had no access_token: {exc}") from exc

    def token(self, force: bool = False) -> str:
        if force or not self._token:
            self._fetch_token()
        return self._token

    # The two ways to send a request: one batch, or many batches.
    def execute(self, program: str, transactions: list[dict],
                max_retries: int = 3, retry_delay: int = 2) -> dict:
        """POST one m3api-rest v2 batch and return the parsed body."""
        payload = {"program": program, "transactions": transactions}
        last = None
        for attempt in range(1, max_retries + 1):
            try:
                r = self.session.post(
                    self.api_url, json=payload, timeout=self.timeout,
                    headers={"Authorization": f"Bearer {self.token()}",
                             "Content-Type": "application/json"})
                # An expired token is worth one silent retry.
                if r.status_code == 401:
                    self.token(force=True)
                    continue
                r.raise_for_status()
                return r.json()
            except Exception as exc:
                last = exc
                if attempt == max_retries:
                    break
                self.token(force=True)
                time.sleep(retry_delay)
        raise M3ApiError(f"{program} call failed after {max_retries} attempts: {last}")

    def execute_many(self, program: str, transaction: str,
                     records: list[dict],
                     selected_columns: list[str] | None = None,
                     batch_size: int | None = None,
                     progress=_noop_progress,
                     label: str = "") -> list[dict]:
        """
        Send many records of one transaction, up to batch_size per HTTP call.

        m3api-rest answers a batch with a 'results' array in the same order as
        the transactions that went in, so result[j] belongs to record[j].
        A short array means M3 dropped the tail - those records are reported as
        'No result returned' rather than silently counted as successes.

        Returns one entry per input record, in order:
            {"record": {...}, "error": "" | "...", "rows": [...]}

        A whole batch that fails to post is attributed to every record in it.
        Because each record is reported on its own, a partial batch is not lost.
        """
        size = max(1, int(batch_size or self.batch_size))
        out = []
        total = len(records)
        progress(0, total, label or f"{program}/{transaction}")

        for start in range(0, total, size):
            batch = records[start:start + size]
            txs = []
            for rec in batch:
                tx = {"transaction": transaction, "record": rec}
                if selected_columns:
                    tx["selectedColumns"] = selected_columns
                txs.append(tx)
            try:
                body = self.execute(program, txs)
                results = (body or {}).get("results", []) or []
                for j, rec in enumerate(batch):
                    if j < len(results):
                        res = results[j] if isinstance(results[j], dict) else {}
                        err = (res.get("errorMessage") or "").strip()
                        out.append({"record": rec, "error": err,
                                    "rows": res.get("records") or []})
                    else:
                        out.append({"record": rec, "error": "No result returned",
                                    "rows": []})
            except Exception as exc:
                for rec in batch:
                    out.append({"record": rec, "error": str(exc), "rows": []})
            progress(min(start + size, total), total,
                     label or f"{program}/{transaction}")
        return out

    @staticmethod
    def _records(body: dict, transaction: str | None = None) -> list[dict]:
        """Pull the record rows out of an m3api-rest v2 response."""
        rows = []
        for res in (body or {}).get("results", []):
            if transaction and res.get("transaction") != transaction:
                continue
            err = (res.get("errorMessage") or "").strip()
            if err:
                code = (res.get("errorCode") or "").strip()
                # "not found" is an empty answer, not a failure.
                if code in ("", "0") and "not found" not in err.lower():
                    raise M3ApiError(err)
                if "not found" in err.lower():
                    continue
                raise M3ApiError(f"{code}: {err}" if code else err)
            rows.extend(res.get("records") or [])
        return rows

    @staticmethod
    def _errors(body: dict) -> list[str]:
        return [(r.get("errorMessage") or "").strip()
                for r in (body or {}).get("results", [])
                if (r.get("errorMessage") or "").strip()]

    # ---- MNS405 role definitions -----------------------------------------
    def list_roles(self, role: str = "") -> list[dict]:
        """MNS405MI/Lst - every role definition in M3 (blank role = all)."""
        body = self.execute("MNS405MI", [
            {"transaction": "Lst",
             "record": {"ROLL": role} if role else {},
             "selectedColumns": ["ROLL", "TX40", "TX15", "TXID", "ROLT"]}])

        return self._records(body, "Lst")

    def get_role(self, role: str) -> dict | None:
        """MNS405MI/Get - one role definition, or None when it does not exist."""
        body = self.execute("MNS405MI", [
            {"transaction": "Get",
             "record": {"ROLL": role},
             "selectedColumns": ["ROLL", "TX40", "TX15", "TXID", "ROLT"]}])

        rows = self._records(body, "Get")
        return rows[0] if rows else None

    def list_role_members(self, role: str) -> list[dict]:
        """MNS410MI/Lst with ROLL only - everyone holding that role."""
        body = self.execute("MNS410MI", [
            {"transaction": "Lst",
             "record": {"ROLL": role},
             "selectedColumns": ["USID", "ROLL", "FVDT", "VTDT", "TXID"]}])
        # A blank ROLL lists everyone, so make sure the answer really is for
        # the role that was asked for.
        return [r for r in self._records(body, "Lst")
                if (r.get("ROLL") or "").strip().upper() == role.strip().upper()]

    def list_users(self) -> list[dict]:
        """MNS150MI/LstUserData - USID to email map."""
        body = self.execute("MNS150MI", [
            {"transaction": "LstUserData",
             "record": {},
             "selectedColumns": ["USID", "TX40", "EMAL", "USTA"]}])

        return self._records(body, "LstUserData")

    def list_function_roles(self, function: str = "", role: str = "") -> list[dict]:
        """
        SES400MI/Lst - function authorisations, one row per function + role.

        Blank input lists everything. Only the identifying columns are asked
        for; SES400 also carries 99 option flags and 24 function-key flags per
        row, which would make the response enormous and are not needed here.
        """
        rec = {}
        if function:
            rec["FNID"] = function
        if role:
            rec["ROLL"] = role
        body = self.execute("SES400MI", [
            {"transaction": "Lst",
             "record": rec,
             "selectedColumns": ["FNID", "ROLL", "CONO", "DIVI", "STAT", "TXID"]}])

        return self._records(body, "Lst")

    def update_function_role(self, function: str, role: str,
                             company: str = "", division: str = "",
                             status: str | None = None) -> dict:
        """
        SES400MI/Upd - change one function authorisation.

        Only the key fields and the fields being changed are sent. SES400 also
        carries 99 option flags and 24 function-key flags; leaving them out of
        an Upd leaves them as they are, which is what we want when the only
        thing changing is the status.
        """
        rec = {"FNID": function, "ROLL": role}
        cono = company or self.company
        if cono:
            rec["CONO"] = cono
        if division:
            rec["DIVI"] = division
        if status is not None:
            rec["STAT"] = status
        return self.execute("SES400MI", [{"transaction": "Upd", "record": rec}])

    def delete_function_role(self, function: str, role: str,
                             company: str = "", division: str = "") -> dict:
        """SES400MI/Dlt - remove one function authorisation."""
        rec = {"FNID": function, "ROLL": role}
        cono = company or self.company
        if cono:
            rec["CONO"] = cono
        if division:
            rec["DIVI"] = division
        return self.execute("SES400MI", [{"transaction": "Dlt", "record": rec}])

    def add_role(self, roll: str, tx40: str, tx15: str = "",
                 rolt: str | None = None) -> dict:
        """MNS405MI/Add - create a role definition in M3."""
        rec = {"ROLL": roll, "TX40": tx40}
        if tx15:
            rec["TX15"] = tx15
        if rolt:
            rec["ROLT"] = rolt
        return self.execute("MNS405MI", [{"transaction": "Add", "record": rec}])

    def add_role_member(self, usid: str, role: str,
                        valid_from: str | None = None,
                        valid_to: str | None = None) -> dict:
        """MNS410MI/Add - connect one user to one role."""
        rec = {"USID": usid, "ROLL": role}
        if valid_from:
            rec["FVDT"] = valid_from
        if valid_to:
            rec["VTDT"] = valid_to
        return self.execute("MNS410MI", [{"transaction": "Add", "record": rec}])

    def delete_role_member(self, usid: str, role: str,
                           delete_permissions: int | None = None) -> dict:
        """MNS410MI/Dlt - take one user out of one role."""
        rec = {"USID": usid, "ROLL": role}
        if delete_permissions is not None:
            rec["DLPR"] = delete_permissions
        return self.execute("MNS410MI", [{"transaction": "Dlt", "record": rec}])


# ---------------------------------------------------------------------------
# Audit trail. Every call that changes anything lands one row per record.
# ---------------------------------------------------------------------------
def _log_m3(conn: sqlite3.Connection, tenant: str, program: str, tx: str,
            outcome: str, run_id: str, role: str | None = None,
            usid: str | None = None, dry_run: bool = False,
            message: str | None = None, payload=None, response=None) -> None:
    conn.execute(
        """
        INSERT INTO M3_Security_M3Log
            (tenant, program, transaction_name, role_name, usid, dry_run,
             outcome, message, payload, response, run_id)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (tenant, program, tx, role, usid, 1 if dry_run else 0, outcome,
         (message or "")[:2000],
         json.dumps(payload)[:4000] if payload is not None else None,
         json.dumps(response)[:4000] if response is not None else None,
         run_id))



def _log_m3_many(conn: sqlite3.Connection, tenant: str, program: str, tx: str,
                 rows: list[dict], run_id: str, dry_run: bool = False) -> None:
    """
    Write one audit row per record in a single executemany.

    rows entries: {"role":…, "usid":…, "outcome":…, "message":…,
                   "payload":…, "response":…}
    """
    conn.executemany(
        """
        INSERT INTO M3_Security_M3Log
            (tenant, program, transaction_name, role_name, usid, dry_run,
             outcome, message, payload, response, run_id)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        [(tenant, program, tx, r.get("role"), r.get("usid"),
          1 if dry_run else 0, r["outcome"], (r.get("message") or "")[:2000],
          json.dumps(r["payload"])[:4000] if r.get("payload") is not None else None,
          json.dumps(r["response"])[:4000] if r.get("response") is not None else None,
          run_id)
         for r in rows])



def sync_many_role_members(conn: sqlite3.Connection, client: M3Client,
                           roles: list[str], run_id: str | None = None,
                           progress=_noop_progress) -> int:
    """
    Refresh M3_Security_M3Members for many roles, batching the MNS410MI/Lst
    calls instead of asking one role at a time.
    """
    tenant = client.tenant
    results = client.execute_many(
        "MNS410MI", "Lst", [{"ROLL": r} for r in roles],
        ["USID", "ROLL", "FVDT", "VTDT", "TXID"],
        progress=progress, label="Reading role members")

    # One result per role, in the order the roles went in.
    cur = conn.cursor()
    total = 0
    for res in results:
        role = res["record"]["ROLL"]
        if res["error"] and "not found" not in res["error"].lower():
            log.warning("members for %s: %s", role, res["error"])
            continue
        rows = [r for r in res["rows"]
                if (r.get("ROLL") or "").strip().upper() == role.strip().upper()]
        cur.execute("DELETE FROM M3_Security_M3Members WHERE tenant = ? AND role_name = ?",
                    (tenant, role))
        cur.executemany(
            "INSERT OR REPLACE INTO M3_Security_M3Members "
            "(tenant, role_name, usid, valid_from, valid_to) VALUES (?, ?, ?, ?, ?)",
            [(tenant, role, (r.get("USID") or "").strip(),
              (r.get("FVDT") or "").strip(), (r.get("VTDT") or "").strip())
             for r in rows if (r.get("USID") or "").strip()])

        cur.execute(
            "UPDATE M3_Security_Roles SET m3_member_count = ?, "
            "m3_checked_at = datetime('now') WHERE tenant = ? AND name = ?",
            (len(rows), tenant, role))

        total += len(rows)
    conn.commit()
    return total


def sync_roles(conn: sqlite3.Connection, client: M3Client,
               with_members: bool = False,
               progress=_noop_progress) -> dict:
    """
    Read every role definition from M3 and flag which of the roles captured
    from the CSV exist there.  Optionally pull each matched role's membership.
    """
    tenant = client.tenant
    run_id = str(uuid.uuid4())
    progress(0, 3, "Reading role definitions from M3")
    m3_roles = client.list_roles()
    by_name = {(r.get("ROLL") or "").strip().upper(): r for r in m3_roles}

    cur = conn.cursor()
    cur.execute(
        "UPDATE M3_Security_Roles SET in_m3 = 0, m3_description = NULL, "
        "m3_name = NULL, m3_role_type = NULL, m3_text_id = NULL, "
        "m3_member_count = NULL, m3_checked_at = datetime('now') "
        "WHERE tenant = ?",
        (tenant,))

    matched = 0
    for row in conn.execute(
            "SELECT role_key, name FROM M3_Security_Roles WHERE tenant = ?", (tenant,)
    ).fetchall():
        m = by_name.get((row["name"] or "").strip().upper())
        if not m:
            continue
        matched += 1
        cur.execute(
            "UPDATE M3_Security_Roles SET in_m3 = 1, m3_description = ?, "
            "m3_name = ?, m3_role_type = ?, m3_text_id = ?, "
            "m3_checked_at = datetime('now') WHERE role_key = ?",
            ((m.get("TX40") or "").strip(), (m.get("TX15") or "").strip(),
             (m.get("ROLT") or "").strip(), (m.get("TXID") or "").strip(),
             row["role_key"]))

    _log_m3(conn, tenant, "MNS405MI", "Lst", "ok", run_id,
            message=f"{len(m3_roles)} roles in M3, {matched} matched locally")
    conn.commit()

    members = 0
    if with_members:
        names = [r["name"] for r in conn.execute(
            "SELECT name FROM M3_Security_Roles WHERE tenant = ? AND in_m3 = 1 ORDER BY name",
            (tenant,)
        ).fetchall()]
        members = sync_many_role_members(conn, client, names, run_id, progress)

    return {
        "tenant": tenant,
        "m3_roles": len(m3_roles),
        "matched": matched,
        "unmatched": len(m3_roles) - matched,
        "members": members,
        "run_id": run_id,
    }


def sync_role_members(conn: sqlite3.Connection, client: M3Client, role: str,
                      run_id: str | None = None) -> list[dict]:
    """Refresh M3_Security_M3Members for one role and return the rows."""
    tenant = client.tenant
    rows = client.list_role_members(role)
    cur = conn.cursor()
    cur.execute("DELETE FROM M3_Security_M3Members WHERE tenant = ? AND role_name = ?",
                (tenant, role))
    cur.executemany(
        "INSERT OR REPLACE INTO M3_Security_M3Members "
        "(tenant, role_name, usid, valid_from, valid_to) VALUES (?, ?, ?, ?, ?)",
        [(tenant, role, (r.get("USID") or "").strip(),
          (r.get("FVDT") or "").strip(), (r.get("VTDT") or "").strip())
         for r in rows if (r.get("USID") or "").strip()])

    cur.execute(
        "UPDATE M3_Security_Roles SET m3_member_count = ?, "
        "m3_checked_at = datetime('now') WHERE tenant = ? AND name = ?",
        (len(rows), tenant, role))

    conn.commit()
    return rows


def sync_users(conn: sqlite3.Connection, client: M3Client) -> int:
    """Refresh the USID -> email map from M3."""
    tenant = client.tenant
    rows = client.list_users()
    cur = conn.cursor()
    cur.execute("DELETE FROM M3_Security_M3Users WHERE tenant = ?", (tenant,))
    cur.executemany(
        "INSERT OR REPLACE INTO M3_Security_M3Users "
        "(tenant, usid, name, email, status) VALUES (?, ?, ?, ?, ?)",
        [(tenant, (r.get("USID") or "").strip(), (r.get("TX40") or "").strip(),
          (r.get("EMAL") or "").strip(), (r.get("USTA") or "").strip())
         for r in rows if (r.get("USID") or "").strip()])

    conn.commit()
    return len(rows)


# M3 field widths. ROLL is the role id itself; TX40 and TX15 are the long and
# short descriptions, and a description longer than TX40 is cut rather than
# refused. They are what decides whether a captured role name can become an M3
# role at all - an Infor OS name like "IFS-ReadOnlyUser" simply does not fit.
ROLL_MAX = 10
TX40_MAX = 40
TX15_MAX = 15


def plan_role_add(name: str, description: str | None,
                  in_m3: int | None = None) -> dict:
    """
    Work out whether one captured role can be created in M3, and with what.

    A role is skipped when the name will not fit or is not a plain M3 role id:
      * longer than 10 characters  - ROLL is A10
      * contains a hyphen          - Infor OS style names, not M3 roles
      * starts with an asterisk    - M3 reserves these

    TX40 and TX15 are the description cut to the field width. A role with no
    description falls back to its own name, since TX40 is mandatory.
    """
    name = (name or "").strip()
    desc = (description or "").strip() or name

    reason = None
    if not name:
        reason = "no name"
    elif name.startswith("*"):
        reason = "starts with *"
    elif "-" in name:
        reason = "contains -"
    elif len(name) > ROLL_MAX:
        reason = f"name is {len(name)} characters, ROLL holds {ROLL_MAX}"
    elif in_m3 == 1:
        reason = "already in M3"

    return {
        "name": name,
        "eligible": reason is None,
        "reason": reason,
        "in_m3": in_m3,
        "ROLL": name,
        "TX40": desc[:TX40_MAX],
        "TX15": desc[:TX15_MAX],
        "truncated": len(desc) > TX40_MAX,
    }


def create_roles(conn: sqlite3.Connection, client: M3Client,
                 role_names: list[str],
                 dry_run: bool = True,
                 confirm: str | None = None,
                 role_type: str | None = None,
                 progress=_noop_progress) -> dict:
    """
    Create captured roles in M3 with MNS405MI/Add.

    dry_run=True (the default) sends nothing - it returns exactly what would be
    created and what would be skipped, with the reason for each skip.
    A real run needs dry_run=False AND confirm == 'CREATE'.
    """
    tenant = client.tenant
    run_id = str(uuid.uuid4())

    # Roles that only ever appear as a functional-role member have no row in
    # M3_Security_Roles, so a plain lookup would report them as unknown. They
    # are named here so the skip reason can say what is actually wrong: the
    # name is real, it just is not in the capture.
    fsr_only = {r[0] for r in conn.execute(
        "SELECT DISTINCT security_role FROM M3_Security_FunctionalRoleMembers "
        "WHERE tenant = ? AND security_role <> '' AND row_state <> 'deleted'",
        (tenant,)).fetchall()}

    # The MNS405 capture is the authority on what M3 already holds.
    in_mns405 = {r[0] for r in conn.execute(
        "SELECT roll FROM M3_Security_M3RoleDefs WHERE tenant = ?",
        (tenant,)).fetchall()}

    plan = []
    for name in role_names:
        row = conn.execute(
            "SELECT name, description, in_m3 FROM M3_Security_Roles WHERE tenant = ? AND name = ?",
            (tenant, name)
        ).fetchone()
        if not row and name in fsr_only:
            plan.append({"name": name, "eligible": False,
                         "reason": "not in the capture", "in_m3": None,
                         "ROLL": name, "TX40": "", "TX15": "", "truncated": False})
            continue
        # MNS405 wins over the stale in_m3 flag on the captured row, which is
        # only as fresh as the last Check M3.
        held = 1 if name in in_mns405 else (row["in_m3"] if row else None)
        plan.append(plan_role_add(row["name"] if row else name,
                                  row["description"] if row else None, held))

    eligible = [p for p in plan if p["eligible"]]
    skipped = [p for p in plan if not p["eligible"]]
    reasons = {}
    for p in skipped:
        reasons[p["reason"]] = reasons.get(p["reason"], 0) + 1
    totals = {"selected": len(plan), "eligible": len(eligible),
              "skipped": len(skipped), "reasons": reasons}

    if dry_run:
        _log_m3(conn, tenant, "MNS405MI", "Add", "preview", run_id, dry_run=True,
                message=f"{len(eligible)} role(s) would be created, "
                        f"{len(skipped)} skipped")
        conn.commit()
        return {"tenant": tenant, "dry_run": True, "plan": plan,
                "eligible": eligible, "skipped": skipped, "totals": totals,
                "created": [], "failed": [], "run_id": run_id,
                "message": f"{len(eligible)} role(s) would be created in M3, "
                           f"{len(skipped)} skipped."}

    if (confirm or "").strip() != "CREATE":
        raise M3ApiError("Type CREATE to confirm creating these roles in M3.")
    if not eligible:
        raise M3ApiError("None of the selected roles can be created in M3.")

    records = []
    for p in eligible:
        rec = {"ROLL": p["ROLL"], "TX40": p["TX40"]}
        if p["TX15"]:
            rec["TX15"] = p["TX15"]
        if role_type:
            rec["ROLT"] = role_type
        records.append(rec)

    results = client.execute_many("MNS405MI", "Add", records,
                                  progress=progress, label="Creating roles")

    created, failed, log_rows, ok_names, ok_defs = [], [], [], [], []
    for p, res in zip(eligible, results):
        if res["error"]:
            failed.append({"name": p["name"], "message": res["error"]})
            log_rows.append({"role": p["name"], "outcome": "error",
                             "message": res["error"], "payload": res["record"]})
            continue
        created.append(p["name"])
        ok_names.append((p["TX40"], p["TX15"], tenant, p["name"]))
        ok_defs.append((tenant, p["ROLL"], p["TX40"], p["TX15"],
                        role_type or ""))
        log_rows.append({"role": p["name"], "outcome": "ok",
                         "payload": res["record"]})

    if ok_names:
        conn.executemany(
            "UPDATE M3_Security_Roles SET in_m3 = 1, m3_description = ?, "
            "m3_name = ?, m3_member_count = 0, m3_checked_at = datetime('now') "
            "WHERE tenant = ? AND name = ?",
            ok_names)
        # Keep the MNS405 capture in step with what was just created, so the
        # Roles tab does not have to be re-read from M3 to show the new rows
        # and the next create_roles() call sees them as already held.
        conn.executemany(
            "INSERT OR REPLACE INTO M3_Security_M3RoleDefs (tenant, roll, tx40, tx15, rolt) VALUES (?, ?, ?, ?, ?)",
            ok_defs)
    _log_m3_many(conn, tenant, "MNS405MI", "Add", log_rows, run_id)
    conn.commit()

    return {"tenant": tenant, "dry_run": False, "plan": plan,
            "eligible": eligible, "skipped": skipped, "totals": totals,
            "created": created, "failed": failed, "run_id": run_id,
            "message": f"{len(created)} role(s) created in M3, "
                       f"{len(failed)} failed, {len(skipped)} skipped."}


# ---------------------------------------------------------------------------
# MNS405 - role definitions.
#
# The MNS405 tab works on its own capture (M3_Security_M3RoleDefs) rather than
# on the IFS roles table, so it shows what M3 holds rather than what the export
# said. Add and Upd share one body: the two differ only in the transaction,
# the confirm word and which fields may be left out. Dlt does more, because a
# deleted role has to be announced to IFS as well as removed from M3.
# ---------------------------------------------------------------------------

def mns405_sync(conn: sqlite3.Connection, client: M3Client,
                progress=_noop_progress) -> dict:
    """Read every role definition from M3 into M3_Security_M3RoleDefs."""
    tenant = client.tenant
    run_id = str(uuid.uuid4())
    progress(0, 2, "Reading role definitions from MNS405")
    rows = client.list_roles()
    progress(1, 2, "Storing")

    cur = conn.cursor()
    cur.execute("DELETE FROM M3_Security_M3RoleDefs WHERE tenant = ?", (tenant,))
    cur.executemany(
        "INSERT OR REPLACE INTO M3_Security_M3RoleDefs "
        "(tenant, roll, tx40, tx15, rolt, txid) VALUES (?, ?, ?, ?, ?, ?)",
        [(tenant, (r.get("ROLL") or "").strip(), (r.get("TX40") or "").strip(),
          (r.get("TX15") or "").strip(), (r.get("ROLT") or "").strip(),
          (r.get("TXID") or "").strip())
         for r in rows if (r.get("ROLL") or "").strip()])

    _log_m3(conn, tenant, "MNS405MI", "Lst", "ok", run_id,
            message=f"{len(rows)} role definition(s) read")
    conn.commit()
    progress(2, 2, "Done")
    return {"tenant": tenant, "roles": len(rows), "run_id": run_id,
            "message": f"{len(rows)} role definition(s) read from MNS405."}


def _mns405_records(rows: list[dict], for_update: bool) -> list[dict]:
    """Turn UI rows into MNS405MI records. Blank optional fields are omitted."""
    out = []
    for r in rows:
        rec = {"ROLL": (r.get("roll") or "").strip()}
        tx40 = (r.get("tx40") or "").strip()
        tx15 = (r.get("tx15") or "").strip()
        rolt = (r.get("rolt") or "").strip()
        # On an Upd an omitted field keeps its value; on an Add TX40 is required.
        if tx40 or not for_update:
            rec["TX40"] = tx40 or rec["ROLL"][:TX40_MAX]
        if tx15:
            rec["TX15"] = tx15[:TX15_MAX]
        if rolt:
            rec["ROLT"] = rolt
        out.append(rec)
    return out


def _mns405_write(conn: sqlite3.Connection, client: M3Client,
                  transaction: str, rows: list[dict],
                  dry_run: bool, confirm: str | None,
                  expect: str, progress) -> dict:
    """Shared Add / Upd body for MNS405 role definitions."""
    tenant = client.tenant
    run_id = str(uuid.uuid4())
    rows = [r for r in rows if (r.get("roll") or "").strip()]
    if not rows:
        raise M3ApiError("No roles given.")
    for r in rows:
        roll = r["roll"].strip()
        if len(roll) > ROLL_MAX:
            raise M3ApiError(
                f"'{roll}' is {len(roll)} characters; ROLL holds {ROLL_MAX}.")

    records = _mns405_records(rows, for_update=(transaction == "Upd"))
    verb = "created" if transaction == "Add" else "updated"

    if dry_run:
        _log_m3(conn, tenant, "MNS405MI", transaction, "preview", run_id,
                dry_run=True, message=f"{len(records)} role(s) would be {verb}")
        conn.commit()
        return {"tenant": tenant, "dry_run": True, "plan": records,
                "done": [], "failed": [], "run_id": run_id,
                "totals": {"roles": len(records)},
                "message": f"{len(records)} role definition(s) would be {verb}."}

    if (confirm or "").strip() != expect:
        raise M3ApiError(f"Type {expect} to confirm.")

    results = client.execute_many("MNS405MI", transaction, records,
                                  progress=progress,
                                  label=f"MNS405 {transaction}")
    done, failed, log_rows = [], [], []
    for rec, res in zip(records, results):
        row = {"role": rec["ROLL"], "payload": rec}
        if res["error"]:
            failed.append({**rec, "message": res["error"]})
            log_rows.append({**row, "outcome": "error", "message": res["error"]})
            continue
        done.append(rec)
        log_rows.append({**row, "outcome": "ok"})

    for rec in done:
        if transaction == "Add":
            conn.execute(
                "INSERT OR REPLACE INTO M3_Security_M3RoleDefs "
                "(tenant, roll, tx40, tx15, rolt) VALUES (?, ?, ?, ?, ?)",
                (tenant, rec["ROLL"], rec.get("TX40", ""),
                 rec.get("TX15", ""), rec.get("ROLT", "")))
            continue
        sets, vals = [], []
        for col, key in (("tx40", "TX40"), ("tx15", "TX15"), ("rolt", "ROLT")):
            if key in rec:
                sets.append(f"{col} = ?")
                vals.append(rec[key])
        if not sets: continue
        conn.execute(
            f"UPDATE M3_Security_M3RoleDefs SET {', '.join(sets)}"
            ", checked_at = datetime('now') WHERE tenant = ? AND roll = ?",
            vals + [tenant, rec["ROLL"]])
    _log_m3_many(conn, tenant, "MNS405MI", transaction, log_rows, run_id)
    conn.commit()

    return {"tenant": tenant, "dry_run": False, "plan": records,
            "done": done, "failed": failed, "run_id": run_id,
            "totals": {"roles": len(records)},
            "message": f"{len(done)} role definition(s) {verb}, {len(failed)} failed."}


def mns405_add(conn, client, rows, dry_run=True, confirm=None,
               progress=_noop_progress) -> dict:
    """MNS405MI/Add - create role definitions."""
    return _mns405_write(conn, client, "Add", rows, dry_run, confirm,
                         "CREATE", progress)


def mns405_update(conn, client, rows, dry_run=True, confirm=None,
                  progress=_noop_progress) -> dict:
    """MNS405MI/Upd - change description, name or role type."""
    return _mns405_write(conn, client, "Upd", rows, dry_run, confirm,
                         "UPDATE", progress)


def mns405_delete(conn: sqlite3.Connection, client: M3Client, keys: list[int],
                  dry_run: bool = True,
                  confirm: str | None = None,
                  emit_bods: bool = True,
                  strip_from_users: bool = True,
                  bod_dir=None,
                  progress=_noop_progress) -> dict:
    """
    MNS405MI/Dlt - remove role definitions. Keys are role_def_key values.

    Deleting a security role has to be announced with a SyncSecurityRoleMaster
    event carrying Operation DELETE, the same document M3 publishes from the
    CMNROL smart rule. One is written per role that M3 actually deleted; roles
    M3 refused do not get one.

    A deleted role is also stripped off the captured IFS users, so the next
    users export carries them without it. Push that export into IFS and nobody
    holds the role any more, which is what lets the BOD delete it.
    """
    tenant = client.tenant
    run_id = str(uuid.uuid4())
    if not keys:
        raise M3ApiError("No roles selected.")
    marks = ",".join("?" * len(keys))
    rows = conn.execute(
        "SELECT role_def_key, roll, tx40, tx15, rolt FROM M3_Security_M3RoleDefs "
        f"WHERE tenant = ? AND role_def_key IN ({marks}) ORDER BY roll",
        [tenant] + list(keys)).fetchall()
    plan = [dict(r) for r in rows]

    bod_cfg, bod_error = None, ""
    if emit_bods:
        try:
            bod_cfg = tenant_bod_config(tenant)
        except BodConfigError as exc:
            bod_error = str(exc)

    if dry_run:
        # How many captured IFS users still hold each role.
        for p in plan:
            p["ifs_users"] = conn.execute(
                "SELECT COUNT(*) FROM M3_Security_UserRoles WHERE tenant = ? "
                "AND role_name = ? AND row_state <> 'deleted'",
                (tenant, p["roll"])).fetchone()[0]
        n_users = sum(p["ifs_users"] for p in plan)
        _log_m3(conn, tenant, "MNS405MI", "Dlt", "preview", run_id, dry_run=True,
                message=f"{len(plan)} role definition(s) would be deleted")
        conn.commit()
        return {"tenant": tenant, "dry_run": True, "plan": plan,
                "done": [], "failed": [], "bods": [], "run_id": run_id,
                "bod_config": bod_cfg, "bod_error": bod_error,
                "totals": {"roles": len(plan), "ifs_users": n_users},
                "message": f"{len(plan)} role definition(s) would be deleted from MNS405, taken off "
                           f"{n_users} captured IFS user record(s), and written out as a SyncSecurityRoleMaster event each."}

    # A real run writes BODs, so refuse before touching M3 if the BOD settings
    # for this tenant are not usable - half a deletion is worse than none.
    if (confirm or "").strip() != "DELETE":
        raise M3ApiError("Type DELETE to confirm.")
    if emit_bods and bod_error:
        # No point deleting in M3 if the event cannot be written afterwards.
        raise M3ApiError(bod_error)

    results = client.execute_many("MNS405MI", "Dlt",
                                  [{"ROLL": p["roll"]} for p in plan],
                                  progress=progress, label="MNS405 Dlt")
    done, failed, log_rows, gone = [], [], [], []
    for p, res in zip(plan, results):
        row = {"role": p["roll"], "payload": res["record"]}
        if res["error"]:
            failed.append({**p, "message": res["error"]})
            log_rows.append({**row, "outcome": "error", "message": res["error"]})
            continue
        done.append(p)
        gone.append((p["role_def_key"],))
        log_rows.append({**row, "outcome": "ok"})
    if gone:
        conn.executemany(
            "DELETE FROM M3_Security_M3RoleDefs WHERE role_def_key = ?", gone)
    _log_m3_many(conn, tenant, "MNS405MI", "Dlt", log_rows, run_id)

    stripped = {"users": 0, "user_links": 0, "assignments": 0}
    if strip_from_users and done:
        progress(0, 1, "Taking the role off the captured IFS users")
        for p in done:
            res = strip_role_members(conn, tenant, p["roll"], commit=False)
            for k in stripped:
                stripped[k] += res.get(k, 0)
        progress(1, 1, "Taking the role off the captured IFS users")

    bods = []
    if emit_bods and done:
        progress(0, 1, "Writing SyncSecurityRoleMaster events")
        kwargs = {"out_dir": bod_dir} if bod_dir else {}

        bods = build_and_write(
            [{"roll": p["roll"], "tx15": p.get("tx15") or "",
              "tx40": p.get("tx40") or ""} for p in done], tenant, **kwargs)
        _log_m3_many(conn, tenant, "SecurityRoleMaster", "Sync",
                     [{"role": b["role"],
                       "outcome": "error" if b["error"] else "ok",
                       "message": b["error"] or b.get("file", ""),
                       "payload": {"EventId": b.get("event_id", ""),
                                   "Sequence": b.get("sequence", ""),
                                   "keyValue": b.get("key_value", "")}}
                      for b in bods], run_id)
        progress(1, 1, "Writing SyncSecurityRoleMaster events")
    conn.commit()

    n_bod = sum(1 for b in bods if not b["error"])
    return {"tenant": tenant, "dry_run": False, "plan": plan,
            "done": done, "failed": failed, "bods": bods, "run_id": run_id,
            "bod_config": bod_cfg, "bod_error": bod_error,
            "stripped": stripped,
            "totals": {"roles": len(plan), "ifs_users": stripped["user_links"]},
            "message": f"{len(done)} role definition(s) deleted, "
                       f"{len(failed)} failed, {n_bod} event(s) written, taken off "
                       f"{stripped['user_links']} captured IFS user record(s) — export the users next."}


# ---------------------------------------------------------------------------
# MNS410 - role per user. Same shape as MNS405: one sync into a local capture,
# a shared Add / Upd body and a Dlt. The key is the pair (ROLL, USID).
# ---------------------------------------------------------------------------
def mns410_sync(conn: sqlite3.Connection, client: M3Client,
                role: str = "", user: str = "",
                progress=_noop_progress) -> dict:
    """
    Read role-per-user rows into M3_Security_M3Members.

    Blank role and user list everything in one call; either one narrows it.
    """
    tenant = client.tenant
    run_id = str(uuid.uuid4())
    progress(0, 2, "Reading roles per user from MNS410")
    rec = {}
    if role:
        rec["ROLL"] = role
    if user:
        rec["USID"] = user
    body = client.execute("MNS410MI", [
        {"transaction": "Lst", "record": rec,
         "selectedColumns": ["USID", "ROLL", "FVDT", "VTDT", "TXID"]}])

    rows = client._records(body, "Lst")
    progress(1, 2, "Storing")

    cur = conn.cursor()
    if role or user:
        if role:
            cur.execute("DELETE FROM M3_Security_M3Members WHERE tenant = ? AND role_name = ?",
                        (tenant, role))
        else:
            cur.execute("DELETE FROM M3_Security_M3Members WHERE tenant = ? AND usid = ?",
                        (tenant, user))
    else:
        cur.execute("DELETE FROM M3_Security_M3Members WHERE tenant = ?", (tenant,))
    cur.executemany(
        "INSERT OR REPLACE INTO M3_Security_M3Members "
        "(tenant, role_name, usid, valid_from, valid_to) VALUES (?, ?, ?, ?, ?)",
        [(tenant, (r.get("ROLL") or "").strip(), (r.get("USID") or "").strip(),
          (r.get("FVDT") or "").strip(), (r.get("VTDT") or "").strip())
         for r in rows
         if (r.get("ROLL") or "").strip() and (r.get("USID") or "").strip()])

    _log_m3(conn, tenant, "MNS410MI", "Lst", "ok", run_id,
            message=f"{len(rows)} role-per-user row(s) read")
    conn.commit()
    progress(2, 2, "Done")
    return {"tenant": tenant, "rows": len(rows),
            "roles": len({(r.get("ROLL") or "").strip() for r in rows}),
            "users": len({(r.get("USID") or "").strip() for r in rows}),
            "run_id": run_id,
            "message": f"{len(rows)} role-per-user row(s) read from MNS410."}


def _mns410_write(conn: sqlite3.Connection, client: M3Client,
                  transaction: str, rows: list[dict],
                  dry_run: bool, confirm: str | None,
                  expect: str, progress) -> dict:
    """Shared Add / Upd body for role per user."""
    tenant = client.tenant
    run_id = str(uuid.uuid4())
    records = []
    for r in rows:
        usid = (r.get("usid") or "").strip()
        roll = (r.get("roll") or "").strip()
        if not usid or not roll:
            raise M3ApiError("Both a user and a role are required.")
        rec = {"USID": usid, "ROLL": roll}
        for key, col in (("FVDT", "valid_from"), ("VTDT", "valid_to")):
            val = (r.get(col) or "").strip()
            if not val: continue
            rec[key] = val
        records.append(rec)
    if not records:
        raise M3ApiError("Nothing to send.")
    verb = "added" if transaction == "Add" else "updated"

    if dry_run:
        _log_m3(conn, tenant, "MNS410MI", transaction, "preview", run_id,
                dry_run=True, message=f"{len(records)} row(s) would be {verb}")
        conn.commit()
        return {"tenant": tenant, "dry_run": True, "plan": records,
                "done": [], "failed": [], "run_id": run_id,
                "totals": {"rows": len(records)},
                "message": f"{len(records)} role-per-user row(s) would be {verb}."}

    if (confirm or "").strip() != expect:
        raise M3ApiError(f"Type {expect} to confirm.")

    results = client.execute_many("MNS410MI", transaction, records,
                                  progress=progress,
                                  label=f"MNS410 {transaction}")
    done, failed, log_rows = [], [], []
    for rec, res in zip(records, results):
        row = {"role": rec["ROLL"], "usid": rec["USID"], "payload": rec}
        if res["error"]:
            failed.append({**rec, "message": res["error"]})
            log_rows.append({**row, "outcome": "error", "message": res["error"]})
            continue
        done.append(rec)
        log_rows.append({**row, "outcome": "ok"})
        if transaction == "Add":
            conn.execute(
                "INSERT OR REPLACE INTO M3_Security_M3Members "
                "(tenant, role_name, usid, valid_from, valid_to) "
                "VALUES (?, ?, ?, ?, ?)",
                (tenant, rec["ROLL"], rec["USID"],
                 rec.get("FVDT", ""), rec.get("VTDT", "")))
            continue
        # An Upd only ever changes the two dates, so the local row is patched
        # with whatever was actually sent rather than rewritten wholesale.

        sets, vals = [], []
        for col, key in (("valid_from", "FVDT"), ("valid_to", "VTDT")):
            if key in rec:
                sets.append(f"{col} = ?")
                vals.append(rec[key])
        if not sets: continue
        conn.execute(
            f"UPDATE M3_Security_M3Members SET {', '.join(sets)}"
            ", checked_at = datetime('now') "
            "WHERE tenant = ? AND role_name = ? AND usid = ?",
            vals + [tenant, rec["ROLL"], rec["USID"]])
    _log_m3_many(conn, tenant, "MNS410MI", transaction, log_rows, run_id)
    conn.commit()
    return {"tenant": tenant, "dry_run": False, "plan": records,
            "done": done, "failed": failed, "run_id": run_id,
            "totals": {"rows": len(records)},
            "message": f"{len(done)} row(s) {verb}, {len(failed)} failed."}


def mns410_add(conn, client, rows, dry_run=True, confirm=None,
               progress=_noop_progress) -> dict:
    """MNS410MI/Add - give a user a role."""
    return _mns410_write(conn, client, "Add", rows, dry_run, confirm,
                         "CREATE", progress)


def mns410_update(conn, client, rows, dry_run=True, confirm=None,
                  progress=_noop_progress) -> dict:
    """MNS410MI/Upd - change the valid-from / valid-to dates."""
    return _mns410_write(conn, client, "Upd", rows, dry_run, confirm,
                         "UPDATE", progress)


def mns410_delete(conn: sqlite3.Connection, client: M3Client, keys: list[int],
                  dry_run: bool = True,
                  confirm: str | None = None,
                  delete_permissions: int | None = None,
                  progress=_noop_progress) -> dict:
    """MNS410MI/Dlt - take users out of roles. Keys are m3_member_key values."""
    tenant = client.tenant
    run_id = str(uuid.uuid4())
    if not keys:
        raise M3ApiError("No rows selected.")
    marks = ",".join("?" * len(keys))
    rows = conn.execute(
        "SELECT m3_member_key, role_name, usid, valid_from, valid_to "
        "FROM M3_Security_M3Members WHERE tenant = ? AND m3_member_key IN ("
        f"{marks}) ORDER BY role_name, usid",
        [tenant] + list(keys)).fetchall()
    plan = [dict(r) for r in rows]

    if dry_run:
        _log_m3(conn, tenant, "MNS410MI", "Dlt", "preview", run_id, dry_run=True,
                message=f"{len(plan)} row(s) would be deleted")
        conn.commit()
        return {"tenant": tenant, "dry_run": True, "plan": plan,
                "done": [], "failed": [], "run_id": run_id,
                "totals": {"rows": len(plan),
                           "roles": len({p["role_name"] for p in plan}),
                           "users": len({p["usid"] for p in plan})},
                "message": f"{len(plan)} role-per-user row(s) would be deleted from MNS410."}

    # DLPR asks M3 to drop the user's permissions along with the role; left
    # unset, M3 keeps whatever it already had for that user.
    if (confirm or "").strip() != "DELETE":
        raise M3ApiError("Type DELETE to confirm.")

    records = []
    for p in plan:
        rec = {"USID": p["usid"], "ROLL": p["role_name"]}
        if delete_permissions is not None:
            rec["DLPR"] = delete_permissions
        records.append(rec)
    results = client.execute_many("MNS410MI", "Dlt", records,
                                  progress=progress, label="MNS410 Dlt")
    done, failed, log_rows, gone = [], [], [], []
    for p, res in zip(plan, results):
        row = {"role": p["role_name"], "usid": p["usid"], "payload": res["record"]}
        if res["error"]:
            failed.append({**p, "message": res["error"]})
            log_rows.append({**row, "outcome": "error", "message": res["error"]})
            continue
        done.append(p)
        gone.append((p["m3_member_key"],))
        log_rows.append({**row, "outcome": "ok"})
    if gone:
        conn.executemany(
            "DELETE FROM M3_Security_M3Members WHERE m3_member_key = ?", gone)
    _log_m3_many(conn, tenant, "MNS410MI", "Dlt", log_rows, run_id)
    conn.commit()
    return {"tenant": tenant, "dry_run": False, "plan": plan,
            "done": done, "failed": failed, "run_id": run_id,
            "totals": {"rows": len(plan)},
            "message": f"{len(done)} row(s) deleted, {len(failed)} failed."}


# SES400 - function authorisations, one row per function + role + company.
# ---------------------------------------------------------------------------
def sync_function_roles(conn: sqlite3.Connection, client: M3Client,
                        progress=_noop_progress) -> dict:
    """
    Pull every function authorisation from SES400 and work out which of the
    roles it references are missing from MNS405.

    MNS405 is read in the same pass, so this does not depend on the Security
    Roles tab having been checked first.
    """
    tenant = client.tenant
    run_id = str(uuid.uuid4())

    progress(0, 3, "Reading function authorisations from SES400")
    rows = client.list_function_roles()
    progress(1, 3, "Reading role definitions from MNS405")
    mns405 = {(r.get("ROLL") or "").strip().upper()
              for r in client.list_roles()}
    progress(2, 3, "Storing")

    cur = conn.cursor()
    cur.execute("DELETE FROM M3_Security_FunctionRoles WHERE tenant = ?", (tenant,))
    payload, roles = [], set()
    for r in rows:
        fnid = (r.get("FNID") or "").strip()
        roll = (r.get("ROLL") or "").strip()
        if not fnid or not roll:
            continue
        roles.add(roll)
        payload.append((tenant, fnid, roll,
                        (r.get("CONO") or "").strip(),
                        (r.get("DIVI") or "").strip(),
                        (r.get("STAT") or "").strip(),
                        (r.get("TXID") or "").strip()))
    cur.executemany(
        "INSERT OR REPLACE INTO M3_Security_FunctionRoles "
        "(tenant, fnid, roll, cono, divi, stat, txid) VALUES (?, ?, ?, ?, ?, ?, ?)",
        payload)

    # Rebuilt from scratch, so a role that lost its last authorisation goes.
    cur.execute("DELETE FROM M3_Security_FunctionRoleStatus WHERE tenant = ?",
                (tenant,))
    cur.executemany(
        "INSERT OR REPLACE INTO M3_Security_FunctionRoleStatus "
        "(tenant, roll, in_mns405, checked_at) VALUES (?, ?, ?, datetime('now'))",
        [(tenant, roll, 1 if roll.upper() in mns405 else 0)
         for roll in sorted(roles)])

    missing = sorted(r for r in roles if r.upper() not in mns405)

    _log_m3(conn, tenant, "SES400MI", "Lst", "ok", run_id,
            message=f"{len(payload)} authorisation(s), "
                    f"{len(roles)} role(s), {len(missing)} missing from MNS405")
    conn.commit()
    progress(3, 3, "Done")

    return {
        "tenant": tenant,
        "authorisations": len(payload),
        "functions": len({p[1] for p in payload}),
        "roles": len(roles),
        "in_mns405": len(roles) - len(missing),
        "missing": missing,
        "missing_count": len(missing),
        "run_id": run_id,
    }


def _refresh_function_role_status(conn: sqlite3.Connection, tenant: str) -> None:
    """Drop status rows for roles that no longer appear in any authorisation."""
    conn.execute(
        """
        DELETE FROM M3_Security_FunctionRoleStatus
         WHERE tenant = ?
           AND roll NOT IN (SELECT DISTINCT roll FROM M3_Security_FunctionRoles
                             WHERE tenant = ?)
        """,
        (tenant, tenant))


# ---------------------------------------------------------------------------
def update_function_role_status(conn: sqlite3.Connection, client: M3Client,
                                keys: list[int],
                                status: str,
                                dry_run: bool = True,
                                confirm: str | None = None,
                                progress=_noop_progress) -> dict:
    """
    Change the status of function authorisations with SES400MI/Upd.

    Only the key fields and STAT are sent, so the option and function-key flags
    on each authorisation are left exactly as they are.

    Rows already at the requested status are reported and not sent - there is
    no point spending a call to write a value that is already there.

    dry_run=True (the default) sends nothing. A real run needs dry_run=False
    AND confirm == 'UPDATE'.
    """
    tenant = client.tenant
    run_id = str(uuid.uuid4())
    status = (status or "").strip()
    if not status:
        raise M3ApiError("A status is required.")
    if status not in FUNCTION_STATUS:
        valid = ", ".join(f"{k} ({v})" for k, v in FUNCTION_STATUS.items())
        raise M3ApiError(f"Status must be one of {valid}; got '{status}'.")
    if not keys:
        raise M3ApiError("No authorisations selected.")

    marks = ",".join("?" * len(keys))
    rows = conn.execute(
        f"""
        SELECT function_role_key, fnid, roll, cono, divi, stat
        FROM M3_Security_FunctionRoles
        WHERE tenant = ? AND function_role_key IN ({marks})
        ORDER BY fnid, roll, cono, divi
        """,
        [tenant] + list(keys)
    ).fetchall()

    todo, unchanged = [], []
    for r in rows:
        entry = {**dict(r), "new_stat": status}
        if (r["stat"] or "").strip() == status:
            unchanged.append(entry)
            continue
        todo.append(entry)

    totals = {
        "selected": len(rows),
        "to_change": len(todo),
        "already": len(unchanged),
        "functions": len({t["fnid"] for t in todo}),
        "roles": len({t["roll"] for t in todo}),
        "from": sorted({(t["stat"] or "").strip() or "(blank)" for t in todo}),
        "to": status,
    }

    if dry_run:
        _log_m3(conn, tenant, "SES400MI", "Upd", "preview", run_id, dry_run=True,
                message=f"{len(todo)} authorisation(s) would move to "
                        f"{function_status_label(status)}")
        conn.commit()
        return {"tenant": tenant, "dry_run": True, "new_status": status,
                "plan": todo, "unchanged": unchanged, "totals": totals,
                "updated": [], "failed": [], "run_id": run_id,
                "message": f"{len(todo)} authorisation(s) would move to "
                           f"{function_status_label(status)}"
                           + (f", {len(unchanged)} already there."
                              if unchanged else ".")}

    if (confirm or "").strip() != "UPDATE":
        raise M3ApiError("Type UPDATE to confirm changing the status.")
    if not todo:
        raise M3ApiError("Every selected authorisation is already at "
                         f"{function_status_label(status)}.")

    records = []
    for t in todo:
        rec = {"FNID": t["fnid"], "ROLL": t["roll"], "STAT": status}
        # Each row keeps the company and division it was read with, so an
        # authorisation is updated where it actually lives.
        if t["cono"] or client.company:
            rec["CONO"] = t["cono"] or client.company
        if t["divi"]:
            rec["DIVI"] = t["divi"]
        records.append(rec)

    results = client.execute_many("SES400MI", "Upd", records,
                                  progress=progress,
                                  label=f"Setting status {status}")

    updated, failed, log_rows, ok = [], [], [], []
    for t, res in zip(todo, results):
        row = {"role": t["roll"], "payload": res["record"]}
        if res["error"]:
            failed.append({**t, "message": res["error"]})
            log_rows.append({**row, "outcome": "error", "message": res["error"]})
            continue
        updated.append(t)
        ok.append((status, t["function_role_key"]))
        log_rows.append({**row, "outcome": "ok"})

    if ok:
        conn.executemany(
            "UPDATE M3_Security_FunctionRoles SET stat = ?, checked_at = datetime('now') WHERE function_role_key = ?",
            ok)
    _log_m3_many(conn, tenant, "SES400MI", "Upd", log_rows, run_id)
    conn.commit()

    return {"tenant": tenant, "dry_run": False, "new_status": status,
            "plan": todo, "unchanged": unchanged, "totals": totals,
            "updated": updated, "failed": failed, "run_id": run_id,
            "message": f"{len(updated)} authorisation(s) moved to "
                       f"{function_status_label(status)}, {len(failed)} failed"
                       + (f", {len(unchanged)} already there." if unchanged else ".")}


def delete_function_roles(conn: sqlite3.Connection, client: M3Client,
                          keys: list[int],
                          dry_run: bool = True,
                          confirm: str | None = None,
                          progress=_noop_progress) -> dict:
    """
    Remove function authorisations from SES400 with SES400MI/Dlt.

    keys are M3_Security_FunctionRoles.function_role_key values, so each delete
    targets exactly the row that was read from M3 - function, role, company and
    division together.

    dry_run=True (the default) sends nothing. A real run needs dry_run=False
    AND confirm == 'DELETE'. Rows M3 accepts are removed from the local copy;
    rows it refuses are left in place and reported, so the tab keeps showing
    what is really still there.
    """
    tenant = client.tenant
    run_id = str(uuid.uuid4())
    if not keys:
        raise M3ApiError("No authorisations selected.")

    marks = ",".join("?" * len(keys))
    rows = conn.execute(
        f"""
        SELECT function_role_key, fnid, roll, cono, divi, stat
        FROM M3_Security_FunctionRoles
        WHERE tenant = ? AND function_role_key IN ({marks})
        ORDER BY fnid, roll, cono, divi
        """,
        [tenant] + list(keys)
    ).fetchall()
    plan = [dict(r) for r in rows]
    totals = {
        "authorisations": len(plan),
        "functions": len({p["fnid"] for p in plan}),
        "roles": len({p["roll"] for p in plan}),
    }

    if dry_run:
        _log_m3(conn, tenant, "SES400MI", "Dlt", "preview", run_id, dry_run=True,
                message=f"{len(plan)} authorisation(s) would be deleted")
        conn.commit()
        return {"tenant": tenant, "dry_run": True, "plan": plan,
                "totals": totals, "deleted": [], "failed": [], "run_id": run_id,
                "message": f"{len(plan)} authorisation(s) would be deleted from SES400, across "
                           f"{totals['functions']} function(s) and "
                           f"{totals['roles']} role(s)."}

    if (confirm or "").strip() != "DELETE":
        raise M3ApiError("Type DELETE to confirm removing these authorisations.")

    records = []
    for p in plan:
        rec = {"FNID": p["fnid"], "ROLL": p["roll"]}
        if p["cono"] or client.company:
            rec["CONO"] = p["cono"] or client.company
        if p["divi"]:
            rec["DIVI"] = p["divi"]
        records.append(rec)

    results = client.execute_many("SES400MI", "Dlt", records,
                                  progress=progress,
                                  label="Deleting authorisations")

    deleted, failed, log_rows, gone = [], [], [], []
    for p, res in zip(plan, results):
        row = {"role": p["roll"], "payload": res["record"]}
        if res["error"]:
            failed.append({**p, "message": res["error"]})
            log_rows.append({**row, "outcome": "error", "message": res["error"]})
            continue
        deleted.append(p)
        gone.append((p["function_role_key"],))
        log_rows.append({**row, "outcome": "ok"})

    if gone:
        conn.executemany(
            "DELETE FROM M3_Security_FunctionRoles WHERE function_role_key = ?",
            gone)
        _refresh_function_role_status(conn, tenant)
    _log_m3_many(conn, tenant, "SES400MI", "Dlt", log_rows, run_id)
    conn.commit()

    return {"tenant": tenant, "dry_run": False, "plan": plan, "totals": totals,
            "deleted": deleted, "failed": failed, "run_id": run_id,
            "message": f"{len(deleted)} authorisation(s) deleted from SES400, "
                       f"{len(failed)} failed."}


def create_missing_function_roles(conn: sqlite3.Connection, client: M3Client,
                                  roles: list[str] | None = None,
                                  dry_run: bool = True,
                                  confirm: str | None = None,
                                  role_type: str | None = None,
                                  progress=_noop_progress) -> dict:
    """
    Create the SES400 roles that MNS405 does not have, via MNS405MI/Add.

    SES400 carries no description, so TX40 and TX15 are both the role name
    (TX15 cut to its 15-character field). These names already exist in M3 as
    valid role ids, so no name filtering is applied - anything M3 refuses is
    reported per role.
    """
    tenant = client.tenant
    run_id = str(uuid.uuid4())

    if roles:
        names = [r.strip() for r in roles if r and r.strip()]
    else:
        names = [r[0] for r in conn.execute(
            "SELECT roll FROM M3_Security_FunctionRoleStatus WHERE tenant = ? "
            "AND COALESCE(in_mns405, 0) = 0 ORDER BY roll",
            (tenant,)).fetchall()]

    plan = [{"name": n,
             "ROLL": n,
             "TX40": n[:TX40_MAX],
             "TX15": n[:TX15_MAX],
             "functions": conn.execute(
                 "SELECT COUNT(DISTINCT fnid) FROM M3_Security_FunctionRoles "
                 "WHERE tenant = ? AND roll = ?",
                 (tenant, n)).fetchone()[0]}
            for n in names]

    if dry_run:
        _log_m3(conn, tenant, "MNS405MI", "Add", "preview", run_id,
                dry_run=True,
                message=f"{len(plan)} SES400 role(s) would be added to MNS405")
        conn.commit()
        return {
            "tenant": tenant,
            "dry_run": True,
            "plan": plan,
            "created": [],
            "failed": [],
            "run_id": run_id,
            "totals": {"roles": len(plan)},
            "message": f"{len(plan)} role(s) from SES400 would be added to "
                       f"MNS405, named and described as the role itself.",
        }

    if (confirm or "").strip() != "CREATE":
        raise M3ApiError("Type CREATE to confirm adding these roles to MNS405.")
    if not plan:
        raise M3ApiError("No missing roles to add.")

    records = []
    for p in plan:
        rec = {"ROLL": p["ROLL"], "TX40": p["TX40"], "TX15": p["TX15"]}
        if role_type:
            rec["ROLT"] = role_type
        records.append(rec)

    results = client.execute_many("MNS405MI", "Add", records,
                                  progress=progress,
                                  label="Adding roles to MNS405")

    created, failed, log_rows, ok = [], [], [], []
    for p, res in zip(plan, results):
        if res["error"]:
            failed.append({"name": p["name"], "message": res["error"]})
            log_rows.append({"role": p["name"], "outcome": "error",
                             "message": res["error"], "payload": res["record"]})
            continue
        created.append(p["name"])
        ok.append((tenant, p["name"]))
        log_rows.append({"role": p["name"], "outcome": "ok",
                         "payload": res["record"]})

    if ok:
        conn.executemany(
            "UPDATE M3_Security_FunctionRoleStatus SET in_mns405 = 1, "
            "checked_at = datetime('now') WHERE tenant = ? AND roll = ?", ok)
        # The role now exists in M3, so the capture's own M3 columns follow.
        conn.executemany(
            "UPDATE M3_Security_Roles SET in_m3 = 1, m3_description = ?, "
            "m3_name = ?, m3_checked_at = datetime('now') "
            "WHERE tenant = ? AND name = ?",
            [(n[:TX40_MAX], n[:TX15_MAX], t, n) for t, n in ok])

    _log_m3_many(conn, tenant, "MNS405MI", "Add", log_rows, run_id)
    conn.commit()

    return {
        "tenant": tenant,
        "dry_run": False,
        "plan": plan,
        "created": created,
        "failed": failed,
        "run_id": run_id,
        "totals": {"roles": len(plan)},
        "message": f"{len(created)} role(s) added to MNS405, "
                   f"{len(failed)} failed.",
    }


def _usid_by_email(conn: sqlite3.Connection, tenant: str) -> dict[str, str]:
    """Lower-cased email -> USID, from the M3 user list (MNS150MI)."""
    return {
        (r["email"] or "").strip().lower(): (r["usid"] or "").strip()
        for r in conn.execute(
            "SELECT usid, email FROM M3_Security_M3Users "
            "WHERE tenant = ? AND COALESCE(email, '') <> ''",
            (tenant,)).fetchall()
    }


def plan_member_add(conn: sqlite3.Connection, tenant: str,
                    role_names: list[str]) -> tuple[list[dict], list[dict], dict]:
    """
    Work out the MNS410MI/Add calls needed to give the captured members of each
    role their role in M3.

    The USID comes from matching the assignment's email address against the M3
    user list. An address M3 does not know is reported rather than guessed at.
    """
    emails = _usid_by_email(conn, tenant)
    todo = []
    skipped = []

    for name in role_names:
        role = conn.execute(
            "SELECT role_key, name, in_m3 FROM M3_Security_Roles "
            "WHERE tenant = ? AND name = ?", (tenant, name)).fetchone()
        if not role:
            skipped.append({"role": name, "email": "", "usid": "",
                            "reason": "role not in the capture"})
            continue
        if role["in_m3"] != 1:
            skipped.append({"role": name, "email": "", "usid": "",
                            "reason": "role is not in M3 - create it first"})
            continue

        held = {r[0].strip().upper() for r in conn.execute(
            "SELECT usid FROM M3_Security_M3Members "
            "WHERE tenant = ? AND role_name = ?", (tenant, name)).fetchall()}
        members = [r[0] for r in conn.execute(
            "SELECT email_id FROM M3_Security_RoleAssignments "
            "WHERE tenant = ? AND role_name = ? AND email_id <> '' "
            "AND row_state <> 'deleted' ORDER BY email_id",
            (tenant, name)).fetchall()]
        if not members:
            skipped.append({"role": name, "email": "", "usid": "",
                            "reason": "no members in the capture"})
            continue

        for email in members:
            usid = emails.get(email.strip().lower())
            if not usid:
                skipped.append({"role": name, "email": email, "usid": "",
                                "reason": "no M3 user with that email"})
                continue
            if usid.upper() in held:
                skipped.append({"role": name, "email": email, "usid": usid,
                                "reason": "already holds the role in M3"})
                continue
            todo.append({"role": name, "email": email, "usid": usid})

    reasons = {}
    for s in skipped:
        reasons[s["reason"]] = reasons.get(s["reason"], 0) + 1

    totals = {
        "roles": len(role_names),
        "roles_to_touch": len({t["role"] for t in todo}),
        "adds": len(todo),
        "skipped": len(skipped),
        "unresolved_emails": len({s["email"] for s in skipped
                                  if s["reason"] == "no M3 user with that email"}),
        "reasons": reasons,
    }
    return (todo, skipped, totals)


def add_role_members(conn: sqlite3.Connection, client: M3Client,
                     role_names: list[str], dry_run: bool = True,
                     confirm: str | None = None,
                     progress=_noop_progress) -> dict:
    """
    Give the captured members of each role that role in M3, via MNS410MI/Add.

    USID is resolved from the member's email address using the M3 user list, so
    run --sync-users (or Check M3) first. dry_run=True sends nothing.
    A real run needs dry_run=False AND confirm == 'ADD'.
    """
    tenant = client.tenant
    run_id = str(uuid.uuid4())

    todo, skipped, totals = plan_member_add(conn, tenant, role_names)

    if dry_run:
        _log_m3(conn, tenant, "MNS410MI", "Add", "preview", run_id,
                dry_run=True,
                message=f"{len(todo)} member(s) would be added across "
                        f"{totals['roles_to_touch']} role(s), "
                        f"{len(skipped)} skipped")
        conn.commit()
        return {
            "tenant": tenant,
            "dry_run": True,
            "todo": todo,
            "skipped": skipped,
            "totals": totals,
            "added": [],
            "failed": [],
            "run_id": run_id,
            "message": f"{len(todo)} member(s) would be added across "
                       f"{totals['roles_to_touch']} role(s), "
                       f"{len(skipped)} skipped.",
        }

    if (confirm or "").strip() != "ADD":
        raise M3ApiError("Type ADD to confirm adding these members in M3.")
    if not todo:
        raise M3ApiError("Nothing to add - every member is already in M3 or "
                         "could not be matched to a USID.")

    results = client.execute_many(
        "MNS410MI", "Add",
        [{"USID": t["usid"], "ROLL": t["role"]} for t in todo],
        progress=progress, label="Adding role members")

    added, failed, log_rows = [], [], []
    for t, res in zip(todo, results):
        row = {"role": t["role"], "usid": t["usid"], "payload": res["record"]}
        if res["error"]:
            failed.append({**t, "message": res["error"]})
            log_rows.append({**row, "outcome": "error", "message": res["error"]})
            continue
        added.append(t)
        log_rows.append({**row, "outcome": "ok"})

    _log_m3_many(conn, tenant, "MNS410MI", "Add", log_rows, run_id)
    conn.commit()

    # Re-read the roles we touched so the local M3 picture matches.
    sync_many_role_members(conn, client, sorted({t["role"] for t in todo}),
                           run_id, progress)

    return {
        "tenant": tenant,
        "dry_run": False,
        "todo": todo,
        "skipped": skipped,
        "totals": totals,
        "added": added,
        "failed": failed,
        "run_id": run_id,
        "message": f"{len(added)} member(s) added in M3, {len(failed)} failed, "
                   f"{len(skipped)} skipped.",
    }


def remove_role_members(conn: sqlite3.Connection, client: M3Client, role: str,
                        dry_run: bool = True, confirm: str | None = None,
                        clear_local: bool = True,
                        delete_permissions: int | None = None,
                        progress=_noop_progress) -> dict:
    """
    Take every user out of one M3 role.

    dry_run=True (the default) touches nothing - it re-reads the membership
    from M3 and reports exactly which USIDs a real run would delete.

    A real run needs dry_run=False AND confirm == the role name.
    """
    tenant = client.tenant
    run_id = str(uuid.uuid4())

    definition = client.get_role(role)
    if not definition:
        raise M3ApiError(f"Role '{role}' does not exist in M3 ({tenant}).")

    members = sync_role_members(conn, client, role, run_id=run_id)
    usids = [(r.get("USID") or "").strip() for r in members
             if (r.get("USID") or "").strip()]

    if dry_run:
        _log_m3(conn, tenant, "MNS410MI", "Dlt", "preview", run_id,
                role=role, dry_run=True,
                message=f"{len(usids)} member(s) would be removed",
                payload={"ROLL": role, "USID": usids})
        conn.commit()
        return {
            "role": role,
            "tenant": tenant,
            "dry_run": True,
            "members": usids,
            "removed": 0,
            "failed": [],
            "run_id": run_id,
            "message": f"{len(usids)} member(s) would be removed from "
                       f"'{role}'.",
        }

    if (confirm or "").strip() != role.strip():
        raise M3ApiError(
            f"Confirmation does not match. To remove every member of "
            f"'{role}' the confirm value must be exactly '{role}'.")

    recs = [{"USID": u, "ROLL": role} for u in usids]
    if delete_permissions is not None:
        for r in recs:
            r["DLPR"] = delete_permissions
    results = client.execute_many("MNS410MI", "Dlt", recs, progress=progress,
                                  label=f"Removing {role} members")

    removed, failed, log_rows = [], [], []
    for usid, res in zip(usids, results):
        row = {"role": role, "usid": usid, "payload": res["record"]}
        if res["error"]:
            failed.append({"usid": usid, "message": res["error"]})
            log_rows.append({**row, "outcome": "error", "message": res["error"]})
            continue
        removed.append(usid)
        log_rows.append({**row, "outcome": "ok"})

    _log_m3_many(conn, tenant, "MNS410MI", "Dlt", log_rows, run_id)
    conn.commit()

    try:
        sync_role_members(conn, client, role, run_id=run_id)
    except M3ApiError as exc:
        log.warning("could not re-read members of %s: %s", role, exc)

    # Only mirror the removal locally when M3 took all of it.
    if clear_local and removed and not failed:
        strip_role_members(conn, tenant, role)

    return {
        "role": role,
        "tenant": tenant,
        "dry_run": False,
        "members": usids,
        "removed": removed,
        "failed": failed,
        "run_id": run_id,
        "message": f"{len(removed)} removed, {len(failed)} failed.",
    }


def _cli_progress(width: int = 34):
    """A single-line progress bar for the command line."""
    state = {"label": None}

    def show(done: int, total: int, message: str = "") -> None:
        if not total:
            return
        if message and message != state["label"]:
            if state["label"]:
                sys.stdout.write("\n")
            state["label"] = message
        frac = min(1.0, done / total)
        filled = int(width * frac)
        sys.stdout.write(
            f"\r  {message[:28]:<28} "
            f"[{'#' * filled}{'.' * (width - filled)}] {done}/{total}")
        sys.stdout.flush()
        if done >= total:
            sys.stdout.write("\n")
            state["label"] = None
            sys.stdout.flush()

    return show


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(
        description="M3 REST calls for the security capture.")
    ap.add_argument("--tenant", help="Tenant id, e.g. ZFQP353QZYV89ZHG_TST")
    ap.add_argument("--ionapi-dir", default=str(DEFAULT_IONAPI_DIR))
    ap.add_argument("--company", default=None,
                    help="CONO to send. Left off, the tenant's entry in "
                         "M3_Security_M3.json is used; with neither, the "
                         "request omits it and M3 uses the service account's "
                         "own default company.")
    ap.add_argument("--division", default=None,
                    help="DIVI to send, same rules")
    ap.add_argument("--companies", action="store_true",
                    help="List the companies and divisions this tenant has "
                         "(MNS095MI/Lst + MNS100MI/LstDivisions) and stop")
    ap.add_argument("--m3user", default=None)
    ap.add_argument("--batch-size", type=int, default=DEFAULT_BATCH_SIZE,
                    help=f"Records per API call (default {DEFAULT_BATCH_SIZE})")
    ap.add_argument("--db", default=None,
                    help=f"SQLite path (default {DEFAULT_DB_PATH})")
    ap.add_argument("--list-ionapi", action="store_true",
                    help="Show every .ionapi file and the tenant it points at")
    ap.add_argument("--probe", action="store_true",
                    help="Authenticate and read one role - use this to check "
                         "setup")
    ap.add_argument("--sync-roles", action="store_true",
                    help="Flag which captured roles exist in M3")
    ap.add_argument("--with-members", action="store_true",
                    help="With --sync-roles, also read each matched role's "
                         "members")
    ap.add_argument("--sync-users", action="store_true",
                    help="Refresh the USID to email map from M3")
    ap.add_argument("--members", metavar="ROLE",
                    help="List the M3 members of a role")
    ap.add_argument("--remove-members", metavar="ROLE",
                    help="Remove every member of a role (preview unless "
                         "--commit)")
    ap.add_argument("--create-roles", nargs="*", metavar="ROLE",
                    help="Create roles in M3 with MNS405MI/Add. Names given "
                         "here, or all captured roles not in M3 when given no "
                         "names. Preview unless --commit")
    ap.add_argument("--role-type", default=None,
                    help="Optional ROLT sent with MNS405MI/Add")
    ap.add_argument("--sync-functions", action="store_true",
                    help="Pull function authorisations from SES400MI/Lst and "
                         "flag which of their roles MNS405 is missing")
    ap.add_argument("--delete-function-roles", nargs="+", metavar="FNID:ROLL",
                    help="Delete SES400 authorisations, each given as "
                         "FNID:ROLL (optionally FNID:ROLL:CONO:DIVI). Preview "
                         "unless --commit")
    ap.add_argument("--set-function-status", metavar="STAT",
                    help="Set STAT on the authorisations named by "
                         "--update-function-roles. 10 = Preliminary, "
                         "20 = Active. Preview unless --commit")
    ap.add_argument("--update-function-roles", nargs="+", metavar="FNID:ROLL",
                    help="Authorisations to update, each FNID:ROLL "
                         "(optionally FNID:ROLL:CONO:DIVI)")
    ap.add_argument("--add-missing-roles", action="store_true",
                    help="Add the SES400 roles MNS405 does not have (name and "
                         "description both the role). Preview unless --commit")
    ap.add_argument("--add-members", nargs="*", metavar="ROLE",
                    help="Give the captured members of these roles the role in "
                         "M3 (MNS410MI/Add), resolving USID from the member's "
                         "email. No names = every captured role that is in M3. "
                         "Preview unless --commit")
    ap.add_argument("--commit", action="store_true",
                    help="Actually perform --remove-members / --create-roles")
    args = ap.parse_args(argv)

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s  %(levelname)-8s  %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
        handlers=[logging.StreamHandler(sys.stdout)],
    )

    if args.list_ionapi:
        for e in list_ionapi_files(args.ionapi_dir):
            print(f"  {e['file']:<28} {e.get('tenant') or e.get('error')}")
        return 0

    if not args.tenant:
        ap.error("--tenant is required (or use --list-ionapi)")

    cfg_cono, cfg_divi = tenant_company(args.tenant)
    company = args.company if args.company is not None else cfg_cono
    division = args.division if args.division is not None else cfg_divi

    client = M3Client(args.tenant, args.ionapi_dir, company, division,
                      args.m3user, batch_size=args.batch_size)
    log.info("Tenant %s -> %s", args.tenant, client.ionapi_path.name)
    log.info("Company %s, division %s, %s record(s) per call",
             client.company or "(service account default)",
             client.division or "(service account default)",
             client.batch_size)

    if args.companies:
        for c in discover_companies(client):
            log.info("   CONO %-4s %s", c["cono"], c["name"])
            for d in c["divisions"]:
                log.info("      DIVI %-4s %s", d["divi"], d["name"])
        return 0

    progress = _cli_progress()

    if args.probe:
        client.token(force=True)
        log.info("Token acquired.")
        roles = client.list_roles()
        log.info("MNS405MI/Lst returned %s role(s). First few:", len(roles))
        for r in roles[:5]:
            log.info("   %-12s %s",
                     (r.get("ROLL") or "").strip(),
                     (r.get("TX40") or "").strip())
        return 0

    conn = connect(args.db)
    log.info("Database: %s", resolve_db_path(args.db))
    try:
        if args.sync_users:
            log.info("Users read from M3: %s", sync_users(conn, client))

        if args.sync_roles:
            res = sync_roles(conn, client, with_members=args.with_members,
                             progress=progress)
            log.info("M3 holds %s role(s); %s of the captured roles exist in "
                     "M3, %s M3 roles are not in the capture.",
                     res["m3_roles"], res["matched"], res["unmatched"])
            if args.with_members:
                log.info("Member rows read: %s", res["members"])

        if args.members:
            rows = sync_role_members(conn, client, args.members)
            log.info("%s has %s member(s) in M3:", args.members, len(rows))
            for r in rows:
                log.info("   %s", (r.get("USID") or "").strip())

        if args.sync_functions:
            res = sync_function_roles(conn, client, progress=progress)
            log.info("SES400: %s authorisation(s) over %s function(s) and "
                     "%s role(s)",
                     res["authorisations"], res["functions"], res["roles"])
            log.info("        %s role(s) exist in MNS405, %s missing",
                     res["in_mns405"], res["missing_count"])
            for name in res["missing"][:20]:
                log.info("          missing %s", name)
            if res["missing_count"] > 20:
                log.info("          ... and %s more", res["missing_count"] - 20)

        def _keys_for(specs):
            """Resolve FNID:ROLL[:CONO[:DIVI]] specs to stored authorisation keys."""
            unknown, keys = [], []
            for spec in specs:
                parts = spec.split(":")
                sql = ("SELECT function_role_key FROM M3_Security_FunctionRoles "
                       "WHERE tenant = ? AND fnid = ? AND roll = ?")
                params = [args.tenant, parts[0],
                          parts[1] if len(parts) > 1 else ""]
                if len(parts) > 2 and parts[2]:
                    sql += " AND cono = ?"
                    params.append(parts[2])
                if len(parts) > 3 and parts[3]:
                    sql += " AND divi = ?"
                    params.append(parts[3])
                found = [r[0] for r in conn.execute(sql, params).fetchall()]
                keys.extend(found)
                if found:
                    continue
                unknown.append(spec)
            for spec in unknown:
                log.warning("no stored authorisation matches %s (run "
                            "--sync-functions first)", spec)
            return keys

        if args.set_function_status:
            if not args.update_function_roles:
                ap.error("--set-function-status needs --update-function-roles")
            keys = _keys_for(args.update_function_roles)
            if keys:
                res = update_function_role_status(
                    conn, client, keys, args.set_function_status,
                    dry_run=not args.commit,
                    confirm="UPDATE" if args.commit else None,
                    progress=progress)
                log.info("%s", res["message"])
                for t in (res["plan"] if res["dry_run"] else []):
                    log.info("   would set %-10s %-10s cono=%s divi=%s  %s -> %s",
                             t["fnid"], t["roll"], t["cono"], t["divi"],
                             t["stat"] or "(blank)", t["new_stat"])
                if res["dry_run"]:
                    log.info("Re-run with --commit to perform this.")
                for f in res["failed"]:
                    log.error("   %s / %s: %s", f["fnid"], f["roll"],
                              f["message"])

        if args.delete_function_roles:
            keys = _keys_for(args.delete_function_roles)
            if keys:
                res = delete_function_roles(
                    conn, client, keys, dry_run=not args.commit,
                    confirm="DELETE" if args.commit else None,
                    progress=progress)
                log.info("%s", res["message"])
                for p in (res["plan"] if res["dry_run"] else []):
                    log.info("   would delete %-10s %-10s cono=%s divi=%s",
                             p["fnid"], p["roll"], p["cono"], p["divi"])
                if res["dry_run"]:
                    log.info("Re-run with --commit to perform this.")
                for f in res["failed"]:
                    log.error("   %s / %s: %s", f["fnid"], f["roll"],
                              f["message"])

        if args.add_missing_roles:
            res = create_missing_function_roles(
                conn, client, dry_run=not args.commit,
                confirm="CREATE" if args.commit else None,
                role_type=args.role_type, progress=progress)
            log.info("%s", res["message"])
            if res["dry_run"]:
                for p in res["plan"][:20]:
                    log.info("   would add %-10s TX40=%-10s TX15=%-10s "
                             "(%s function(s))",
                             p["ROLL"], p["TX40"], p["TX15"], p["functions"])
                if len(res["plan"]) > 20:
                    log.info("   ... and %s more", len(res["plan"]) - 20)
                log.info("Re-run with --commit to perform this.")
            else:
                for f in res["failed"]:
                    log.error("   %s: %s", f["name"], f["message"])

        if args.create_roles is not None:
            names = args.create_roles or [r[0] for r in conn.execute(
                "SELECT name FROM M3_Security_Roles WHERE tenant = ? "
                "AND COALESCE(in_m3, 0) = 0 AND row_state <> 'deleted' "
                "ORDER BY name",
                (args.tenant,))]
            log.info("%s captured role(s) considered", len(names))
            res = create_roles(conn, client, names,
                               dry_run=not args.commit,
                               confirm="CREATE" if args.commit else None,
                               role_type=args.role_type, progress=progress)
            log.info("%s", res["message"])
            for reason, n in sorted(res["totals"]["reasons"].items(),
                                    key=lambda x: -x[1]):
                log.info("   skipped %5s  %s", n, reason)
            if res["dry_run"]:
                for p in res["eligible"][:20]:
                    log.info("   would add %-10s TX40=%-40s TX15=%s",
                             p["ROLL"], p["TX40"], p["TX15"])
                if len(res["eligible"]) > 20:
                    log.info("   ... and %s more", len(res["eligible"]) - 20)
                log.info("Re-run with --commit to perform this.")
            else:
                for f in res["failed"]:
                    log.error("   %s: %s", f["name"], f["message"])

        if args.add_members is not None:
            names = args.add_members or [r[0] for r in conn.execute(
                "SELECT name FROM M3_Security_Roles WHERE tenant = ? "
                "AND in_m3 = 1 AND row_state <> 'deleted' ORDER BY name",
                (args.tenant,))]
            log.info("%s role(s) considered", len(names))
            res = add_role_members(conn, client, names,
                                   dry_run=not args.commit,
                                   confirm="ADD" if args.commit else None,
                                   progress=progress)
            log.info("%s", res["message"])
            for reason, n in sorted(res["totals"]["reasons"].items(),
                                    key=lambda x: -x[1]):
                log.info("   skipped %5s  %s", n, reason)
            if res["dry_run"]:
                for t in res["todo"][:20]:
                    log.info("   would add %-10s -> %-10s (%s)",
                             t["usid"], t["role"], t["email"])
                if len(res["todo"]) > 20:
                    log.info("   ... and %s more", len(res["todo"]) - 20)
                log.info("Re-run with --commit to perform this.")
            else:
                for f in res["failed"]:
                    log.error("   %s -> %s: %s", f["usid"], f["role"],
                              f["message"])

        if args.remove_members:
            res = remove_role_members(
                conn, client, args.remove_members,
                dry_run=not args.commit,
                confirm=args.remove_members if args.commit else None,
                progress=progress)
            log.info("%s", res["message"])
            if res["dry_run"]:
                for u in res["members"]:
                    log.info("   would remove %s", u)
                log.info("Re-run with --commit to perform this.")
            else:
                for f in res["failed"]:
                    log.error("   %s: %s", f["usid"], f["message"])
    finally:
        conn.close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
