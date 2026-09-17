"""
SheetSecurity_M3Api - live M3 calls behind the Doppio API Sheet security tool.

EXT124MI is the custom program the API Sheet's own security lives in (the
same program SEC_GrantAccess.py and the Xtend workbook write through):

    LstUsrInfo   list every security record
    GetUsrInfo   read one record
    AddUsrInfo   create one record
    UpdUsrInfo   change one record
    DelUsrInfo   remove one record

A record's natural key is (PCID, TNNM, AUTH) - PCID is the person/login id,
TNNM the customer tenant name the record grants against, AUTH the access
level (1 user grant, 20 tenant registration, 99 pending). HASH, M3ID and
UMSG are the value fields carried on top of that key; M3ID holds an M3 USID
and UMSG an email address once a record has been matched to a real M3 user.

Because EXT124MI is a private extension, it has no metadata in the shared
M3 API catalogue used elsewhere in this repo - LstUsrInfo/GetUsrInfo are
therefore called with no selectedColumns filter, so whatever field layout
the program actually returns comes through untouched rather than a
hardcoded guess of it.

Connection details come from an .ionapi file, same convention as every
other Infor tool in this repo: every file in the ionapi directory is read
and the one whose "ti" equals the requested tenant wins.

guess_m3_user() is the "best guess" half of the tool: given a TNNM value and
a PCID, it resolves that tenant's own .ionapi file (TNNM is a display name
like "NUTRACORP TST", not necessarily an exact "ti"), connects to it, reads
MNS150MI/LstUserData, and scores every M3 user against the PCID to suggest
which USID/EMAL the record's M3ID/UMSG should carry.
"""

from __future__ import annotations

import base64
import datetime
import difflib
import json
import logging
import re
import time
from pathlib import Path

import requests

log = logging.getLogger("SheetSecurity_M3Api")

BASE_DIR = Path(__file__).parent.resolve()
# ionapi/ is the shared repo-root folder of .ionapi tenant-credential files
# used by every Infor automation tool in this repo - not package-local.
DEFAULT_IONAPI_DIR = BASE_DIR.parent.parent / "ionapi"
DEFAULT_TENANT = "DOPPIO_DEM"
DEFAULT_TIMEOUT = 60

# EXT124MI record shape - the key that identifies a row, then the fields
# carried on top of it. Any other field a live LstUsrInfo happens to return
# is still shown (see api_security() in the Flask app) - this list only
# decides display and form order.
EXT124_KEY_FIELDS = ("PCID", "TNNM", "AUTH")
EXT124_VALUE_FIELDS = ("HASH", "M3ID", "UMSG")
EXT124_FIELD_ORDER = EXT124_KEY_FIELDS + EXT124_VALUE_FIELDS

# AUTH is a plain M3 field with no built-in value list; these are the codes
# seen in actual use elsewhere in this repo (SEC_GrantAccess.py, Doppio.bas).
# 0/1 are a pair - 1 grants a user access to a tenant, 0 is that same grant
# blocked rather than deleted, which is why both show up together on the
# Users tab in the Flask app.
AUTH_LABELS = {
    "0": "0 — User access blocked",
    "1": "1 — User access",
    "20": "20 — Tenant registration",
    "99": "99 — Pending request",
}

# ---------------------------------------------------------------------------
# EXPORTMI/Select over EXTXSM - the raw table behind EXT124MI.
#
# LstUsrInfo returns PCID/TNNM/AUTH/HASH/M3ID/UMSG only. The list view also
# wants the audit trail EXT124MI doesn't expose - when a record was
# registered (EXRGDT/EXRGTM), when it was last changed (EXLMDT/EXLMTS) and
# by what change (EXCHNO/EXCHID) - which means reading the table directly
# through EXPORTMI/Select instead, the same call SEC_GrantAccess.py and
# Doppio.bas already make against this table.
# ---------------------------------------------------------------------------

EXPORTMI_SEP = "^"
EXTXSM_QUERY = ("EXPCID,EXTNNM,EXM3ID,EXAUTH,EXUMSG,EXLMTS,EXLMDT,EXRGDT,EXRGTM,"
                "EXCHNO,EXCHID,EXHASH from EXTXSM")

# EXTXSM's own column names, folded down to the names the rest of this app
# already uses (PCID/TNNM/AUTH/HASH/M3ID/UMSG match LstUsrInfo's; the rest
# are new).
EXTXSM_FIELD_MAP = {
    "EXPCID": "PCID", "EXTNNM": "TNNM", "EXM3ID": "M3ID", "EXAUTH": "AUTH",
    "EXUMSG": "UMSG", "EXLMTS": "LMTS", "EXLMDT": "LMDT", "EXRGDT": "RGDT",
    "EXRGTM": "RGTM", "EXCHNO": "CHNO", "EXCHID": "CHID", "EXHASH": "HASH",
}
# USERNAME/DOMAIN/COMPUTERNAME aren't real EXTXSM columns - the Flask app
# decodes them out of an AUTH=99 record's HASH (see _decode_review_hash())
# and slots them in here so the Review tab shows them in a sensible spot.
EXTXSM_FIELD_ORDER = ("PCID", "TNNM", "AUTH", "M3ID", "UMSG",
                      "USERNAME", "DOMAIN", "COMPUTERNAME",
                      "RGDT", "RGTM", "LMDT", "LMTS", "CHNO", "CHID", "HASH")


class M3ApiError(RuntimeError):
    """Anything that goes wrong talking to M3."""


def _parse_exportmi_rows(records: list[dict], field_map: dict[str, str] | None = None) -> list[dict]:
    """
    Turn EXPORTMI/Select's REPL rows (called with HDRS=1, so the first row is
    a header of column names) into a list of dicts, one per data row.
    field_map renames raw table columns (e.g. "EXPCID") to the names the rest
    of this app already uses ("PCID"); a column not in it keeps its raw name.
    """
    repl_rows = [r.get("REPL", "") for r in records if r.get("REPL")]
    if not repl_rows:
        return []
    field_map = field_map or {}
    header = [c.strip() for c in repl_rows[0].rstrip(EXPORTMI_SEP).split(EXPORTMI_SEP)]
    columns = [field_map.get(c, c) for c in header]

    out = []
    for raw in repl_rows[1:]:
        values = raw.rstrip(EXPORTMI_SEP).split(EXPORTMI_SEP)
        values += [""] * (len(columns) - len(values))
        out.append({columns[i]: values[i].strip() for i in range(len(columns))})
    return out


def format_java_timestamp(value) -> str:
    """
    EXLMTS is a Java timestamp - milliseconds since the Unix epoch, the same
    unit System.currentTimeMillis() returns. Blank or zero means the record
    has never been changed since it was registered.
    """
    raw = str(value or "").strip()
    if not raw or raw == "0":
        return ""
    try:
        millis = int(float(raw))
    except ValueError:
        return raw
    try:
        return (datetime.datetime.fromtimestamp(millis / 1000, tz=datetime.timezone.utc)
                .strftime("%Y-%m-%d %H:%M:%S UTC"))
    except (OverflowError, OSError, ValueError):
        return raw


# ---------------------------------------------------------------------------
# .ionapi discovery
# ---------------------------------------------------------------------------

def list_ionapi_files(ionapi_dir: str | Path = DEFAULT_IONAPI_DIR) -> list[dict]:
    """Every readable .ionapi file with the tenant ('ti') it points at."""
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


def resolve_ionapi(tenant: str, ionapi_dir: str | Path = DEFAULT_IONAPI_DIR) -> Path:
    """Find the .ionapi file whose 'ti' matches the tenant, exactly."""
    matches = [e for e in list_ionapi_files(ionapi_dir)
               if (e.get("tenant") or "").strip().upper() == tenant.strip().upper()]
    if not matches:
        known = ", ".join(sorted(
            e["tenant"] for e in list_ionapi_files(ionapi_dir) if e.get("tenant")
        )) or "none"
        raise M3ApiError(
            f"No .ionapi file in {ionapi_dir} has ti='{tenant}'. Tenants available: {known}")
    if len(matches) > 1:
        log.warning("Several .ionapi files match %s (%s) - using %s",
                    tenant, ", ".join(m["file"] for m in matches), matches[0]["file"])
    return Path(matches[0]["path"])


def _normalize_name(s: str | None) -> str:
    return " ".join(str(s or "").replace("_", " ").split()).strip().lower()


def resolve_ionapi_fuzzy(name: str, ionapi_dir: str | Path = DEFAULT_IONAPI_DIR) -> Path:
    """
    Find the .ionapi file for a tenant name that may not exactly equal any
    file's 'ti' - TNNM values are display names like "NUTRACORP TST" while
    the file/ti pair is usually "NUTRACORP_TST".
    """
    files = list_ionapi_files(ionapi_dir)
    target = _normalize_name(name)
    if not target:
        raise M3ApiError("A tenant name is required.")

    for f in files:
        if _normalize_name(f.get("tenant")) == target:
            return Path(f["path"])
    for f in files:
        if _normalize_name(Path(f["file"]).stem) == target:
            return Path(f["path"])

    candidates = [f for f in files
                  if target in _normalize_name(f.get("tenant"))
                  or target in _normalize_name(Path(f["file"]).stem)]
    if len(candidates) == 1:
        return Path(candidates[0]["path"])
    if len(candidates) > 1:
        names = ", ".join(sorted(c["file"] for c in candidates))
        raise M3ApiError(
            f"'{name}' matches more than one .ionapi file ({names}) - "
            f"use a more specific TNNM value.")

    known = ", ".join(sorted(f["file"] for f in files)) or "none"
    raise M3ApiError(f"No .ionapi file matches tenant '{name}'. Files available: {known}")


# ---------------------------------------------------------------------------
# The client itself.
# ---------------------------------------------------------------------------

class M3Client:
    """Minimal, non-interactive M3 REST client for EXT124MI / MNS150MI."""

    def __init__(self, tenant: str, ionapi_dir: str | Path = DEFAULT_IONAPI_DIR,
                 timeout: int = DEFAULT_TIMEOUT, _ionapi_path: Path | None = None):
        self.tenant = tenant
        self.timeout = timeout
        # _ionapi_path lets a caller that already resolved (or just wrote) an
        # exact file hand it over directly, rather than resolve_ionapi()
        # re-deriving it from 'ti' - which, when two files happen to share a
        # 'ti' (see for_tenant_name/ensure_ionapi_file), could pick the other
        # one instead of the file that was actually just resolved.
        self.ionapi_path = _ionapi_path or resolve_ionapi(tenant, ionapi_dir)
        self._cfg = json.loads(self.ionapi_path.read_text())
        for key in ("ti", "iu", "pu", "ot", "ci", "cs", "saak", "sask"):
            if self._cfg.get(key):
                continue
            raise M3ApiError(
                f"{self.ionapi_path.name} is missing '{key}' - it does not look "
                f"like a service-account .ionapi file.")

        self._token = None
        self.session = requests.Session()
        # Set True by for_tenant_name() when it had to extract a fresh
        # .ionapi from a security tenant's EXT124MI - see ensure_ionapi_file().
        self.extracted_ionapi = False

    @classmethod
    def for_tenant_name(cls, name: str, ionapi_dir: str | Path = DEFAULT_IONAPI_DIR,
                        timeout: int = DEFAULT_TIMEOUT,
                        security_tenant: str | None = None) -> "M3Client":
        """
        Build a client for a tenant identified loosely - see
        resolve_ionapi_fuzzy. When no .ionapi file for it exists yet and
        security_tenant is given, one is extracted from security_tenant's own
        EXT124MI AUTH=20 record and saved - see ensure_ionapi_file(). The
        resulting client's .extracted_ionapi says whether that happened.
        """
        path, extracted = ensure_ionapi_file(name, ionapi_dir, security_tenant)
        cfg = json.loads(path.read_text())
        client = cls(cfg["ti"], ionapi_dir, timeout, _ionapi_path=path)
        client.extracted_ionapi = extracted
        return client

    @property
    def api_url(self) -> str:
        iu = self._cfg["iu"].rstrip("/")
        return (f"{iu}/{self._cfg['ti']}/M3/m3api-rest/v2/execute"
                f"?maxrecs=0&extendedresult=true&righttrim=true")

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
        except requests.RequestException as exc:
            raise M3ApiError(
                f"Token request failed for {self.tenant} ({self.ionapi_path.name}): {exc}"
            ) from exc
        except (KeyError, ValueError) as exc:
            raise M3ApiError(f"Token response had no access_token: {exc}") from exc

    def token(self, force: bool = False) -> str:
        if force or not self._token:
            self._fetch_token()
        return self._token

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
                if "not found" in err.lower():
                    continue
                raise M3ApiError(f"{code}: {err}" if code else err)
            rows.extend(res.get("records") or [])
        return rows

    # ---- EXT124MI - the API Sheet's own security ---------------------
    def list_usr_info(self) -> list[dict]:
        """EXT124MI/LstUsrInfo - every security record."""
        body = self.execute("EXT124MI", [{"transaction": "LstUsrInfo", "record": {}}])
        return self._records(body, "LstUsrInfo")

    def get_usr_info(self, pcid: str, tnnm: str, auth: str) -> dict | None:
        """EXT124MI/GetUsrInfo - one record, or None."""
        rec = {"PCID": pcid, "TNNM": tnnm, "AUTH": str(auth)}
        body = self.execute("EXT124MI", [{"transaction": "GetUsrInfo", "record": rec}])
        rows = self._records(body, "GetUsrInfo")
        return rows[0] if rows else None

    def add_usr_info(self, pcid: str, tnnm: str, auth: str,
                     hash_val: str = "", m3id: str = "", umsg: str = "") -> dict:
        """EXT124MI/AddUsrInfo - create a security record."""
        rec = {"PCID": pcid, "TNNM": tnnm, "AUTH": str(auth)}
        if hash_val:
            rec["HASH"] = hash_val
        if m3id:
            rec["M3ID"] = m3id
        if umsg:
            rec["UMSG"] = umsg
        return self.execute("EXT124MI", [{"transaction": "AddUsrInfo", "record": rec}])

    def update_usr_info(self, pcid: str, tnnm: str, auth: str, **changes) -> dict:
        """EXT124MI/UpdUsrInfo - change HASH / M3ID / UMSG on an existing record."""
        rec = {"PCID": pcid, "TNNM": tnnm, "AUTH": str(auth)}
        for key in EXT124_VALUE_FIELDS:
            if key in changes and changes[key] is not None:
                rec[key] = changes[key]
        return self.execute("EXT124MI", [{"transaction": "UpdUsrInfo", "record": rec}])

    def delete_usr_info(self, pcid: str, tnnm: str, auth: str) -> dict:
        """EXT124MI/DelUsrInfo - remove a security record."""
        rec = {"PCID": pcid, "TNNM": tnnm, "AUTH": str(auth)}
        return self.execute("EXT124MI", [{"transaction": "DelUsrInfo", "record": rec}])

    # ---- EXPORTMI - EXTXSM read straight off the table -----------------
    def list_extxsm(self) -> list[dict]:
        """
        EXPORTMI/Select over EXTXSM - every security record, with the
        registration/last-changed audit columns LstUsrInfo does not carry.

        EXLMTS comes back as a Java timestamp (milliseconds since the Unix
        epoch) - formatted here so the list never shows a bare 13-digit
        number.
        """
        body = self.execute("EXPORTMI", [{
            "transaction": "Select",
            "record": {"QERY": EXTXSM_QUERY, "SEPC": EXPORTMI_SEP, "HDRS": "1"},
            "selectedColumns": ["QERY", "SEPC", "HDRS", "REPL"],
        }])
        rows = _parse_exportmi_rows(self._records(body, "Select"), EXTXSM_FIELD_MAP)
        for row in rows:
            if "LMTS" in row:
                row["LMTS"] = format_java_timestamp(row["LMTS"])
        return rows

    # ---- MNS150MI - the M3 user directory -----------------------------
    def list_user_data(self) -> list[dict]:
        """MNS150MI/LstUserData - USID / name / email / status for every M3 user."""
        body = self.execute("MNS150MI", [
            {"transaction": "LstUserData", "record": {},
             "selectedColumns": ["USID", "TX40", "EMAL", "USTA"]}])
        return self._records(body, "LstUserData")


# ---------------------------------------------------------------------------
# Best-guess USID / email lookup for a PCID on a given tenant.
# ---------------------------------------------------------------------------

def _login_key(s: str | None) -> str:
    """Fold a login-ish string down to bare alphanumerics for comparison."""
    return "".join(ch for ch in str(s or "").lower() if ch.isalnum())


def _match_score(term_key: str, usid: str, local: str, email: str, name: str) -> float:
    """How well one search term (already folded via _login_key) fits one M3 user."""
    if not term_key:
        return 0.0
    if usid and _login_key(usid) == term_key:
        return 1.0
    if local and _login_key(local) == term_key:
        return 0.97
    if email and _login_key(email) == term_key:
        return 0.97
    return max(
        difflib.SequenceMatcher(None, term_key, _login_key(usid)).ratio() if usid else 0.0,
        difflib.SequenceMatcher(None, term_key, _login_key(local)).ratio() if local else 0.0,
        difflib.SequenceMatcher(None, term_key, _login_key(email)).ratio() if email else 0.0,
        difflib.SequenceMatcher(None, term_key, _login_key(name)).ratio() if name else 0.0,
    )


def guess_m3_user(tenant_name: str, pcid: str,
                  ionapi_dir: str | Path = DEFAULT_IONAPI_DIR,
                  security_tenant: str | None = None,
                  hint: str | None = None) -> dict:
    """
    Best guess at the M3 user a PCID belongs to on the tenant named by TNNM.

    Connects with that tenant's own .ionapi (resolved loosely - TNNM is a
    display name, not necessarily an exact 'ti'). If no file for it exists
    yet and security_tenant is given (normally whichever tenant's security
    table is currently open), one is extracted on the spot from that
    tenant's own AUTH=20 EXT124MI record and saved - there should already be
    one, since that record is exactly what registered this TNNM in the first
    place. Reads MNS150MI/LstUserData and scores every user against PCID: an
    exact match on USID or email scores highest, everything else falls back
    to a fuzzy string match against USID, email and full name.

    hint is an optional override search term - a name or email typed by
    hand for when PCID is not a real login id (or is simply wrong). When
    hint is given, it replaces PCID entirely for scoring rather than being
    blended with it, so a bad PCID can't drag down a search the user has
    deliberately pointed elsewhere. At least one of pcid / hint must be
    given.

    Candidates are returned sorted by score; 'best' is set only when the top
    score clears a plausibility bar.
    """
    hint_key = _login_key(hint)
    term_key = hint_key or _login_key(pcid)
    if not term_key:
        raise M3ApiError("Enter a PCID, or a name/email hint, to look up an M3 user.")

    client = M3Client.for_tenant_name(tenant_name, ionapi_dir,
                                      security_tenant=security_tenant)
    users = client.list_user_data()

    scored = []
    for u in users:
        usid = (u.get("USID") or "").strip()
        name = (u.get("TX40") or "").strip()
        email = (u.get("EMAL") or "").strip()
        if not (usid or email):
            continue
        local = email.split("@", 1)[0] if "@" in email else email

        score = _match_score(term_key, usid, local, email, name)
        if score >= 0.4:
            scored.append({"usid": usid, "name": name, "email": email,
                          "status": (u.get("USTA") or "").strip(),
                          "score": round(score, 3)})

    scored.sort(key=lambda r: -r["score"])
    best = scored[0] if scored and scored[0]["score"] >= 0.6 else None
    return {"tenant": client.tenant, "ionapi": client.ionapi_path.name,
            "extracted": client.extracted_ionapi,
            "candidates": scored[:8], "best": best}


# ---------------------------------------------------------------------------
# HASH decode / encode.
#
# HASH is not actually encrypted - it is a base64-wrapped JSON blob, the same
# scheme SEC_GrantAccess.py's _decode_exhash()/_encode_ionapi() use: an
# AUTH=20 record's HASH is an encoded .ionapi file, an AUTH=99 record's HASH
# is workbook telemetry ({"userName":..., "userDomain":..., ...}). Decoding
# it is just base64 + json.loads; encoding is json.dumps (compact) + base64.
# ---------------------------------------------------------------------------

_LEGAL_ESCAPE = re.compile(r'\\(?:["\\/bfnrt]|u[0-9A-Fa-f]{4})')


def _escape_bare_backslashes(text: str) -> str:
    """
    Double every backslash that is not the start of a legal JSON escape
    (\\\\ \\" \\/ \\b \\f \\n \\r \\t \\uXXXX).

    Windows paths in this telemetry ("C:\\Users\\Name") show up with bare
    single backslashes, which is invalid JSON on its own. A naive fix would
    treat any "\\u" as the start of a unicode escape and leave it alone, but
    a lowercase network path segment - "\\\\fileserver\\users\\jdoe" - has a
    "\\u" that is not followed by four hex digits, which is exactly the
    "Invalid \\uXXXX escape" error this works around: only a backslash
    followed by four real hex digits counts as \\uXXXX, so "\\users" still
    gets escaped like any other bare backslash.
    """
    out = []
    i, n = 0, len(text)
    while i < n:
        if text[i] == "\\":
            m = _LEGAL_ESCAPE.match(text, i)
            if m:
                out.append(m.group())
                i = m.end()
                continue
            out.append("\\\\")
            i += 1
            continue
        out.append(text[i])
        i += 1
    return "".join(out)


def decode_hash_blob(value: str) -> dict:
    """
    Base64-decode a HASH value and parse it as JSON.

    Mirrors SEC_GrantAccess._decode_exhash(): padded to a multiple of 4
    before decoding, and on a JSON parse failure, bare backslashes are
    escaped (see _escape_bare_backslashes) and the parse is retried -
    Windows profile paths like "C:\\Users\\Name" show up unescaped in these
    blobs.
    """
    value = (value or "").strip()
    if not value:
        raise M3ApiError("Nothing to decode.")

    padded = value + "=" * (-len(value) % 4)
    try:
        raw_bytes = base64.b64decode(padded)
    except Exception as exc:
        raise M3ApiError(f"HASH is not valid base64: {exc}") from exc

    for encoding in ("utf-8", "latin-1"):
        try:
            text = raw_bytes.decode(encoding)
            break
        except Exception:
            continue
    else:
        text = raw_bytes.decode("utf-8", errors="replace")

    try:
        return json.loads(text)
    except json.JSONDecodeError:
        fixed = _escape_bare_backslashes(text)
        try:
            return json.loads(fixed)
        except json.JSONDecodeError as exc:
            raise M3ApiError(
                f"HASH decoded but is not valid JSON: {exc}") from exc


def encode_hash_blob(data) -> str:
    """
    JSON-encode (compact, matching what M3 already stores) then base64-encode
    - the inverse of decode_hash_blob(). `data` may be a dict/list already, or
    a JSON string (e.g. edited by hand, pretty-printed or not - json.loads
    ignores the whitespace either way), which is parsed first.

    Trailing '=' padding is stripped from the result. Real HASH values never
    carry it (a stored value's length is routinely not a multiple of 4) -
    the VBA encoder behind these blobs omits it, and decode_hash_blob() /
    SEC_GrantAccess._decode_exhash() both re-pad before decoding to
    compensate. Leaving padding on here would produce a HASH that looks
    different from every other row's for no functional reason.
    """
    if isinstance(data, str):
        try:
            data = json.loads(data)
        except json.JSONDecodeError as exc:
            raise M3ApiError(f"Not valid JSON: {exc}") from exc
    text = json.dumps(data, separators=(",", ":"))
    return base64.b64encode(text.encode("utf-8")).decode("utf-8").rstrip("=")


# ---------------------------------------------------------------------------
# Extracting AUTH=20 records back into real .ionapi files.
#
# An AUTH=20 record is EXTXSM's own note that "this tenant is registered" -
# its HASH is exactly an encoded .ionapi file (see decode_hash_blob and
# SEC_GrantAccess._encode_ionapi/check_and_add_customer_tenant). Extracting
# it recreates the credential file that produced it in the first place.
# ---------------------------------------------------------------------------

# Same fields M3Client.__init__ requires of a service-account .ionapi file -
# a decoded blob missing any of these is not one, whatever else it is.
REQUIRED_IONAPI_FIELDS = ("ti", "iu", "pu", "ot", "ci", "cs", "saak", "sask")

_UNSAFE_FILENAME_CHARS = re.compile(r"[^A-Za-z0-9_]+")


def looks_like_ionapi(data) -> bool:
    return isinstance(data, dict) and all(
        str(data.get(k) or "").strip() for k in REQUIRED_IONAPI_FIELDS)


def safe_ionapi_stem(name: str) -> str:
    """A tenant identifier, folded down to characters safe in a filename."""
    return _UNSAFE_FILENAME_CHARS.sub("_", str(name or "").strip()).strip("_").upper()


def _find_type20_record(security_tenant: str, tenant_name: str,
                        ionapi_dir: str | Path) -> dict | None:
    """The AUTH=20 record for tenant_name, read from security_tenant's own EXT124MI."""
    client = M3Client(security_tenant, ionapi_dir=ionapi_dir)
    # AUTH=20 records are keyed PCID==TNNM==the tenant name in every case seen
    # so far, so try the direct Get first - it is one call instead of a full Lst.
    row = client.get_usr_info(tenant_name, tenant_name, "20")
    if row and str(row.get("HASH") or "").strip():
        return row
    target = _normalize_name(tenant_name)
    for r in client.list_usr_info():
        if str(r.get("AUTH") or "").strip() != "20":
            continue
        if _normalize_name(r.get("TNNM")) == target:
            return r
    return None


def ensure_ionapi_file(tenant_name: str, ionapi_dir: str | Path = DEFAULT_IONAPI_DIR,
                       security_tenant: str | None = None) -> tuple[Path, bool]:
    """
    Resolve tenant_name to an .ionapi file, extracting and saving one from
    security_tenant's own EXT124MI if no file for it exists yet.

    Returns (path, extracted) - extracted is True only when this call wrote
    a new file. Without security_tenant this behaves exactly like
    resolve_ionapi_fuzzy (raises if nothing matches on disk), which is the
    right behaviour for anything that isn't already connected to a security
    tenant to fall back to.
    """
    try:
        return resolve_ionapi_fuzzy(tenant_name, ionapi_dir), False
    except M3ApiError as exc:
        if not security_tenant:
            raise
        # `as exc` unbinds exc once the except block ends (Python deletes it
        # to avoid a reference cycle), so its message has to be saved here.
        not_on_disk_message = str(exc)

    row = _find_type20_record(security_tenant, tenant_name, ionapi_dir)
    if row is None:
        raise M3ApiError(
            f"{not_on_disk_message} {security_tenant} has no AUTH=20 "
            f"registration for '{tenant_name}' either.")

    hash_val = str(row.get("HASH") or "").strip()
    data = decode_hash_blob(hash_val)
    if not looks_like_ionapi(data):
        raise M3ApiError(
            f"{security_tenant}'s AUTH=20 record for '{tenant_name}' does not "
            f"decode to a usable .ionapi.")

    stem = safe_ionapi_stem(row.get("TNNM") or data.get("ti") or row.get("PCID"))
    ionapi_dir = Path(ionapi_dir)
    ionapi_dir.mkdir(parents=True, exist_ok=True)
    path = ionapi_dir / f"{stem}.ionapi"
    path.write_text(json.dumps(data, indent=4))
    return path, True


def extract_type20_ionapi(rows: list[dict], ionapi_dir: str | Path,
                          overwrite: bool = False, dry_run: bool = True,
                          selected_keys=None) -> dict:
    """
    Write every AUTH=20 record's HASH out as a real .ionapi file.

    dry_run=True (the default) touches nothing - it returns the plan for
    every AUTH=20 record: which file it would write, or why it was skipped
    (no HASH, HASH does not decode to a complete .ionapi payload, the file
    already exists and overwrite is False, or it was left unchecked - see
    selected_keys). A real run needs dry_run=False; existing files are left
    alone unless overwrite=True.

    selected_keys, when given, is an iterable of (PCID, TNNM) pairs - the
    ones the caller actually wants exported. A record still appears in the
    plan either way (so the browser can show every AUTH=20 record and let
    the user tick which ones to write), but one whose (PCID, TNNM) is not in
    selected_keys is reported skipped as "Not selected for export" and never
    written. None (the default) means everything eligible is a candidate,
    which is what a first preview - before anything has been picked - wants.

    The filename is the record's own TNNM (the customer tenant name it
    grants access to), falling back to the decoded blob's 'ti' then PCID if
    TNNM is somehow blank.
    """
    ionapi_dir = Path(ionapi_dir)
    plan = []
    wanted = None
    if selected_keys is not None:
        wanted = {(str(p).strip(), str(t).strip()) for p, t in selected_keys}

    for r in rows:
        if str(r.get("AUTH") or "").strip() != "20":
            continue
        pcid = str(r.get("PCID") or "").strip()
        tnnm = str(r.get("TNNM") or "").strip()
        entry = {"pcid": pcid, "tnnm": tnnm, "ti": "", "file": None,
                 "action": "skip", "reason": None}

        hash_val = str(r.get("HASH") or "").strip()
        if not hash_val:
            entry["reason"] = "No HASH on this record."
            plan.append(entry)
            continue
        try:
            data = decode_hash_blob(hash_val)
        except M3ApiError as exc:
            entry["reason"] = str(exc)
            plan.append(entry)
            continue
        if not looks_like_ionapi(data):
            entry["reason"] = "Decoded JSON is not a complete service-account .ionapi payload."
            plan.append(entry)
            continue

        entry["ti"] = data.get("ti", "")
        stem = safe_ionapi_stem(tnnm or data.get("ti") or pcid)
        if not stem:
            entry["reason"] = "No usable name (TNNM, ti and PCID are all blank)."
            plan.append(entry)
            continue

        filename = f"{stem}.ionapi"
        path = ionapi_dir / filename
        entry["file"] = filename

        if wanted is not None and (pcid, tnnm) not in wanted:
            entry["reason"] = "Not selected for export."
            plan.append(entry)
            continue

        if path.exists() and not overwrite:
            entry["reason"] = "File already exists (enable overwrite to replace it)."
            plan.append(entry)
            continue

        entry["action"] = "overwrite" if path.exists() else "write"
        entry["reason"] = None
        if not dry_run:
            ionapi_dir.mkdir(parents=True, exist_ok=True)
            path.write_text(json.dumps(data, indent=4))
        plan.append(entry)

    totals = {
        "type20": sum(1 for r in rows if str(r.get("AUTH") or "").strip() == "20"),
        "written": sum(1 for p in plan if p["action"] in ("write", "overwrite")),
        "skipped": sum(1 for p in plan if p["action"] == "skip"),
    }
    return {"dry_run": dry_run, "plan": plan, "totals": totals}
