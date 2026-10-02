"""
SheetSecurity_App - Flask front end for the Doppio API Sheet security table.

One page over EXTXSM, the table EXT124MI (the custom M3 program the API
Sheet's own access control lives in) reads and writes. The list is read
straight off the table via EXPORTMI/Select - list_extxsm() - rather than
EXT124MI/LstUsrInfo, because the table carries audit columns (when a record
was registered, last changed, and by what change) LstUsrInfo does not
return. Add / edit / delete still go through EXT124MI itself (AddUsrInfo /
UpdUsrInfo / DelUsrInfo). Nothing is cached locally - every action talks
straight to M3.

The front end splits that one list into three tabs by AUTH: Tenants (20 -
a tenant registration), Users (1 access granted, 0 that same access
blocked) and Review (99 - someone who has attempted to use the sheet
without a grant yet). Review has a "Delete all" action, since that tab is
the one meant to be cleared out wholesale once its requests are handled.

A record's real key is (PCID, TNNM) - not (PCID, TNNM, AUTH) as the tabs
might suggest. GetUsrInfo/AddUsrInfo/DelUsrInfo all match on PCID+TNNM alone
and ignore whatever AUTH is passed alongside them (confirmed against a live
tenant: GetUsrInfo with a wrong AUTH still returns the one row for that
PCID+TNNM). AUTH, HASH, M3ID and UMSG are all just value fields on top of
that key, and AUTH is changed with a plain UpdUsrInfo exactly like the
others - see api_security_update(). HASH is often very long, so the list
only ever sends a truncated copy of it to the browser - GetUsrInfo is called
for the full value when a row is opened for editing.

The "Find M3 user" action on the add/edit form is the other half of the
tool: given the record's TNNM and PCID it connects to *that* tenant's own
.ionapi, reads MNS150MI/LstUserData, and offers a best guess for which USID
(-> M3ID) and email (-> UMSG) the record should carry.
"""

from __future__ import annotations

import argparse
import ipaddress
import logging
import os
import sys
import threading
import time
import traceback
from pathlib import Path

import requests

from flask import Flask, jsonify, render_template, request

from SheetSecurity_M3Api import (
    AUTH_LABELS,
    DEFAULT_IONAPI_DIR,
    DEFAULT_TENANT,
    EXT124_KEY_FIELDS,
    EXTXSM_FIELD_ORDER,
    M3ApiError,
    M3Client,
    decode_hash_blob,
    encode_hash_blob,
    extract_type20_ionapi,
    guess_from_ifs_export,
    guess_m3_user,
    list_ionapi_files,
    match_ionapi_name,
    parse_ifs_export,
    resolve_tenant_registration,
    score_m3_users,
)
from SheetSecurity_M3Api import _login_key

# The two UMSG values a Tenants-tab (AUTH=20) row can be tagged with once
# its M3ID has been resolved - see _resolve_tenant_row() and api_tenants().
UMSG_MANAGING_SYSTEM = "Managing System"
UMSG_CONNECTION_ERROR = "Connection Error"

BASE_DIR = Path(__file__).parent.resolve()

log = logging.getLogger("SheetSecurity_App")

app = Flask(__name__, template_folder=str(BASE_DIR / "templates"))
app.config["SHEET_SECURITY_IONAPI_DIR"] = DEFAULT_IONAPI_DIR
# Flask's jsonify() alphabetises object keys by default. That silently
# reordered a decoded HASH's fields, so a "Decrypt" -> "Encrypt" round trip
# with no edits produced a byte-different HASH from the one M3 actually has
# (same JSON, different key order -> different base64). Keeping insertion
# order means an unedited round trip reproduces the original HASH exactly.
app.json.sort_keys = False

# HASH values can run to several KB (an encoded .ionapi blob); the list view
# only ever needs enough of it to show something is there.
HASH_PREVIEW_LEN = 24

# An uploaded IFS export lives only in memory for as long as this process
# runs, like every other piece of tenant state here - see api_ifs_export_*()
# and the merge into api_guess() below.
_ifs_export: dict = {"filename": None, "rows": []}

# AbuseIPDB, for the Review tab's PUBLICIP. The public check page
# (abuseipdb.com/check/<ip>) sits behind Cloudflare and refuses anything that
# is not a browser, so the score comes from their v2 API instead
# (docs.abuseipdb.com - GET /api/v2/check, key in the Key header), which
# needs a key. It is looked for in --abuseipdb-key, then ABUSEIPDB_API_KEY,
# then ABUSEIPDB_KEY_FILE - outside the repo on purpose, since ionapi/ is
# committed - which the Review tab's "AbuseIPDB key" button writes. Without
# one the Review tab still links each IP to its check page, just unscored.
# Scores are cached in memory for ABUSEIPDB_CACHE_SECONDS, because the free
# plan allows 1,000 checks a day and every reload of the list would otherwise
# spend one per IP.
ABUSEIPDB_CHECK_URL = "https://api.abuseipdb.com/api/v2/check"
ABUSEIPDB_CACHE_SECONDS = 6 * 3600
ABUSEIPDB_MAX_AGE_DAYS = 90
ABUSEIPDB_KEY_FILE = Path.home() / ".config" / "doppio" / "abuseipdb.key"


def _read_abuse_key_file() -> str:
    try:
        return ABUSEIPDB_KEY_FILE.read_text(encoding="utf-8").strip()
    except OSError:
        return ""


app.config["ABUSEIPDB_API_KEY"] = (os.environ.get("ABUSEIPDB_API_KEY", "").strip()
                                   or _read_abuse_key_file())
_abuse_cache: dict[str, tuple[float, dict]] = {}
_abuse_lock = threading.Lock()


@app.errorhandler(M3ApiError)
def _handle_m3(exc):
    return jsonify(status="error", message=str(exc)), 400


@app.errorhandler(Exception)
def _handle(exc):
    if hasattr(exc, "code") and isinstance(getattr(exc, "code"), int):
        return jsonify(status="error", message=str(exc)), exc.code
    log.error("%s", traceback.format_exc())
    return jsonify(status="error", message=str(exc),
                   trace=traceback.format_exc()), 500


def _ionapi_dir() -> Path:
    return app.config["SHEET_SECURITY_IONAPI_DIR"]


def _client(tenant: str) -> M3Client:
    tenant = (tenant or "").strip()
    if not tenant:
        raise M3ApiError("A tenant is required.")
    return M3Client(tenant, ionapi_dir=_ionapi_dir())


def _order_columns(rows: list[dict]) -> list[str]:
    """Known EXTXSM fields first, in a sensible order, then anything else
    EXPORTMI/Select happened to return, alphabetically - so the table never
    hides a field this app does not already know the name of."""
    seen = set()
    for r in rows:
        seen.update(r.keys())
    known = [c for c in EXTXSM_FIELD_ORDER if c in seen]
    extra = sorted(c for c in seen if c not in EXTXSM_FIELD_ORDER)
    return known + extra


def _decode_review_hash(hash_val: str) -> dict:
    """
    An AUTH=99 record's HASH isn't an .ionapi like a tenant registration's -
    it's workbook telemetry ({"userName":..., "userDomain":...,
    "computerName":..., "publicIP":..., ...}), so the Review tab shows those
    fields instead of the raw blob. Anything that fails to decode (or doesn't look
    like this shape) is left blank rather than breaking the whole list.
    """
    if not hash_val:
        return {}
    try:
        data = decode_hash_blob(hash_val)
    except M3ApiError:
        return {}
    if not isinstance(data, dict):
        return {}
    return {"USERNAME": data.get("userName") or "",
            "DOMAIN": data.get("userDomain") or "",
            "COMPUTERNAME": data.get("computerName") or "",
            "PUBLICIP": str(data.get("publicIP") or "").strip()}


def _abuse_request(ip: str, key: str) -> requests.Response:
    return requests.get(
        ABUSEIPDB_CHECK_URL,
        headers={"Key": key, "Accept": "application/json"},
        params={"ipAddress": ip, "maxAgeInDays": ABUSEIPDB_MAX_AGE_DAYS},
        timeout=10)


def _abuse_error(resp: requests.Response, body: dict) -> str:
    detail = "; ".join(e.get("detail", "") for e in body.get("errors", []))
    if resp.status_code == 429:
        # The free plan's 1,000 a day. Retry-After is in seconds.
        wait = resp.headers.get("Retry-After")
        detail = "Daily AbuseIPDB limit reached" + (
            f" - try again in {int(wait) // 60} min" if wait and wait.isdigit() else "")
    return f"AbuseIPDB {resp.status_code}: {detail or resp.reason}"


def _abuse_check(ip: str) -> dict:
    """
    One IP's AbuseIPDB result: {"score": 0-100, "reports": n, "country",
    "isp", "usage"} or {"error": "..."}. A private or malformed address is
    answered here without spending a check.
    """
    try:
        addr = ipaddress.ip_address(ip)
    except ValueError:
        return {"error": "Not an IP address."}
    if not addr.is_global:
        return {"error": "Private or reserved address - nothing to check."}

    now = time.time()
    with _abuse_lock:
        hit = _abuse_cache.get(ip)
        if hit and now - hit[0] < ABUSEIPDB_CACHE_SECONDS:
            return hit[1]

    try:
        resp = _abuse_request(ip, app.config["ABUSEIPDB_API_KEY"])
    except requests.RequestException as exc:
        return {"error": f"AbuseIPDB unreachable: {exc}"}
    body = {}
    try:
        body = resp.json()
    except ValueError:
        pass
    if resp.status_code != 200:
        return {"error": _abuse_error(resp, body)}

    data = body.get("data") or {}
    result = {"score": data.get("abuseConfidenceScore"),
              "reports": data.get("totalReports"),
              "country": data.get("countryCode") or "",
              "isp": data.get("isp") or "",
              "usage": data.get("usageType") or ""}
    with _abuse_lock:
        _abuse_cache[ip] = (now, result)
    return result


def _row_key(data: dict) -> tuple[str, str, str]:
    pcid = str(data.get("PCID") or "").strip()
    tnnm = str(data.get("TNNM") or "").strip()
    auth = str(data.get("AUTH") or "").strip()
    missing = [f for f, v in (("PCID", pcid), ("TNNM", tnnm), ("AUTH", auth)) if not v]
    if missing:
        raise M3ApiError(f"{', '.join(missing)} {'is' if len(missing) == 1 else 'are'} required.")
    return pcid, tnnm, auth


def _resolve_tenant_row(client: M3Client, row: dict) -> dict:
    """
    Resolve one Tenants-tab (AUTH=20) row - see resolve_tenant_registration()
    in SheetSecurity_M3Api.py - and write whatever comes back straight to
    the record with a plain UpdUsrInfo through `client` (the currently open
    security tenant this row actually lives on; the target tenant's own
    connection, built inside resolve_tenant_registration(), is only ever
    used to ask it questions, never to write). Called from the "Get Default
    Users" button (api_resolve_tenant_defaults()) for every row whose M3ID
    is still blank, so a row is only ever resolved once - the next run
    skips straight past a non-blank M3ID - and "Assign to tenant" can read
    M3ID straight off this row instead of making its own live lookup.

    A failed write (e.g. a stale record, or M3 unreachable a second time) is
    swallowed rather than raised - the row just keeps its unresolved blank
    M3ID and gets tried again on the next load, same as if this had never
    run. Returns the fields actually written, to merge into the row so the
    browser sees them without a second round trip.
    """
    pcid, tnnm = row.get("PCID", ""), row.get("TNNM", "")
    if not tnnm:
        return {}
    result = resolve_tenant_registration(pcid, tnnm, ionapi_dir=_ionapi_dir(),
                                         security_tenant=client.tenant)
    fields = {}
    if result["error"]:
        fields["UMSG"] = UMSG_CONNECTION_ERROR
    else:
        if result["m3id"]:
            fields["M3ID"] = result["m3id"]
        if result["managing_system"]:
            fields["UMSG"] = UMSG_MANAGING_SYSTEM
    if not fields:
        return {}
    try:
        body = client.update_usr_info(pcid, tnnm, "20", **fields)
        client._records(body, "UpdUsrInfo")
    except M3ApiError:
        return {}
    return fields


@app.route("/")
def index():
    return render_template("SheetSecurity_Index.html")


@app.route("/api/tenants")
def api_tenants():
    """
    The tenant dropdown lists every tenant with a readable .ionapi file in
    the shared ionapi/ folder - not just Managing Systems (tenants that
    carry their own EXT124MI/EXTXSM security table). DOPPIO_DEM is always
    included and always the default, whether or not it has its own .ionapi
    file on disk.
    """
    tenants = {DEFAULT_TENANT}
    tenants.update(
        (e.get("tenant") or "").strip() for e in list_ionapi_files(_ionapi_dir())
        if e.get("tenant"))
    tenants.discard("")

    tenants = sorted(tenants)
    return jsonify(status="success", tenants=tenants,
                   default=DEFAULT_TENANT if DEFAULT_TENANT in tenants
                           else (tenants[0] if tenants else ""),
                   auth_labels=AUTH_LABELS,
                   key_fields=list(EXT124_KEY_FIELDS),
                   error=None)


# Where TNNM links to on the Tenants and Users tabs - the tenant's Ming.le
# portal, by the 'ti' in its .ionapi.
MINGLE_PORTAL_URL = "https://mingle-portal.inforcloudsuite.com/v2/{ti}"


def _tnnm_ti(row: dict, files: list[dict]) -> str:
    """
    The 'ti' a record's TNNM belongs to, or "" when it cannot be told.

    A Tenants-tab (AUTH=20) record's HASH *is* the .ionapi, so its own 'ti'
    is read first. Anything else - and a tenant whose HASH will not decode -
    is matched to an .ionapi file in the shared folder by name, the same
    loose match "Find M3 user" connects with; an ambiguous name gets no link
    rather than a guess.
    """
    if str(row.get("AUTH") or "").strip() == "20" and row.get("HASH"):
        try:
            ti = str(decode_hash_blob(row["HASH"]).get("ti") or "").strip()
        except (M3ApiError, AttributeError):
            ti = ""
        if ti:
            return ti
    hits = match_ionapi_name(row.get("TNNM") or "", files)
    return str(hits[0].get("tenant") or "").strip() if len(hits) == 1 else ""


@app.route("/api/security")
def api_security_list():
    tenant = request.args.get("tenant", "")
    client = _client(tenant)
    rows = client.list_extxsm()
    files = list_ionapi_files(_ionapi_dir())

    out = []
    for r in rows:
        row = {k: ("" if v is None else v) for k, v in r.items()}
        if str(row.get("AUTH") or "").strip() in ("20", "1", "0"):
            ti = _tnnm_ti(row, files)
            if ti:
                row["TNNM_URL"] = MINGLE_PORTAL_URL.format(ti=ti)
        if str(row.get("AUTH") or "").strip() == "99":
            row.update(_decode_review_hash(row.get("HASH")))
        if "HASH" in row and len(row["HASH"]) > HASH_PREVIEW_LEN:
            row["HASH_full_length"] = len(row["HASH"])
            row["HASH"] = row["HASH"][:HASH_PREVIEW_LEN] + "…"
        out.append(row)

    return jsonify(status="success", tenant=tenant, total=len(out),
                   columns=_order_columns(out), rows=out)


@app.route("/api/abuseipdb", methods=["POST"])
def api_abuseipdb():
    """
    AbuseIPDB scores for the Review tab's PUBLICIP column, asked for after
    the list has loaded so a slow or missing AbuseIPDB never holds the list
    up. `configured` is false when there is no API key, and nothing is
    looked up.
    """
    data = request.get_json(force=True) or {}
    ips = sorted({str(ip).strip() for ip in (data.get("ips") or []) if str(ip).strip()})
    if not app.config["ABUSEIPDB_API_KEY"]:
        return jsonify(status="success", configured=False, results={})
    return jsonify(status="success", configured=True,
                   results={ip: _abuse_check(ip) for ip in ips})


@app.route("/api/abuseipdb/key", methods=["GET", "POST"])
def api_abuseipdb_key():
    """
    GET: whether a key is set. POST {key}: check it against AbuseIPDB (one
    lookup of 8.8.8.8, out of the daily allowance) and, if it is accepted,
    save it to ABUSEIPDB_KEY_FILE, readable by this user only. The key is
    never sent back to the browser.
    """
    if request.method == "GET":
        return jsonify(status="success",
                       configured=bool(app.config["ABUSEIPDB_API_KEY"]),
                       key_file=str(ABUSEIPDB_KEY_FILE))
    key = str((request.get_json(force=True) or {}).get("key") or "").strip()
    if not key:
        return jsonify(status="error", message="Paste an API key."), 400
    try:
        resp = _abuse_request("8.8.8.8", key)
    except requests.RequestException as exc:
        return jsonify(status="error", message=f"AbuseIPDB unreachable: {exc}"), 502
    if resp.status_code != 200:
        try:
            body = resp.json()
        except ValueError:
            body = {}
        return jsonify(status="error", message=_abuse_error(resp, body)), 400

    ABUSEIPDB_KEY_FILE.parent.mkdir(parents=True, exist_ok=True)
    fd = os.open(ABUSEIPDB_KEY_FILE, os.O_WRONLY | os.O_CREAT | os.O_TRUNC, 0o600)
    with os.fdopen(fd, "w", encoding="utf-8") as fh:
        fh.write(key + "\n")
    os.chmod(ABUSEIPDB_KEY_FILE, 0o600)
    app.config["ABUSEIPDB_API_KEY"] = key
    with _abuse_lock:
        _abuse_cache.clear()
    return jsonify(status="success", configured=True,
                   key_file=str(ABUSEIPDB_KEY_FILE),
                   remaining=resp.headers.get("X-RateLimit-Remaining"))


@app.route("/api/tenants/resolve-defaults", methods=["POST"])
def api_resolve_tenant_defaults():
    """
    "Get Default Users" - the Tenants tab's own button for what resolving a
    registration (see _resolve_tenant_row()) used to do automatically on
    every load. Runs it for every AUTH=20 row on this tenant whose M3ID is
    still blank; a row that already has one is left alone and counted as
    already resolved rather than touched again.
    """
    data = request.get_json(force=True) or {}
    client = _client(data.get("tenant"))
    rows = [r for r in client.list_extxsm() if str(r.get("AUTH") or "").strip() == "20"]
    pending = [r for r in rows if not str(r.get("M3ID") or "").strip()]

    resolved = managing_system = connection_errors = unresolved = 0
    for r in pending:
        fields = _resolve_tenant_row(client, r)
        if fields.get("UMSG") == UMSG_CONNECTION_ERROR:
            connection_errors += 1
        elif fields.get("M3ID") or fields.get("UMSG") == UMSG_MANAGING_SYSTEM:
            resolved += 1
            if fields.get("UMSG") == UMSG_MANAGING_SYSTEM:
                managing_system += 1
        else:
            unresolved += 1

    return jsonify(status="success", total=len(rows),
                   already_resolved=len(rows) - len(pending), checked=len(pending),
                   resolved=resolved, managing_system=managing_system,
                   connection_errors=connection_errors, unresolved=unresolved)


@app.route("/api/security/get", methods=["POST"])
def api_security_get():
    """Re-read one record - used to pull the full HASH into the edit form."""
    data = request.get_json(force=True) or {}
    client = _client(data.get("tenant"))
    pcid, tnnm, auth = _row_key(data)
    row = client.get_usr_info(pcid, tnnm, auth)
    if row is None:
        return jsonify(status="error", message="Record not found."), 404
    return jsonify(status="success", row=row)


@app.route("/api/security", methods=["POST"])
def api_security_add():
    data = request.get_json(force=True) or {}
    client = _client(data.get("tenant"))
    pcid, tnnm, auth = _row_key(data)
    fields = data.get("fields") or {}
    body = client.add_usr_info(
        pcid, tnnm, auth,
        hash_val=str(fields.get("HASH") or ""),
        m3id=str(fields.get("M3ID") or ""),
        umsg=str(fields.get("UMSG") or ""))
    client._records(body, "AddUsrInfo")
    return jsonify(status="success",
                   message=f"Security record added for {pcid} on {tnnm}.")


@app.route("/api/security", methods=["PUT"])
def api_security_update():
    """
    Update a record's value fields, including AUTH.

    AUTH looks like part of the record's key from the tabs, but it isn't
    (see the module docstring) - EXT124MI matches on PCID+TNNM alone, so
    changing AUTH is a plain UpdUsrInfo just like HASH/M3ID/UMSG. There is
    no need to add a row under the new AUTH and delete the old one, and
    doing so is actively wrong: DelUsrInfo matches on PCID+TNNM too, so it
    would delete the very row the update just changed.
    """
    data = request.get_json(force=True) or {}
    client = _client(data.get("tenant"))
    pcid, tnnm, auth = _row_key(data)
    fields = data.get("fields") or {}
    body = client.update_usr_info(pcid, tnnm, auth, **{
        k: str(fields[k]) for k in ("HASH", "M3ID", "UMSG") if k in fields})
    client._records(body, "UpdUsrInfo")
    return jsonify(status="success",
                   message=f"Security record updated for {pcid} on {tnnm}.")


@app.route("/api/security", methods=["DELETE"])
def api_security_delete():
    data = request.get_json(force=True) or {}
    client = _client(data.get("tenant"))
    pcid, tnnm, auth = _row_key(data)
    body = client.delete_usr_info(pcid, tnnm, auth)
    client._records(body, "DelUsrInfo")
    return jsonify(status="success",
                   message=f"Security record deleted for {pcid} on {tnnm}.")


@app.route("/api/security/bulk", methods=["DELETE"])
def api_security_delete_bulk():
    """
    Delete many records with one batched EXT124MI/DelUsrInfo call - used by
    the Review tab's "Delete all". The browser sends the rows in chunks so it
    can draw a progress bar; each chunk is a single m3api-rest request.
    """
    data = request.get_json(force=True) or {}
    client = _client(data.get("tenant"))
    keys = [_row_key(r) for r in (data.get("rows") or [])]
    errors = client.delete_usr_info_many(keys)
    failed = [{"PCID": k[0], "TNNM": k[1], "error": e}
              for k, e in zip(keys, errors) if e]
    return jsonify(status="success", deleted=len(keys) - len(failed),
                   failed=failed)


@app.route("/api/guess", methods=["POST"])
def api_guess():
    """
    Best-guess USID/email for a PCID, from two sources: MNS150MI on the
    record's own TNNM tenant, and - if one has been imported - the uploaded
    IFS export (see api_ifs_export_upload()). Candidates from both are
    merged and re-sorted by score, each tagged with where it came from
    ("m3" or "ifs_export"), so a login the target tenant doesn't know about
    can still be matched from the IFS export, and vice versa.

    `tenant` is the security tenant currently open in the app (e.g.
    DOPPIO_DEM) - when TNNM has no .ionapi file on disk yet, its own AUTH=20
    EXT124MI record on that tenant is extracted and saved to supply one. If
    that fails entirely (no .ionapi, no AUTH=20 record to extract one from,
    a bad connection, ...) it is not fatal so long as the IFS export has
    something to offer instead - the failure is reported back as `m3_error`
    rather than raised, and only raised for real if there is nothing else to
    fall back on.

    `hint` is an optional override search term - a name or email typed by
    hand for when PCID is not a real login id (or is simply wrong). When
    given, it replaces PCID entirely for the search rather than blending
    with it.
    """
    data = request.get_json(force=True) or {}
    tnnm = (data.get("tnnm") or "").strip()
    pcid = (data.get("pcid") or "").strip()
    hint = (data.get("hint") or "").strip()
    security_tenant = (data.get("tenant") or "").strip() or None
    if not tnnm:
        return jsonify(status="error", message="TNNM (tenant) is required."), 400
    if not pcid and not hint:
        return jsonify(status="error",
                       message="Enter a PCID, or a name/email hint, to look up an M3 user."), 400

    m3_error = None
    try:
        result = guess_m3_user(tnnm, pcid, ionapi_dir=_ionapi_dir(),
                               security_tenant=security_tenant, hint=hint)
    except M3ApiError as exc:
        result = {"tenant": tnnm, "ionapi": None, "extracted": False, "candidates": []}
        m3_error = str(exc)
    for c in result["candidates"]:
        c["source"] = "m3"

    ifs_rows = _ifs_export["rows"]
    if ifs_rows:
        ifs_result = guess_from_ifs_export(pcid, ifs_rows, hint=hint)
        for c in ifs_result["candidates"]:
            c["source"] = "ifs_export"
        result["candidates"] = result["candidates"] + ifs_result["candidates"]
        result["candidates"].sort(key=lambda r: -r["score"])
    elif m3_error:
        raise M3ApiError(m3_error)

    result["best"] = (result["candidates"][0]
                      if result["candidates"] and result["candidates"][0]["score"] >= 0.6
                      else None)
    if m3_error:
        result["m3_error"] = m3_error
    return jsonify(status="success", **result)


# The Users tab's "Auto-match M3 users": a match is only proposed when the
# best candidate scores over this (the same score "Find M3 user" shows), and
# no candidate for a *different* M3 user also clears it - two near-perfect
# answers is a question for a person, not something to pick between.
AUTO_MATCH_THRESHOLD = 0.6


@app.route("/api/users/match", methods=["POST"])
def api_users_match():
    """
    "Find M3 user" for many Users-tab records at once - all on one TNNM,
    which is how the page sends them, so that tenant's MNS150MI user list is
    read once rather than once per record. The uploaded IFS export, if any,
    is scored too, exactly as /api/guess does. Nothing is written; the page
    shows the proposals and sends the ones kept to /api/security/bulk-update.

    Each row comes back with `match` ({usid, email, name, score, source}) when
    it is safe to apply, or `skip` saying why not.
    """
    data = request.get_json(force=True) or {}
    tnnm = str(data.get("tnnm") or "").strip()
    rows = data.get("rows") or []
    threshold = float(data.get("threshold") or AUTO_MATCH_THRESHOLD)
    if not tnnm:
        return jsonify(status="error", message="TNNM is required."), 400

    users, m3_error, ionapi = [], None, None
    try:
        client = M3Client.for_tenant_name(tnnm, _ionapi_dir(),
                                          security_tenant=(data.get("tenant") or None))
        users = client.list_user_data()
        ionapi = client.ionapi_path.name
    except M3ApiError as exc:
        m3_error = str(exc)
    ifs_rows = _ifs_export["rows"]
    if m3_error and not ifs_rows:
        return jsonify(status="success", tnnm=tnnm, m3_error=m3_error, ionapi=None,
                       rows=[{**r, "skip": f"M3 lookup failed: {m3_error}"} for r in rows])

    out = []
    for r in rows:
        pcid = str(r.get("PCID") or "").strip()
        row = {k: r.get(k, "") for k in ("PCID", "TNNM", "AUTH", "M3ID", "UMSG")}
        term = _login_key(pcid)
        if not term:
            out.append({**row, "skip": "No PCID to search by."})
            continue
        cands = [{**c, "source": "m3"} for c in score_m3_users(users, term)]
        if ifs_rows:
            cands += [{**c, "source": "ifs_export"}
                      for c in guess_from_ifs_export(pcid, ifs_rows)["candidates"]]
        cands.sort(key=lambda c: -c["score"])
        if not cands:
            out.append({**row, "skip": "No candidates."})
            continue
        best = cands[0]
        if best["score"] <= threshold:
            out.append({**row, "best": best,
                        "skip": f"Best match only {best['score'] * 100:.0f}%."})
            continue
        rivals = {c["usid"].upper() for c in cands
                  if c["score"] > threshold and c["usid"]} - {best["usid"].upper()}
        # Two users over the bar is settled by the M3ID the record already
        # carries, when it is one of them - somebody chose it on purpose.
        current = str(row["M3ID"] or "").strip().upper()
        if rivals and current and current in rivals | {best["usid"].upper()}:
            best = next(c for c in cands
                        if c["score"] > threshold and c["usid"].upper() == current)
            rivals = set()
        if rivals:
            out.append({**row, "best": best,
                        "skip": "Ambiguous - also over the bar: " + ", ".join(sorted(rivals))})
            continue
        if not best["usid"]:
            out.append({**row, "best": best, "skip": "Best match has no M3 user id."})
            continue
        if (best["usid"] == str(row["M3ID"]).strip()
                and (best["email"] or "") == str(row["UMSG"]).strip()):
            out.append({**row, "best": best, "skip": "Already set."})
            continue
        out.append({**row, "match": best})
    res = {"tnnm": tnnm, "ionapi": ionapi, "rows": out}
    if m3_error:
        res["m3_error"] = m3_error
    return jsonify(status="success", **res)


@app.route("/api/security/bulk-update", methods=["POST"])
def api_security_update_bulk():
    """
    Write many M3ID / UMSG changes with one batched EXT124MI/UpdUsrInfo call
    - the Users tab's auto-match "Apply". Only M3ID and UMSG are taken from
    each row; AUTH is sent as the record already has it, and HASH is left
    alone.
    """
    data = request.get_json(force=True) or {}
    client = _client(data.get("tenant"))
    updates = []
    for r in data.get("rows") or []:
        pcid, tnnm, auth = _row_key(r)
        updates.append({"PCID": pcid, "TNNM": tnnm, "AUTH": auth,
                        "M3ID": str(r.get("M3ID") or "").strip(),
                        "UMSG": str(r.get("UMSG") or "").strip()})
    errors = client.update_usr_info_many(updates)
    failed = [{"PCID": u["PCID"], "TNNM": u["TNNM"], "error": e}
              for u, e in zip(updates, errors) if e]
    return jsonify(status="success", updated=len(updates) - len(failed),
                   failed=failed)


@app.route("/api/ifs-export")
def api_ifs_export_status():
    return jsonify(status="success", filename=_ifs_export["filename"],
                   total=len(_ifs_export["rows"]))


@app.route("/api/ifs-export", methods=["POST"])
def api_ifs_export_upload():
    """
    Import an IFS user export CSV - see parse_ifs_export() /
    guess_from_ifs_export() in SheetSecurity_M3Api.py. Kept in memory only;
    re-uploading replaces it, and there is no per-tenant scoping since one
    IFS instance's export covers every tenant under it.
    """
    file = request.files.get("file")
    if not file or not file.filename:
        return jsonify(status="error", message="No file given."), 400
    raw = file.read()
    try:
        text = raw.decode("utf-8-sig")
    except UnicodeDecodeError:
        text = raw.decode("latin-1")
    rows = parse_ifs_export(text)
    _ifs_export["filename"] = file.filename
    _ifs_export["rows"] = rows
    return jsonify(status="success", filename=file.filename, total=len(rows))


@app.route("/api/ifs-export", methods=["DELETE"])
def api_ifs_export_clear():
    _ifs_export["filename"] = None
    _ifs_export["rows"] = []
    return jsonify(status="success")


@app.route("/api/hash/decode", methods=["POST"])
def api_hash_decode():
    """HASH isn't encrypted, just base64(JSON) - unwrap it for inspection."""
    data = request.get_json(force=True) or {}
    return jsonify(status="success", json=decode_hash_blob(str(data.get("hash") or "")))


@app.route("/api/hash/encode", methods=["POST"])
def api_hash_encode():
    """The inverse of /api/hash/decode - JSON (object or text) back to HASH."""
    data = request.get_json(force=True) or {}
    if "json" not in data or data["json"] in (None, ""):
        return jsonify(status="error", message="No JSON given to encode."), 400
    return jsonify(status="success", hash=encode_hash_blob(data["json"]))


@app.route("/api/extract-ionapi", methods=["POST"])
def api_extract_ionapi():
    """
    Write every AUTH=20 record's HASH out as a real .ionapi file in the
    shared ionapi/ folder. Without confirm=true this is a dry run - the plan
    comes back (write / overwrite / skip, and why) and nothing is touched on
    disk.

    `keys`, when given, is a list of [PCID, TNNM] pairs - only those records
    are written (or considered "would be written" during a dry run);
    everything else still appears in the plan, reported as not selected.
    Omitting it (the initial preview, before anything has been picked) means
    every eligible AUTH=20 record is a candidate.
    """
    data = request.get_json(force=True) or {}
    client = _client(data.get("tenant"))
    overwrite = bool(data.get("overwrite"))
    confirm = bool(data.get("confirm"))
    keys = data.get("keys")
    selected = [(k[0], k[1]) for k in keys] if keys is not None else None
    rows = client.list_usr_info()
    result = extract_type20_ionapi(rows, _ionapi_dir(), overwrite=overwrite,
                                   dry_run=not confirm, selected_keys=selected)
    return jsonify(status="success", ionapi_dir=str(_ionapi_dir()), **result)


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="Doppio API Sheet security front end.")
    ap.add_argument("--host", default="127.0.0.1")
    # 5060/5061 are Chrome's blocked SIP ports; 5057-5059 are the other
    # packages in this repo (m3_security, adp_concur, mig_sync); 5063 is
    # ERP_Concur; 5064 is deploy_packages; 5065 is launcher.
    ap.add_argument("--port", type=int, default=5062)
    ap.add_argument("--ionapi-dir", default=str(DEFAULT_IONAPI_DIR),
                    help="Folder holding the .ionapi files")
    ap.add_argument("--abuseipdb-key", default=None,
                    help="AbuseIPDB API key for the Review tab's IP scores "
                         "(default: $ABUSEIPDB_API_KEY, then "
                         f"{ABUSEIPDB_KEY_FILE})")
    ap.add_argument("--debug", action="store_true")
    args = ap.parse_args(argv)
    if args.abuseipdb_key:
        app.config["ABUSEIPDB_API_KEY"] = args.abuseipdb_key.strip()

    logging.basicConfig(
        level=logging.DEBUG if args.debug else logging.INFO,
        format="%(asctime)s  %(levelname)-8s  %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
        handlers=[logging.StreamHandler(sys.stdout)])

    app.config["SHEET_SECURITY_IONAPI_DIR"] = Path(args.ionapi_dir)
    log.info("ION API  : %s", args.ionapi_dir)
    log.info("AbuseIPDB: %s", "API key set" if app.config["ABUSEIPDB_API_KEY"]
             else "no API key - IPs are linked but not scored")
    log.info("Open     : http://%s:%s", args.host, args.port)
    app.run(host=args.host, port=args.port, debug=args.debug, threaded=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
