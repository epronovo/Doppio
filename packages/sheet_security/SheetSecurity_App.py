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
import logging
import sys
import traceback
from pathlib import Path

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
    parse_ifs_export,
    resolve_tenant_registration,
)

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
    "computerName":..., ...}), so the Review tab shows those three fields
    instead of the raw blob. Anything that fails to decode (or doesn't look
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
            "COMPUTERNAME": data.get("computerName") or ""}


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
    The tenant dropdown lists only Managing Systems - tenants that carry
    their own EXT124MI/EXTXSM security table, not every customer tenant a
    Managing System happens to have an .ionapi file for on disk (extracted
    for "Find M3 user" or "Extract Type 20", say). DOPPIO_DEM is the
    bootstrap: its own Tenants tab is where every other Managing System gets
    tagged UMSG="Managing System", once "Get Default Users" has been run
    there (see _resolve_tenant_row() / api_resolve_tenant_defaults()), so
    that tab's AUTH=20 rows are the source of this list rather than
    list_ionapi_files(). DOPPIO_DEM itself is always included and always
    the default, whether or not it happens to carry a self-registration row.
    """
    tenants = {DEFAULT_TENANT}
    error = None
    try:
        rows = _client(DEFAULT_TENANT).list_extxsm()
        tenants.update(
            (r.get("TNNM") or "").strip() for r in rows
            if str(r.get("AUTH") or "").strip() == "20"
            and str(r.get("UMSG") or "").strip() == UMSG_MANAGING_SYSTEM)
        tenants.discard("")
    except M3ApiError as exc:
        error = str(exc)

    tenants = sorted(tenants)
    return jsonify(status="success", tenants=tenants,
                   default=DEFAULT_TENANT if DEFAULT_TENANT in tenants
                           else (tenants[0] if tenants else ""),
                   auth_labels=AUTH_LABELS,
                   key_fields=list(EXT124_KEY_FIELDS),
                   error=error)


@app.route("/api/security")
def api_security_list():
    tenant = request.args.get("tenant", "")
    client = _client(tenant)
    rows = client.list_extxsm()

    out = []
    for r in rows:
        row = {k: ("" if v is None else v) for k, v in r.items()}
        if str(row.get("AUTH") or "").strip() == "99":
            row.update(_decode_review_hash(row.get("HASH")))
        if "HASH" in row and len(row["HASH"]) > HASH_PREVIEW_LEN:
            row["HASH_full_length"] = len(row["HASH"])
            row["HASH"] = row["HASH"][:HASH_PREVIEW_LEN] + "…"
        out.append(row)

    return jsonify(status="success", tenant=tenant, total=len(out),
                   columns=_order_columns(out), rows=out)


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
    ap.add_argument("--debug", action="store_true")
    args = ap.parse_args(argv)

    logging.basicConfig(
        level=logging.DEBUG if args.debug else logging.INFO,
        format="%(asctime)s  %(levelname)-8s  %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
        handlers=[logging.StreamHandler(sys.stdout)])

    app.config["SHEET_SECURITY_IONAPI_DIR"] = Path(args.ionapi_dir)
    log.info("ION API  : %s", args.ionapi_dir)
    log.info("Open     : http://%s:%s", args.host, args.port)
    app.run(host=args.host, port=args.port, debug=args.debug, threaded=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
