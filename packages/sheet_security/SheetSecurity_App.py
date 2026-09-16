"""
SheetSecurity_App - Flask front end for the Doppio API Sheet security table.

One page over EXT124MI, the custom M3 program the API Sheet's own access
control lives in: a sortable/filterable list (LstUsrInfo), and add / edit /
delete against a live tenant (AddUsrInfo / UpdUsrInfo / DelUsrInfo). Nothing
is cached locally - every action talks straight to M3, and the table always
shows what LstUsrInfo just returned.

A record's key is (PCID, TNNM, AUTH); HASH / M3ID / UMSG are the value
fields on top of it. HASH is often very long, so the list only ever sends a
truncated copy of it to the browser - GetUsrInfo is called for the full
value when a row is opened for editing.

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
    EXT124_FIELD_ORDER,
    EXT124_KEY_FIELDS,
    M3ApiError,
    M3Client,
    decode_hash_blob,
    encode_hash_blob,
    extract_type20_ionapi,
    guess_m3_user,
    list_ionapi_files,
)

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
    """Known EXT124MI fields first, in a sensible order, then anything else
    LstUsrInfo happened to return, alphabetically - so the table never hides
    a field this app does not already know the name of."""
    seen = set()
    for r in rows:
        seen.update(r.keys())
    known = [c for c in EXT124_FIELD_ORDER if c in seen]
    extra = sorted(c for c in seen if c not in EXT124_FIELD_ORDER)
    return known + extra


def _row_key(data: dict) -> tuple[str, str, str]:
    pcid = str(data.get("PCID") or "").strip()
    tnnm = str(data.get("TNNM") or "").strip()
    auth = str(data.get("AUTH") or "").strip()
    missing = [f for f, v in (("PCID", pcid), ("TNNM", tnnm), ("AUTH", auth)) if not v]
    if missing:
        raise M3ApiError(f"{', '.join(missing)} {'is' if len(missing) == 1 else 'are'} required.")
    return pcid, tnnm, auth


@app.route("/")
def index():
    return render_template("SheetSecurity_Index.html")


@app.route("/api/tenants")
def api_tenants():
    files = list_ionapi_files(_ionapi_dir())
    tenants = sorted({f["tenant"] for f in files if f.get("tenant")})
    return jsonify(status="success", tenants=tenants,
                   default=DEFAULT_TENANT if DEFAULT_TENANT in tenants
                           else (tenants[0] if tenants else ""),
                   auth_labels=AUTH_LABELS,
                   key_fields=list(EXT124_KEY_FIELDS))


@app.route("/api/security")
def api_security_list():
    tenant = request.args.get("tenant", "")
    client = _client(tenant)
    rows = client.list_usr_info()

    out = []
    for r in rows:
        row = {k: ("" if v is None else v) for k, v in r.items()}
        if "HASH" in row and len(row["HASH"]) > HASH_PREVIEW_LEN:
            row["HASH_full_length"] = len(row["HASH"])
            row["HASH"] = row["HASH"][:HASH_PREVIEW_LEN] + "…"
        out.append(row)

    return jsonify(status="success", tenant=tenant, total=len(out),
                   columns=_order_columns(rows), rows=out)


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
    Best-guess USID/email for a PCID, read from the record's own TNNM tenant.

    `tenant` is the security tenant currently open in the app (e.g.
    DOPPIO_DEM) - when TNNM has no .ionapi file on disk yet, its own AUTH=20
    EXT124MI record on that tenant is extracted and saved to supply one.

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
    result = guess_m3_user(tnnm, pcid, ionapi_dir=_ionapi_dir(),
                           security_tenant=security_tenant, hint=hint)
    return jsonify(status="success", **result)


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
