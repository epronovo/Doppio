# SEC_GrantAccess.py
"""
SEC_GrantAccess.py
------------------
Security access management utility for M3 multi-tenant environments.

Purpose:
    Manages which Customer Tenants and users are authorized to use the
    Xtend workbook.  All authorization records live in EXTXSM on the
    Master Tenant and are maintained via EXT124MI.AddUsrInfo.

Workflow:
    1. Force-prompt for Master Tenant (always asks — no caching).
    2. Prompt for a partial Customer Tenant name, list matching ionapi
       files, and let the user pick one.
    3. Check whether the Customer Tenant already has an EXAUTH=20 record
       in EXTXSM (Master Tenant).  If not, add it.
    4. Find EXAUTH=99 "pending" records created this month.  These are
       users who ran the workbook but have not yet been granted access.
    5. For each pending record: decode the EXHASH blob, look up the user
       in the Customer Tenant's CMNUSR table, confirm with the operator,
       then call EXT124MI.AddUsrInfo to grant access (AUTH=20).

Reference pattern:
    MIG_SyncTranslData.py — REST API calls / EXPORTMI parsing approach.
    SOURCE  = Master Tenant  (EXTXSM reads, EXT124MI writes)
    DEST    = Customer Tenant (CMNUSR reads)

Usage:
    python SEC_GrantAccess.py
"""

import base64
import json
import os
import re
import requests
from pathlib import Path

from InforMI import (
    CONFIG,
    get_ion_token,
    post_to_m3,
)
from UserDefaults import load_user_defaults, save_user_defaults

# ─────────────────────────────────────────────────────────────────────────────
# Constants
# ─────────────────────────────────────────────────────────────────────────────

EXPORTMI_SEP = "^"


# =============================================================================
# CONFIG snapshot / restore  (switch between Master ↔ Customer tenants)
# =============================================================================

def _snapshot() -> dict:
    return dict(CONFIG)


def _restore(snap: dict) -> None:
    CONFIG.clear()
    CONFIG.update(snap)


# =============================================================================
# URL helpers
# =============================================================================

def _global_url() -> str:
    """Return an M3 API URL with no &cono / &divi (EXTXSM is global)."""
    return (
        f"{CONFIG['iu']}/{CONFIG['ti']}/M3/m3api-rest/v2/execute"
        f"?maxrecs=0&extendedresult=true&righttrim=true"
    )


def _apply_global_url() -> None:
    CONFIG["api_url"] = _global_url()


def _token_global() -> None:
    """Fetch/refresh the ION OAuth token then force the global URL."""
    get_ion_token()
    _apply_global_url()


# =============================================================================
# Step 1 – Master Tenant setup  (always forced — no caching)
# =============================================================================

def _list_ionapi_files(ionapi_dir: Path) -> list[str]:
    """Return sorted list of .ionapi filenames in *ionapi_dir*."""
    files = sorted(
        [f for f in os.listdir(ionapi_dir) if f.endswith(".ionapi")],
        key=str.lower,
    )
    if not files:
        raise FileNotFoundError(f"No .ionapi files found in: {ionapi_dir}")
    return files


def _pick_ionapi(ionapi_dir: Path, label: str, candidates: list[str] | None = None) -> Path:
    """
    Display *candidates* (or all ionapi files when None) and return the
    chosen Path.  Always prompts — no saved default is applied.
    """
    files = candidates if candidates is not None else _list_ionapi_files(ionapi_dir)

    print(f"\n  Select the {label} ION API file:")
    for i, f in enumerate(files, start=1):
        print(f"    {i:2}. {f}")

    while True:
        try:
            choice = int(input(f"\n  Enter choice (1-{len(files)}): "))
            if 1 <= choice <= len(files):
                return ionapi_dir / files[choice - 1]
            print("  ❌  Invalid choice — try again.")
        except ValueError:
            print("  ❌  Please enter a number.")


def setup_master_tenant(ionapi_dir: Path) -> dict:
    """
    Always prompt the user to select the Master Tenant ionapi file.
    Returns a CONFIG snapshot that can be restored later.
    """
    print(f"\n{'─' * 60}")
    print(f"  📡  Step 1 – Configure MASTER Tenant")
    print(f"{'─' * 60}")
    print("  (All EXT124MI writes and EXTXSM reads target this tenant.)")

    ionapi_path = _pick_ionapi(ionapi_dir, "MASTER Tenant")
    CONFIG["tenant"]   = str(ionapi_path)
    CONFIG["company"]  = "100"
    CONFIG["division"] = ""
    _token_global()

    snap = _snapshot()
    stem = Path(ionapi_path).stem
    print(f"\n  ✅  Master Tenant ready: {stem}  (ti={CONFIG['ti']})")
    return snap


# =============================================================================
# Step 2 – Customer Tenant selection (partial-name filter on ionapi files)
# =============================================================================

def select_customer_tenant(ionapi_dir: Path) -> tuple[Path, str]:
    """
    Prompt for a partial Customer Tenant name, filter matching ionapi files,
    and return (selected_path, tenant_stem).

    *tenant_stem* is the ionapi filename without the .ionapi extension
    (e.g. "BETTERBEING_PRD") and is used as TNNM / PCID in API calls.
    """
    print(f"\n{'─' * 60}")
    print(f"  📡  Step 2 – Select Customer Tenant")
    print(f"{'─' * 60}")

    partial = input("  Enter Customer Tenant (partial name): ").strip().upper()
    if not partial:
        raise ValueError("Customer Tenant name cannot be empty.")

    all_files = _list_ionapi_files(ionapi_dir)
    matches   = [f for f in all_files if partial in f.upper()]

    if not matches:
        raise ValueError(f"No ionapi files found matching '{partial}'.")

    print(f"\n  Tenants matching '{partial}':")
    for i, f in enumerate(matches, start=1):
        print(f"    {i:2}. {f}  →  stem: {Path(f).stem}")

    if len(matches) == 1:
        selected_file = matches[0]
        print(f"\n  (Auto-selected — only one match: {selected_file})")
    else:
        selected_path_tmp = _pick_ionapi(ionapi_dir, f"Customer Tenant (match for '{partial}')", matches)
        selected_file = selected_path_tmp.name

    selected_path = ionapi_dir / selected_file
    tenant_stem   = Path(selected_file).stem.replace("_", " ")   # e.g. "BETTERBEING PRD"

    print(f"\n  ✅  Customer Tenant selected: {selected_file}")
    return selected_path, tenant_stem


def setup_customer_tenant(ionapi_path: Path) -> dict:
    """
    Point CONFIG at the Customer Tenant ionapi file and return a snapshot.
    Used later for CMNUSR user-lookup queries.
    """
    with open(ionapi_path, "r") as fh:
        ionapi_data = json.load(fh)

    CONFIG["tenant"]   = str(ionapi_path)
    CONFIG["company"]  = ionapi_data.get("company", "100")
    CONFIG["division"] = ionapi_data.get("division", "")
    _token_global()

    snap = _snapshot()
    print(f"  ✅  Customer Tenant configured: {ionapi_path.stem}  (ti={CONFIG['ti']})")
    return snap


# =============================================================================
# EXPORTMI.Select — generic helper (mirrors MIG_SyncTranslData pattern)
# =============================================================================

def _fetch_exportmi(
    session: requests.Session,
    query:   str,
    label:   str,
) -> list[dict]:
    """
    POST EXPORTMI.Select with *query* and return a list of row-dicts.
    The first REPL row is the column-name header (HDRS=1).
    Handles 401 token-refresh automatically.
    """
    payload = {
        "program": "EXPORTMI",
        "transactions": [{
            "transaction": "Select",
            "record": {
                "QERY": query,
                "SEPC": EXPORTMI_SEP,
                "HDRS": "1",
            },
            "selectedColumns": ["QERY", "SEPC", "HDRS", "REPL"],
        }],
    }

    headers = {
        "Authorization": f"Bearer {CONFIG['access_token']}",
        "Content-Type": "application/json",
    }

    resp = session.post(CONFIG["api_url"], json=payload, headers=headers)
    if resp.status_code == 401:
        _token_global()
        headers["Authorization"] = f"Bearer {CONFIG['access_token']}"
        resp = session.post(CONFIG["api_url"], json=payload, headers=headers)

    resp.raise_for_status()
    data = resp.json()

    repl_rows: list[str] = []
    for result in data.get("results", []):
        for record in result.get("records", []):
            val = record.get("REPL", "")
            if val:
                repl_rows.append(val)

    if not repl_rows:
        print(f"  ℹ️   No rows returned from EXPORTMI ({label}).")
        return []

    # Row 0 is the header
    col_names = [c.strip() for c in repl_rows[0].rstrip(EXPORTMI_SEP).split(EXPORTMI_SEP)]

    parsed: list[dict] = []
    for raw in repl_rows[1:]:
        values = raw.rstrip(EXPORTMI_SEP).split(EXPORTMI_SEP)
        values += [""] * (len(col_names) - len(values))
        row = {col_names[idx]: values[idx].strip() for idx in range(len(col_names))}
        parsed.append(row)

    print(f"  ✅  {len(parsed):,} record(s) retrieved ({label}).")
    return parsed


# =============================================================================
# EXT124MI.AddUsrInfo — generic helper
# =============================================================================

def _add_usr_info(
    session:  requests.Session,
    pcid:     str,
    tnnm:     str,
    auth:     str   = "20",
    hash_val: str   = "",
    m3id:     str   = "",
) -> bool:
    """
    Call EXT124MI.AddUsrInfo on the currently-active tenant (Master).
    Returns True on success, False on any error.
    """
    record: dict = {
        "PCID": pcid,
        "TNNM": tnnm,
        "AUTH": auth,
    }
    if hash_val:
        record["HASH"] = hash_val
    if m3id:
        record["M3ID"] = m3id

    payload = {
        "program": "EXT124MI",
        "transactions": [{
            "transaction": "AddUsrInfo",
            "record": record,
        }],
    }

    try:
        result = post_to_m3(payload, session)
        for res in result.get("results", []):
            err = res.get("errorMessage", "").strip() if isinstance(res, dict) else ""
            if err:
                print(f"  ❌  EXT124MI.AddUsrInfo error: {err}")
                return False
        tag = f"M3ID={m3id}" if m3id else f"HASH={'(encoded)' if hash_val else '—'}"
        print(f"  ✅  EXT124MI.AddUsrInfo → PCID={pcid}  TNNM={tnnm}  AUTH={auth}  {tag}")
        return True
    except Exception as exc:
        print(f"  ❌  EXT124MI.AddUsrInfo failed: {exc}")
        return False


# =============================================================================
# EXHASH encode / decode helpers
# =============================================================================

def _encode_ionapi(ionapi_path: Path) -> str:
    """
    Parse the ionapi file as JSON, re-serialize it to a compact string,
    then base64-encode that string as UTF-8.

    Unwrapping through json.load / json.dumps normalises whitespace and
    ensures the encoded blob contains clean JSON regardless of how the
    original file was formatted.
    """
    with open(ionapi_path, "r", encoding="utf-8") as fh:
        data = json.load(fh)
    text = json.dumps(data, separators=(",", ":"))
    return base64.b64encode(text.encode("utf-8")).decode("utf-8")


def _decode_exhash(exhash: str) -> dict:
    """
    Base64-decode an EXHASH value and parse as JSON.

    Mirrors the VBA Base64Decode padding logic exactly:
        Do While Len(encodedText) Mod 4 <> 0
            encodedText = encodedText & "="
        Loop
    In Python:  exhash + "=" * (-len(exhash) % 4)

    Known quirk – Windows user-profile paths (e.g. "C:\\Users\\JTippets")
    are stored in the JSON blob with bare backslashes, which are technically
    invalid JSON escape sequences.  The VBA side never calls a JSON parser so
    it never notices; Python's json.loads does.  We handle this by catching
    JSONDecodeError and escaping any bare backslash that is not already part
    of a legal JSON escape sequence before retrying.

    Legal JSON escapes:  \\\\ \\" \\/ \\b \\f \\n \\r \\t \\uXXXX
    Any other \\X is escaped to \\\\X.

    Example decoded payload:
        {"userName":"JTippets", "userDomain":"NUTRACORP",
         "computerName":"JTIPP-12070", "localIP":"192.168.23.122",
         "publicIP":"161.115.98.67", "userProfile":"C:\\Users\\JTippets",
         "osName":"Microsoft Windows 11 Pro", "osVersion":"10.0.26100",
         "sheetVersion":"v2.02"}

    Returns an empty dict on any unrecoverable failure (with a diagnostic
    print so the caller can see the raw value).
    """
    if not exhash:
        return {}

    # ── Step 1: pad to next multiple of 4 (mirrors VBA Do-While) ─────────── #
    padded = exhash + "=" * (-len(exhash) % 4)

    # ── Step 2: base64-decode bytes → UTF-8 text ─────────────────────────── #
    try:
        raw_bytes = base64.b64decode(padded)
    except Exception as exc:
        print(f"  ⚠️   _decode_exhash: base64 decode failed — {exc}")
        print(f"        raw value: {exhash[:60]}…")
        return {}

    for encoding in ("utf-8", "latin-1"):
        try:
            text = raw_bytes.decode(encoding)
            break
        except Exception:
            continue
    else:
        text = raw_bytes.decode("utf-8", errors="replace")

    # ── Step 3: parse JSON — with backslash-escape fix on first failure ───── #
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        # Windows paths like C:\Users\Name contain bare backslashes that are
        # invalid JSON.  Escape any \ not already followed by a legal
        # JSON escape character: \ " / b f n r t u
        fixed = re.sub(r'\\(?![\\"/bfnrtu])', r'\\\\', text)
        try:
            return json.loads(fixed)
        except json.JSONDecodeError as exc:
            print(f"  ⚠️   _decode_exhash: JSON parse failed after backslash-fix — {exc}")
            print(f"        raw text (first 120 chars): {text[:120]}")
            return {}


# =============================================================================
# Step 3 – Check / add Customer Tenant in EXTXSM (EXAUTH=20)
# =============================================================================

def check_and_add_customer_tenant(
    session:       requests.Session,
    master_snap:   dict,
    customer_path: Path,
    tenant_stem:   str,
) -> str:
    """
    Verify an EXAUTH=20 record exists for *tenant_stem* in EXTXSM.
    If missing, add it via EXT124MI.AddUsrInfo with the encoded ionapi.
    Always runs against the Master Tenant.

    Returns the EXPCID from the EXAUTH=20 record — this is the value to use
    as TNNM when granting individual user access.
    """
    _restore(master_snap)
    _apply_global_url()

    print(f"\n{'─' * 60}")
    print(f"  🔍  Step 3 – Check EXTXSM for Customer Tenant: {tenant_stem}")
    print(f"{'─' * 60}")

    query = (
        f"EXPCID from EXTXSM "
        f"where EXTNNM = '{tenant_stem}' and EXAUTH = 20"
    )
    rows = _fetch_exportmi(session, query, label="EXTXSM EXAUTH=20 check")

    if rows:
        expcid = rows[0].get("EXPCID", tenant_stem)
        print(f"  ✅  Customer Tenant '{tenant_stem}' already registered "
              f"(EXPCID={expcid}, EXAUTH=20).  No action needed.")
        return expcid

    # ── Not found — add it ───────────────────────────────────────────────── #
    print(f"  ⚠️   Customer Tenant '{tenant_stem}' NOT found in EXTXSM.  Adding…")

    with open(customer_path, "r") as fh:
        ionapi_data = json.load(fh)

    pcid     = tenant_stem   # both PCID and TNNM use the space-version stem
    hash_val = _encode_ionapi(customer_path)

    _restore(master_snap)
    _apply_global_url()

    success = _add_usr_info(
        session  = session,
        pcid     = pcid,
        tnnm     = tenant_stem,
        auth     = "20",
        hash_val = hash_val,
    )

    if not success:
        print(f"  ❌  Could not register Customer Tenant '{tenant_stem}' in EXTXSM.")

    return pcid


# =============================================================================
# Step 4 – Fetch EXAUTH=99 pending records (current month)
# =============================================================================

def fetch_pending_99_records(
    session:     requests.Session,
    master_snap: dict,
) -> list[dict]:
    """
    Query EXTXSM on the Master Tenant for EXAUTH=99 records created
    on or after the first day of the current month (yyyymm00 format).

    Returns a list of dicts with keys EXPCID, EXTNNM, EXHASH.
    """
    _restore(master_snap)
    _apply_global_url()

    print(f"\n{'─' * 60}")
    print(f"  🔍  Step 4 – Pending '99' records")
    print(f"{'─' * 60}")

    query = (
        f"EXPCID,EXTNNM,EXHASH from EXTXSM "
        f"where EXAUTH = 99"
    )
    rows = _fetch_exportmi(session, query, label="EXTXSM EXAUTH=99 pending")

    if not rows:
        return []

    # ── Display summary table ─────────────────────────────────────────────── #
    print(f"\n  {'#':<4} {'EXPCID':<25} {'EXTNNM':<30} {'User':<20} {'Domain':<20}")
    print(f"  {'─'*4} {'─'*25} {'─'*30} {'─'*20} {'─'*20}")
    for idx, row in enumerate(rows, start=1):
        info   = _decode_exhash(row.get("EXHASH", ""))
        user   = info.get("userName",   "?")
        domain = info.get("userDomain", "?")
        print(f"  {idx:<4} {row.get('EXPCID',''):<25} "
              f"{row.get('EXTNNM',''):<30} {user:<20} {domain:<20}")

    return rows


# =============================================================================
# Step 5a – Look up TNNM for the EXAUTH=20 record matching EXPCID
# =============================================================================

def get_tnnm_for_pcid(
    session:     requests.Session,
    master_snap: dict,
    expcid:      str,
) -> str:
    """
    Retrieve EXTNNM from the EXAUTH=20 EXTXSM record whose EXPCID matches
    *expcid*.  Falls back to *expcid* itself if no record is found.
    """
    _restore(master_snap)
    _apply_global_url()

    query = (
        f"EXTNNM from EXTXSM "
        f"where EXPCID = '{expcid}' and EXAUTH = 20"
    )
    rows = _fetch_exportmi(session, query, label=f"EXTXSM TNNM lookup (EXPCID={expcid})")
    return rows[0].get("EXTNNM", expcid) if rows else expcid


# =============================================================================
# Step 5b – Look up M3 user in Customer Tenant (CMNUSR / JUTX40)
# =============================================================================

def lookup_m3_user(
    session:       requests.Session,
    customer_snap: dict,
    search_term:   str,
) -> dict | None:
    """
    Switch to Customer Tenant, retrieve CMNUSR records, filter by
    *search_term* inside JUTX40 (case-insensitive partial match), display
    the results and let the operator pick one.

    Returns a dict with keys JUTX40 and JUUSID, or None if no match found.
    """
    _restore(customer_snap)
    _apply_global_url()

    print(f"\n  🔍  Looking up M3 users matching '{search_term}' "
          f"in Customer Tenant (CMNUSR)…")

    query   = "JUTX40,JUUSID from CMNUSR"
    all_usr = _fetch_exportmi(session, query, label="CMNUSR all users")

    if not all_usr:
        print("  ⚠️   No CMNUSR records returned from Customer Tenant.")
        return None

    needle  = search_term.lower()
    matches = [r for r in all_usr if needle in r.get("JUTX40", "").lower()]

    if not matches:
        print(f"  ⚠️   No users found in JUTX40 matching '{search_term}'.")
        # Offer manual search
        again = input("  Enter a different search term (or press Enter to skip): ").strip()
        if not again:
            return None
        needle  = again.lower()
        matches = [r for r in all_usr if needle in r.get("JUTX40", "").lower()]
        if not matches:
            print(f"  ⚠️   Still no match for '{again}'. Skipping this record.")
            return None

    print(f"\n  {'#':<4} {'JUTX40 (Full Name)':<40} {'JUUSID':<20}")
    print(f"  {'─'*4} {'─'*40} {'─'*20}")
    for idx, row in enumerate(matches, start=1):
        print(f"  {idx:<4} {row.get('JUTX40',''):<40} {row.get('JUUSID',''):<20}")

    if len(matches) == 1:
        print(f"\n  (Auto-selected — only one match.)")
        return matches[0]

    while True:
        try:
            pick = int(input(f"\n  Select user (1-{len(matches)}): "))
            if 1 <= pick <= len(matches):
                return matches[pick - 1]
            print("  ❌  Invalid choice — try again.")
        except ValueError:
            print("  ❌  Please enter a number.")


# =============================================================================
# Step 5 – Select and process one pending '99' record at a time
# =============================================================================

def _process_one(
    session:       requests.Session,
    master_snap:   dict,
    customer_snap: dict,
    row:           dict,
    tnnm:          str,
) -> str:
    """
    Process a single pending record.  Returns 'granted', 'skipped', or 'failed'.
    """
    expcid = row.get("EXPCID", "")
    extnnm = row.get("EXTNNM", "")
    exhash = row.get("EXHASH", "")

    info        = _decode_exhash(exhash)
    user_name   = info.get("userName",     "")
    user_domain = info.get("userDomain",   "")
    computer    = info.get("computerName", "")
    local_ip    = info.get("localIP",      "")
    public_ip   = info.get("publicIP",     "")
    sheet_ver   = info.get("sheetVersion", "")

    print(f"\n  ┌─ Selected pending request ──────────────────────────────")
    print(f"  │  EXPCID      : {expcid}")
    print(f"  │  EXTNNM      : {extnnm}")
    print(f"  │  User        : {user_name}  @  {user_domain}")
    print(f"  │  Computer    : {computer}")
    print(f"  │  Local IP    : {local_ip}   Public IP: {public_ip}")
    print(f"  │  Sheet ver   : {sheet_ver}")
    print(f"  └──────────────────────────────────────────────────────────")

    if not user_name:
        print("  ⚠️   Cannot decode user from EXHASH.")
        return "skipped"

    # ── Look up the M3 user in Customer Tenant ────────────────────────── #
    m3_user = lookup_m3_user(session, customer_snap, user_name)

    if not m3_user:
        print(f"  ⚠️   No M3 user resolved for '{user_name}'.")
        return "skipped"

    juusid = m3_user.get("JUUSID", "")
    jutx40 = m3_user.get("JUTX40", "")
    print(f"\n  ✅  Resolved M3 user: {jutx40}  (USID: {juusid})")

    # ── Confirm with operator ─────────────────────────────────────────── #
    # PCID  = EXPCID from the 99 record (the user's Windows login name)
    # TNNM  = EXTNNM of the EXAUTH=20 customer tenant record (passed in)
    # M3ID  = JUUSID from CMNUSR
    # AUTH  = 1  (user-level grant)
    print(f"\n  ── Proposed EXT124MI.AddUsrInfo ─────────────────────────")
    print(f"     PCID  = {expcid}")
    print(f"     TNNM  = {tnnm}")
    print(f"     M3ID  = {juusid}")
    print(f"     AUTH  = 1")
    print(f"  ─────────────────────────────────────────────────────────")

    confirm = input("  Grant access? [y/N]: ").strip().lower()
    if confirm != "y":
        print("  ℹ️   Skipped.")
        return "skipped"

    # ── Switch back to Master Tenant and write the record ─────────────── #
    _restore(master_snap)
    _apply_global_url()

    success = _add_usr_info(
        session = session,
        pcid    = expcid,
        tnnm    = tnnm,
        auth    = "1",
        m3id    = juusid,
    )

    if success:
        print(f"  ✅  Access granted: {user_name} → {juusid} on tenant {expcid}.")
        return "granted"
    else:
        print(f"  ❌  Failed to grant access for {user_name}.")
        return "failed"


def process_pending_records(
    session:       requests.Session,
    master_snap:   dict,
    customer_snap: dict,
    tenant_stem:   str,
) -> None:
    """
    Fetch pending EXAUTH=99 records, display the numbered list, and let the
    operator pick one to process.  Loops until the operator chooses to quit.

    *tenant_stem* is the EXTNNM from the EXAUTH=20 customer tenant record
    (e.g. "BETTERBEING_PRD") and is passed as TNNM to EXT124MI.AddUsrInfo.
    """
    pending = fetch_pending_99_records(session, master_snap)

    if not pending:
        print(f"\n  ✅  No pending access requests this month.  Nothing to do.")
        return

    granted = skipped = failed = 0

    while True:
        # ── Re-display the current list ───────────────────────────────── #
        print(f"\n{'─' * 60}")
        print(f"  Step 5 – Pending access requests  ({len(pending)} total)")
        print(f"{'─' * 60}")
        print(f"  {'#':<4} {'EXPCID':<25} {'EXTNNM':<30} {'User':<20} {'Domain':<15}")
        print(f"  {'─'*4} {'─'*25} {'─'*30} {'─'*20} {'─'*15}")
        for idx, row in enumerate(pending, start=1):
            info   = _decode_exhash(row.get("EXHASH", ""))
            user   = info.get("userName",   "?")
            domain = info.get("userDomain", "?")
            print(f"  {idx:<4} {row.get('EXPCID',''):<25} "
                  f"{row.get('EXTNNM',''):<30} {user:<20} {domain:<15}")
        print(f"  {'─'*4} {'─'*25} {'─'*30} {'─'*20} {'─'*15}")

        # ── Prompt for selection ──────────────────────────────────────── #
        raw = input(f"\n  Select a record to process (1-{len(pending)}), "
                    f"or press Enter to quit: ").strip()
        if not raw:
            break

        try:
            pick = int(raw)
            if not (1 <= pick <= len(pending)):
                print("  ❌  Number out of range — try again.")
                continue
        except ValueError:
            print("  ❌  Please enter a number.")
            continue

        result = _process_one(session, master_snap, customer_snap, pending[pick - 1], tnnm=tenant_stem)

        if result == "granted":
            granted += 1
            # Remove from list so it no longer appears as pending
            pending.pop(pick - 1)
        elif result == "skipped":
            skipped += 1
        else:
            failed += 1

        if not pending:
            print("\n  ✅  All pending requests have been processed.")
            break

    # ── Session summary ───────────────────────────────────────────────── #
    if granted or skipped or failed:
        print(f"\n{'═' * 60}")
        print(f"  Session summary")
        print(f"{'═' * 60}")
        print(f"  Granted  : {granted}")
        print(f"  Skipped  : {skipped}")
        print(f"  Failed   : {failed}")
        print(f"{'═' * 60}")


# =============================================================================
# Main entry point
# =============================================================================

def grant_access() -> None:
    ionapi_dir = Path(__file__).parent / "ionapi"

    print(f"\n{'═' * 60}")
    print(f"  SEC_GrantAccess – M3 Tenant Access Management")
    print(f"{'═' * 60}")

    # ── Step 1: Master Tenant ──────────────────────────────────────────── #
    master_snap = setup_master_tenant(ionapi_dir)

    # ── Step 2: Customer Tenant ────────────────────────────────────────── #
    customer_path, tenant_stem = select_customer_tenant(ionapi_dir)

    # Configure Customer Tenant (save snapshot for CMNUSR lookups later)
    _restore(master_snap)   # ensure clean state before switching
    customer_snap = setup_customer_tenant(customer_path)

    # ── Step 3–5: All API calls ────────────────────────────────────────── #

    with requests.Session() as session:

        # Step 3 – verify / register Customer Tenant in EXTXSM.
        check_and_add_customer_tenant(
            session       = session,
            master_snap   = master_snap,
            customer_path = customer_path,
            tenant_stem   = tenant_stem,
        )

        # Step 4a – test grant with operator's own ID before processing the queue.
        print(f"\n{'─' * 60}")
        print(f"  Step 4a – Test access grant with operator ID")
        print(f"{'─' * 60}")

        defaults     = load_user_defaults()
        default_pcid = defaults.get("operator_pcid", "ericpronovost")
        raw_pcid     = input(f"  Operator PCID (default: {default_pcid}): ").strip()
        operator_pcid = raw_pcid if raw_pcid else default_pcid

        defaults["operator_pcid"] = operator_pcid
        save_user_defaults(defaults)

        # Look up the operator's M3 user in the Customer Tenant
        m3_user = lookup_m3_user(session, customer_snap, operator_pcid)

        if m3_user:
            juusid = m3_user.get("JUUSID", "")
            jutx40 = m3_user.get("JUTX40", "")

            print(f"\n  ── Proposed test EXT124MI.AddUsrInfo ────────────────")
            print(f"     PCID  = {operator_pcid}")
            print(f"     TNNM  = {tenant_stem}")
            print(f"     M3ID  = {juusid}")
            print(f"     AUTH  = 1")
            print(f"  ─────────────────────────────────────────────────────")

            confirm = input("  Proceed with test grant? [y/N]: ").strip().lower()
            if confirm == "y":
                _restore(master_snap)
                _apply_global_url()
                _add_usr_info(
                    session = session,
                    pcid    = operator_pcid,
                    tnnm    = tenant_stem,
                    auth    = "1",
                    m3id    = juusid,
                )
            else:
                print("  ℹ️   Test skipped.")
        else:
            print(f"  ⚠️   No M3 user found for '{operator_pcid}' — skipping test.")


        # ── Step 4b – Force prompt for a second ID (chetan) ────────────────── #
        print(f"\n{'─' * 60}")
        print(f"  Step 4b – Force test grant for second ID")
        print(f"{'─' * 60}")

        raw_second = input(f"  Second PCID to process (default: chetan): ").strip()
        second_pcid = raw_second if raw_second else "chetan"

        m3_user_2 = lookup_m3_user(session, customer_snap, second_pcid)

        if m3_user_2:
            juusid_2 = m3_user_2.get("JUUSID", "")
            jutx40_2 = m3_user_2.get("JUTX40", "")

            print(f"\n  ── Proposed test EXT124MI.AddUsrInfo ────────────────")
            print(f"     PCID  = {second_pcid}")
            print(f"     TNNM  = {tenant_stem}")
            print(f"     M3ID  = {juusid_2}")
            print(f"     AUTH  = 1")
            print(f"  ─────────────────────────────────────────────────────")

            confirm_2 = input("  Proceed with second test grant? [y/N]: ").strip().lower()
            if confirm_2 == "y":
                _restore(master_snap)
                _apply_global_url()
                _add_usr_info(
                    session = session,
                    pcid    = second_pcid,
                    tnnm    = tenant_stem,
                    auth    = "1",
                    m3id    = juusid_2,
                )
            else:
                print("  ℹ️   Second test skipped.")
        else:
            print(f"  ⚠️   No M3 user found for '{second_pcid}' — skipping second test.")


        # Step 5 – pending '99' records → grant access.
        # TNNM = tenant_stem (e.g. "NUTRACORP TST") — the _ → space stem,
        # which is what was stored as TNNM on the EXAUTH=20 tenant record.
        process_pending_records(
            session       = session,
            master_snap   = master_snap,
            customer_snap = customer_snap,
            tenant_stem   = tenant_stem,
        )

    print(f"\n  SEC_GrantAccess complete.\n")


# =============================================================================
# Entry point
# =============================================================================

if __name__ == "__main__":
    grant_access()
