"""
MIG_Api - self-contained auth/HTTP module for the MIG Sync package.

This mirrors the pattern used by M3_Security_M3Api.py and ADP_Concur's own
API module: it does NOT import anything from the repo-root InforMI.py /
UserDefaults.py / APIBatchLogger.py / config.py.  Those modules are built
around input()-prompting CLI tools with a single global CONFIG dict and a
1-hour "reuse the last answer" cache - useful for a human running one script
at a time, wrong for a web app that needs two independent, named tenants
(SOURCE and DEST) alive at once with no prompting at all.

Everything the 9 MIG_Sync* routines need from "the old InforMI.py" is
reimplemented here against an explicit Tenant object instead of a module
global, so two tenants can be authenticated and held in memory side by side.

Deliberate simplification vs. the CLI scripts' InforMI.post_to_m3(): there is
no APIBatchLogger / sqlite call-logging here. The CLI tools logged every M3
call to a local sqlite table; this app does not, by design (see MIG_README.md)
- mig_sync is a thin, no-persistence wrapper, and audit-logging every call
would need a database this package deliberately does not have.
"""

from __future__ import annotations

import json
import time
from dataclasses import dataclass, field
from pathlib import Path

import requests

# ---------------------------------------------------------------------------
# Shared repo-root resource.
#
# ionapi/ is NOT package-local. It is the same folder used by every Infor
# automation tool in this repo (MIG, M3 Security, ADP<->Concur, ...). Do not
# copy .ionapi files into packages/mig_sync/ - read them from here.
#   packages/mig_sync/MIG_Api.py -> parents[2] -> repo root -> /ionapi
# ---------------------------------------------------------------------------
DEFAULT_IONAPI_DIR = Path(__file__).resolve().parents[2] / "ionapi"

DEFAULT_TIMEOUT = 120


class M3ApiError(Exception):
    """Anything that goes wrong talking to M3 (auth, HTTP, or a bad .ionapi)."""


@dataclass
class Tenant:
    """
    One authenticated M3 connection (SOURCE or DEST, or the single tenant a
    one-tenant routine needs). Held server-side, keyed by a uuid4 tenant_id -
    the access_token never goes to the browser.
    """
    ionapi_path: Path
    label: str = ""
    company: str = ""
    division: str = ""
    access_token: str = ""
    iu: str = ""
    ti: str = ""
    global_scope: bool = False

    @property
    def api_url(self) -> str:
        """
        The M3 REST execute URL for this tenant.

        Mirrors InforMI.get_ion_token()'s CONFIG['api_url'] construction:
            {iu}/{ti}/M3/m3api-rest/v2/execute
                ?maxrecs=0&extendedresult=true&righttrim=true&cono=...&divi=...

        global_scope=True (translation data - MBMTRN/MBMTRD are not company or
        division scoped, see MIG_SyncTranslData._global_url() in the original
        CLI script) drops cono/divi from the URL entirely rather than sending
        them blank.
        """
        iu = self.iu.rstrip("/")
        base = f"{iu}/{self.ti}/M3/m3api-rest/v2/execute?maxrecs=0&extendedresult=true&righttrim=true"
        if self.global_scope:
            return base
        divi_param = f"&divi={self.division}" if self.division else ""
        return f"{base}&cono={self.company}{divi_param}"


def list_ionapi_files(ionapi_dir: Path = DEFAULT_IONAPI_DIR) -> list[str]:
    """Sorted (case-insensitive) .ionapi filenames in the shared ionapi dir."""
    d = Path(ionapi_dir)
    if not d.is_dir():
        return []
    return sorted((f.name for f in d.glob("*.ionapi")), key=str.lower)


def authenticate(tenant: Tenant, session: requests.Session | None = None) -> None:
    """
    Read tenant.ionapi_path and do the OAuth2 password-grant POST, exactly as
    InforMI.get_ion_token() does. Sets tenant.access_token, tenant.iu,
    tenant.ti in place. Raises M3ApiError on any failure.
    """
    try:
        ionapi = json.loads(Path(tenant.ionapi_path).read_text())
    except FileNotFoundError as exc:
        raise M3ApiError(f"ionapi file not found: {tenant.ionapi_path}") from exc
    except json.JSONDecodeError as exc:
        raise M3ApiError(f"{tenant.ionapi_path} is not valid JSON: {exc}") from exc

    for key in ("ti", "iu", "pu", "ot", "ci", "cs", "saak", "sask"):
        if not ionapi.get(key):
            raise M3ApiError(
                f"{Path(tenant.ionapi_path).name} is missing '{key}' - it does "
                f"not look like a service-account .ionapi file.")

    tenant.ti = ionapi["ti"]
    tenant.iu = ionapi["iu"].rstrip("/")

    token_url = ionapi["pu"] + ionapi["ot"]
    payload = {
        "client_id": ionapi["ci"],
        "client_secret": ionapi["cs"],
        "grant_type": "password",
        "username": ionapi["saak"],
        "password": ionapi["sask"],
    }
    headers = {"Content-Type": "application/x-www-form-urlencoded"}

    own_session = session is None
    sess = session or requests.Session()
    try:
        resp = sess.post(token_url, data=payload, headers=headers,
                          timeout=DEFAULT_TIMEOUT)
        resp.raise_for_status()
        tenant.access_token = resp.json()["access_token"]
    except requests.RequestException as exc:
        raise M3ApiError(
            f"Token request failed for {tenant.label or tenant.ti} "
            f"({Path(tenant.ionapi_path).name}): {exc}") from exc
    except (KeyError, ValueError) as exc:
        raise M3ApiError(f"Token response had no access_token: {exc}") from exc
    finally:
        if own_session:
            sess.close()


def post_to_m3(tenant: Tenant, payload: dict, session: requests.Session,
               max_retries: int = 3, retry_delay: float = 2) -> dict:
    """
    POST one m3api-rest v2 batch to tenant.api_url.

    Same retry/401-handling behavior as InforMI.post_to_m3(): on HTTP 401,
    re-authenticate and retry, up to max_retries, sleeping retry_delay between
    attempts. Raises M3ApiError if every attempt fails.
    """
    last_exc: Exception | None = None
    for attempt in range(1, max_retries + 1):
        try:
            headers = {
                "Authorization": f"Bearer {tenant.access_token}",
                "Content-Type": "application/json",
            }
            resp = session.post(tenant.api_url, json=payload, headers=headers,
                                 timeout=DEFAULT_TIMEOUT)
            if resp.status_code == 401:
                authenticate(tenant, session)
                continue
            resp.raise_for_status()
            return resp.json()
        except Exception as exc:  # noqa: BLE001 - reported via M3ApiError below
            last_exc = exc
            if attempt == max_retries:
                break
            try:
                authenticate(tenant, session)
            except M3ApiError:
                pass
            time.sleep(retry_delay)
    raise M3ApiError(f"M3 API request failed after {max_retries} attempt(s): {last_exc}")


def upload_file_to_m3(tenant: Tenant, file_path: Path,
                       session: requests.Session) -> tuple[bool, str]:
    """
    PUT file_path to the M3 FileImport area via the File Management REST API,
    exactly as every script's own upload_file_to_m3() does. One retry on 401.
    Returns (True, message) on HTTP 200/201/204, else (False, message).
    """
    file_path = Path(file_path)
    filename = file_path.name
    upload_url = (f"{tenant.iu}/{tenant.ti}"
                  f"/M3/foundation-rest/file-management/v1/file/FileImport/{filename}")
    headers = {
        "Authorization": f"Bearer {tenant.access_token}",
        "Content-Type": "application/octet-stream",
    }
    file_bytes = file_path.read_bytes()

    resp = session.put(upload_url, data=file_bytes, headers=headers,
                        timeout=DEFAULT_TIMEOUT)
    if resp.status_code == 401:
        authenticate(tenant, session)
        headers["Authorization"] = f"Bearer {tenant.access_token}"
        resp = session.put(upload_url, data=file_bytes, headers=headers,
                            timeout=DEFAULT_TIMEOUT)

    if resp.status_code in (200, 201, 204):
        return True, f"Upload succeeded (HTTP {resp.status_code}): {filename}"
    return False, f"Upload failed (HTTP {resp.status_code}): {resp.text[:200]}"


def process_file_in_m3(tenant: Tenant, filename: str,
                        session: requests.Session) -> tuple[bool, str]:
    """
    EVS100MI.ImportFile - tells M3 to process an already-uploaded file.
    Same success/failure semantics as the scripts' own process_file_in_m3().
    """
    payload = {
        "program": "EVS100MI",
        "transactions": [{"transaction": "ImportFile", "record": {"FNAM": filename}}],
    }
    try:
        result = post_to_m3(tenant, payload, session)
    except M3ApiError as exc:
        return False, f"EVS100MI.ImportFile failed: {exc}"

    for res in result.get("results", []):
        err = (res.get("errorMessage") or "").strip() if isinstance(res, dict) else ""
        if err:
            return False, f"EVS100MI.ImportFile error: {err}"
    return True, f"EVS100MI.ImportFile processed successfully: {filename}"
