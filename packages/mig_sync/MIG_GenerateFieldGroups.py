# MIG_GenerateFieldGroups.py
"""
MIG_GenerateFieldGroups.py
----------------------
Runs CMS005MI.GenStandard (ST02=1) against a single tenant and waits
for the background job to complete.

Steps
-----
1.  Select ionapi file + company/division and authenticate.
2.  Call CMS005MI.GenStandard → capture the returned BJNO.
3.  Poll MNS320MI.Get every minute:
      · No record returned  →  job completed normally.
      · STAT != "20"        →  job ended in error; display full result.

Usage:
    python MIG_GenerateFieldGroups.py
"""

import time
import requests
from pathlib import Path

from InforMI import (
    CONFIG,
    get_ion_token,
    post_to_m3,
    select_ionapi_file,
    prompt_for_company_division,
)
from UserDefaults import load_user_defaults, save_user_defaults

# =============================================================================
# Constants
# =============================================================================

# How long to wait between job-status polls (seconds).
POLL_INTERVAL = 60

# MNS320MI STAT values that mean "keep waiting".
# 00 = not yet started, 20 = running.
STATUS_WAIT = frozenset({"00", "20"})

# Fields to retrieve when checking background-job status.
JOB_STATUS_FIELDS = [
    "BJNO", "STAT", "JOBQ", "JBPR", "JNA", "JNU",
    "CONO", "DIVI", "RGDT", "RGTM", "LMDT", "LMTM", "CHID",
]


# =============================================================================
# Helpers
# =============================================================================

def norm(v) -> str:
    if v is None:
        return ""
    return str(v).strip()


# =============================================================================
# Step 2 – GenStandard
# =============================================================================

def gen_standard(session: requests.Session) -> str:
    """
    Calls CMS005MI.GenStandard with ST02=1.
    Returns the BJNO string, or raises RuntimeError on failure.
    """
    payload = {
        "program": "CMS005MI",
        "transactions": [{
            "transaction": "GenStandard",
            "record": {"ST02": "1"},
            "selectedColumns": ["BJNO"],
        }],
    }

    headers = {
        "Authorization": f"Bearer {CONFIG['access_token']}",
        "Content-Type": "application/json",
    }

    response = session.post(CONFIG["api_url"], json=payload, headers=headers)
    if response.status_code == 401:
        get_ion_token()
        headers["Authorization"] = f"Bearer {CONFIG['access_token']}"
        response = session.post(CONFIG["api_url"], json=payload, headers=headers)

    response.raise_for_status()
    data = response.json()

    for result in data.get("results", []):
        err = result.get("errorMessage", "").strip()
        if err:
            raise RuntimeError(f"CMS005MI.GenStandard error: {err}")
        for record in result.get("records", []):
            bjno = record.get("BJNO", "").strip()
            if bjno:
                return bjno

    raise RuntimeError("CMS005MI.GenStandard did not return a BJNO.")


# =============================================================================
# Step 3 – Poll MNS320MI.Get until job completes
# =============================================================================

def poll_job_status(bjno: str) -> None:
    """
    Polls MNS320MI.Get for the given BJNO every POLL_INTERVAL seconds.
    A fresh HTTP session is opened for each attempt to avoid stale connections.

    Termination:
      · No records returned      →  completed normally.
      · STAT not in STATUS_WAIT  →  ended in error; display full result.
    """
    print(f"\n⏳  Waiting for background job BJNO={bjno} to complete …")
    print(f"    (polling every {POLL_INTERVAL}s  |  STAT=00 not started, STAT=20 running)\n")

    attempt = 0
    while True:
        attempt += 1
        headers = {
            "Authorization": f"Bearer {CONFIG['access_token']}",
            "Content-Type": "application/json",
        }
        payload = {
            "program": "MNS320MI",
            "transactions": [{
                "transaction": "Get",
                "record": {"BJNO": bjno},
                "selectedColumns": JOB_STATUS_FIELDS,
            }],
        }

        try:
            # Fresh session each attempt — prevents stale-connection errors
            # that occur when the server closes an idle socket after 60 s.
            with requests.Session() as poll_session:
                response = poll_session.post(CONFIG["api_url"], json=payload, headers=headers)
                if response.status_code == 401:
                    get_ion_token()
                    headers["Authorization"] = f"Bearer {CONFIG['access_token']}"
                    response = poll_session.post(CONFIG["api_url"], json=payload, headers=headers)
                response.raise_for_status()
                data = response.json()
        except Exception as exc:
            print(f"  ⚠️   Poll attempt {attempt} failed: {exc}")
            print(f"       Retrying in {POLL_INTERVAL}s …")
            time.sleep(POLL_INTERVAL)
            continue

        # Collect returned records
        records: list[dict] = []
        for result in data.get("results", []):
            for record in result.get("records", []):
                records.append(record)

        if not records:
            # Empty response → job no longer in queue → completed normally
            print(f"  ✅  Job BJNO={bjno} completed normally.")
            return

        rec = records[0]
        stat = norm(rec.get("STAT", ""))

        print(f"  🔄  [{attempt}] STAT={stat:<4}  "
              f"JBPR={norm(rec.get('JBPR','')):<12}  "
              f"JOBQ={norm(rec.get('JOBQ',''))}")

        if stat not in STATUS_WAIT:
            # Job ended with an unexpected status → error
            print(f"\n  ❌  Job BJNO={bjno} ended with STAT={stat}")
            print(f"\n{'─' * 60}")
            for field in JOB_STATUS_FIELDS:
                val = norm(rec.get(field, ""))
                if val:
                    print(f"    {field:<8} : {val}")
            print(f"{'─' * 60}\n")
            return

        print(f"       Job still running. Next check in {POLL_INTERVAL}s …")
        time.sleep(POLL_INTERVAL)


# =============================================================================
# Main
# =============================================================================

def run_gen_standard() -> None:
    ionapi_dir = Path(__file__).parent / "ionapi"

    # Force prompts by removing the "last used within 1 hour" timestamps.
    # Keeps saved values (ionapi file, company, division) as convenient defaults.
    defaults = load_user_defaults()
    defaults.pop("last_ionapi_prompt_time", None)
    defaults.pop("last_prompt_time", None)
    save_user_defaults(defaults)

    # Step 1 – Authenticate
    print(f"\n{'─' * 60}")
    print(f"  📡  CMS005MI.GenStandard")
    print(f"{'─' * 60}")
    input("  Press Enter to select the ionapi file …")
    CONFIG["tenant"] = select_ionapi_file(ionapi_dir)
    prompt_for_company_division()
    get_ion_token()
    print(f"  ✅  Ready: {CONFIG.get('tenant', '')}  "
          f"CONO={CONFIG.get('company', '')}  DIVI={CONFIG.get('division', '')}\n")

    # Step 2 – Submit GenStandard
    print("▶   CMS005MI.GenStandard (ST02=1) …")
    start_time = time.time()
    with requests.Session() as session:
        try:
            bjno = gen_standard(session)
        except Exception as exc:
            print(f"  ❌  {exc}")
            return

    print(f"  ✅  Job submitted. BJNO={bjno}")

    # Step 3 – Poll until done (manages its own sessions internally)
    poll_job_status(bjno)

    elapsed = time.time() - start_time
    hours, rem = divmod(int(elapsed), 3600)
    minutes, seconds = divmod(rem, 60)
    print(f"\n⏱️   Total runtime: {hours:02d}:{minutes:02d}:{seconds:02d} (hh:mm:ss)")


# =============================================================================
# Entry Point
# =============================================================================

if __name__ == "__main__":
    run_gen_standard()
