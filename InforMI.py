# InforMI.py
import json
import requests
import time
import os
import time

from datetime import datetime, timedelta
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, as_completed
from tqdm import tqdm
from APIBatchLogger import APIBatchLogger
from UserDefaults import load_user_defaults, save_user_defaults

# =============================================================================
# Configuration
# =============================================================================

CONFIG = {
    "company": "",
    "division": "",
    "tenant": None,
    "api_url": "",
    "access_token": "",
    "iu": "",
    "ti": ""
}

# =============================================================================
# User Selection: IONAPI + Company/Division
# =============================================================================
def select_ionapi_file(folder_path):
    defaults = load_user_defaults()

    # ✅ Filter and sort .ionapi files alphabetically (case-insensitive)
    files = sorted([f for f in os.listdir(folder_path) if f.endswith(".ionapi")], key=str.lower)
    if not files:
        raise FileNotFoundError("No .ionapi files found in folder.")

    saved_file = defaults.get("ionapi_file")
    last_prompt = defaults.get("last_ionapi_prompt_time")
    now = datetime.now()

    # ✅ Use saved file if last prompt < 1 hour ago
    if saved_file and saved_file in files and last_prompt:
        last_prompt_dt = datetime.fromisoformat(last_prompt)
        if now - last_prompt_dt < timedelta(hours=1):
            return Path(folder_path) / saved_file

    # Otherwise, ask user
    if saved_file and saved_file in files:
        use_saved = input(f"Use saved ION API file '{saved_file}'? (Y/n): ").strip().lower()
        if use_saved in ["", "y", "yes"]:
            defaults["last_ionapi_prompt_time"] = now.isoformat()
            save_user_defaults(defaults)
            return Path(folder_path) / saved_file

    print("\n🔹 Select an ION API file:")
    for i, file in enumerate(files, 1):
        print(f"{i}. {file}")

    while True:
        try:
            choice = int(input(f"\nEnter your choice (1-{len(files)}): "))
            if 1 <= choice <= len(files):
                selected = files[choice - 1]
                defaults["ionapi_file"] = selected
                defaults["last_ionapi_prompt_time"] = now.isoformat()
                save_user_defaults(defaults)
                return Path(folder_path) / selected
            else:
                print("❌ Invalid choice. Try again.")
        except ValueError:
            print("❌ Please enter a valid number.")


def get_m3_mdprest_base_url():
    """
    Returns the base MDPrest URL for the tenant in the selected ionapi file.
    Assumes CONFIG['tenant'] is the path to the .ionapi file.
    """
    if "tenant" not in CONFIG or not CONFIG["tenant"]:
        raise ValueError("CONFIG['tenant'] is not set. Did you call select_ionapi_file()?")

    ionapi_path = Path(CONFIG["tenant"])
    if not ionapi_path.is_file():
        raise FileNotFoundError(f"ionapi file not found: {CONFIG['tenant']}")

    with open(ionapi_path, "r") as f:
        ionapi_json = json.load(f)

    tenant_id = ionapi_json.get("ti")
    if not tenant_id:
        raise ValueError(f"Tenant ID 'ti' not found in ionapi file {CONFIG['tenant']}")

    ion_api_url = CONFIG.get("ion_api_url", "https://mingle-ionapi.inforcloudsuite.com").rstrip("/")

    return f"{ion_api_url}/{tenant_id}/M3/mdprest/les"


def prompt_for_company_division():
    defaults = load_user_defaults()

    # check last prompt time
    last_prompt = defaults.get("last_prompt_time")
    now = datetime.now()

    if last_prompt:
        last_prompt_dt = datetime.fromisoformat(last_prompt)
        if now - last_prompt_dt < timedelta(hours=1):
            # reuse saved defaults without prompting
            CONFIG["company"] = defaults.get("company", "910")
            CONFIG["division"] = defaults.get("division", "088")
            CONFIG["local_db_name"] = defaults.get("local_db_name", "asr_uat2.db")
            CONFIG["batch_size"] = int(defaults.get("batch_size", 500))
            return  # ✅ skip interactive prompt

    # otherwise prompt as normal
    default_company = defaults.get("company", "100")
    default_division = defaults.get("division", "500")
    default_db_name = defaults.get("local_db_name", "doppio.db")
    default_batch = defaults.get("batch_size", 100)

    db_name = input(f"Enter local database name (default: {default_db_name}): ").strip()
    company = input(f"Enter company code (default: {default_company}): ").strip()
    division_input = input(f"Enter division code (default: {default_division}, enter a single space for blank): ")
    batch_size = input(f"Enter batch size (default: {default_batch}): ").strip()

    if division_input == "":
        division = default_division
    elif division_input.strip() == "":
        division = ""  # explicit blank division
    else:
        division = division_input.strip()

    CONFIG["company"] = company if company else default_company
    CONFIG["division"] = division
    CONFIG["local_db_name"] = db_name if db_name else default_db_name
    CONFIG["batch_size"] = int(batch_size) if batch_size else int(default_batch)

    # save back to defaults
    defaults["company"] = CONFIG["company"]
    defaults["division"] = CONFIG["division"]
    defaults["local_db_name"] = CONFIG["local_db_name"]
    defaults["batch_size"] = CONFIG["batch_size"]
    defaults["last_prompt_time"] = now.isoformat()  # ⏰ save current time

    save_user_defaults(defaults)
    print("")

# =============================================================================
# Authentication
# =============================================================================

def get_ion_token():
    """
    Loads the .ionapi configuration file and fetches an OAuth2 token
    using the Resource Owner Password Credentials (ROPC) flow.
    Sets the access token and API URL in the CONFIG dictionary.
    """
    with open(CONFIG["tenant"], 'r') as f:
        ionapi = json.load(f)
    
    CONFIG['ti'] = ionapi['ti']
    CONFIG["iu"] = ionapi["iu"].rstrip("/")
    
    token_url = ionapi['pu'] + ionapi['ot']
    _divi = CONFIG.get("division", "")
    _divi_param = f"&divi={_divi}" if _divi else ""
    CONFIG['api_url'] = (
        f"{ionapi['iu']}/{ionapi['ti']}/M3/m3api-rest/v2/execute"
        f"?maxrecs=0&extendedresult=true&righttrim=true"
        f"&cono={CONFIG['company']}{_divi_param}"
    )

    payload = {
        'client_id': ionapi['ci'],
        'client_secret': ionapi['cs'],
        'grant_type': 'password',
        'username': ionapi['saak'],
        'password': ionapi['sask']
    }

    headers = {'Content-Type': 'application/x-www-form-urlencoded'}

    try:
        with requests.Session() as client:
            response = client.post(token_url, data=payload, headers=headers)
        response.raise_for_status()
        CONFIG["access_token"] = response.json()['access_token']
        # print(f"✅ Token acquired successfully.")
    except requests.RequestException as e:
        raise RuntimeError(f"Token request failed: {e}")

# =============================================================================
# M3 API Posting
# =============================================================================

# Initialize this in your if __name__ == "__main__": block
logger = None 

def post_to_m3(payload, session, max_retries=3, retry_delay=2):
    program = payload.get("program")
    # Taking the first transaction name for the log record
    tx_name = payload.get("transactions", [{}])[0].get("transaction", "UNKNOWN")
    
    for attempt in range(1, max_retries + 1):
        try:
            headers = {
                "Authorization": f"Bearer {CONFIG['access_token']}",
                "Content-Type": "application/json"
            }
            response = session.post(CONFIG["api_url"], json=payload, headers=headers)
            
            if response.status_code == 401:
                get_ion_token()
                continue

            response.raise_for_status()
            res_json = response.json()
            
            # ✅ LOG SUCCESS (AFTER)
            if logger:
                logger.log(program, tx_name, payload, res_json, "SUCCESS")
                
            return res_json

        except Exception as e:
            # ✅ LOG FAILURE
            if logger:
                logger.log(program, tx_name, payload, {"error": str(e)}, f"FAILED_ATTEMPT_{attempt}")
            
            if attempt == max_retries:
                raise RuntimeError(f"M3 API request failed: {e}")
            
            get_ion_token()
            time.sleep(retry_delay)

def post_batch_to_m3(program, transaction_name, records, selected_columns, batch_size=500):
    for i in range(0, len(records), batch_size):
        batch = records[i:i + batch_size]
        transactions = [
            {
                "transaction": transaction_name,
                "record": rec,
                "selectedColumns": selected_columns
            } for rec in batch
        ]
        payload = {"program": program, "transactions": transactions}
        yield post_to_m3(payload, session=requests.Session())

def post_many_payloads(payloads, max_workers=5):
    results = []
    with requests.Session() as session:
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = {
                executor.submit(post_to_m3, payload, session): payload
                for payload in payloads
            }
            for future in as_completed(futures):
                try:
                    result = future.result()
                    results.append(result)
                except Exception as e:
                    tqdm.write(f"Error with payload {futures[future]}: {e}")
    return results

# =============================================================================
# Initialization: Run Before Anything Else
# =============================================================================

if __name__ == "__main__":
    ionapi_dir = Path(__file__).parent / "ionapi"
    CONFIG["tenant"] = select_ionapi_file(ionapi_dir)
    prompt_for_company_division()
    get_ion_token()

    # ✅ Initialize the Logger
    logger = APIBatchLogger(CONFIG["local_db_name"])
    
    print(f"🚀 Logging active. Saving to {CONFIG['local_db_name']}")