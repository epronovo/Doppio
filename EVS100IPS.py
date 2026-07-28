# =============================================================================
# Imports
# =============================================================================
import sys
import os
import re
import json
import time
import subprocess
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, as_completed

import requests
import pandas as pd
from tqdm import tqdm

from UserDefaults import load_user_defaults, save_user_defaults

# =============================================================================
# Configuration
# =============================================================================
CONFIG = {
    "company": "",
    "division": "",
    "tenant": None,
    "api_url": "",
    "access_token": ""
}

# SOAP_TEMPLATE_PATH = Path(__file__).parent / "ips" / "MPD_PPS044_ADD.ips"
# SOAP_ENDPOINT = "https://mingle-ionapi.inforcloudsuite.com/W8DY5WFWZLXDHJPA_PRD/M3/ips/service/MPD_PPS044_Add"
SOAP_TEMPLATE_PATH = Path(__file__).parent / "ips" / "MPD_PPS044_UPD.ips"
SOAP_ENDPOINT = "https://mingle-ionapi.inforcloudsuite.com/W8DY5WFWZLXDHJPA_PRD/M3/ips/service/MPD_PPS044_Upd"
LOG_DIR = Path(__file__).parent / "soap_logs"
LOG_DIR.mkdir(exist_ok=True)

# =============================================================================
# User Selection: IONAPI + Company/Division
# =============================================================================
def select_ionapi_file(folder_path):
    defaults = load_user_defaults()
    files = [f for f in os.listdir(folder_path) if f.endswith(".ionapi")]
    if not files:
        raise FileNotFoundError("No .ionapi files found in folder.")

    saved_file = defaults.get("ionapi_file")
    if saved_file and saved_file in files:
        return Path(folder_path) / saved_file

    print("\n🔹 Select an ION API file:")
    for i, file in enumerate(files, 1):
        print(f"{i}. {file}")

    choice = int(input(f"\nEnter your choice (1-{len(files)}): "))
    selected = files[choice - 1]
    defaults["ionapi_file"] = selected
    save_user_defaults(defaults)
    return Path(folder_path) / selected


def prompt_for_company_division():
    defaults = load_user_defaults()
    default_company = defaults.get("company", "910")
    default_division = defaults.get("division", "088")

    company = input(f"Enter company code (default: {default_company}): ").strip()
    division_input = input(f"Enter division code (default: {default_division}, enter a single space for blank): ")

    if division_input == "":
        division = default_division
    elif division_input.strip() == "":
        division = ""  # explicit blank division
    else:
        division = division_input.strip()

    CONFIG["company"] = company if company else default_company
    CONFIG["division"] = division

    defaults["company"] = CONFIG["company"]
    defaults["division"] = CONFIG["division"]
    save_user_defaults(defaults)

# =============================================================================
# Authentication
# =============================================================================
def get_ion_token():
    with open(CONFIG["tenant"], 'r') as f:
        ionapi = json.load(f)

    token_url = ionapi['pu'] + ionapi['ot']
    payload = {
        'client_id': ionapi['ci'],
        'client_secret': ionapi['cs'],
        'grant_type': 'password',
        'username': ionapi['saak'],
        'password': ionapi['sask']
    }
    headers = {'Content-Type': 'application/x-www-form-urlencoded'}

    with requests.Session() as client:
        response = client.post(token_url, data=payload, headers=headers)
    response.raise_for_status()
    CONFIG["access_token"] = response.json()['access_token']
    print(f"✅ Token acquired successfully.")

# =============================================================================
# SOAP Posting
# =============================================================================
def load_soap_template():
    with open(SOAP_TEMPLATE_PATH, "r", encoding="utf-8") as f:
        return f.read()


def build_soap_payload(template, record):
    xml = template

    def as_str(value):
        """Convert Excel/pandas values to clean strings for SOAP."""
        if value is None:
            return ""
        if isinstance(value, float):
            if value.is_integer():
                return str(int(value))
            return f"{value:.10g}"  # avoid scientific notation, trim trailing zeros
        return str(value)

    # company / division from CONFIG
    xml = xml.replace(
        "<cred:company>decimal(3,0)</cred:company>",
        f"<cred:company>{as_str(CONFIG['company'])}</cred:company>"
    )
    xml = xml.replace(
        "<cred:division>string(3)</cred:division>",
        f"<cred:division>{as_str(CONFIG['division'])}</cred:division>"
    )

    # Excel-driven fields
    replacements = {
        "<add:ItemNumber>string(15)</add:ItemNumber>":
            f"<add:ItemNumber>{as_str(record.get('ItemNumber'))}</add:ItemNumber>",
        "<add:Supplier>string(10)</add:Supplier>":
            f"<add:Supplier>{as_str(record.get('Supplier'))}</add:Supplier>",
        "<add:ServiceProcess>string(3)</add:ServiceProcess>":
            f"<add:ServiceProcess>{as_str(record.get('ServiceProcess'))}</add:ServiceProcess>",
        "<add:Service>string(20)</add:Service>":
            f"<add:Service>{as_str(record.get('Service'))}</add:Service>",
        "<add:Warehouse>string(3)</add:Warehouse>":
            f"<add:Warehouse>{as_str(record.get('Warehouse'))}</add:Warehouse>",
        "<add:SupplyLeadTime>decimal(3,0)</add:SupplyLeadTime>":
            f"<add:SupplyLeadTime>{as_str(record.get('SupplyLeadTime'))}</add:SupplyLeadTime>",
    }

    # Replace all placeholders in one go and remove empty lines
    for before, after in replacements.items():
        value_inside = after.split(">")[1].split("<")[0]
        if value_inside == "":
            xml = xml.replace(before, "")
        else:
            xml = xml.replace(before, after)

    # Compact XML (remove whitespace between tags)
    xml = re.sub(r">\s+<", "><", xml)
    return xml

import xml.etree.ElementTree as ET

def post_soap_request(xml_payload, session, index):
    request_file = LOG_DIR / f"request_{index}.xml"
    response_file = LOG_DIR / f"response_{index}.xml"
    
    request_file.write_text(xml_payload, encoding="utf-8")

    curl_cmd = [
        "curl",
        "-s", "-S",
        "-X", "POST",
        SOAP_ENDPOINT,
        "-H", f"Authorization: Bearer {CONFIG['access_token']}",
        "-H", "Content-Type: application/xml; charset=utf-8",
        "-H", "Accept: application/xml",
        "--data-binary", f"@{request_file}"
    ]

    try:
        result = subprocess.run(curl_cmd, capture_output=True, text=True, check=True)
        response_text = result.stdout
    except subprocess.CalledProcessError as e:
        raise RuntimeError(f"❌ SOAP request {index} failed:\n{e.stderr}")

    response_file.write_text(response_text, encoding="utf-8")

    # --- Check for faultstring ---
    keep_logs = False
    try:
        root = ET.fromstring(response_text)
        ns = {"soap": "http://schemas.xmlsoap.org/soap/envelope/"}
        faultstring = root.find(".//soap:Fault/faultstring", ns)

        if faultstring is not None and faultstring.text.strip():
            fault_text = faultstring.text.strip()
            # only keep logs if faultstring does NOT include "already exists"
            if "already exists" not in fault_text.lower():
                keep_logs = True
    except ET.ParseError:
        pass  # if not valid XML, leave files for inspection

    if not keep_logs:
        # delete both files
        try:
            request_file.unlink(missing_ok=True)
            response_file.unlink(missing_ok=True)
        except Exception as e:
            print(f"⚠️ Could not delete log files for request {index}: {e}")

    return response_text

def post_excel_rows_to_soap(excel_file, max_workers=10):
    df = pd.read_excel(excel_file)
    template = load_soap_template()

    results = []
    with requests.Session() as session:
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = {
                executor.submit(post_soap_request, build_soap_payload(template, row.to_dict()), session, i): (i, row.to_dict())
                for i, (_, row) in enumerate(df.iterrows(), start=1)
            }

            for future in tqdm(as_completed(futures), total=len(futures), desc="SOAP Requests"):
                idx, row_dict = futures[future]
                try:
                    results.append(future.result())
                except Exception as e:
                    tqdm.write(f"Error with row {idx}: {row_dict} → {e}")
    return results

# =============================================================================
# Main
# =============================================================================
if __name__ == "__main__":
    excel_file = "MPD_PPS044_Add_v2.xlsx"

    ionapi_dir = Path(__file__).parent / "ionapi"
    CONFIG["tenant"] = select_ionapi_file(ionapi_dir)
    prompt_for_company_division()
    get_ion_token()

    print(f"\n🚀 Processing Excel file: {excel_file}")
    results = post_excel_rows_to_soap(excel_file, max_workers=5)

    print(f"\n✅ Completed {len(results)} SOAP requests.\n")

    results_file = Path(__file__).parent / "EVS100IPS.results.txt"
    with open(results_file, "w", encoding="utf-8") as f:
        for i, result in enumerate(results, start=1):
            header = "=" * 80 + f"\n📦 SOAP Response #{i}:\n"
            footer = "\n" + "=" * 80 + "\n\n"
            # print(header + result + footer)  # still print to console
            f.write(header + result + footer)

    print(f"📝 All responses saved to {results_file}")