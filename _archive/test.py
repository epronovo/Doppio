import pandas as pd
import requests
from concurrent.futures import ThreadPoolExecutor

EXCEL_FILE = r"\\Mac\Home\Downloads\MPD_PPS044_Add.xlsx" 

url = "https://mingle-ionapi.inforcloudsuite.com/W8DY5WFWZLXDHJPA_TST/M3/ips/service/PPS040WS"
headers = {
    'accept': 'application/xml; charset=UTF-8',
    'Content-Type': 'application/xml; charset=UTF-8',
    'Authorization': 'Bearer eyJraWQiOiJrZzo5MGYwZTI4Ny0wOTlhLTRjNzAtYWM5Yi0zMmRjN2E0ZTJhZDYiLCJhbGciOiJSUzI1NiJ9.eyJTZXJ2aWNlQWNjb3VudCI6Ilc4RFk1V0ZXWkxYREhKUEFfVFNUIzBseUd1WjVNdWJQZWRkU3ozd3hsTGNrd1pCeTcwN3czMjdQMkxoS2VNTjVmUl8teDBkWGJfZ0tZdVpOR1lfYS1fWDVRRk5KeUViOW8xR05jUV8xSG5RIiwiVGVuYW50IjoiVzhEWTVXRldaTFhESEpQQV9UU1QiLCJJZGVudGl0eTIiOiI4YzM3YmY0ZC1kZjczLTQyNGEtODNiNi0zMWRhYjdiYTg2NDIiLCJFbmZvcmNlU2NvcGVzRm9yQ2xpZW50IjoiMCIsImdyYW50X2lkIjoiOTkxNTJiZTktZGNhYi00ZGZkLWJhNDYtZDlhYzYwYzM1MzU2IiwiSW5mb3JTVFNJc3N1ZWRUeXBlIjoiQVMiLCJjbGllbnRfaWQiOiJXOERZNVdGV1pMWERISlBBX1RTVH50eUY3UVhic21jcnBXNnZ5dmFaVFdYMVljUUdMcDQwb2V2TWNMU0h5ZzFzIiwianRpIjoiZGVjZjUwZmQtYzg2MS00Y2E0LWFhNTktY2FkY2UyNjU1ZjUyIiwiaWF0IjoxNzQ5ODQwODc3LCJuYmYiOjE3NDk4NDA4NzcsImV4cCI6MTc0OTg0NDQ3NywiaXNzIjoiaHR0cHM6Ly9taW5nbGUtc3NvLmluZm9yY2xvdWRzdWl0ZS5jb206NDQzIiwiYXVkIjoiaHR0cHM6Ly9taW5nbGUtaW9uYXBpLmluZm9yY2xvdWRzdWl0ZS5jb20ifQ.WIb5FC_A0Q2w2cKsTOa655mYAnRabGWKNPXO3nGXtLkyu1_A1EKcPFGtHhTf8LRcMq4mZ3RmE2X3eDDspJk4_70Qzi95UZF9aDs-3Nl7y8goPFYNDUQtzChiDrQIn3ZxIK04--YVM1b1f1Wd3IKFBFocfZ3Uwj33BWcJJEv5QCwINx-2Je0j3yQdfdeEccnygU_J1u3O4U32Dvimn6vF74W9lAeWvfvdlnc_L-dG0LkgcV67VIODeIjOFtcGm3HTnhwijYNOi-VX-cSMvX8ujdfoSqYoP7qHK1RbPfpB5ezHN_SJlgOSegWmD163Lu_WVwhgA1V2LeG9kDe8Hxsbcg'  # Replace with your token
}

def build_payload(item_number, supplier, warehouse, supply_lead_time):
    return f"""<?xml version="1.0" encoding="UTF-8"?>
<SOAP-ENV:Envelope xmlns:SOAP-ENV="http://schemas.xmlsoap.org/soap/envelope/" xmlns:chg="http://schemas.infor.com/ips/PPS040WS/Add" xmlns:cred="http://lawson.com/ws/credentials">
  <SOAP-ENV:Header>
    <cred:lws>
      <cred:company>820</cred:company>
      <cred:division>088</cred:division>
    </cred:lws>
  </SOAP-ENV:Header>
  <SOAP-ENV:Body>
    <chg:Add>
      <chg:PPS040>
        <chg:ItemNumber>{item_number}</chg:ItemNumber>
        <chg:Supplier>{supplier}</chg:Supplier>
        <chg:PPS044>
          <chg:Warehouse>{warehouse}</chg:Warehouse>
          <chg:SupplyLeadTime>{supply_lead_time}</chg:SupplyLeadTime>
        </chg:PPS044>
      </chg:PPS040>
    </chg:Add>
  </SOAP-ENV:Body>
</SOAP-ENV:Envelope>"""

def send_request(row):
    print(f"Processing ItemNumber: {row['ItemNumber']}, Supplier: {row['Supplier']}, Warehouse: {row['Warehouse']}")
    payload = build_payload(row['ItemNumber'], row['Supplier'], row['Warehouse'], row['SupplyLeadTime'])
    response = requests.post(url, headers=headers, data=payload)
    return {
        'ItemNumber': row['ItemNumber'],
        'Supplier': row['Supplier'],
        'Warehouse': row['Warehouse'],
        'SupplyLeadTime': row['SupplyLeadTime'],
        'status': response.status_code,
        'response': response.text
    }

def main():
    df = pd.read_excel(EXCEL_FILE)
    with ThreadPoolExecutor(max_workers=10) as executor:
        results = list(executor.map(send_request, [row for _, row in df.iterrows()]))
    results_df = pd.DataFrame(results)
    results_df.to_excel("MPD_PPS044_Add_results.xlsx", index=False)

if __name__ == "__main__":
    main()
