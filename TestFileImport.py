import requests

url = "https://mingle-ionapi.inforcloudsuite.com/U1SOURCE_TST/M3/foundation-rest/file-management/v1/file%2FFileImport%2FCTOS_Frame.xlsx"

headers = {
    "accept": "application/json",
    "Content-Type": "application/octet-stream",
    "Authorization": "Bearer eyJraWQiOiJrZzo0MjFmZTVkNC02ZGRmLTQwMjAtYmNlZS1iNjkzMWUxN2NkMWQiLCJhbGciOiJSUzI1NiJ9.eyJTZXJ2aWNlQWNjb3VudCI6IlUxU09VUkNFX1RTVCNlVC1yTVEtNS1nWVZnSThESkhqWVBPalZRXzVQZ3k2YlJSZFQtMjg2SEo2b2U3N2E0RjVJZWx5bWlxaWhWdU9CRG5LVHVvQTBDZHJMQ3ZGVjZYSVZmUSIsIlRlbmFudCI6IlUxU09VUkNFX1RTVCIsIklkZW50aXR5MiI6Ijg0ODZkOTA5LTVkNGMtNDQxOS05OTEzLTU3MTljZDRmOGZjOCIsIkVuZm9yY2VTY29wZXNGb3JDbGllbnQiOiIwIiwiZ3JhbnRfaWQiOiI5NzMyZGY0ZS05ZTU2LTQ5NTMtODVlZS00ZjY3ZjUwNzEzMjkiLCJJbmZvclNUU0lzc3VlZFR5cGUiOiJBUyIsImNsaWVudF9pZCI6IlUxU09VUkNFX1RTVH56dHFiVldpYkt4V3pFRFU5dWhYVW1BTm8xM3Y3LW1NNm9ocWUyS0x4QS0wIiwianRpIjoiNTk5ZTY1OTItNTI0Ni00ZDA5LWJlNDItYzVlOWVjZDE1ODk2IiwiaWF0IjoxNzU4ODI0NDAyLCJuYmYiOjE3NTg4MjQ0MDIsImV4cCI6MTc1ODgyODAwMiwiaXNzIjoiaHR0cHM6Ly9taW5nbGUtc3NvLmluZm9yY2xvdWRzdWl0ZS5jb206NDQzIiwiYXVkIjoiaHR0cHM6Ly9taW5nbGUtaW9uYXBpLmluZm9yY2xvdWRzdWl0ZS5jb20ifQ.QWFQ4R2G7fHfOFSyKaRbaUMrVHZAnKFKd65gUtCCdcpJi0LlxZGRyfv_Hd0tp3fmpGeobkvYBB9iIsao00X0Hb5TCiqK--2R_0CvSv8gBSIae4AEQS1v5iuazzvseOexxUO8uyBK0dZEIthSMP7sm4rrdXgJitcRFwiyEsTJXIcAl8nIl22cW8c7kaq2itsnU4HPEIENSK-RaTgGi-IJEgFMaQwMQCvhGbxNT345NEm-RxOkyOTYkSKX_0DXjpa-VFkkg09_lWOZdzwsDox6a9kRGtKqlcl3ylJ4NC_Lo5uCsJW63oJxdTKzrKJlm0euvLo8D46wYLpnjW1PXkZx2Q"
}

# Open file in binary mode and send as body
with open("/Users/ericpronovost/Downloads/CTOS_Frame.xlsx", "rb") as f:
    response = requests.put(url, headers=headers, data=f)

print("Status Code:", response.status_code)
print("Response:", response.text)

# === Second call: only if upload succeeded ===
if response.status_code == 201:
    execute_url = (
        "https://mingle-ionapi.inforcloudsuite.com/"
        "U1SOURCE_TST/M3/m3api-rest/v2/execute"
        "?maxrecs=20000&extendedresult=true&m3user=EPRONOVOST"
        "&righttrim=true&cono=501&divi=USA"
    )

    execute_headers = {
        "accept": "application/json; charset=UTF-8",
        "Content-Type": "application/json; charset=UTF-8",
        "Authorization": "Bearer eyJraWQiOiJrZzo0MjFmZTVkNC02ZGRmLTQwMjAtYmNlZS1iNjkzMWUxN2NkMWQiLCJhbGciOiJSUzI1NiJ9.eyJTZXJ2aWNlQWNjb3VudCI6IlUxU09VUkNFX1RTVCNlVC1yTVEtNS1nWVZnSThESkhqWVBPalZRXzVQZ3k2YlJSZFQtMjg2SEo2b2U3N2E0RjVJZWx5bWlxaWhWdU9CRG5LVHVvQTBDZHJMQ3ZGVjZYSVZmUSIsIlRlbmFudCI6IlUxU09VUkNFX1RTVCIsIklkZW50aXR5MiI6Ijg0ODZkOTA5LTVkNGMtNDQxOS05OTEzLTU3MTljZDRmOGZjOCIsIkVuZm9yY2VTY29wZXNGb3JDbGllbnQiOiIwIiwiZ3JhbnRfaWQiOiI5NzMyZGY0ZS05ZTU2LTQ5NTMtODVlZS00ZjY3ZjUwNzEzMjkiLCJJbmZvclNUU0lzc3VlZFR5cGUiOiJBUyIsImNsaWVudF9pZCI6IlUxU09VUkNFX1RTVH56dHFiVldpYkt4V3pFRFU5dWhYVW1BTm8xM3Y3LW1NNm9ocWUyS0x4QS0wIiwianRpIjoiNTk5ZTY1OTItNTI0Ni00ZDA5LWJlNDItYzVlOWVjZDE1ODk2IiwiaWF0IjoxNzU4ODI0NDAyLCJuYmYiOjE3NTg4MjQ0MDIsImV4cCI6MTc1ODgyODAwMiwiaXNzIjoiaHR0cHM6Ly9taW5nbGUtc3NvLmluZm9yY2xvdWRzdWl0ZS5jb206NDQzIiwiYXVkIjoiaHR0cHM6Ly9taW5nbGUtaW9uYXBpLmluZm9yY2xvdWRzdWl0ZS5jb20ifQ.QWFQ4R2G7fHfOFSyKaRbaUMrVHZAnKFKd65gUtCCdcpJi0LlxZGRyfv_Hd0tp3fmpGeobkvYBB9iIsao00X0Hb5TCiqK--2R_0CvSv8gBSIae4AEQS1v5iuazzvseOexxUO8uyBK0dZEIthSMP7sm4rrdXgJitcRFwiyEsTJXIcAl8nIl22cW8c7kaq2itsnU4HPEIENSK-RaTgGi-IJEgFMaQwMQCvhGbxNT345NEm-RxOkyOTYkSKX_0DXjpa-VFkkg09_lWOZdzwsDox6a9kRGtKqlcl3ylJ4NC_Lo5uCsJW63oJxdTKzrKJlm0euvLo8D46wYLpnjW1PXkZx2Q"  # second token
    }

    execute_payload = {
        "program": "EVS100MI",
        "transactions": [
            {
                "transaction": "ImportFile",
                "record": {"FNAM": "CTOS_Frame.xlsx"},
                "selectedColumns": ["FNAM"]
            }
        ]
    }

    execute_response = requests.post(
        execute_url, headers=execute_headers, json=execute_payload, timeout=300
    )

    print("Execute Status:", execute_response.status_code)
    print("Execute Response:", execute_response.text)
else:
    print("Skipping second call because upload did not return 201.")