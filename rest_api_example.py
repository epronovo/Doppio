#!/usr/bin/env python3
"""
rest_api_example.py - Minimal, standalone example of the REST calls used by
etl_datalake.py: OAuth2 token, Data Lake ping/version/dataobjects, and one
M3 REST API (MNS120MI.Get) call.

Uses only the Python standard library. Point IONAPI_FILE at a .ionapi file
downloaded from the ION API gateway for the environment you want to hit.
"""

import json
import urllib.parse
import urllib.request
import zlib

IONAPI_FILE = "/Users/ericpronovost/Doppio/ionapi/DOPPIO_DEM.ionapi"
M3_CONO = "001"


def load_config(path):
    with open(path) as fh:
        return json.load(fh)


# 1. OAuth2 resource-owner (service account) grant -> access_token
def get_access_token(cfg):
    token_url = cfg["pu"] + cfg["ot"]
    data = urllib.parse.urlencode({
        "grant_type": "password",
        "username": cfg["saak"],
        "password": cfg["sask"],
        "client_id": cfg["ci"],
        "client_secret": cfg["cs"],
    }).encode()
    req = urllib.request.Request(
        token_url, data=data, method="POST",
        headers={"Content-Type": "application/x-www-form-urlencoded"})
    with urllib.request.urlopen(req, timeout=60) as resp:
        return json.loads(resp.read().decode())["access_token"]


def datalake_base_url(cfg):
    return f"{cfg['iu'].rstrip('/')}/{cfg['ti']}/DATAFABRIC/datalake/v2"


def m3_base_url(cfg):
    return (f"{cfg['iu'].rstrip('/')}/{cfg['ti']}/M3/m3api-rest/v2/execute"
            f"?maxrecs=100&extendedresult=true&righttrim=true&cono={M3_CONO}")


def api_get(url, token, accept="application/json", deflate=False):
    headers = {"accept": accept, "Authorization": f"Bearer {token}"}
    if deflate:
        headers["Accept-Encoding"] = "deflate"
    req = urllib.request.Request(url, headers=headers)
    with urllib.request.urlopen(req, timeout=120) as resp:
        body = resp.read()
        enc = resp.headers.get("Content-Encoding", "")
    if enc == "deflate" or deflate:
        for wbits in (zlib.MAX_WBITS, -zlib.MAX_WBITS, zlib.MAX_WBITS | 16):
            try:
                return zlib.decompress(body, wbits)
            except zlib.error:
                continue
    return body


def api_post_json(url, token, payload):
    data = json.dumps(payload).encode()
    headers = {
        "accept": "application/json; charset=UTF-8",
        "Content-Type": "application/json; charset=UTF-8",
        "Authorization": f"Bearer {token}",
    }
    req = urllib.request.Request(url, data=data, method="POST", headers=headers)
    with urllib.request.urlopen(req, timeout=120) as resp:
        return json.loads(resp.read().decode())


# 2. Ping the Data Lake
def ping(base, token):
    body = api_get(f"{base}/ping", token, accept="*/*").decode().strip()
    return body.strip('"').upper() == "OK"


# 3. Current Data Lake build/version
def get_version(base, token):
    return api_get(f"{base}/version", token, accept="*/*").decode().strip()


# 4. List data objects changed since a given date (first page only, for the example)
def list_dataobjects(base, token, since, records=50):
    flt = urllib.parse.quote(f'dl_document_date ge "{since}"')
    url = (f"{base}/dataobjects?filter={flt}"
           f"&sort={urllib.parse.quote('dl_document_date:asc')}"
           f"&page=1&records={records}")
    return json.loads(api_get(url, token).decode())


# 5. Download one data object (NDJSON, deflate-compressed)
def download_dataobject(base, token, dl_id):
    raw = api_get(f"{base}/dataobjects/{dl_id}", token,
                   accept="application/octet-stream", deflate=True)
    return [json.loads(line) for line in raw.decode("utf-8", errors="replace").splitlines() if line.strip()]


# 6. M3 REST API: run a single program/transaction (here, MNS120MI.Get)
def m3_get_table_keys(cfg, token, table):
    payload = {
        "program": "MNS120MI",
        "transactions": [{"transaction": "Get", "record": {"FILE": f"{table}00"}}],
    }
    return api_post_json(m3_base_url(cfg), token, payload)


def main():
    cfg = load_config(IONAPI_FILE)
    token = get_access_token(cfg)
    base = datalake_base_url(cfg)

    print("Ping OK:", ping(base, token))
    print("Version:", get_version(base, token))

    listing = list_dataobjects(base, token, since="2026-07-01T00:00:00.000Z")
    print(f"Objects found: {listing.get('numFound', 0)}")

    if listing.get("fields"):
        first = listing["fields"][0]
        records = download_dataobject(base, token, first["dl_id"])
        print(f"Downloaded {first['dl_document_name']}: {len(records)} record(s)")

    keys = m3_get_table_keys(cfg, token, "CSYTAB")
    print("M3 MNS120MI.Get result:", keys)


if __name__ == "__main__":
    main()
