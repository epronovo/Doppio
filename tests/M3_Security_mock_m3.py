"""Stand-in for the ION token endpoint + m3api-rest v2, for offline testing."""
import json
import re
import threading
from http.server import BaseHTTPRequestHandler, HTTPServer

TENANT = "ZFQP353QZYV89ZHG_TST"

# Roles that "exist in M3". Deliberately a mix: some match the CSV capture,
# one (ZZM3ONLY) exists only in M3.
M3_ROLES = [
    {"ROLL": "ADMIN",     "TX40": "Administrator",    "TX15": "ADMIN",   "TXID": "1", "ROLT": "0"},
    {"ROLL": "APAUDRO",   "TX40": "AP Auditor RO",    "TX15": "*AUDITOR", "TXID": "2", "ROLT": "0"},
    {"ROLL": "APAUDRW",   "TX40": "AP Auditor RW",    "TX15": "*AUDITOR", "TXID": "3", "ROLT": "0"},
    {"ROLL": "APDIVRW",   "TX40": "AP Division RW",   "TX15": "*DIVFIN", "TXID": "4", "ROLT": "0"},
    {"ROLL": "ZZM3ONLY",  "TX40": "Only in M3",       "TX15": "ZZ",      "TXID": "5", "ROLT": "0"},
]

MEMBERS = {
    "APDIVRW": [
        {"USID": "ADATKHILE", "ROLL": "APDIVRW", "FVDT": "20240101", "VTDT": "", "TXID": "0"},
        {"USID": "ATHAR",     "ROLL": "APDIVRW", "FVDT": "20240101", "VTDT": "", "TXID": "0"},
        {"USID": "MDRAGOVIC", "ROLL": "APDIVRW", "FVDT": "20240101", "VTDT": "", "TXID": "0"},
        {"USID": "GHOSTUSER", "ROLL": "APDIVRW", "FVDT": "20240101", "VTDT": "", "TXID": "0"},
    ],
    "APAUDRO": [
        {"USID": "ATHAR", "ROLL": "APAUDRO", "FVDT": "20240101", "VTDT": "", "TXID": "0"},
    ],
    "ADMIN": [],
    "APAUDRW": [],
    "ZZM3ONLY": [{"USID": "ATHAR", "ROLL": "ZZM3ONLY", "FVDT": "", "VTDT": "", "TXID": "0"}],
}

USERS = [
    {"USID": "ADATKHILE", "TX40": "Akash Datkhile", "EMAL": "Akash.Datkhile@chcheli.com", "USTA": "20"},
    {"USID": "ATHAR",     "TX40": "Athar",          "EMAL": "athar@doppiogroup.com",      "USTA": "20"},
    {"USID": "MDRAGOVIC", "TX40": "Milan Dragovic", "EMAL": "milan.dragovic@chcheli.com", "USTA": "20"},
    {"USID": "GHOSTUSER", "TX40": "Ghost User",     "EMAL": "ghost@chcheli.com",          "USTA": "20"},
]

# SES400 function authorisations. Deliberately references roles that MNS405
# does have (APAUDRO, ADMIN), roles it does NOT (GHOSTROLE, ORPHAN2), and
# spreads a couple of rows over a second division.
FUNCTION_ROLES = [
    {"FNID": "CRS610",  "ROLL": "ADMIN",     "CONO": "001", "DIVI": "100", "STAT": "20", "TXID": "0"},
    {"FNID": "CRS610",  "ROLL": "APAUDRO",   "CONO": "001", "DIVI": "100", "STAT": "20", "TXID": "0"},
    {"FNID": "CRS610",  "ROLL": "GHOSTROLE", "CONO": "001", "DIVI": "100", "STAT": "20", "TXID": "0"},
    {"FNID": "MMS001",  "ROLL": "ADMIN",     "CONO": "001", "DIVI": "100", "STAT": "20", "TXID": "0"},
    {"FNID": "MMS001",  "ROLL": "ORPHAN2",   "CONO": "001", "DIVI": "100", "STAT": "20", "TXID": "0"},
    {"FNID": "OIS100",  "ROLL": "APAUDRO",   "CONO": "001", "DIVI": "100", "STAT": "10", "TXID": "0"},
    {"FNID": "OIS100",  "ROLL": "GHOSTROLE", "CONO": "001", "DIVI": "200", "STAT": "20", "TXID": "0"},
    {"FNID": "PPS200",  "ROLL": "ORPHAN2",   "CONO": "001", "DIVI": "200", "STAT": "10", "TXID": "0"},
    {"FNID": "PPS200",  "ROLL": "APDIVRW",   "CONO": "001", "DIVI": "100", "STAT": "20", "TXID": "0"},
]

# Role names the mock refuses to create, to exercise the failure path
FAIL_ROLE_ADD = {"ORPHAN2"}
FAIL_ROLE_UPDATE = {"ZZM3ONLY"}
FAIL_ROLE_DELETE = {"ADMIN"}
FAIL_UPDATE = {"GHOSTUSER"}

# Authorisations the mock refuses to delete, to exercise the failure path
FAIL_FUNCTION_DELETE = {("MMS001", "ADMIN")}
FAIL_FUNCTION_UPDATE = {("PPS200", "APDIVRW")}

# USIDs the mock refuses to delete / add, to exercise the failure paths
FAIL_DELETE = {"GHOSTUSER"}
FAIL_ADD = {"MDRAGOVIC"}

CALLS = []


class Handler(BaseHTTPRequestHandler):
    def log_message(self, *a):
        pass

    def _send(self, obj, code=200):
        body = json.dumps(obj).encode()
        self.send_response(code)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_POST(self):
        n = int(self.headers.get("Content-Length", 0))
        raw = self.rfile.read(n)

        if self.path.startswith("/token/"):
            return self._send({"access_token": "mock-token", "expires_in": 3600})

        if "/m3api-rest/v2/execute" not in self.path:
            return self._send({"error": "bad path"}, 404)

        auth = self.headers.get("Authorization", "")
        if auth != "Bearer mock-token":
            return self._send({"error": "unauthorized"}, 401)

        payload = json.loads(raw)
        program = payload["program"]
        CALLS.append(payload)
        results = []
        for t in payload["transactions"]:
            tx = t["transaction"]
            rec = t.get("record") or {}
            recs, err = [], ""

            if program == "MNS405MI" and tx == "Lst":
                roll = (rec.get("ROLL") or "").strip()
                recs = [r for r in M3_ROLES if not roll or r["ROLL"].startswith(roll)]
            elif program == "MNS405MI" and tx == "Get":
                hit = [r for r in M3_ROLES if r["ROLL"] == (rec.get("ROLL") or "").strip()]
                if hit:
                    recs = hit
                else:
                    err = "Role does not exist. Record not found"
            elif program == "MNS410MI" and tx == "Lst":
                roll = (rec.get("ROLL") or "").strip()
                usid = (rec.get("USID") or "").strip()
                if roll:
                    recs = list(MEMBERS.get(roll, []))
                else:                       # blank input lists everything
                    recs = [m for rows in MEMBERS.values() for m in rows]
                if usid:
                    recs = [m for m in recs if m["USID"] == usid]
                if not recs:
                    err = "No records found. Record not found"
            elif program == "SES400MI" and tx == "Lst":
                fn = (rec.get("FNID") or "").strip()
                rl = (rec.get("ROLL") or "").strip()
                recs = [r for r in FUNCTION_ROLES
                        if (not fn or r["FNID"] == fn)
                        and (not rl or r["ROLL"] == rl)]
                if not recs:
                    err = "No records found. Record not found"
            elif program == "SES400MI" and tx == "Upd":
                fn = (rec.get("FNID") or "").strip()
                rl = (rec.get("ROLL") or "").strip()
                co = (rec.get("CONO") or "").strip()
                dv = (rec.get("DIVI") or "").strip()
                if (fn, rl) in FAIL_FUNCTION_UPDATE:
                    err = "Authorisation is locked"
                else:
                    hit = [r for r in FUNCTION_ROLES
                           if r["FNID"] == fn and r["ROLL"] == rl
                           and (not co or r["CONO"] == co)
                           and (not dv or r["DIVI"] == dv)]
                    if not hit:
                        err = "Record does not exist"
                    else:
                        for r in hit:
                            # only the supplied fields change; the option flags
                            # this mock does not carry are left alone
                            if "STAT" in rec:
                                r["STAT"] = rec["STAT"]
            elif program == "SES400MI" and tx == "Dlt":
                fn = (rec.get("FNID") or "").strip()
                rl = (rec.get("ROLL") or "").strip()
                co = (rec.get("CONO") or "").strip()
                dv = (rec.get("DIVI") or "").strip()
                if (fn, rl) in FAIL_FUNCTION_DELETE:
                    err = "Authorisation is locked"
                else:
                    before = len(FUNCTION_ROLES)
                    FUNCTION_ROLES[:] = [
                        r for r in FUNCTION_ROLES
                        if not (r["FNID"] == fn and r["ROLL"] == rl
                                and (not co or r["CONO"] == co)
                                and (not dv or r["DIVI"] == dv))
                    ]
                    if len(FUNCTION_ROLES) == before:
                        err = "Record does not exist"
            elif program == "MNS405MI" and tx == "Add":
                roll = (rec.get("ROLL") or "").strip()
                if any(r["ROLL"] == roll for r in M3_ROLES):
                    err = "Role already exists"
                elif roll in FAIL_ROLE_ADD:
                    err = "Role id not permitted"
                elif len(roll) > 10:
                    err = "ROLL too long"
                else:
                    M3_ROLES.append({"ROLL": roll,
                                     "TX40": rec.get("TX40", ""),
                                     "TX15": rec.get("TX15", ""),
                                     "TXID": "9", "ROLT": rec.get("ROLT", "0")})
                    MEMBERS.setdefault(roll, [])
            elif program == "MNS405MI" and tx == "Upd":
                roll = (rec.get("ROLL") or "").strip()
                hit = [r for r in M3_ROLES if r["ROLL"] == roll]
                if not hit:
                    err = "Role does not exist"
                elif roll in FAIL_ROLE_UPDATE:
                    err = "Role is locked"
                else:
                    for key in ("TX40", "TX15", "ROLT"):
                        if key in rec:          # omitted fields are left alone
                            hit[0][key] = rec[key]
            elif program == "MNS405MI" and tx == "Dlt":
                roll = (rec.get("ROLL") or "").strip()
                if roll in FAIL_ROLE_DELETE:
                    err = "Role is in use"
                elif not any(r["ROLL"] == roll for r in M3_ROLES):
                    err = "Role does not exist"
                else:
                    M3_ROLES[:] = [r for r in M3_ROLES if r["ROLL"] != roll]
                    MEMBERS.pop(roll, None)
            elif program == "MNS410MI" and tx == "Upd":
                usid = (rec.get("USID") or "").strip()
                roll = (rec.get("ROLL") or "").strip()
                hit = [m for m in MEMBERS.get(roll, []) if m["USID"] == usid]
                if not hit:
                    err = "Record does not exist"
                elif usid in FAIL_UPDATE:
                    err = "User is locked"
                else:
                    for key in ("FVDT", "VTDT"):
                        if key in rec:
                            hit[0][key] = rec[key]
            elif program == "MNS410MI" and tx == "Add":
                usid = (rec.get("USID") or "").strip()
                roll = (rec.get("ROLL") or "").strip()
                if not any(r["ROLL"] == roll for r in M3_ROLES):
                    err = "Role does not exist"
                elif not any(u["USID"] == usid for u in USERS):
                    err = "User does not exist"
                elif any(m["USID"] == usid for m in MEMBERS.get(roll, [])):
                    err = "User already connected to role"
                elif usid in FAIL_ADD:
                    err = "User is locked"
                else:
                    MEMBERS.setdefault(roll, []).append(
                        {"USID": usid, "ROLL": roll,
                         "FVDT": rec.get("FVDT", ""), "VTDT": rec.get("VTDT", ""),
                         "TXID": "0"})
            elif program == "MNS410MI" and tx == "Dlt":
                usid = (rec.get("USID") or "").strip()
                roll = (rec.get("ROLL") or "").strip()
                if usid in FAIL_DELETE:
                    err = "USID is locked and cannot be removed"
                else:
                    MEMBERS[roll] = [m for m in MEMBERS.get(roll, []) if m["USID"] != usid]
            elif program == "MNS150MI" and tx == "LstUserData":
                recs = list(USERS)
            else:
                err = f"Unknown transaction {program}/{tx}"

            results.append({"transaction": tx, "errorMessage": err,
                            "errorCode": "" if not err else "1",
                            "records": recs})
        return self._send({"results": results,
                           "nrOfSuccessfullTransactions":
                               sum(1 for r in results if not r["errorMessage"])})


def serve(port=5099):
    srv = HTTPServer(("127.0.0.1", port), Handler)
    threading.Thread(target=srv.serve_forever, daemon=True).start()
    return srv


if __name__ == "__main__":
    serve()
    import time
    while True:
        time.sleep(1)
