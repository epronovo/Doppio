# sheet_security — Doppio API Sheet security manager

A small Flask app for managing the API Sheet's own access-control table,
**EXTXSM**, which the custom M3 program `EXT124MI` (the same program
`SEC_GrantAccess.py` and the Xtend workbook write through) reads and writes.
It lists every record straight off the table with `EXPORTMI/Select`, and
adds / edits / deletes them with `EXT124MI`'s `AddUsrInfo` / `UpdUsrInfo` /
`DelUsrInfo` — live against whichever tenant you pick, no local copy or
export/import step.

The list splits into three tabs by `AUTH`:

| Tab | AUTH | What it holds |
|-----|------|----------------|
| **Tenants** | `20` | A customer tenant's registration — its `HASH` is an encoded `.ionapi` file (see Extract Type 20, below) |
| **Users** | `1` / `0` | A user's access to a tenant — `1` granted, `0` that same grant blocked |
| **Review** | `99` | Someone who has attempted to use the sheet without a grant yet — has a **Delete all** action, since this tab is meant to be cleared out once its requests are handled |

| File | Role |
|------|------|
| `SheetSecurity_App.py` | Flask routes |
| `SheetSecurity_M3Api.py` | `M3Client` (EXT124MI + EXPORTMI + MNS150MI over m3api-rest), the "find M3 user" guesser, HASH decode/encode, and the AUTH=20 → `.ionapi` extractor |
| `templates/SheetSecurity_Index.html` | The page — table, add/edit panel |

## Quick start

```bash
pip install -r requirements.txt
python SheetSecurity_App.py      # http://127.0.0.1:5062/
```

The tenant dropdown lists every `.ionapi` file in the repo-root `ionapi/`
folder (the same shared folder every other Infor tool in this repo uses),
defaulting to `DOPPIO_DEM`. Switching tenants re-reads the list from that
tenant's own EXT124MI.

## The record shape

A security record's key is **PCID + TNNM + AUTH**:

- `PCID` — the login / person id the record is for
- `TNNM` — the customer tenant name it grants access to (e.g. `NUTRACORP TST`)
- `AUTH` — the access level (`1` user access, `0` that access blocked, `20`
  tenant registration, `99` pending request — see `AUTH_LABELS` in
  `SheetSecurity_M3Api.py`)

`HASH`, `M3ID` and `UMSG` are the value fields carried on top of that key.
`M3ID` holds an M3 USID and `UMSG` an email address once a record has been
matched to a real M3 user. The key fields are locked once a record is
opened for editing — changing PCID, TNNM or AUTH means deleting the record
and adding a new one, since that is what actually changes on EXT124MI.

The list also carries EXTXSM's own audit columns, which `EXT124MI/LstUsrInfo`
does not return: `RGDT` / `RGTM` (registered), `LMDT` / `LMTS` (last
changed) and `CHNO` / `CHID` (the change number and who made it). `LMTS` is
a Java timestamp — milliseconds since the Unix epoch — so it is formatted
into a plain UTC date/time (`format_java_timestamp()` in
`SheetSecurity_M3Api.py`) rather than shown as a 13-digit number.

Each tab hides the columns that are not useful in a list (`TAB_HIDDEN_COLUMNS`
in `templates/SheetSecurity_Index.html`) — they are still on the add/edit
panel, just not worth a column in a table this narrow:

| Tab | Hidden from the list |
|-----|----------------------|
| Tenants | `M3ID`, `RGDT`, `RGTM`, `LMDT`, `HASH` |
| Users | `RGDT`, `RGTM`, `LMDT`, `HASH` |
| Review | `UMSG`, `RGDT`, `RGTM`, `LMDT`, `HASH` |

`HASH` ends up hidden everywhere, for three different reasons: a tenant's is
an `.ionapi` blob (opened on the edit panel, or written out wholesale by
Extract Type 20), a user grant's is normally empty, and a review request's is
decoded server-side instead — see below.

Opening any record for edit decodes its `HASH` automatically (the same call
"Decrypt HASH → JSON" makes) so the JSON is already there rather than making
that a second click; a HASH that fails to decode just leaves the box
collapsed with a note, rather than an error dialog.

## The Review tab's HASH

An AUTH=99 record's `HASH` isn't an `.ionapi` — it's workbook telemetry, a
JSON blob describing whoever the Xtend workbook let make a request it had no
grant for:

```json
{
  "userName": "chrita", "userDomain": "HMABIQR", "computerName": "JANEADAM",
  "localIP": "192.168.0.120", "publicIP": "72.153.216.187",
  "userProfile": "C:\\Users\\chrita",
  "osName": "Microsoft Windows 10 Enterprise", "osVersion": "10.0.15063",
  "sheetVersion": "v1.27"
}
```

Rather than show that raw blob, the Review list decodes it server-side
(`_decode_review_hash()` in `SheetSecurity_App.py`) and shows `USERNAME`,
`DOMAIN` and `COMPUTERNAME` as their own columns — who tried to use the
sheet and from where, at a glance. A blob that fails to decode leaves those
columns blank rather than breaking the list; the full HASH is still on the
edit panel if it needs a closer look.

Decoding a Windows path like `C:\Users\Name` out of this blob means fixing
up backslashes that aren't valid JSON on their own — `decode_hash_blob()` in
`SheetSecurity_M3Api.py` walks the text and doubles every backslash that
isn't the start of a real escape sequence. A lowercase path segment like
`\users\` needs care here: a naive check would see `\u` and assume it starts
a `\uXXXX` unicode escape, then fail with "Invalid \uXXXX escape" the moment
what follows isn't four hex digits — which is exactly the bug this walk
fixes, by only treating `\u` as an escape when it actually is followed by
four hex digits.

EXT124MI is a private extension with no entry in the shared M3 API
metadata this repo otherwise draws on, so `EXPORTMI/Select` is called with
no fixed column list beyond the query itself, and `GetUsrInfo` (used to
re-read one record) with no `selectedColumns` filter — whatever field
layout the program actually returns comes through as-is. The table's
columns are built from that response rather than a hardcoded list, so an
unexpected field is shown rather than silently dropped. `HASH` can run long
(it is sometimes an encoded `.ionapi` blob); the list only ever sends a
truncated copy to the browser, and the full value is re-read with
`GetUsrInfo` when a row is opened for editing.

## Find M3 user

On the add/edit panel, **Find M3 user** takes the record's `TNNM` and
`PCID`, resolves `TNNM` to *that* tenant's own `.ionapi` file (matched
loosely — `TNNM` is a display name like `NUTRACORP TST`, not necessarily an
exact `ti`), connects to it, and reads `MNS150MI.LstUserData`. Every M3
user is scored against the PCID — an exact match on USID or email scores
highest, everything else falls back to a fuzzy string match against USID,
email and full name — and the results are listed by score with **Use** to
fill `M3ID` (USID) and `UMSG` (email). The top match auto-fills when its
score clears a plausibility bar; the rest stay one click away in case the
guess is wrong.

PCID is often a Windows login, not a real M3 identifier, so it does not
always score well against anyone. The **Search hint** field takes a second
term by hand — a name or an email — and every candidate is scored against
*both* PCID and the hint, keeping whichever fits better; a good hint can
surface the right person even when PCID alone would not have. At least one
of PCID / hint has to be given, but not both.

If `TNNM` has no `.ionapi` file on disk yet, it doesn't just fail: there
should already be an AUTH=20 record for that same `TNNM` on the tenant
currently open in the app (that record is exactly what registered the
tenant in the first place - see Extract Type 20, below), so its `HASH` is
decoded and saved as a real `.ionapi` file on the spot, then used
immediately for the lookup. The status line says "extracted just now" when
that happened. This only ever adds a new file; it never overwrites one that
was already there.

## Decrypt / Encrypt HASH

HASH isn't actually encrypted — it's base64(compact JSON), the same scheme
`SEC_GrantAccess._decode_exhash()` / `_encode_ionapi()` use. **Decrypt HASH
→ JSON** unwraps it into a pretty-printed, editable JSON box; **Encrypt
JSON → HASH** re-flattens whatever is in that box (`json.loads` ignores the
whitespace either way, so pretty vs. flat text encodes identically) and
writes the result back into HASH — nothing is saved to M3 until Save.

Two details worth knowing since they change the actual bytes:

- The result is base64 **without** trailing `=` padding, because real HASH
  values never carry it (the original VBA encoder omits it, and decoding
  always re-pads first) — padding it would make an edited row's HASH look
  different from every other row's for no reason.
- `app.json.sort_keys = False` is set at startup. Flask's `jsonify()`
  alphabetises object keys by default, which would silently reorder a
  decoded HASH's fields; with it off, decoding and re-encoding an
  unmodified value reproduces the original HASH byte-for-byte.

## Extract Type 20 → .ionapi files

An AUTH=20 record is EXTXSM's own note that a customer tenant is
registered, and its HASH is exactly an encoded `.ionapi` file. The toolbar's
**Extract Type 20 → .ionapi files** button decodes every AUTH=20 record for
the current tenant and writes each one out as a real `.ionapi` file in the
shared repo-root `ionapi/` folder, named after the record's own TNNM
(falling back to the decoded blob's `ti`, then PCID).

It is dry-run by default: opening the panel shows the plan — write /
overwrite / skip and why (no HASH, HASH isn't a complete service-account
payload, or the file already exists) — and nothing touches disk until
**Write files** is clicked. Existing files are left alone unless
**Overwrite files that already exist** is checked, and overwriting asks for
one more confirmation naming the files it would replace. If two records
decode to the same `ti`, the first one to write wins and the second is
reported as skipped ("already exists") rather than silently clobbering it.

## What else this app does not do

Beyond that one dry-run/confirm workflow, there is no local cache and no
other bulk operations — every add, edit and delete is one live M3 call,
applied as soon as you click Save or Delete (with a plain confirm prompt on
delete). That matches EXT124MI itself: there is nothing to export or
re-import, because the sheet reads this table directly.
