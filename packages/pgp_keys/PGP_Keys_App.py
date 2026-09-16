"""
PGP_Keys_App - Flask front end for generating OpenPGP key pairs with GnuPG,
and for proving a pair actually works with a round-trip encrypt/decrypt test.

Every gpg invocation goes through run_gpg() below, which is also what feeds
the on-page terminal: each call's argv (never a passphrase - those always
travel over stdin, never argv) and output are handed back to the browser so
the user can see exactly what ran.

The keyring is package-local, not the user's real ~/.gnupg - see
_ensure_gnupg_home() for why a short-path symlink stands in for it.

Every route returns JSON; the page itself is templates/PGP_Keys_Index.html.
"""
from __future__ import annotations

import argparse
import logging
import re
import secrets
import shlex
import shutil
import subprocess
import tempfile
import threading
import time
from pathlib import Path

from flask import Flask, Response, jsonify, render_template, request

BASE_DIR = Path(__file__).parent.resolve()

# The real keyring - persists across restarts, never committed (see .gitignore).
GNUPGHOME_REAL = BASE_DIR / "gnupg_home"

# gpg-agent binds a unix-domain socket under --homedir, and those have a
# ~108-char path limit. A checkout of this repo can land anywhere, so the
# path to GNUPGHOME_REAL is not reliably short enough. A short-named symlink
# in the system temp dir is what's actually passed to gpg as --homedir;
# gpg follows it straight through to the real directory for every file it
# reads or writes, so this is transparent other than keeping the socket path
# short.
GNUPGHOME_LINK = Path(tempfile.gettempdir()) / "doppio_pgp_keys_gnupg"

GPG_BIN = "gpg"
GPG_TIMEOUT = 30

log = logging.getLogger("PGP_Keys_App")

app = Flask(__name__, template_folder=str(BASE_DIR / "templates"))
app.config["MAX_CONTENT_LENGTH"] = 25 * 1024 * 1024

# Decrypted file downloads never touch disk - the bytes sit here just long
# enough for the browser to fetch them once. Token -> (data, filename, expiry).
_DOWNLOAD_LOCK = threading.Lock()
_DOWNLOADS: dict[str, tuple[bytes, str, float]] = {}
_DOWNLOAD_TTL_SECONDS = 600

# The three key shapes offered in the UI. "usage" is what --quick-generate-key
# gives the primary key; RSA primaries are sign/cert-only there, so an
# encryption subkey is added as a second step (see generate_key()). Ed25519's
# "default" usage already includes a cv25519 encryption subkey in one step.
KEY_TYPES = {
    "ed25519": {"label": "Ed25519 / Curve25519 (modern, recommended)",
                "algo": "default", "usage": "default", "subkey_algo": None},
    "rsa3072": {"label": "RSA 3072-bit", "algo": "rsa3072", "usage": "sign",
                "subkey_algo": "rsa3072"},
    "rsa4096": {"label": "RSA 4096-bit", "algo": "rsa4096", "usage": "sign",
                "subkey_algo": "rsa4096"},
}
EXPIRE_CHOICES = {"0": "never", "1y": "1 year", "2y": "2 years", "5y": "5 years"}


# ---------------------------------------------------------------- plumbing


def _ensure_gnupg_home() -> None:
    """Create the real keyring dir and point the short symlink at it."""
    GNUPGHOME_REAL.mkdir(mode=0o700, exist_ok=True)
    GNUPGHOME_REAL.chmod(0o700)
    # Cache TTLs at 0 disable gpg-agent's passphrase cache: without this, a
    # correct decrypt unlocks the key in the agent and a later decrypt of a
    # different message with the WRONG passphrase would still succeed
    # (silently ignoring the passphrase we sent) as long as the agent still
    # had the cached one - misleading for a tool whose whole point is to
    # demonstrate encrypt/decrypt actually working.
    agent_conf = GNUPGHOME_REAL / "gpg-agent.conf"
    if not agent_conf.exists():
        agent_conf.write_text(
            "allow-loopback-pinentry\ndefault-cache-ttl 0\nmax-cache-ttl 0\n")

    if GNUPGHOME_LINK.is_symlink() or GNUPGHOME_LINK.exists():
        if GNUPGHOME_LINK.resolve() != GNUPGHOME_REAL.resolve():
            if GNUPGHOME_LINK.is_symlink():
                GNUPGHOME_LINK.unlink()
            else:
                raise RuntimeError(
                    f"{GNUPGHOME_LINK} exists and is not the expected symlink; "
                    "remove it and retry")
    if not GNUPGHOME_LINK.exists():
        GNUPGHOME_LINK.symlink_to(GNUPGHOME_REAL)


def gpg_available() -> bool:
    return shutil.which(GPG_BIN) is not None


def run_gpg(args: list[str], stdin_data: str | None = None) -> dict:
    """Run one gpg command against our keyring and capture it for the terminal.

    Secrets (passphrases) are only ever passed via stdin_data, never in args,
    so the displayed command line is always safe to show as-is.
    """
    full_args = [GPG_BIN, "--homedir", str(GNUPGHOME_LINK), *args]
    try:
        proc = subprocess.run(
            full_args, input=stdin_data, capture_output=True, text=True,
            timeout=GPG_TIMEOUT,
        )
        stdout, stderr, returncode = proc.stdout, proc.stderr, proc.returncode
    except subprocess.TimeoutExpired:
        stdout, stderr, returncode = "", "gpg timed out", -1
    except FileNotFoundError:
        stdout, stderr, returncode = "", "gpg is not installed or not on PATH", -1

    return {
        "command": " ".join(shlex.quote(a) for a in full_args),
        "stdout": stdout,
        "stderr": stderr,
        "returncode": returncode,
        "ok": returncode == 0,
    }


def run_gpg_raw(args: list[str], stdin_bytes: bytes = b"") -> dict:
    """Like run_gpg, but for commands whose input/output is arbitrary bytes
    (an imported key file, an encrypted upload) rather than text - subprocess's
    text=True mode would try to decode that as UTF-8 and corrupt it.

    Only the passphrase - always text - ever gets prepended to stdin_bytes, so
    the displayed command line is still safe to show as-is.
    """
    full_args = [GPG_BIN, "--homedir", str(GNUPGHOME_LINK), *args]
    try:
        proc = subprocess.run(
            full_args, input=stdin_bytes, capture_output=True,
            timeout=GPG_TIMEOUT,
        )
        stdout, stderr, returncode = proc.stdout, proc.stderr, proc.returncode
    except subprocess.TimeoutExpired:
        stdout, stderr, returncode = b"", b"gpg timed out", -1
    except FileNotFoundError:
        stdout, stderr, returncode = b"", b"gpg is not installed or not on PATH", -1

    return {
        "command": " ".join(shlex.quote(a) for a in full_args),
        "stdout_bytes": stdout,
        "stderr": stderr.decode("utf-8", "replace"),
        "returncode": returncode,
        "ok": returncode == 0,
    }


def _command_log(raw: dict, stdout_display: str = "") -> dict:
    """The JSON-safe shape the terminal panel expects - never raw bytes."""
    return {
        "command": raw["command"],
        "stdout": stdout_display,
        "stderr": raw["stderr"],
        "returncode": raw["returncode"],
        "ok": raw["ok"],
    }


def list_secret_fingerprints() -> set[str]:
    result = run_gpg(["--list-secret-keys", "--with-colons"])
    return {
        line.split(":")[9] for line in result["stdout"].splitlines()
        if line.startswith("fpr:")
    }


# gpg's own pretty-printer (the "sec   rsa4096" you see in --list-secret-keys)
# has no --with-colons equivalent - the colon pub/sub record gives the numeric
# algorithm id (field 3), the bit length (field 2), and, for ECC, the curve
# name (field 16). This rebuilds the same short label from those.
_ALGO_NAMES = {"1": "rsa", "2": "rsa", "3": "rsa", "16": "elgamal",
               "17": "dsa", "18": "ecdh", "19": "ecdsa", "22": "eddsa"}


def _algo_label(algo_num: str, keylength: str, curve: str) -> str:
    if curve:
        return curve
    name = _ALGO_NAMES.get(algo_num, f"algo{algo_num}")
    return f"{name}{keylength}"


def build_key_list() -> list[dict]:
    """Merge --list-keys and --list-secret-keys colon output into one row
    per primary key: fingerprint, uid, algo, dates, capability flags, and
    whether an encryption-capable subkey exists (needed for the Encrypt tab)."""
    secret_fprs = list_secret_fingerprints()
    pub = run_gpg(["--list-keys", "--with-colons"])

    keys: list[dict] = []
    current: dict | None = None
    for line in pub["stdout"].splitlines():
        field = line.split(":")
        kind = field[0]
        if kind == "pub":
            if current is not None:
                keys.append(current)
            curve = field[16] if len(field) > 16 else ""
            current = {
                "fingerprint": None,
                "key_id": field[4],
                "algo": _algo_label(field[3], field[2], curve),
                "created": field[5],
                "expires": field[6] or None,
                "capabilities": field[11],
                "uid": None,
                "has_secret": False,
                "can_encrypt": "e" in field[11].lower(),
            }
        elif kind == "fpr" and current is not None and current["fingerprint"] is None:
            current["fingerprint"] = field[9]
            current["has_secret"] = field[9] in secret_fprs
        elif kind == "uid" and current is not None and current["uid"] is None:
            current["uid"] = field[9]
        elif kind == "sub" and current is not None and "e" in field[11].lower():
            current["can_encrypt"] = True
    if current is not None:
        keys.append(current)
    return keys


def find_key(fingerprint: str) -> dict | None:
    # A key that was just generated can occasionally miss the very next
    # --list-keys if keyboxd hasn't flushed yet; a couple of quick retries
    # is cheaper than making the caller re-request.
    for attempt in range(3):
        for key in build_key_list():
            if key["fingerprint"] == fingerprint:
                return key
        if attempt < 2:
            time.sleep(0.15)
    return None


def _output_filename(uploaded_name: str) -> str:
    """Best-effort plaintext filename for a decrypted upload: drop a
    .gpg/.pgp/.asc suffix if there is one, otherwise mark it decrypted so it's
    never saved over the original encrypted file by accident."""
    base = Path(uploaded_name).name or "decrypted-file"
    for ext in (".gpg", ".pgp", ".asc"):
        if base.lower().endswith(ext):
            return base[: -len(ext)] or "decrypted-file"
    return f"{base}.decrypted"


def _store_download(data: bytes, filename: str) -> str:
    token = secrets.token_urlsafe(16)
    now = time.time()
    with _DOWNLOAD_LOCK:
        for expired in [t for t, (_, _, exp) in _DOWNLOADS.items() if exp < now]:
            del _DOWNLOADS[expired]
        _DOWNLOADS[token] = (data, filename, now + _DOWNLOAD_TTL_SECONDS)
    return token


@app.errorhandler(Exception)
def _handle(exc):
    log.exception("request failed")
    return jsonify(ok=False, error=str(exc)), 500


# -------------------------------------------------------------------- pages


@app.route("/")
def index():
    return render_template("PGP_Keys_Index.html")


@app.route("/api/status")
def api_status():
    if not gpg_available():
        return jsonify(ok=False, installed=False,
                        error="gpg was not found on PATH. Install GnuPG "
                              "(macOS: brew install gnupg) and reload.")
    _ensure_gnupg_home()
    version = run_gpg(["--version"])
    first_line = version["stdout"].splitlines()[0] if version["stdout"] else "unknown"
    return jsonify(ok=True, installed=True, version=first_line,
                   homedir=str(GNUPGHOME_REAL))


@app.route("/api/keys")
def api_keys():
    _ensure_gnupg_home()
    return jsonify(ok=True, keys=build_key_list())


@app.route("/api/import", methods=["POST"])
def api_import():
    """Import a key someone else generated - pasted armor text or an
    uploaded key file, public or private, armored or binary. --import reads
    all of those the same way, so the two input modes just pick where the
    bytes come from."""
    keyfile = request.files.get("keyfile")
    armored = (request.form.get("armored") or "").strip()

    if keyfile and keyfile.filename:
        data = keyfile.read()
    elif armored:
        data = armored.encode("utf-8")
    else:
        return jsonify(ok=False, error="Paste an armored key or choose a file."), 400

    _ensure_gnupg_home()
    raw = run_gpg_raw(["--batch", "--yes", "--import"], stdin_bytes=data)
    cmd = _command_log(raw)
    if not raw["ok"]:
        return jsonify(ok=False, error="Import failed - is this a valid OpenPGP key?",
                        commands=[cmd])
    return jsonify(ok=True, commands=[cmd], keys=build_key_list())


@app.route("/api/generate", methods=["POST"])
def api_generate():
    data = request.get_json(force=True)
    name = (data.get("name") or "").strip()
    email = (data.get("email") or "").strip()
    comment = (data.get("comment") or "").strip()
    key_type = data.get("key_type") or "ed25519"
    expire = data.get("expire") or "1y"
    passphrase = data.get("passphrase") or ""

    if not name or not email:
        return jsonify(ok=False, error="Name and email are required."), 400
    if key_type not in KEY_TYPES:
        return jsonify(ok=False, error=f"Unknown key type: {key_type}"), 400
    if expire not in EXPIRE_CHOICES:
        return jsonify(ok=False, error=f"Unknown expiration: {expire}"), 400

    spec = KEY_TYPES[key_type]
    uid = f"{name} ({comment}) <{email}>" if comment else f"{name} <{email}>"

    _ensure_gnupg_home()
    commands = []

    gen = run_gpg(
        ["--batch", "--pinentry-mode", "loopback", "--passphrase-fd", "0",
         "--quick-generate-key", uid, spec["algo"], spec["usage"], expire],
        stdin_data=passphrase,
    )
    commands.append(gen)
    if not gen["ok"]:
        return jsonify(ok=False, error="Key generation failed.", commands=commands)

    # gpg always writes a revocation cert right after generating a primary
    # key, named after its fingerprint - pull it from there rather than
    # diffing --list-secret-keys before/after, which can race a keyboxd
    # flush and miss the key that was just created.
    match = re.search(r"openpgp-revocs\.d/([0-9A-F]+)\.rev", gen["stderr"])
    fingerprint = match.group(1) if match else None
    if fingerprint is None:
        return jsonify(ok=False, error="Key was generated but its fingerprint "
                                        "could not be found.", commands=commands)

    if spec["subkey_algo"]:
        subkey = run_gpg(
            ["--batch", "--pinentry-mode", "loopback", "--passphrase-fd", "0",
             "--quick-add-key", fingerprint, spec["subkey_algo"], "encr", expire],
            stdin_data=passphrase,
        )
        commands.append(subkey)
        if not subkey["ok"]:
            return jsonify(ok=False, error="Primary key was created, but adding "
                                            "the encryption subkey failed.",
                            commands=commands)

    return jsonify(ok=True, commands=commands, key=find_key(fingerprint))


@app.route("/api/delete", methods=["POST"])
def api_delete():
    data = request.get_json(force=True)
    fingerprint = (data.get("fingerprint") or "").strip()
    if not fingerprint:
        return jsonify(ok=False, error="fingerprint is required"), 400

    _ensure_gnupg_home()
    result = run_gpg(["--batch", "--yes", "--delete-secret-and-public-keys",
                       fingerprint])
    return jsonify(ok=result["ok"], commands=[result],
                    error=None if result["ok"] else "Delete failed.")


@app.route("/api/export")
def api_export():
    """Public-key export only - no secret material, so no passphrase needed."""
    fingerprint = (request.args.get("fingerprint") or "").strip()
    if not fingerprint:
        return jsonify(ok=False, error="fingerprint is required"), 400

    _ensure_gnupg_home()
    result = run_gpg(["--armor", "--export", fingerprint])
    if not result["ok"]:
        return jsonify(ok=False, error="Export failed.", commands=[result])
    return jsonify(ok=True, commands=[result], armored=result["stdout"])


@app.route("/api/export-private", methods=["POST"])
def api_export_private():
    # --export-secret-keys re-encodes the actual key material, which makes
    # gpg-agent want to unlock it - without --pinentry-mode loopback plus a
    # passphrase fed over stdin it fails trying to spawn an interactive
    # pinentry ("Inappropriate ioctl for device") since this process has no
    # controlling terminal.
    data = request.get_json(force=True)
    fingerprint = (data.get("fingerprint") or "").strip()
    passphrase = data.get("passphrase") or ""
    if not fingerprint:
        return jsonify(ok=False, error="fingerprint is required"), 400

    _ensure_gnupg_home()
    result = run_gpg(
        ["--batch", "--pinentry-mode", "loopback", "--passphrase-fd", "0",
         "--armor", "--export-secret-keys", fingerprint],
        stdin_data=passphrase,
    )
    if not result["ok"]:
        return jsonify(ok=False, error="Export failed - check the passphrase.",
                        commands=[result])
    return jsonify(ok=True, commands=[result], armored=result["stdout"])


@app.route("/api/encrypt", methods=["POST"])
def api_encrypt():
    data = request.get_json(force=True)
    fingerprint = (data.get("fingerprint") or "").strip()
    message = data.get("message") or ""
    if not fingerprint or not message:
        return jsonify(ok=False, error="fingerprint and message are required"), 400

    _ensure_gnupg_home()
    result = run_gpg(
        ["--batch", "--yes", "--armor", "--trust-model", "always",
         "--recipient", fingerprint, "--encrypt"],
        stdin_data=message,
    )
    if not result["ok"]:
        return jsonify(ok=False, error="Encryption failed.", commands=[result])
    return jsonify(ok=True, commands=[result], ciphertext=result["stdout"])


@app.route("/api/decrypt", methods=["POST"])
def api_decrypt():
    data = request.get_json(force=True)
    ciphertext = data.get("ciphertext") or ""
    passphrase = data.get("passphrase") or ""
    if not ciphertext.strip():
        return jsonify(ok=False, error="ciphertext is required"), 400

    _ensure_gnupg_home()
    result = run_gpg(
        ["--batch", "--yes", "--pinentry-mode", "loopback",
         "--passphrase-fd", "0", "--decrypt"],
        # gpg reads the passphrase and the armored message from the same
        # stdin stream (--passphrase-fd 0), passphrase line first.
        stdin_data=f"{passphrase}\n{ciphertext}",
    )
    if not result["ok"]:
        return jsonify(ok=False, error="Decryption failed - wrong passphrase, "
                                        "or this isn't a message our keyring "
                                        "can open.", commands=[result])
    return jsonify(ok=True, commands=[result], plaintext=result["stdout"])


@app.route("/api/decrypt-file", methods=["POST"])
def api_decrypt_file():
    """Decrypt an uploaded file - a real .gpg/.pgp/.asc file, not just a
    pasted armored message - using whatever private key in the keyring
    matches it. Binary-safe throughout: run_gpg_raw, not run_gpg, since a
    non-armored ciphertext (or the plaintext coming back out) is not
    guaranteed to be valid UTF-8."""
    upload = request.files.get("file")
    passphrase = request.form.get("passphrase") or ""
    if not upload or not upload.filename:
        return jsonify(ok=False, error="Choose an encrypted file first."), 400

    ciphertext = upload.read()
    _ensure_gnupg_home()
    raw = run_gpg_raw(
        ["--batch", "--yes", "--pinentry-mode", "loopback", "--passphrase-fd", "0",
         "--decrypt"],
        # Same trick as /api/decrypt: passphrase line, then the ciphertext
        # bytes, both over the fd --passphrase-fd points at.
        stdin_bytes=passphrase.encode("utf-8") + b"\n" + ciphertext,
    )
    plaintext = raw["stdout_bytes"]
    cmd = _command_log(
        raw, stdout_display=f"[{len(plaintext)} bytes of decrypted output]" if raw["ok"] else "")
    if not raw["ok"]:
        return jsonify(ok=False, error="Decryption failed - wrong passphrase, or this "
                                        "isn't a message our keyring can open.",
                        commands=[cmd])

    filename = _output_filename(upload.filename)
    token = _store_download(plaintext, filename)

    # Show it inline too when it's plausibly text, so the common case (a
    # decrypted note or config file) doesn't require a download round trip
    # just to see what it says.
    text_preview = None
    if len(plaintext) <= 200_000:
        try:
            text_preview = plaintext.decode("utf-8")
        except UnicodeDecodeError:
            pass

    return jsonify(ok=True, commands=[cmd], filename=filename, size=len(plaintext),
                    download_token=token, text_preview=text_preview)


@app.route("/api/download/<token>")
def api_download(token):
    # One-time: popped on read, so a link can't be replayed once used, and a
    # forgotten tab doesn't leave decrypted plaintext sitting around after
    # _DOWNLOAD_TTL_SECONDS anyway.
    with _DOWNLOAD_LOCK:
        entry = _DOWNLOADS.pop(token, None)
    if entry is None or entry[2] < time.time():
        return jsonify(ok=False, error="This download link has expired."), 404
    data, filename, _ = entry
    return Response(data, mimetype="application/octet-stream",
                     headers={"Content-Disposition": f'attachment; filename="{filename}"'})


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", type=int, default=8799)
    parser.add_argument("--debug", action="store_true")
    args = parser.parse_args()

    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
    _ensure_gnupg_home()
    app.run(host=args.host, port=args.port, debug=args.debug)


if __name__ == "__main__":
    main()
