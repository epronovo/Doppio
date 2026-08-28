# Getting NG\M0042 ready for Concur PGP

**Host:** NGUSMSRV0019 · **Account:** NG\M0042 · **Written:** 2026-08-24

GnuPG keyrings are per-user. Everything below has to happen inside M0042's keyring
(`C:\Users\M0042\AppData\Roaming\gnupg`), because that is the account Task Scheduler runs the SFTP
scripts as. Setting this up under your own login looks like it worked and changes nothing that matters.

Two separate things are needed, for two different directions:

| | Needs | Why |
|---|---|---|
| **Outbound** (we send Concur a PO) | Concur's public key, **trusted** | encrypt to them |
| **Inbound** (Concur sends us invoices) | a **Barnes secret key**, and Concur holding its public half | decrypt what they send |

Outbound can be done today. Inbound needs a key pair that does not exist yet — `concur_exported_key_pair.key`
is not a PGP key despite the name, it is the PEM RSA SFTP login key, and gpg rejects it with
*"no valid OpenPGP data found"*.

---

## Step 0 — how you will act as M0042

This is the part that actually blocks people. Pick one:

**A. You have the password.** Simplest.

```
runas /user:NG\M0042 cmd.exe
```

**B. You do not have the password.** Task Scheduler already holds M0042's stored credentials. Point an
existing M0042 task at a setup `.cmd` temporarily, "Run" it on demand, then put the action back. Needs
local admin, no password. Do it on a stage task, not one that moves money, and put it back the same
sitting.

**C. Read-only look, no password, no admin over the account.** Copy the keyring and inspect the copy:

```
robocopy "C:\Users\M0042\AppData\Roaming\gnupg" "C:\Temp\m0042-keyring" /E /COPY:DAT /R:0
Test-GpgEncryption.ps1 -GpgHome C:\Temp\m0042-keyring
```

**Never point `--homedir` at the live folder** — gpg writes `trustdb.gpg` even on a read, and a file
created under your identity inside M0042's profile can lock the service account out of its own keyring.
Route C diagnoses; it cannot set anything up.

> If `C:\Users\M0042` does not exist, the account has never had its profile loaded. A task set to
> "Run whether user is logged on or not" creates it on first run.

---

## Step 1 — run the setup script

`Setup-ConcurPgp.ps1` does steps 2–5 and is safe to re-run. Start in report mode — it changes nothing
and tells you exactly what is present and what is missing:

```powershell
.\Setup-ConcurPgp.ps1
```

Then, as M0042:

```powershell
.\Setup-ConcurPgp.ps1 -Apply -GenerateKey -BackupPath "\\<vault-share>\concur-pgp"
```

`-GenerateKey` is deliberately a separate switch. Key generation is the one step you cannot undo by
re-running, and a second Barnes key created by accident would decrypt nothing Concur already holds a
public key for.

The rest of this document is what the script does and why, so you can do it by hand or check its work.

---

## Step 2 — generate the Barnes key pair

```
gpg --batch --gen-key barnes-key.params
```

with `barnes-key.params`:

```
Key-Type: RSA
Key-Length: 4096
Key-Usage: sign
Subkey-Type: RSA
Subkey-Length: 4096
Subkey-Usage: encrypt
Name-Real: Barnes Group Concur Integration
Name-Email: concur-integration@barnesgroupinc.com
Expire-Date: 3y
%no-protection
%commit
```

**On `%no-protection` (no passphrase).** This departs from the HSBC pattern, on purpose. A passphrase
stored in a file that the same account can read protects nothing — anyone who can read the keyring can
read the passphrase beside it. It also drags in `gpg-agent`, which does not behave predictably in a
non-interactive service session. The real control is NTFS permissions on the keyring folder. If you
would rather keep a passphrase, replace `%no-protection` with `Passphrase: <value>` and point
`App_PgpPassPhraseLocation` at a file containing it — `Decrypt-Files` handles both.

**On `Expire-Date: 3y`.** An unattended integration key with no expiry is one fewer outage; an expiring
one is one fewer stale credential. Three years, plus a calendar reminder set *now*.
`Test-GpgEncryption.ps1` warns at 45 days, but do not rely on someone reading a warning in a transcript.

Lock the folder down afterwards:

```
icacls "C:\Users\M0042\AppData\Roaming\gnupg" /inheritance:r /grant:r "M0042:(OI)(CI)F" /grant:r "Administrators:(OI)(CI)F"
```

---

## Step 3 — back the key up before anything else touches it

```
gpg --armor --export-secret-keys <barnes-fpr> > barnes_concur_SECRET.asc
gpg --export-ownertrust                       > barnes_ownertrust.txt
```

Also copy the revocation certificate gpg wrote at generation time, from
`%APPDATA%\gnupg\openpgp-revocs.d\<FPR>.rev`.

**If this key is lost, every file Concur ever encrypted to us is permanently unreadable** — including
everything already sitting in the archive folder. This is the highest-consequence step in the runbook and
it takes ten seconds.

Put it in a vault or a restricted share. **Not** `E:\KeysEncryptions\` — that tree already travels with
the code, which is finding #4 in the SFTP technical reference. `barnes_concur_SECRET.asc` is the private
key in the clear; move it off the box and delete the local copy.

Verified: a backup taken this way restores into an empty keyring and still decrypts ciphertext produced
before the restore.

---

## Step 4 — import Concur's key and make it trusted

```
gpg --import "E:\KeysEncryptions\Concur\Dev\concursolutions.asc"
gpg --lsign-key 2584FD81F647ECBFAA1EE1DDBAA83C168C51C89E
```

**Importing alone is not enough.** gpg refuses to encrypt to a key it cannot validate:

```
gpg: 59A289D9E8D85E80: There is no assurance this key belongs to the named user
gpg: encryption failed: Unusable public key
```

**`--lsign-key` only works once M0042 has a secret key of its own.** On an empty keyring it dies with
*"gpg: no default secret key"* and the trust problem stays exactly where it was. That is why Step 2 comes
first. If you only need outbound and have not generated the Barnes key yet, set ownertrust instead — it
works with no secret key present:

```
echo 2584FD81F647ECBFAA1EE1DDBAA83C168C51C89E:6: | gpg --import-ownertrust
```

Check it took — validity should be `f` (lsign) or `u` (ownertrust), not `-` or `q`:

```
gpg --with-colons --list-keys 2584FD81F647ECBFAA1EE1DDBAA83C168C51C89E
```

**Do not reach for `--trust-model always`.** It works, and it switches off trust checking for every
recipient this account will ever encrypt to.

Concur's key, for reference: rsa4096, created 2024-05-09, **no expiry**, encryption subkey
`59A289D9E8D85E80`, uid `concursolutions.asc <sapconcur-file-transfer@sap.com>`.

---

## Step 5 — send Concur our public key

```
gpg --armor --export <barnes-fpr> > barnes_concur_public.asc
```

Open it and confirm the header says `PUBLIC KEY BLOCK`. If it says `PRIVATE`, you exported the wrong
half — delete it and start over. The setup script checks this for you.

Inbound decrypt does nothing until Concur is actually encrypting to this key, so this is the step that
gates the whole inbound half. Ask them at the same time:

1. does the MFT endpoint accept PGP-encrypted inbound files at all
2. what remote path do inbound files go to (`App_SendRemotePath`, currently a `/in/` placeholder)

---

## Step 6 — verify, and do not fool yourself

```powershell
gpgconf --kill gpg-agent
.\Test-GpgEncryption.ps1 -Recipient 2584FD81F647ECBFAA1EE1DDBAA83C168C51C89E -RoundTrip
```

**Kill the agent first.** `gpg-agent` caches passphrases, and a cached one makes a decrypt test pass
while proving nothing. Observed directly while testing this: a deliberate wrong-passphrase decrypt
*succeeded* and wrote correct plaintext, purely from cache. With a cold agent the same command failed
properly with `Bad passphrase`, exit 2, no output file.

Also worth knowing:

- **PATH.** The scripts call bare `gpg`. If two `gpg.exe` are installed — Gpg4win and GnuPG both do this —
  which one runs can differ between your interactive session and the scheduled task.
  `Test-GpgEncryption.ps1` lists every copy it finds.
- **Encrypt failure is silent-ish.** With `--batch`, a trust failure writes no output file and returns
  non-zero. Both the send and get scripts check exit code *and* file existence, which is what catches it.
- Run `Test-GpgEncryption.ps1` as M0042 and compare against your own run. Different keyrings, different
  answers, and only M0042's is the one production uses.

---

## Quick state check

| Symptom | Meaning |
|---|---|
| `encryption failed: Unusable public key` | Concur's key imported but not trusted — Step 4 |
| `gpg: no default secret key` from `--lsign-key` | no Barnes key yet — do Step 2 first |
| `decryption failed: No secret key` | the file was encrypted to a key we do not hold |
| `no valid OpenPGP data found` | you pointed gpg at `concur_exported_key_pair.key` — that is an SSH key |
| decrypt test passes but you do not believe it | `gpgconf --kill gpg-agent` and run it again |
| keyring is empty and you are not M0042 | expected — see Step 0 |
