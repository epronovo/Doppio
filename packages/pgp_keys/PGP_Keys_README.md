# pgp_keys — GnuPG key manager

A small Flask app for generating OpenPGP key pairs with the real `gpg`
binary, importing keys someone else generated, listing what's in the
keyring, exporting keys, running a live encrypt/decrypt round trip to prove
a pair actually works, and decrypting a real uploaded `.gpg`/`.pgp`/`.asc`
file. Every `gpg` invocation is echoed to an on-page terminal panel — the
exact argv (minus passphrases, which never touch argv) plus its
stdout/stderr — so nothing happens off-screen.

| File | Role |
|------|------|
| `PGP_Keys_App.py` | Flask routes + the `gpg` subprocess wrappers (`run_gpg`, `run_gpg_raw`) |
| `templates/PGP_Keys_Index.html` | The page — Generate / Keys / Test Encrypt & Decrypt / Decrypt a File tabs, plus the terminal |
| `gnupg_home/` | This app's own keyring (git-ignored, created on first run) |

## Quick start

```bash
pip install -r requirements.txt
brew install gnupg          # macOS, if gpg isn't already on PATH
python PGP_Keys_App.py      # http://127.0.0.1:8799/
```

If `gpg` isn't found, the page says so on load (Generate is disabled) rather
than failing a request later.

## Its own keyring, not yours

This tool never touches `~/.gnupg`. Everything it generates lives in
`gnupg_home/` inside this package, so it's safe to mint throwaway test keys
here without any risk to a real keyring on the machine.

`gpg-agent` binds a unix-domain socket under `--homedir`, and those have a
~108-character path limit — a repo checkout can land somewhere long enough to
blow past that. So `gnupg_home/` is never handed to `gpg` directly; instead
`_ensure_gnupg_home()` keeps a short-named symlink (`$TMPDIR/doppio_pgp_keys_gnupg`)
pointing at it, and that symlink is what `--homedir` actually gets. gpg
follows it straight through for every file it touches — the only thing the
symlink buys is a short path for the socket.

## Why RSA keys get a second command

`gpg --quick-generate-key` with an explicit algorithm (`rsa3072`, `rsa4096`)
and usage `sign` only produces a sign/certify-capable primary key — no
encryption subkey comes along for free the way it does for algorithm
`default` (which produces an Ed25519 primary with a Curve25519 encryption
subkey in one shot). So the RSA paths in `generate_key()` run a second
`--quick-add-key ... encr` to add one; without it the key would show up in
the list but be unusable in the Test tab (gpg reports "Unusable public key").

## Importing a key you already have

The Keys tab takes either pasted armor text or an uploaded key file (armored
or binary, public or private — `gpg --import` reads all four the same way)
and runs a plain `--import` against the local keyring. Once a private key is
imported, it behaves exactly like one generated here: it shows up with a
"private key" pill, and anything encrypted to it can be opened on the
Test tab or the Decrypt a File tab.

## Decrypting an uploaded file

The Decrypt a File tab is for a real file — someone's `secret.txt.gpg`, not
just a pasted armored message. `run_gpg_raw()` handles it as raw bytes end
to end (`run_gpg`'s `text=True` subprocess mode would try to decode binary
ciphertext as UTF-8 and corrupt it): the passphrase and the ciphertext bytes
both go over the fd `--passphrase-fd` points at, and the decrypted bytes
come back on stdout untouched.

Decrypted bytes are never written to disk — they sit in an in-memory,
one-time-download token (`_DOWNLOADS`, `/api/download/<token>`) for up to
10 minutes, and a UTF-8-decodable result also gets a same-page preview so
the common case (a decrypted note or small text file) doesn't need a
download round trip at all.

## Security notes

- Passphrases are passed to `gpg` only via `--passphrase-fd 0` / stdin, never
  as a command-line argument — so they never appear in `ps`, in the terminal
  panel, or in a shell history.
- Deleting a key (`--delete-secret-and-public-keys`) does not require its
  passphrase; the UI asks for a confirmation click instead.
- Exporting a private key shows the armored block in a copy/paste prompt
  rather than downloading it, so nothing sensitive lands in a Downloads
  folder by default.

## Branding

Same doppiogroup.com look as `packages/adp_concur` — near-black chrome header,
white content, Doppio red (`#d40814`) as the one accent, Poppins/Nunito Sans.
The terminal panel is the deliberate exception: a real dark console (macOS
traffic-light dots included) so the `gpg` commands it prints read the way a
terminal actually looks, rather than blending into the light theme.
