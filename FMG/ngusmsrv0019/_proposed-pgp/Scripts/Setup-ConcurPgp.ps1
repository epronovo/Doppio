<#
.SYNOPSIS
    One-time GnuPG setup for the Concur interface, run as the service account.

.DESCRIPTION
    Puts the NG\M0042 keyring into the state the Concur scripts need:

      1. a Barnes PGP key pair          -> lets us DECRYPT what Concur sends us
      2. Concur's public key, trusted   -> lets us ENCRYPT what we send Concur
      3. a backup of 1                  -> without it, losing the profile makes
                                           every archived Concur file unreadable
      4. our public key, exported       -> this is what you hand to Concur

    Safe to re-run. Every step checks for its own result first and skips if it is
    already done, so this doubles as a state report: run it with no switches and
    it tells you what is present and what is missing without changing anything.

    GnuPG keyrings are PER-USER. Running this as yourself sets up YOUR keyring,
    which the scheduled tasks never touch. It must run as the account in
    -ExpectedAccount or it is pointless; the script refuses by default.

.PARAMETER Apply
    Actually make changes. Without it the script only reports.

.PARAMETER GenerateKey
    Create the Barnes key pair if absent. Separate from -Apply on purpose: key
    generation is the one step you cannot undo by re-running, and a SECOND key
    silently created here would decrypt nothing Concur already holds a key for.

.PARAMETER ConcurPublicKey
    Concur's ASCII-armoured public key.

.PARAMETER BackupPath
    Directory for the secret-key backup, ownertrust and revocation certificate.
    Point this at a vault or a restricted share. NOT E:\KeysEncryptions - that
    tree already travels with the code, which is finding #4 in the SFTP
    technical reference.

.PARAMETER PublicKeyOutPath
    Where to write the Barnes PUBLIC key to send to Concur.

.PARAMETER ExpectedAccount
    The account that runs the scheduled tasks. Refuses to run as anyone else.

.PARAMETER Force
    Run even when the current account is not -ExpectedAccount. For rehearsing on
    a throwaway login only.

.EXAMPLE
    # report only - always start here
    .\Setup-ConcurPgp.ps1

.EXAMPLE
    # full setup
    .\Setup-ConcurPgp.ps1 -Apply -GenerateKey

.NOTES
    Verify afterwards with:
        gpgconf --kill gpg-agent      # see the WARNING below, this matters
        .\Test-GpgEncryption.ps1 -Recipient 2584FD81F647ECBFAA1EE1DDBAA83C168C51C89E -RoundTrip

    WARNING - gpg-agent caches passphrases. A decrypt test can pass using a
    cached passphrase from an earlier run and tell you nothing. Kill the agent
    before any test you intend to believe.
#>
[CmdletBinding()]
param(
    [switch]$Apply,
    [switch]$GenerateKey,
    [string]$ConcurPublicKey  = 'E:\KeysEncryptions\Concur\Dev\concursolutions.asc',
    [string]$BackupPath,
    [string]$PublicKeyOutPath = 'E:\KeysEncryptions\Concur\barnes_concur_public.asc',
    [string]$ExpectedAccount  = 'M0042',
    [switch]$Force
)

$ErrorActionPreference = 'Continue'

# Concur's key, from concursolutions.asc: rsa4096, created 2024-05-09, no expiry,
# uid "concursolutions.asc <sapconcur-file-transfer@sap.com>".
# Pinned by fingerprint so a second key with the same address cannot be picked.
$ConcurFingerprint = '2584FD81F647ECBFAA1EE1DDBAA83C168C51C89E'

$BarnesName  = 'Barnes Group Concur Integration'
$BarnesEmail = 'concur-integration@barnesgroupinc.com'

function Say {
    param(
        [ValidateSet('OK','MISSING','DID','SKIP','WARN','FAIL','INFO')][string]$Status,
        [string]$Message
    )
    $color = @{ OK='Green'; MISSING='Yellow'; DID='Cyan'; SKIP='DarkGray'
                WARN='Yellow'; FAIL='Red'; INFO='Gray' }[$Status]
    Write-Host ("  [{0}] {1}" -f $Status.PadRight(7), $Message) -ForegroundColor $color
}

function Section { param([string]$T) Write-Host ''; Write-Host "== $T" -ForegroundColor White }

Write-Host ''
Write-Host '=========================================================' -ForegroundColor White
Write-Host ' Concur PGP setup' -ForegroundColor White
Write-Host '=========================================================' -ForegroundColor White

# ------------------------------------------------------------ identity
Section 'Identity'
Say INFO ("Machine    : {0}" -f $env:COMPUTERNAME)
Say INFO ("Running as : {0}\{1}" -f $env:USERDOMAIN, $env:USERNAME)

if ($env:USERNAME -ne $ExpectedAccount) {
    if ($Force) {
        Say WARN ("Not {0}. -Force given, continuing - this sets up the WRONG keyring." -f $ExpectedAccount)
    }
    else {
        Say FAIL ("This must run as {0}. Keyrings are per-user; setting up yours does nothing" -f $ExpectedAccount)
        Say FAIL 'for the scheduled tasks. Use runas, or point an existing M0042 scheduled task'
        Say FAIL 'at this script. Override with -Force only to rehearse.'
        exit 1
    }
}
else {
    Say OK ("Running as {0}." -f $ExpectedAccount)
}

if (-not $Apply) {
    Say INFO 'REPORT ONLY. Nothing will be changed. Re-run with -Apply to act.'
}

# ------------------------------------------------------------ gpg
Section 'gpg.exe'
$gpgCmds = @(Get-Command gpg -All -ErrorAction SilentlyContinue |
             Where-Object { $_.CommandType -eq 'Application' })
if ($gpgCmds.Count -eq 0) {
    Say FAIL 'No gpg.exe on PATH. The Concur scripts call bare "gpg" and cannot work.'
    exit 1
}
Say OK ("Using {0}" -f $gpgCmds[0].Source)
if (($gpgCmds.Source | Sort-Object -Unique).Count -gt 1) {
    Say WARN ("{0} different gpg.exe on PATH. Which one runs can differ between your" -f $gpgCmds.Count)
    Say WARN 'session and the scheduled task. Test-GpgEncryption.ps1 lists them.'
}
$homeLine = & gpg --version 2>&1 | Where-Object { $_ -match '^Home:\s*(.+)$' }
if ($homeLine -and $homeLine -match '^Home:\s*(.+)$') { Say INFO ("Keyring    : {0}" -f $Matches[1].Trim()) }

# ------------------------------------------------- 1. the Barnes key pair
Section '1. Barnes PGP key pair (needed to DECRYPT inbound)'

$secColons = & gpg --batch --with-colons --list-secret-keys $BarnesEmail 2>$null
$haveSecret = @($secColons | Where-Object { $_ -like 'sec:*' }).Count -gt 0
$barnesFpr = $null
if ($haveSecret) {
    $barnesFpr = ($secColons | Where-Object { $_ -like 'fpr:*' } | Select-Object -First 1) -split ':' |
                 Select-Object -Index 9
    Say OK ("Present: {0}  <{1}>" -f $barnesFpr, $BarnesEmail)
}
else {
    Say MISSING ("No secret key for <{0}>. Inbound decrypt cannot work without one." -f $BarnesEmail)
    Say INFO 'Note concur_exported_key_pair.key is NOT this - it is a PEM RSA (SSH) key.'

    if ($Apply -and $GenerateKey) {
        # No passphrase, deliberately. A passphrase in a file that the same account
        # can read protects nothing, and it drags in gpg-agent, which does not
        # behave predictably in a non-interactive service session. The real control
        # is NTFS permissions on the keyring folder - see the end of this script.
        $params = @(
            'Key-Type: RSA'
            'Key-Length: 4096'
            'Key-Usage: sign'
            'Subkey-Type: RSA'
            'Subkey-Length: 4096'
            'Subkey-Usage: encrypt'
            "Name-Real: $BarnesName"
            "Name-Email: $BarnesEmail"
            'Expire-Date: 3y'
            '%no-protection'
            '%commit'
        )
        $paramFile = Join-Path $env:TEMP ("barnes-key-{0}.params" -f [guid]::NewGuid().ToString('N').Substring(0,8))
        try {
            $params | Set-Content -Path $paramFile -Encoding ASCII
            & gpg --batch --gen-key $paramFile 2>&1 | ForEach-Object { Say INFO ("gpg: {0}" -f $_) }
        }
        finally {
            Remove-Item $paramFile -Force -ErrorAction SilentlyContinue
        }

        $secColons = & gpg --batch --with-colons --list-secret-keys $BarnesEmail 2>$null
        if (@($secColons | Where-Object { $_ -like 'sec:*' }).Count -gt 0) {
            $barnesFpr = ($secColons | Where-Object { $_ -like 'fpr:*' } | Select-Object -First 1) -split ':' |
                         Select-Object -Index 9
            $haveSecret = $true
            Say DID ("Generated {0}" -f $barnesFpr)
            Say WARN 'Expires in 3 years. Test-GpgEncryption.ps1 warns at 45 days - put a'
            Say WARN 'calendar reminder in now, do not rely on noticing the warning.'
        }
        else {
            Say FAIL 'Key generation did not produce a secret key.'
        }
    }
    elseif ($Apply) {
        Say SKIP 'Add -GenerateKey to create it. Held back on purpose - a duplicate key here'
        Say SKIP 'would decrypt nothing Concur already holds a public key for.'
    }
}

# ------------------------------------------------------------ 2. backup
Section '2. Backup of the Barnes secret key'

if (-not $haveSecret) {
    Say SKIP 'Nothing to back up yet.'
}
elseif (-not $BackupPath) {
    Say WARN 'No -BackupPath given, so no backup was taken.'
    Say WARN 'Lose this key and every file Concur ever encrypted to us becomes'
    Say WARN 'permanently unreadable, including everything already archived.'
    Say INFO 'Re-run with -BackupPath pointed at a vault or restricted share.'
}
else {
    $secretOut = Join-Path $BackupPath 'barnes_concur_SECRET.asc'
    $trustOut  = Join-Path $BackupPath 'barnes_ownertrust.txt'

    if (Test-Path $secretOut) {
        Say OK ("Backup already present: {0}" -f $secretOut)
    }
    elseif ($Apply) {
        if (-not (Test-Path $BackupPath)) {
            New-Item -Path $BackupPath -ItemType Directory -Force | Out-Null
        }
        & gpg --batch --armor --export-secret-keys $barnesFpr 2>$null | Set-Content -Path $secretOut -Encoding ASCII
        & gpg --batch --export-ownertrust 2>$null | Set-Content -Path $trustOut -Encoding ASCII

        if ((Test-Path $secretOut) -and (Get-Item $secretOut).Length -gt 0) {
            Say DID ("Wrote {0}" -f $secretOut)
            Say DID ("Wrote {0}" -f $trustOut)
            Say WARN 'This file IS the private key, in the clear. Move it to a vault and'
            Say WARN 'delete this copy. Do not leave it on E:\.'
            $revoc = Join-Path $env:APPDATA 'gnupg\openpgp-revocs.d'
            if (Test-Path $revoc) {
                Say INFO ("Also back up the revocation certificate in {0}" -f $revoc)
            }
        }
        else {
            Say FAIL 'Backup export produced nothing.'
        }
    }
    else {
        Say MISSING ("Would write {0}" -f $secretOut)
    }
}

# ------------------------------------------- 3. Concur's key, imported+trusted
Section "3. Concur's public key (needed to ENCRYPT outbound)"

$pubColons = & gpg --batch --with-colons --list-keys $ConcurFingerprint 2>$null
$haveConcur = @($pubColons | Where-Object { $_ -like 'pub:*' }).Count -gt 0

if ($haveConcur) {
    Say OK ("Imported: {0}" -f $ConcurFingerprint)
}
else {
    Say MISSING 'Concur key not in this keyring.'
    if (-not (Test-Path $ConcurPublicKey)) {
        Say FAIL ("And the source file is missing: {0}" -f $ConcurPublicKey)
    }
    elseif ($Apply) {
        & gpg --batch --quiet --import $ConcurPublicKey 2>&1 | ForEach-Object { Say INFO ("gpg: {0}" -f $_) }
        $pubColons = & gpg --batch --with-colons --list-keys $ConcurFingerprint 2>$null
        $haveConcur = @($pubColons | Where-Object { $_ -like 'pub:*' }).Count -gt 0
        if ($haveConcur) { Say DID ("Imported {0}" -f $ConcurFingerprint) }
        else { Say FAIL ("Import did not yield {0}. Wrong file?" -f $ConcurFingerprint) }
    }
    else {
        Say MISSING ("Would import {0}" -f $ConcurPublicKey)
    }
}

if ($haveConcur) {
    # Import alone is NOT enough. gpg refuses to encrypt to a key it cannot
    # validate: "encryption failed: Unusable public key". Validity 'f' or better
    # is what we are after.
    $pubLine  = $pubColons | Where-Object { $_ -like 'pub:*' } | Select-Object -First 1
    $validity = ($pubLine -split ':')[1]

    if ($validity -in @('f','u','m')) {
        Say OK ("Trusted (validity '{0}') - encrypt will not fail on trust." -f $validity)
    }
    else {
        Say MISSING ("Validity is '{0}'. Encrypt WILL fail: 'Unusable public key'." -f $validity)

        if ($Apply -and $haveSecret) {
            # --lsign-key needs a secret key in the keyring. That is why this step
            # comes after key generation: on a keyring with no secret key it dies
            # with "gpg: no default secret key" and the trust problem stays.
            & gpg --batch --yes --quiet --lsign-key $ConcurFingerprint 2>&1 |
                ForEach-Object { Say INFO ("gpg: {0}" -f $_) }

            $pubColons = & gpg --batch --with-colons --list-keys $ConcurFingerprint 2>$null
            $pubLine   = $pubColons | Where-Object { $_ -like 'pub:*' } | Select-Object -First 1
            $validity  = ($pubLine -split ':')[1]
            if ($validity -in @('f','u','m')) { Say DID ("Locally signed; validity now '{0}'." -f $validity) }
            else { Say FAIL ("Still '{0}' after lsign." -f $validity) }
        }
        elseif ($Apply) {
            # Fallback for an outbound-only setup, before the Barnes key exists.
            Say INFO 'No secret key yet, so --lsign-key would fail ("no default secret key").'
            Say INFO 'Setting ownertrust instead - this works with an empty keyring:'
            "${ConcurFingerprint}:6:" | & gpg --batch --import-ownertrust 2>&1 |
                ForEach-Object { Say INFO ("gpg: {0}" -f $_) }
            $pubColons = & gpg --batch --with-colons --list-keys $ConcurFingerprint 2>$null
            $pubLine   = $pubColons | Where-Object { $_ -like 'pub:*' } | Select-Object -First 1
            Say DID ("Ownertrust set; validity now '{0}'." -f ($pubLine -split ':')[1])
        }
        else {
            Say MISSING 'Would fix trust (lsign if a secret key exists, else ownertrust).'
        }
        Say INFO 'Do NOT "fix" this with --trust-model always - it switches off trust'
        Say INFO 'checking for every recipient this account will ever encrypt to.'
    }
}

# ------------------------------------------- 4. export our public key
Section '4. Barnes public key to hand to Concur'

if (-not $haveSecret) {
    Say SKIP 'No Barnes key to export yet.'
}
elseif (Test-Path $PublicKeyOutPath) {
    Say OK ("Already exported: {0}" -f $PublicKeyOutPath)
}
elseif ($Apply) {
    $dir = Split-Path -Path $PublicKeyOutPath -Parent
    if ($dir -and -not (Test-Path $dir)) { New-Item -Path $dir -ItemType Directory -Force | Out-Null }
    & gpg --batch --armor --export $barnesFpr 2>$null | Set-Content -Path $PublicKeyOutPath -Encoding ASCII

    if ((Test-Path $PublicKeyOutPath) -and (Get-Item $PublicKeyOutPath).Length -gt 0) {
        # Cheap guard against handing a partner the wrong half of the pair.
        if (Select-String -Path $PublicKeyOutPath -Pattern 'PRIVATE KEY' -Quiet) {
            Say FAIL 'Export contains a PRIVATE KEY block. Do NOT send this. Delete it.'
        }
        else {
            Say DID ("Wrote {0} - this is the file Concur needs." -f $PublicKeyOutPath)
        }
    }
    else {
        Say FAIL 'Export produced nothing.'
    }
}
else {
    Say MISSING ("Would write {0}" -f $PublicKeyOutPath)
}

# ------------------------------------------------------------ summary
Section 'Next'

if ($haveSecret -and $haveConcur) {
    Say OK 'Keyring looks ready. Verify it for real:'
    Say INFO '    gpgconf --kill gpg-agent'
    Say INFO ("    .\Test-GpgEncryption.ps1 -Recipient {0} -RoundTrip" -f $ConcurFingerprint)
    Say WARN 'Kill the agent first. It caches passphrases, and a cached one makes a'
    Say WARN 'decrypt test pass while proving nothing.'
    Write-Host ''
    Say INFO 'Then: send Concur the public key, confirm they accept PGP inbound, and'
    Say INFO 'confirm their inbound remote path before the first real send.'
}
else {
    Say WARN 'Not ready yet - see the missing items above.'
}

$keyringDir = Join-Path $env:APPDATA 'gnupg'
Write-Host ''
Say INFO ("Lock the keyring down: {0} should be readable by {1} and admins only." -f $keyringDir, $env:USERNAME)
Say INFO '    icacls "%APPDATA%\gnupg" /inheritance:r /grant:r "%USERNAME%:(OI)(CI)F" /grant:r "Administrators:(OI)(CI)F"'
Write-Host ''
