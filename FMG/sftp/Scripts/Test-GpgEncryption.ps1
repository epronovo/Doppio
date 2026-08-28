<#
.SYNOPSIS
    Non-destructive GnuPG diagnostic for the HSBC SFTP interface.

.DESCRIPTION
    Answers the questions that actually break this interface, without touching any
    real data:

      1. Which gpg.exe runs, and are there competing copies earlier on PATH?
      2. Which keyring is in scope? (Keyrings are PER-USER. The scheduled tasks run
         as NG\M0042, so its keyring - not yours - is the one that matters.)
      3. Is the HSBC recipient key present, unexpired, and trusted enough to encrypt
         WITHOUT a prompt? HSBC-US_send.ps1 calls gpg with no --batch, so an
         untrusted key makes it hang rather than fail.
      4. Does an encrypt - and optionally a decrypt - actually succeed?

    Everything runs in a temp folder that is removed on exit. No keyring is modified,
    no production directory is read or written, nothing is uploaded.

.PARAMETER Recipient
    PGP recipient to test. Defaults to the address hard-coded in HSBC-US_send.ps1.

.PARAMETER GpgPath
    Pin a specific gpg.exe instead of letting PATH decide. Use this to prove which
    of two installs behaves differently.

.PARAMETER GpgHome
    Inspect a different keyring via --homedir, instead of the current user's.
    Use this when you cannot log on as the service account: take a COPY of its
    gnupg folder and point this at the copy. Never point it at the live folder -
    gpg writes to trustdb.gpg on read, and files created under your identity in
    the service account's profile can lock it out.

.PARAMETER RoundTrip
    Also test the inbound path: encrypt to our OWN secret key and decrypt it back.
    (We cannot decrypt what we send to HSBC - we do not hold their private key.)

.PARAMETER PassPhrase
    Passphrase for our secret key, required by -RoundTrip if the key has one.
    Passed to gpg over stdin, never on the command line.

.EXAMPLE
    .\Test-GpgEncryption.ps1

.EXAMPLE
    # Run as the service account and diff the output against your own run:
    runas /user:NG\M0042 C:\Temp\run-gpgtest.cmd

.EXAMPLE
    # No password for the service account? Inspect a COPY of its keyring instead.
    # (Admin rights on the box are enough; no credentials needed.)
    robocopy "C:\Users\M0042\AppData\Roaming\gnupg" "C:\Temp\m0042-keyring" /E /COPY:DAT /R:0
    .\Test-GpgEncryption.ps1 -GpgHome C:\Temp\m0042-keyring

.EXAMPLE
    .\Test-GpgEncryption.ps1 -RoundTrip -PassPhrase 'secret'

.NOTES
    Exit code 0 = all checks passed, 1 = at least one failure.
#>

[CmdletBinding()]
param(
    [string]$Recipient = 'cmbitconnectdigitalmappingvalidation@hsbc.co.in',
    [string]$GpgPath,
    [string]$GpgHome,
    [switch]$RoundTrip,
    [string]$PassPhrase
)

$ErrorActionPreference = 'Continue'
$script:Failures     = 0
$script:KeyringEmpty = $false

function Write-Result {
    param(
        [ValidateSet('PASS', 'FAIL', 'WARN', 'INFO')][string]$Status,
        [string]$Message
    )
    $color = @{ PASS = 'Green'; FAIL = 'Red'; WARN = 'Yellow'; INFO = 'Gray' }[$Status]
    Write-Host ("  [{0}] {1}" -f $Status.PadRight(4), $Message) -ForegroundColor $color
    if ($Status -eq 'FAIL') { $script:Failures++ }
}

function Write-Section {
    param([string]$Title)
    Write-Host ''
    Write-Host "== $Title" -ForegroundColor Cyan
}

function Convert-GpgDate {
    # gpg --with-colons emits epoch seconds, or yyyyMMddTHHmmss on some builds.
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $null }
    if ($Value -match '^\d+$') {
        return [DateTimeOffset]::FromUnixTimeSeconds([int64]$Value).LocalDateTime
    }
    if ($Value -match '^(\d{8})T') {
        return [datetime]::ParseExact($Matches[1], 'yyyyMMdd', $null)
    }
    return $null
}

function Get-CanonicalPath {
    # PATH entries can contain '..' traversals - Gpg4win writes one that hops out
    # of its own folder into the sibling GnuPG folder - so two different strings
    # can be the same binary. Resolve before comparing.
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return $Path }
    try { return (Resolve-Path -LiteralPath $Path -ErrorAction Stop).ProviderPath } catch { }
    try { return [System.IO.Path]::GetFullPath($Path) } catch { }
    return $Path
}

$ValidityText = @{
    'o' = 'unknown'; 'i' = 'INVALID'; 'd' = 'disabled'; 'r' = 'REVOKED'
    'e' = 'EXPIRED'; '-' = 'unknown'; 'q' = 'undefined'; 'n' = 'never trusted'
    'm' = 'marginal'; 'f' = 'full'; 'u' = 'ultimate'
}

Write-Host ''
Write-Host '=========================================================' -ForegroundColor White
Write-Host ' GnuPG encryption test - HSBC interface' -ForegroundColor White
Write-Host '=========================================================' -ForegroundColor White

# ---------------------------------------------------------------- 1. Context
Write-Section 'Context'
Write-Result INFO ("Machine     : {0}" -f $env:COMPUTERNAME)
Write-Result INFO ("Running as  : {0}\{1}" -f $env:USERDOMAIN, $env:USERNAME)
Write-Result INFO ("Profile     : {0}" -f $env:APPDATA)
Write-Result INFO ("PowerShell  : {0}" -f $PSVersionTable.PSVersion)

# ------------------------------------------------------- 2. Locate gpg.exe
Write-Section 'gpg.exe resolution'

$onPath = @(Get-Command gpg -All -ErrorAction SilentlyContinue |
            Where-Object { $_.CommandType -eq 'Application' } |
            Select-Object -ExpandProperty Source)

$knownDirs = @(
    'C:\Program Files\GnuPG\bin\gpg.exe'
    'C:\Program Files (x86)\GnuPG\bin\gpg.exe'
    'C:\Program Files\Gpg4win\bin\gpg.exe'
    'C:\Program Files (x86)\Gpg4win\bin\gpg.exe'
) | Where-Object { Test-Path $_ }

$onPathCanon = @($onPath | ForEach-Object { (Get-CanonicalPath $_).ToLowerInvariant() })
$distinctBinaries = @($onPathCanon | Select-Object -Unique)

if ($onPath.Count -eq 0) {
    Write-Result FAIL 'No gpg.exe found on PATH. The HSBC scripts call bare "gpg" and will fail.'
} else {
    Write-Result INFO ("gpg.exe on PATH ({0} entr(y/ies), first one wins):" -f $onPath.Count)
    for ($i = 0; $i -lt $onPath.Count; $i++) {
        Write-Result INFO ("    {0}. {1}" -f ($i + 1), $onPath[$i])
        $canon = Get-CanonicalPath $onPath[$i]
        if ($canon -ne $onPath[$i]) {
            Write-Result INFO ("       resolves to {0}" -f $canon)
        }
    }
    if ($distinctBinaries.Count -gt 1) {
        Write-Result WARN ("{0} DIFFERENT gpg.exe binaries on PATH. Which one runs depends on PATH" -f $distinctBinaries.Count)
        Write-Result WARN 'order, which can differ between your session and the scheduled task.'
    } else {
        Write-Result PASS 'One gpg.exe binary on PATH (all entries resolve to the same file).'
    }
}

foreach ($d in $knownDirs) {
    if ($onPathCanon -notcontains (Get-CanonicalPath $d).ToLowerInvariant()) {
        Write-Result INFO ("Installed but NOT on PATH: {0}" -f $d)
    }
}

if ($GpgPath) {
    if (-not (Test-Path $GpgPath)) {
        Write-Result FAIL ("-GpgPath does not exist: {0}" -f $GpgPath)
        exit 1
    }
    $gpg = $GpgPath
    Write-Result INFO ("Pinned via -GpgPath: {0}" -f $gpg)
} elseif ($onPath.Count -gt 0) {
    $gpg = $onPath[0]
} else {
    exit 1
}

# Every gpg call below is splatted with this, so --homedir applies uniformly.
$gpgArgs = @()
if ($GpgHome) {
    if (-not (Test-Path $GpgHome)) {
        Write-Result FAIL ("-GpgHome does not exist: {0}" -f $GpgHome)
        exit 1
    }
    $gpgArgs = @('--homedir', $GpgHome)
    Write-Result WARN ("Inspecting keyring at {0} (not this user's)." -f $GpgHome)
    Write-Result WARN 'Make sure that is a COPY - gpg writes to trustdb.gpg even on read.'
}

# ------------------------------------------------------ 3. Version / homedir
Write-Section 'GnuPG environment'

$versionOut = & $gpg @gpgArgs --version 2>&1
if ($LASTEXITCODE -ne 0) {
    Write-Result FAIL ("'{0} --version' failed with exit code {1}." -f $gpg, $LASTEXITCODE)
    exit 1
}
Write-Result PASS ("Using : {0}" -f $gpg)
Write-Result INFO ("Version: {0}" -f (($versionOut | Select-Object -First 1) -replace '^gpg \(GnuPG\) ', ''))

$homeDir = ($versionOut | Where-Object { $_ -match '^Home:\s*(.+)$' } |
            ForEach-Object { $Matches[1].Trim() } | Select-Object -First 1)
if (-not $homeDir) {
    $cfg = & $gpg @gpgArgs --with-colons --list-config homedir 2>$null
    if ($cfg -match '^cfg:homedir:(.+)$') { $homeDir = $Matches[1] }
}
if ($homeDir) {
    Write-Result INFO ("Keyring: {0}" -f $homeDir)
    Write-Result INFO 'Keyrings are per-user. The scheduled tasks use the SERVICE ACCOUNT keyring.'
} else {
    Write-Result WARN 'Could not determine the GnuPG home directory.'
}

# -------------------------------------------------- 4. Inspect recipient key
Write-Section ("Recipient key: {0}" -f $Recipient)

$colons = & $gpg @gpgArgs --batch --with-colons --list-keys $Recipient 2>&1
if ($LASTEXITCODE -ne 0) {
    Write-Result FAIL 'Recipient key is NOT in this keyring. Encryption cannot succeed.'

    # Distinguish "this key is missing" from "this keyring has nothing in it at all",
    # which usually just means you are not the account that runs the scheduled tasks.
    $allKeys  = & $gpg @gpgArgs --batch --with-colons --list-keys 2>$null
    $pubCount = @($allKeys | Where-Object { $_ -like 'pub:*' }).Count
    if ($pubCount -eq 0) {
        $script:KeyringEmpty = $true
        Write-Result INFO 'This keyring contains NO public keys at all.'
    } else {
        Write-Result INFO ("Keyring holds {0} other public key(s), just not this one." -f $pubCount)
    }

    Write-Result INFO 'Import it as the account that runs the task, e.g.:'
    Write-Result INFO '    gpg --import "E:\KeysEncryptions\HSBC\theirs\production\hsbc_bis_prd_pgp_pub_key202408"'
} else {
    $pub = $colons | Where-Object { $_ -like 'pub:*' } | Select-Object -First 1
    if ($pub) {
        $f        = $pub -split ':'
        $validity = $ValidityText[$f[1]]
        if (-not $validity) { $validity = $f[1] }
        $expires  = Convert-GpgDate $f[6]

        Write-Result PASS ("Found key {0}, created {1:yyyy-MM-dd}" -f $f[4], (Convert-GpgDate $f[5]))
        Write-Result INFO ("Owner trust / validity: {0}" -f $validity)

        if ($expires) {
            $days = [int]($expires - (Get-Date)).TotalDays
            if ($days -lt 0)      { Write-Result FAIL ("Primary key EXPIRED on {0:yyyy-MM-dd}." -f $expires) }
            elseif ($days -lt 45) { Write-Result WARN ("Primary key expires {0:yyyy-MM-dd} ({1} days)." -f $expires, $days) }
            else                  { Write-Result PASS ("Primary key valid until {0:yyyy-MM-dd} ({1} days)." -f $expires, $days) }
        } else {
            Write-Result PASS 'Primary key has no expiry.'
        }
    }

    # The encryption SUBKEY is what actually gets used - check it separately.
    $encSubs = $colons | Where-Object { $_ -like 'sub:*' -and ($_ -split ':')[11] -match 'e' }
    if ($encSubs) {
        foreach ($s in $encSubs) {
            $sf   = $s -split ':'
            $sExp = Convert-GpgDate $sf[6]
            if ($sExp) {
                $sDays = [int]($sExp - (Get-Date)).TotalDays
                if ($sDays -lt 0)      { Write-Result FAIL ("Encryption subkey {0} EXPIRED on {1:yyyy-MM-dd}." -f $sf[4], $sExp) }
                elseif ($sDays -lt 45) { Write-Result WARN ("Encryption subkey {0} expires {1:yyyy-MM-dd} ({2} days)." -f $sf[4], $sExp, $sDays) }
                else                   { Write-Result PASS ("Encryption subkey {0} valid until {1:yyyy-MM-dd} ({2} days)." -f $sf[4], $sExp, $sDays) }
            } else {
                Write-Result PASS ("Encryption subkey {0} has no expiry." -f $sf[4])
            }
        }
    } else {
        Write-Result WARN 'No encryption-capable subkey listed; the primary key must carry the "e" capability.'
    }
}

# --------------------------------------------------------- 5. Encrypt test
Write-Section 'Encrypt test (production flags)'

$work = Join-Path $env:TEMP ("gpgtest_" + [guid]::NewGuid().ToString('N').Substring(0, 8))
try {
    New-Item -Path $work -ItemType Directory -Force | Out-Null
    $plain = Join-Path $work 'sample.txt'
    "GPG test payload {0:yyyy-MM-dd HH:mm:ss} from {1}" -f (Get-Date), $env:USERNAME |
        Set-Content -Path $plain -Encoding ASCII

    $cipher = Join-Path $work 'sample.pgp'

    # --batch makes an untrusted key FAIL instead of PROMPT. Production omits it,
    # so a failure here is exactly the case where HSBC-US_send.ps1 would hang.
    & $gpg @gpgArgs --batch --yes --encrypt --recipient $Recipient --output $cipher $plain 2>&1 |
        ForEach-Object { Write-Result INFO ("gpg: {0}" -f $_) }

    if ($LASTEXITCODE -eq 0 -and (Test-Path $cipher)) {
        $size = (Get-Item $cipher).Length
        Write-Result PASS ("Encrypted successfully ({0} bytes)." -f $size)
        Write-Result PASS 'Key is trusted - the production command will not prompt.'
    } else {
        Write-Result FAIL 'Encryption failed with --batch.'

        # Distinguish "key untrusted" from "key missing or broken".
        & $gpg @gpgArgs --batch --yes --trust-model always --encrypt --recipient $Recipient `
               --output $cipher $plain 2>&1 | Out-Null

        if ($LASTEXITCODE -eq 0 -and (Test-Path $cipher)) {
            Write-Result WARN 'Succeeds with --trust-model always, so this is a TRUST problem.'
            Write-Result WARN 'HSBC-US_send.ps1 has no --batch, so it will PROMPT and hang the task.'
            Write-Result INFO ("Fix: gpg --lsign-key {0}   (or add --trust-model always to the script)" -f $Recipient)
        } else {
            Write-Result FAIL 'Also fails with --trust-model always - the key or gpg install is the problem.'
        }
    }

    # ----------------------------------------------------- 6. Round trip
    if ($RoundTrip) {
        Write-Section 'Decrypt test (our own key)'

        $secColons = & $gpg @gpgArgs --batch --with-colons --list-secret-keys 2>&1
        $sec = $secColons | Where-Object { $_ -like 'sec:*' } | Select-Object -First 1

        if (-not $sec) {
            Write-Result FAIL 'No secret key in this keyring - cannot test the decrypt path.'
        } else {
            $selfKey = ($sec -split ':')[4]
            Write-Result INFO ("Using our secret key {0}" -f $selfKey)

            $selfCipher = Join-Path $work 'self.pgp'
            $selfPlain  = Join-Path $work 'self.out'

            & $gpg @gpgArgs --batch --yes --trust-model always --encrypt --recipient $selfKey `
                   --output $selfCipher $plain 2>&1 | Out-Null

            if ($LASTEXITCODE -ne 0 -or -not (Test-Path $selfCipher)) {
                Write-Result FAIL 'Could not encrypt to our own key.'
            } else {
                # Passphrase via stdin, NOT on the command line - unlike production,
                # where it is briefly visible in the process table.
                if ($PassPhrase) {
                    $PassPhrase | & $gpg @gpgArgs --quiet --batch --yes --pinentry-mode loopback `
                        --passphrase-fd 0 -o $selfPlain -d $selfCipher 2>&1 |
                        ForEach-Object { Write-Result INFO ("gpg: {0}" -f $_) }
                } else {
                    & $gpg @gpgArgs --quiet --batch --yes --pinentry-mode loopback `
                        -o $selfPlain -d $selfCipher 2>&1 |
                        ForEach-Object { Write-Result INFO ("gpg: {0}" -f $_) }
                }

                if ($LASTEXITCODE -eq 0 -and (Test-Path $selfPlain)) {
                    $before = Get-Content $plain -Raw
                    $after  = Get-Content $selfPlain -Raw
                    if ($before -eq $after) {
                        Write-Result PASS 'Round trip succeeded - decrypt path is healthy.'
                    } else {
                        Write-Result FAIL 'Decrypted content does not match the original.'
                    }
                } else {
                    Write-Result FAIL 'Decryption failed. Check the passphrase and the secret key.'
                }
            }
        }
    }
}
finally {
    if (Test-Path $work) { Remove-Item $work -Recurse -Force -ErrorAction SilentlyContinue }
}

# ------------------------------------------------------------- 7. Summary
Write-Section 'Summary'
if ($script:Failures -eq 0) {
    Write-Result PASS 'All checks passed.'
} else {
    Write-Result FAIL ("{0} check(s) failed." -f $script:Failures)
}

if ($script:KeyringEmpty) {
    Write-Host ''
    Write-Host ("  NOTE: {0}\{1} has an EMPTY keyring." -f $env:USERDOMAIN, $env:USERNAME) -ForegroundColor Yellow
    Write-Host '  If this is not the account the scheduled tasks run as, the failures above' -ForegroundColor Yellow
    Write-Host '  are EXPECTED and say nothing about production. Re-run as that account.' -ForegroundColor Yellow
}

Write-Host ''
Write-Host 'Run this again as NG\M0042 and compare - the keyring is per-user.' -ForegroundColor DarkGray
Write-Host ''

exit ([int]($script:Failures -gt 0))
