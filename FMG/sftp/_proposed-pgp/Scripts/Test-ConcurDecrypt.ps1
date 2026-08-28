<#
.SYNOPSIS
    Offline test of Decrypt-Files in the proposed ConcurGetFiles.ps1.

.DESCRIPTION
    Runs the REAL Decrypt-Files function against a throwaway keyring and a
    scratch directory. Never touches Concur, the network, your own keyring, or
    anything under E:\.

    How it gets at the function without running the script: ConcurGetFiles.ps1
    calls main() at the bottom, so dot-sourcing it would try to reach Concur.
    Instead this parses the file, pulls out just the Decrypt-Files function
    definition, and defines that here. So it exercises the actual shipped code,
    not a copy that can drift away from it. Parsing the whole file also
    syntax-checks it for free - a parse error fails the run before any test.

    Everything happens under a temp folder and a temp GNUPGHOME, both removed on
    exit. The key it generates exists only for the length of the run.

.PARAMETER ScriptPath
    The ConcurGetFiles.ps1 under test.

.PARAMETER KeepArtifacts
    Leave the scratch folder behind so you can inspect what happened.

.EXAMPLE
    .\Test-ConcurDecrypt.ps1

.EXAMPLE
    .\Test-ConcurDecrypt.ps1 -ScriptPath ..\Applications\ConcurImport\_scripts\ConcurGetFiles.ps1 -KeepArtifacts

.NOTES
    Needs gpg on PATH. Nothing else.
    Exit code 0 = all passed, 1 = at least one failure.
#>
[CmdletBinding()]
param(
    [string]$ScriptPath = (Join-Path $PSScriptRoot '..\Applications\ConcurImport\_scripts\ConcurGetFiles.ps1'),
    [switch]$KeepArtifacts
)

$ErrorActionPreference = 'Continue'
$script:Pass = 0
$script:Fail = 0

function Assert {
    param([string]$Name, [bool]$Condition, [string]$Detail = '')
    if ($Condition) {
        $script:Pass++
        Write-Host ("  [PASS] {0}" -f $Name) -ForegroundColor Green
    }
    else {
        $script:Fail++
        Write-Host ("  [FAIL] {0}" -f $Name) -ForegroundColor Red
        if ($Detail) { Write-Host ("         {0}" -f $Detail) -ForegroundColor DarkRed }
    }
}
function Case { param([string]$T) Write-Host ''; Write-Host "-- $T" -ForegroundColor Cyan }

Write-Host ''
Write-Host '=========================================================' -ForegroundColor White
Write-Host ' Decrypt-Files - offline test' -ForegroundColor White
Write-Host '=========================================================' -ForegroundColor White

# ------------------------------------------------------------ preconditions
if (-not (Get-Command gpg -ErrorAction SilentlyContinue)) {
    Write-Host '  gpg is not on PATH. Cannot run.' -ForegroundColor Red
    exit 1
}
if (-not (Test-Path $ScriptPath)) {
    Write-Host ("  Script under test not found: {0}" -f $ScriptPath) -ForegroundColor Red
    exit 1
}
$ScriptPath = (Resolve-Path $ScriptPath).Path
Write-Host ("  Under test : {0}" -f $ScriptPath) -ForegroundColor Gray
Write-Host ("  gpg        : {0}" -f (Get-Command gpg).Source) -ForegroundColor Gray

# ------------------------------------------- parse + lift out Decrypt-Files
Case 'Parse the script under test'

$parseErrors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile(
           $ScriptPath, [ref]$null, [ref]$parseErrors)

if ($parseErrors -and $parseErrors.Count -gt 0) {
    Assert 'ConcurGetFiles.ps1 parses without errors' $false (($parseErrors | ForEach-Object { $_.Message }) -join '; ')
    Write-Host ''
    Write-Host '  Cannot continue with a script that does not parse.' -ForegroundColor Red
    exit 1
}
Assert 'ConcurGetFiles.ps1 parses without errors' $true

$fnAst = $ast.FindAll({
            param($n)
            $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
            $n.Name -eq 'Decrypt-Files'
         }, $true) | Select-Object -First 1

Assert 'Decrypt-Files is present in the script' ($null -ne $fnAst)
if (-not $fnAst) { exit 1 }

# Define the real function in this session, verbatim.
. ([scriptblock]::Create($fnAst.Extent.Text))
Assert 'Decrypt-Files loaded into this session' ($null -ne (Get-Command Decrypt-Files -ErrorAction SilentlyContinue))

# ------------------------------------------------------------ sandbox
$root       = Join-Path ([System.IO.Path]::GetTempPath()) ("concurtest_" + [guid]::NewGuid().ToString('N').Substring(0,8))
$gpgHome    = Join-Path $root 'gnupg'
$inbox      = Join-Path $root 'inbox'
$archive    = Join-Path $root 'archive'
$savedHome  = $env:GNUPGHOME

New-Item -Path $gpgHome -ItemType Directory -Force | Out-Null
New-Item -Path $inbox   -ItemType Directory -Force | Out-Null
$env:GNUPGHOME = $gpgHome

# These two are what Decrypt-Files closes over in the real script.
$SavePath = $inbox
$config   = '{}' | ConvertFrom-Json
$config | Add-Member NoteProperty App_PgpPassPhraseLocation ''      -Force
$config | Add-Member NoteProperty App_EncryptedArchivePath  $archive -Force

function New-TestKey {
    param([string]$Email, [string]$Passphrase)
    $lines = @(
        'Key-Type: RSA', 'Key-Length: 3072', 'Key-Usage: sign'
        'Subkey-Type: RSA', 'Subkey-Length: 3072', 'Subkey-Usage: encrypt'
        'Name-Real: Concur Decrypt Test', "Name-Email: $Email", 'Expire-Date: 0'
    )
    if ($Passphrase) { $lines += "Passphrase: $Passphrase" } else { $lines += '%no-protection' }
    $lines += '%commit'

    $pf = Join-Path $root ("key-{0}.params" -f [guid]::NewGuid().ToString('N').Substring(0,6))
    $lines | Set-Content -Path $pf -Encoding ASCII
    & gpg --batch --quiet --gen-key $pf 2>$null | Out-Null
    Remove-Item $pf -Force -ErrorAction SilentlyContinue

    $colons = & gpg --batch --with-colons --list-secret-keys $Email 2>$null
    return (($colons | Where-Object { $_ -like 'fpr:*' } | Select-Object -First 1) -split ':')[9]
}

function New-EncryptedFile {
    param([string]$Recipient, [string]$Name, [string]$Content, [string]$Extension = '.pgp')
    $plain = Join-Path $root 'staging.tmp'
    Set-Content -Path $plain -Value $Content -Encoding ASCII -NoNewline
    $out = Join-Path $inbox ($Name + $Extension)
    & gpg --batch --yes --quiet --trust-model always --encrypt --recipient $Recipient --output $out $plain 2>$null | Out-Null
    Remove-Item $plain -Force -ErrorAction SilentlyContinue
    return $out
}

function Reset-Inbox {
    Get-ChildItem -Path $inbox -File -ErrorAction SilentlyContinue | Remove-Item -Force
    if (Test-Path $archive) { Remove-Item $archive -Recurse -Force }
}

function Get-ArchiveDir { Join-Path $archive (Get-Date -Format 'yyyy-MM-dd') }

try {
    Case 'Set up a throwaway keyring'
    $fpr = New-TestKey -Email 'nopass@example.invalid'
    Assert 'Throwaway key generated' (-not [string]::IsNullOrWhiteSpace($fpr)) "fpr='$fpr'"
    Write-Host ("         key $fpr (temporary, in $gpgHome)") -ForegroundColor DarkGray

    # ---------------------------------------------------------------- 1
    Case '1. Empty inbox is a no-op'
    Reset-Inbox
    $out = Decrypt-Files *>&1
    Assert 'Reports nothing to do' (($out -join ' ') -match 'No encrypted files')
    Assert 'Creates no archive folder' (-not (Test-Path (Get-ArchiveDir)))

    # ---------------------------------------------------------------- 2
    Case '2. Plaintext files are left completely alone (deploy-before-cutover safety)'
    Reset-Inbox
    $plainPath = Join-Path $inbox 'concur_invoice_20260824.txt'
    Set-Content -Path $plainPath -Value 'INVOICE|1|Barnes' -Encoding ASCII -NoNewline
    $before = (Get-Item $plainPath).LastWriteTimeUtc
    $out2 = Decrypt-Files *>&1
    Assert 'Plaintext not picked up as a candidate at all' (($out2 -join ' ') -match 'No encrypted files') `
        ("output was: " + ($out2 -join ' '))
    Assert 'Plaintext file still present' (Test-Path $plainPath)
    Assert 'Plaintext contents unchanged' ((Get-Content $plainPath -Raw) -eq 'INVOICE|1|Barnes')
    Assert 'Plaintext not even rewritten' ((Get-Item $plainPath).LastWriteTimeUtc -eq $before)

    # ---------------------------------------------------------------- 3
    Case '3. Happy path: invoice.txt.pgp -> invoice.txt, original archived'
    Reset-Inbox
    $body = "INVOICE|12345|Barnes`nINVOICE|12346|Barnes"
    $enc  = New-EncryptedFile -Recipient $fpr -Name 'concur_invoice_20260824.txt' -Content $body
    Assert 'Encrypted fixture created' (Test-Path $enc)

    $null = Decrypt-Files *>&1
    $dec = Join-Path $inbox 'concur_invoice_20260824.txt'
    Assert 'Decrypted file exists with the .pgp stripped' (Test-Path $dec)
    Assert 'Decrypted contents match the original' (((Get-Content $dec -Raw) -replace "`r`n","`n") -eq $body) `
        ("got: " + (Get-Content $dec -Raw))
    Assert 'Encrypted original removed from the inbox' (-not (Test-Path $enc))
    Assert 'Encrypted original archived under today''s date' `
        (Test-Path (Join-Path (Get-ArchiveDir) 'concur_invoice_20260824.txt.pgp'))

    # ---------------------------------------------------------------- 4
    Case '4. .gpg extension is handled too'
    Reset-Inbox
    $null = New-EncryptedFile -Recipient $fpr -Name 'concur_po.xml' -Content '<po/>' -Extension '.gpg'
    $null = Decrypt-Files *>&1
    Assert 'concur_po.xml.gpg decrypted to concur_po.xml' (Test-Path (Join-Path $inbox 'concur_po.xml'))

    # ---------------------------------------------------------------- 5
    Case '5. Mixed batch: encrypted decrypted, plaintext untouched'
    Reset-Inbox
    Set-Content -Path (Join-Path $inbox 'already_plain.txt') -Value 'PLAIN' -Encoding ASCII -NoNewline
    $null = New-EncryptedFile -Recipient $fpr -Name 'needs_decrypt.txt' -Content 'SECRET'
    $null = Decrypt-Files *>&1
    Assert 'Plaintext survived untouched' ((Get-Content (Join-Path $inbox 'already_plain.txt') -Raw) -eq 'PLAIN')
    Assert 'Encrypted one was decrypted' ((Test-Path (Join-Path $inbox 'needs_decrypt.txt')) -and
                                          ((Get-Content (Join-Path $inbox 'needs_decrypt.txt') -Raw).Trim() -eq 'SECRET'))

    # ---------------------------------------------------------------- 6
    Case '6. Garbage ciphertext fails safely and is kept for retry'
    Reset-Inbox
    $junk = Join-Path $inbox 'corrupt.txt.pgp'
    $bytes = New-Object byte[] 400
    (New-Object Random).NextBytes($bytes)
    [System.IO.File]::WriteAllBytes($junk, $bytes)

    $null = Decrypt-Files *>&1
    Assert 'No plaintext produced' (-not (Test-Path (Join-Path $inbox 'corrupt.txt')))
    Assert 'Bad .pgp left in place for the next run' (Test-Path $junk)
    Assert 'Bad .pgp NOT archived' (-not (Test-Path (Join-Path (Get-ArchiveDir) 'corrupt.txt.pgp')))

    # ---------------------------------------------------------------- 7
    Case '7. Truncated ciphertext leaves no partial plaintext behind'
    # The realistic half-transferred-file case. gpg can exit non-zero AFTER
    # having already written part of the plaintext - a partial invoice reaching
    # the SSIS packages would be a silent data bug, so the partial must go.
    Reset-Inbox
    # Must be INCOMPRESSIBLE. Repetitive text compresses to a ciphertext small
    # enough that truncation yields no output at all, which silently skips the
    # very behaviour this case is here to check.
    $rnd = New-Object byte[] 150000
    (New-Object Random).NextBytes($rnd)
    $bigBody = [Convert]::ToBase64String($rnd)
    $bigEnc  = New-EncryptedFile -Recipient $fpr -Name 'big.txt' -Content $bigBody
    $all     = [System.IO.File]::ReadAllBytes($bigEnc)
    $half    = New-Object byte[] ([int]($all.Length * 0.6))
    [Array]::Copy($all, $half, $half.Length)
    [System.IO.File]::WriteAllBytes($bigEnc, $half)

    $null = Decrypt-Files *>&1
    $partial = Join-Path $inbox 'big.txt'
    $partialDetail = ''
    if (Test-Path $partial) { $partialDetail = "partial is $((Get-Item $partial).Length) bytes" }
    Assert 'No partial plaintext left on disk' (-not (Test-Path $partial)) $partialDetail
    Assert 'Truncated .pgp kept for retry' (Test-Path $bigEnc)

    # ---------------------------------------------------------------- 8
    Case '8. Existing target is never clobbered'
    # gpg is called with --yes, which WOULD overwrite. The Test-Path guard in
    # Decrypt-Files is what stops it, so this asserts the guard, not gpg.
    Reset-Inbox
    $enc2 = New-EncryptedFile -Recipient $fpr -Name 'collide.txt' -Content 'NEW CONTENT'
    Set-Content -Path (Join-Path $inbox 'collide.txt') -Value 'PRE-EXISTING' -Encoding ASCII -NoNewline

    $null = Decrypt-Files *>&1
    Assert 'Pre-existing file untouched' ((Get-Content (Join-Path $inbox 'collide.txt') -Raw) -eq 'PRE-EXISTING')
    Assert 'Colliding .pgp left in place, not archived' (Test-Path $enc2)

    # ---------------------------------------------------------------- 9
    Case '9. Passphrase-protected key, passphrase read from a file'
    Reset-Inbox
    $pass    = 'S0me-Real-Passphrase!'
    $fprPass = New-TestKey -Email 'withpass@example.invalid' -Passphrase $pass
    Assert 'Passphrase-protected key generated' (-not [string]::IsNullOrWhiteSpace($fprPass))

    $passFile = Join-Path $root 'keyphrase.txt'
    Set-Content -Path $passFile -Value $pass -Encoding ASCII -NoNewline
    $config.App_PgpPassPhraseLocation = $passFile

    $null = New-EncryptedFile -Recipient $fprPass -Name 'protected.txt' -Content 'PROTECTED PAYLOAD'
    # gpg-agent caches passphrases; a stale cache would make this pass for the
    # wrong reason. Start cold.
    & gpgconf --homedir $gpgHome --kill gpg-agent 2>$null | Out-Null

    $null = Decrypt-Files *>&1
    $decP = Join-Path $inbox 'protected.txt'
    Assert 'Decrypted using the passphrase file' `
        ((Test-Path $decP) -and ((Get-Content $decP -Raw).Trim() -eq 'PROTECTED PAYLOAD'))

    Case '10. Wrong passphrase fails, and fails loudly'
    Reset-Inbox
    Set-Content -Path $passFile -Value 'definitely-not-the-passphrase' -Encoding ASCII -NoNewline
    $null = New-EncryptedFile -Recipient $fprPass -Name 'protected2.txt' -Content 'PROTECTED PAYLOAD'
    & gpgconf --homedir $gpgHome --kill gpg-agent 2>$null | Out-Null

    $null = Decrypt-Files *>&1
    Assert 'No plaintext written with a wrong passphrase' (-not (Test-Path (Join-Path $inbox 'protected2.txt')))
    Assert 'Encrypted file kept for retry' (Test-Path (Join-Path $inbox 'protected2.txt.pgp'))

    $config.App_PgpPassPhraseLocation = ''
}
finally {
    & gpgconf --homedir $gpgHome --kill gpg-agent 2>$null | Out-Null
    $env:GNUPGHOME = $savedHome

    if ($KeepArtifacts) {
        Write-Host ''
        Write-Host ("  Artifacts kept in {0}" -f $root) -ForegroundColor Yellow
    }
    else {
        Start-Sleep -Milliseconds 300   # let gpg-agent release its socket
        Remove-Item $root -Recurse -Force -ErrorAction SilentlyContinue
    }
}

Write-Host ''
Write-Host '=========================================================' -ForegroundColor White
if ($script:Fail -eq 0) {
    Write-Host (" All {0} checks passed." -f $script:Pass) -ForegroundColor Green
}
else {
    Write-Host (" {0} passed, {1} FAILED." -f $script:Pass, $script:Fail) -ForegroundColor Red
}
Write-Host '=========================================================' -ForegroundColor White
Write-Host ''

exit ([int]($script:Fail -gt 0))
