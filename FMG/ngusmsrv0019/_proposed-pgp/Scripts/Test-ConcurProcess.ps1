<#
.SYNOPSIS
    Offline test of Process_Files in the proposed ConcurGetFiles.ps1.

.DESCRIPTION
    Runs the REAL Process_Files function against a scratch directory - no gpg,
    no keyring, no network, nothing under E:\. Same lift-the-function-out-of-
    the-file approach as Test-ConcurDecrypt.ps1, for the same reason: it
    exercises the actual shipped code (and syntax-checks the whole file) rather
    than a copy that can drift.

.PARAMETER ScriptPath
    The ConcurGetFiles.ps1 under test.

.PARAMETER KeepArtifacts
    Leave the scratch folder behind so you can inspect what happened.

.NOTES
    Needs nothing but PowerShell itself.
    Exit code 0 = all passed, 1 = at least one failure.
#>
[CmdletBinding()]
param(
    [string]$ScriptPath = (Join-Path $PSScriptRoot '..\Applications\ConcurImport\_scripts\ConcurGetFiles.ps1'),
    [switch]$KeepArtifacts
)

$ErrorActionPreference = 'Continue'
$script:PassCount = 0
$script:FailCount = 0

function Assert {
    param([string]$Name, [bool]$Condition, [string]$Detail = '')
    if ($Condition) {
        $script:PassCount++
        Write-Host ("  [PASS] {0}" -f $Name) -ForegroundColor Green
    }
    else {
        $script:FailCount++
        Write-Host ("  [FAIL] {0}" -f $Name) -ForegroundColor Red
        if ($Detail) { Write-Host ("         {0}" -f $Detail) -ForegroundColor DarkRed }
    }
}
function Case { param([string]$T) Write-Host ''; Write-Host "-- $T" -ForegroundColor Cyan }

Write-Host ''
Write-Host '=========================================================' -ForegroundColor White
Write-Host ' Process_Files - offline test' -ForegroundColor White
Write-Host '=========================================================' -ForegroundColor White

# ------------------------------------------------------------ preconditions
if (-not (Test-Path $ScriptPath)) {
    Write-Host ("  Script under test not found: {0}" -f $ScriptPath) -ForegroundColor Red
    exit 1
}
$ScriptPath = (Resolve-Path $ScriptPath).Path
Write-Host ("  Under test : {0}" -f $ScriptPath) -ForegroundColor Gray

# ------------------------------------------- parse + lift out Process_Files
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
            $n.Name -eq 'Process_Files'
         }, $true) | Select-Object -First 1

Assert 'Process_Files is present in the script' ($null -ne $fnAst)
if (-not $fnAst) { exit 1 }

# Define the real function in this session, verbatim.
. ([scriptblock]::Create($fnAst.Extent.Text))
Assert 'Process_Files loaded into this session' ($null -ne (Get-Command Process_Files -ErrorAction SilentlyContinue))

# ------------------------------------------------------------ sandbox
$root = Join-Path ([System.IO.Path]::GetTempPath()) ("concurprocesstest_" + [guid]::NewGuid().ToString('N').Substring(0,8))
New-Item -Path $root -ItemType Directory -Force | Out-Null

# This is what Process_Files closes over in the real script.
$SavePath = $root
$teDir      = Join-Path $root 'ConcurTravelExpense'
$invoiceDir = Join-Path $root 'ConcurInvoice\import'

function Reset-Save {
    Get-ChildItem -Path $SavePath -File -ErrorAction SilentlyContinue | Remove-Item -Force
    foreach ($d in @($teDir, $invoiceDir)) {
        if (Test-Path $d) { Remove-Item $d -Recurse -Force }
    }
}

try {
    # ---------------------------------------------------------------- 1
    Case '1. Empty save path is a no-op, but still creates both destinations'
    Reset-Save
    $out = Process_Files *>&1
    Assert 'Reports nothing to do' (($out -join ' ') -match 'No files to route')
    Assert 'ConcurTravelExpense created' (Test-Path $teDir)
    Assert 'ConcurInvoice\import created' (Test-Path $invoiceDir)

    # ---------------------------------------------------------------- 2
    Case '2. T&E extract routes to ConcurTravelExpense'
    Reset-Save
    Set-Content -Path (Join-Path $SavePath 'CONCURSAEEXTRACT20260910.TXT') -Value 'DETAIL|x' -NoNewline
    $null = Process_Files *>&1
    Assert 'Moved into ConcurTravelExpense' (Test-Path (Join-Path $teDir 'CONCURSAEEXTRACT20260910.TXT'))
    Assert 'No longer in the save path root' (-not (Test-Path (Join-Path $SavePath 'CONCURSAEEXTRACT20260910.TXT')))

    # ---------------------------------------------------------------- 3
    Case '3. T&E match is case-insensitive, like the SSIS side'
    Reset-Save
    Set-Content -Path (Join-Path $SavePath 'concursaeextract20260910.txt') -Value 'DETAIL|x' -NoNewline
    $null = Process_Files *>&1
    Assert 'Lower-cased filename still routed' (Test-Path (Join-Path $teDir 'concursaeextract20260910.txt'))

    # ---------------------------------------------------------------- 4
    Case '4. Invoice header + detail route to ConcurInvoice\import'
    Reset-Save
    Set-Content -Path (Join-Path $SavePath 'invoice_header.txt') -Value 'H|1' -NoNewline
    Set-Content -Path (Join-Path $SavePath 'invoice_detail.txt') -Value 'D|1' -NoNewline
    $null = Process_Files *>&1
    Assert 'invoice_header.txt routed' (Test-Path (Join-Path $invoiceDir 'invoice_header.txt'))
    Assert 'invoice_detail.txt routed' (Test-Path (Join-Path $invoiceDir 'invoice_detail.txt'))

    # ---------------------------------------------------------------- 5
    Case '5. Mixed batch: each file goes to its own destination'
    Reset-Save
    Set-Content -Path (Join-Path $SavePath 'CONCURSAEEXTRACT20260910.TXT') -Value 'DETAIL|x' -NoNewline
    Set-Content -Path (Join-Path $SavePath 'invoice_header.txt') -Value 'H|1' -NoNewline
    Set-Content -Path (Join-Path $SavePath 'invoice_detail.txt') -Value 'D|1' -NoNewline
    $null = Process_Files *>&1
    Assert 'T&E file in its own folder' (Test-Path (Join-Path $teDir 'CONCURSAEEXTRACT20260910.TXT'))
    Assert 'Invoice files in their own folder' (
        (Test-Path (Join-Path $invoiceDir 'invoice_header.txt')) -and
        (Test-Path (Join-Path $invoiceDir 'invoice_detail.txt')))
    Assert 'Save path root left empty' ((Get-ChildItem -Path $SavePath -File).Count -eq 0)

    # ---------------------------------------------------------------- 6
    Case '6. Unrecognised file name is left in place, not guessed at'
    Reset-Save
    Set-Content -Path (Join-Path $SavePath 'some_other_file.txt') -Value 'x' -NoNewline
    $out6 = Process_Files *>&1
    Assert 'Still sitting in the save path' (Test-Path (Join-Path $SavePath 'some_other_file.txt'))
    Assert 'Not routed to either destination' (
        (-not (Test-Path (Join-Path $teDir 'some_other_file.txt'))) -and
        (-not (Test-Path (Join-Path $invoiceDir 'some_other_file.txt'))))
    Assert 'Explains why it was skipped' (($out6 -join ' ') -match 'does not match a known Concur extract pattern')

    # ---------------------------------------------------------------- 7
    Case '7. A still-encrypted leftover (failed decrypt) is not routed'
    Reset-Save
    Set-Content -Path (Join-Path $SavePath 'CONCURSAEEXTRACT20260910.TXT.pgp') -Value 'not really pgp' -NoNewline
    $null = Process_Files *>&1
    Assert 'Encrypted file left for the next Decrypt-Files retry' (Test-Path (Join-Path $SavePath 'CONCURSAEEXTRACT20260910.TXT.pgp'))
    Assert 'Nothing routed to ConcurTravelExpense' ((Get-ChildItem -Path $teDir -File).Count -eq 0)

    # ---------------------------------------------------------------- 8
    Case '8. Existing target is never clobbered'
    Reset-Save
    New-Item -Path $teDir -ItemType Directory -Force | Out-Null
    Set-Content -Path (Join-Path $teDir 'CONCURSAEEXTRACT20260910.TXT') -Value 'PRE-EXISTING' -NoNewline
    Set-Content -Path (Join-Path $SavePath 'CONCURSAEEXTRACT20260910.TXT') -Value 'NEW' -NoNewline
    $null = Process_Files *>&1
    Assert 'Pre-existing destination file untouched' (
        (Get-Content (Join-Path $teDir 'CONCURSAEEXTRACT20260910.TXT') -Raw) -eq 'PRE-EXISTING')
    Assert 'Source left in place rather than silently dropped' (Test-Path (Join-Path $SavePath 'CONCURSAEEXTRACT20260910.TXT'))

    # ---------------------------------------------------------------- 9
    Case '9. Re-running after a successful route is a clean no-op'
    Reset-Save
    Set-Content -Path (Join-Path $SavePath 'invoice_header.txt') -Value 'H|1' -NoNewline
    $null = Process_Files *>&1
    $out9 = Process_Files *>&1
    Assert 'Second run finds nothing left to route' (($out9 -join ' ') -match 'No files to route')
    Assert 'Routed file still where the first run put it' (Test-Path (Join-Path $invoiceDir 'invoice_header.txt'))
}
finally {
    if ($KeepArtifacts) {
        Write-Host ''
        Write-Host ("  Artifacts kept in {0}" -f $root) -ForegroundColor Yellow
    }
    else {
        Remove-Item $root -Recurse -Force -ErrorAction SilentlyContinue
    }
}

Write-Host ''
Write-Host '=========================================================' -ForegroundColor White
if ($script:FailCount -eq 0) {
    Write-Host (" All {0} checks passed." -f $script:PassCount) -ForegroundColor Green
}
else {
    Write-Host (" {0} passed, {1} FAILED." -f $script:PassCount, $script:FailCount) -ForegroundColor Red
}
Write-Host '=========================================================' -ForegroundColor White
Write-Host ''

exit ([int]($script:FailCount -gt 0))
