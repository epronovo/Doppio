<#
.SYNOPSIS
    PGP-encrypt the Concur Purchase Order extract and upload it to Concur.

.DESCRIPTION
    Outbound half of the Concur interface. Mirrors HSBC-US_send.ps1 in shape:

        Encrypt-Files   gpg --encrypt to Concur's public key -> <name>_<ts>.pgp
                        original moved to processed\<yyyy-MM-dd>\ only on success
        Send-Files      upload every *.pgp, archive on a $true result

    Differences from HSBC-US_send.ps1, all deliberate:

      * --batch is passed to gpg. Without it, an untrusted recipient key sends gpg
        to the tty for a confirmation - in a scheduled task with no interactive
        desktop that is at best an opaque failure and at worst a stuck task. With
        --batch it fails immediately and visibly:
            gpg: encryption failed: Unusable public key
        and writes no output file, which the Test-Path check below catches.

        TRUST SETUP, one time, as the account that runs the task (NG\M0042):
            gpg --import "E:\KeysEncryptions\Concur\Dev\concursolutions.asc"
            echo 2584FD81F647ECBFAA1EE1DDBAA83C168C51C89E:6: | gpg --import-ownertrust
        Verified: importing alone is NOT enough - encrypt still fails on trust.
        Note `gpg --lsign-key` is the usual advice and it does NOT work here: it
        needs a secret key in the keyring ("gpg: no default secret key"), and
        M0042 has none for Concur. Set ownertrust instead. Do not "fix" this by
        adding --trust-model always - that works, but it switches off trust
        checking for every recipient this account will ever encrypt to.
      * the recipient comes from config, per environment. HSBC hard-codes one
        address for both, so HSBC stage traffic is encrypted to the production key.
      * a failed upload is Write-Error, not Write-Host, so it actually shows up as
        a failure in the transcript and to Task Scheduler.

.PARAMETER Environment
    stage | production. There is no production Concur connection config yet -
    Scripts\WinScpFTP\Concur\config\ has a stage folder only.

.NOTES
    STATUS: not yet run against Concur. Two things must be confirmed with Concur
    before this goes anywhere near production:
      1. does their MFT endpoint accept PGP-encrypted inbound files at all
      2. what remote path do inbound files go to (App_SendRemotePath)
#>
param (
    [Parameter(Mandatory=$true)]
    [ValidateSet("stage","production")]
    [string]$Environment
)

$basePath = Split-Path -Path $PSScriptRoot -Parent
$configPath = Join-Path $basePath "$Environment\configuration\concurextracts.json"
write-host $configPath
if (-not (Test-Path $configPath)) {
    Write-Error "Config file not found: $configPath"
    exit 1
}

$config = Get-Content $configPath | ConvertFrom-Json

$dataDirectory = $config.App_DataDirectory
$processedPath = $config.App_ProcessedPath
$recipient     = $config.App_PgpRecipient
$remotePath    = $config.App_SendRemotePath

Start-Transcript -Path "$PSScriptRoot\concursend.log" -Append


function Encrypt-Files {

    if (-not $recipient) {
        Write-Error "No PGP recipient configured. Set App_PgpRecipient in $configPath."
        return
    }

    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $files = @(Get-ChildItem -Path $dataDirectory -File |
               Where-Object { $_.Extension -in ".dat", ".xml", ".txt", ".csv" })

    if ($files.Count -eq 0) {
        Write-Host "No files to encrypt"
        return
    }

    foreach ($file in $files) {

        Write-Host "Encrypting $($file.FullName)"

        $_pgpFile = Join-Path $dataDirectory ($file.Name + "_$timestamp.pgp")

        # --batch: fail on an untrusted key instead of prompting and hanging.
        gpg --batch --yes --encrypt --recipient $recipient `
            --output $_pgpFile $file.FullName

        if ($LASTEXITCODE -ne 0 -or -not (Test-Path $_pgpFile)) {
            Write-Error "Encryption failed for $($file.FullName). Not sending it."
            if (Test-Path $_pgpFile) { Remove-Item $_pgpFile -Force -ErrorAction SilentlyContinue }
            continue
        }

        Write-Host "Encryption successful: $_pgpFile"

        # Original goes to processed only after a good encrypt.
        $folderName = Get-Date -Format "yyyy-MM-dd"
        $_folderPath = Join-Path -Path $processedPath -ChildPath $folderName

        if (-not (Test-Path $_folderPath)) {
            New-Item -Path $_folderPath -ItemType Directory -Force | Out-Null
        }

        Move-Item $file.FullName -Destination $_folderPath -Force -ErrorAction Stop
    }

}


function Send-Files {

    $pgpFiles = @(Get-ChildItem -Path $dataDirectory -Filter "*.pgp" -File)

    if ($pgpFiles.Count -eq 0) {
        Write-Host "No encrypted files to send"
        return
    }

    . "E:\Scripts\WinScpFTP\Concur\ConcurConnect.ps1"

    foreach ($pgpFile in $pgpFiles) {

        $_pgpFileName = $pgpFile.Name
        $_remotePath = $remotePath.TrimEnd('/') + "/" + $_pgpFileName

        Write-Host "Sending $_pgpFileName to Concur"

        try {
            $result = SendFile -LocalFile $pgpFile.FullName -RemotePath $_remotePath -Environment $Environment

            if ($result) {

                $folderName = Get-Date -Format "yyyy-MM-dd"
                $_folderPath = Join-Path -Path $processedPath -ChildPath $folderName

                if (-not (Test-Path $_folderPath)) {
                    New-Item -Path $_folderPath -ItemType Directory -Force | Out-Null
                }

                Move-Item $pgpFile.FullName -Destination $_folderPath -Force -ErrorAction Stop

                Write-Host "Sent and archived $_pgpFileName"
            }
            else {
                # Left in the data directory on purpose - the next run retries it.
                Write-Error "Unable to send file $_pgpFileName"
            }
        }
        catch {
            Write-Error "An error occurred sending $_pgpFileName : $_"
        }

    }

}


function main {
    try {
        Encrypt-Files
        Send-Files
    }
    catch {
        Write-Error $_
        Stop-Transcript
        exit 1
    }
}


main

Stop-Transcript -ErrorAction SilentlyContinue
