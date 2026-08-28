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
