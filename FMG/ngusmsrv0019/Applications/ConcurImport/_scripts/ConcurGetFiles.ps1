param (
    [Parameter(Mandatory=$true)]
    [ValidateSet("stage","production")]
    [string]$Environment
)

$basePath = Split-Path -Path $PSScriptRoot -Parent
$configPath = Join-Path $basePath "$Environment\configuration\concurimport.json"
write-host $configPath
if (-not (Test-Path $configPath)) {
    Write-Error "Config file not found: $configPath"
    exit 1
}

$config = Get-Content $configPath | ConvertFrom-Json
$SavePath = $config.App_IncomingSavePath

Start-Transcript -Path "$PSScriptRoot\concurget.log" -Append

function Get_Files {

. "E:\Scripts\WinScpFTP\Concur\ConcurConnect.ps1"			
    GetFiles -SavePath $SavePath -Environment $Environment

}

function Decrypt-Files {

    $encrypted = @(Get-ChildItem -Path $SavePath -File |
                   Where-Object { $_.Extension -in ".pgp", ".gpg" })

    if ($encrypted.Count -eq 0) {
        Write-Host "No encrypted files to decrypt"
        return
    }

    $passPhrase = $null
    $passPhraseLocation = $config.App_PgpPassPhraseLocation
    if ($passPhraseLocation) {
        if (-not (Test-Path $passPhraseLocation)) {
            Write-Error "PGP passphrase file not found: $passPhraseLocation"
            return
        }
        $passPhrase = (Get-Content $passPhraseLocation -Raw).Trim()
    }

    $archiveRoot = $config.App_EncryptedArchivePath
    if (-not $archiveRoot) {
        $archiveRoot = Join-Path $SavePath "encrypted"
    }
    $archivePath = Join-Path $archiveRoot (Get-Date -Format "yyyy-MM-dd")

    if (-not (Test-Path $archivePath)) {
        New-Item -Path $archivePath -ItemType Directory -Force | Out-Null
        Write-Host "Created Directory $archivePath"
    }

    foreach ($file in $encrypted) {

        Write-Host "Decrypting $($file.FullName)"

        # Strip only the .pgp/.gpg wrapper - concur_invoice.txt.pgp -> concur_invoice.txt
        $outputFile = Join-Path $file.DirectoryName $file.BaseName

        if (Test-Path $outputFile) {
            Write-Error "Decrypt target already exists, skipping: $outputFile"
            continue
        }

        if ($passPhrase) {
            $passPhrase | gpg --quiet --batch --yes --pinentry-mode loopback `
                              --passphrase-fd 0 `
                              -o $outputFile `
                              -d $file.FullName
        }
        else {
            gpg --quiet --batch --yes --pinentry-mode loopback `
                -o $outputFile `
                -d $file.FullName
        }

        if ($LASTEXITCODE -ne 0 -or -not (Test-Path $outputFile)) {
            Write-Error "Decryption failed for $($file.Name). Leaving it in place to retry next run."
            if (Test-Path $outputFile) { Remove-Item $outputFile -Force -ErrorAction SilentlyContinue }
            continue
        }

        if ((Get-Item $outputFile).Length -eq 0) {
            Write-Error "Decryption produced an empty file for $($file.Name). Leaving it in place to retry next run."
            Remove-Item $outputFile -Force -ErrorAction SilentlyContinue
            continue
        }

        Write-Host "Decryption successful: $outputFile"

        # Archive the encrypted original only after a good decrypt.
        Move-Item $file.FullName -Destination $archivePath -Force -ErrorAction Stop
    }

}


function main {
    try{
        Get_Files
        #Decrypt-Files
		#Process_Files
    }
    catch{
        Write-Error $_
        Stop-Transcript
        exit 1
    }
}


main

Stop-Transcript -ErrorAction SilentlyContinue
