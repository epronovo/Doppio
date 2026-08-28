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


<#
    Decrypt-Files
    -------------
    Decrypts every PGP file Concur delivered, in place, so that whatever consumes
    $SavePath afterwards (the ConcurImport SSIS packages) sees exactly the same
    filenames it sees today, minus the .pgp/.gpg extension.

      concur_invoice_20260824.txt.pgp  ->  concur_invoice_20260824.txt

    Files that are NOT .pgp/.gpg are left completely alone, so this function is a
    no-op against the plaintext files Concur sends today. That makes it safe to
    deploy BEFORE Concur switches encryption on, and it keeps working through a
    mixed period where some files are encrypted and some are not.

    The encrypted original is archived, never deleted - a bad decrypt must not
    lose a file we can no longer re-fetch.

    BLOCKED ON KEY MATERIAL. Decrypting inbound files needs a Barnes PGP SECRET
    key, and we do not have one for Concur. KeysEncryptions\Concur\Dev\
    concur_exported_key_pair.key is NOT a PGP key pair despite the name - it is a
    PEM RSA private key (the SFTP login key, the same one as DevConcurSFTP_*.ppk),
    and keyphrase.txt is byte-identical to App_KeyPassPhrase, i.e. the SSH
    passphrase. gpg rejects the file outright: "no valid OpenPGP data found".
    So before this function can do anything: generate a Barnes PGP key pair, give
    Concur the public half, and point App_PgpPassPhraseLocation at ITS passphrase.
    The code below is correct and harmless until then - it no-ops on plaintext.

    Deliberately different from HSBC-US_get.ps1:
      * every test below is against the file actually being processed, not a
        differently-named sibling (the HSBC ACH branch tests $file when it means
        $outputFile, which is why every ACH ack is misfiled)
      * the passphrase goes to gpg over stdin, not on the command line, where it
        is visible in the process table to anyone with Task Manager
      * a failed decrypt leaves the .pgp in place and continues, so the next run
        retries it
#>
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
        Decrypt-Files
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
