param (
    [Parameter(Mandatory=$true)]
    [ValidateSet("stage","production")]
    [string]$Environment
)

$basePath = Split-Path -Path $PSScriptRoot -Parent
$configPath = Join-Path $basePath "\configuration\$Environment.json"
write-host $configPath
if (-not (Test-Path $configPath)) {
    Write-Error "Config file not found: $configPath"
    exit 1
}

$config = Get-Content $configPath | ConvertFrom-Json

$dataDirectory = $config.App_DataDirectory
$processedPath = $config.App_ProcessedPath


Start-Transcript -Path "$PSScriptRoot\send_log.log" -Append
#$dataDirectory = "E:\Applications\HSBC\stage\"
#$processedPath = "E:\Applications\HSBC\stage\processed\"



function Encrypt-Files {
    
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
	$files = Get-ChildItem -Path $dataDirectory -File | where {$_.Extension -in ".dat",".xml",".txt",""}

    foreach ($file in $files) {
        write-host "Encrypting " $file.FullName
        $_pgpFile = Join-Path $dataDirectory ($file.BaseName + "_$timestamp.pgp")

        gpg --encrypt --recipient "cmbitconnectdigitalmappingvalidation@hsbc.co.in" `
            --output $_pgpFile $file.FullName

        if ($LASTEXITCODE -ne 0 -or -not (Test-Path $_pgpFile)) {
            Write-Error "Encryption failed for $($file.FullName)"
            continue
        }

        Write-Host "Encryption successful: $_pgpFile"

        # Move original file to processed only after encryption
        $folderName = Get-Date -Format "yyyy-MM-dd"
        $_folderPath = Join-Path -Path $processedPath -ChildPath $folderName

        if (-not(Test-Path $_folderPath)) {
            New-Item -Path $_folderPath -ItemType Directory | Out-Null
        }

        Move-Item $file.FullName -Destination $_folderPath
    }

}

function Send-Files {

    $remotePath ="/"
    $pgpFiles = Get-ChildItem -Path $dataDirectory -Filter "*.pgp" -File

     foreach ($pgpFile in $pgpFiles){
        
        $_pgpfileName = $pgpFile.Name
        $_remotePath = $remotePath.TrimEnd('/') + "/" + $_pgpfileName

        Write-Host "Sending $_pgpfileName to HSBC"
        
        try {
            . "E:\Scripts\WinScpFTP\HSBC\HSBC-US.ps1"			
            $result = SendFile -LocalFile $pgpFile.FullName -RemotePath $_remotePath -Environment $Environment

            if ($result){

                $folderName = Get-Date -Format "yyyy-MM-dd"
                $_folderPath = Join-Path -Path $processedPath -ChildPath $folderName

                if (-not(Test-Path $_folderPath)) {
                    New-Item -Path $_folderPath -ItemType Directory | Out-Null
                }

                Move-Item $pgpFile.FullName -Destination $_folderPath

                Write-Host "Sent and archived $_pgpfileName"
            }
            else {
                Write-Host "Unable to send file $_pgpfileName"
            }
        }
        catch {
            Write-Host "An error occurred: $_"
        }

     }

}



function old_main{

	$remotePath ="/"
    $folderName = Get-Date -Format "yyyy-MM-dd"
    $_folderPath = Join-Path -Path $processedPath -ChildPath $folderName
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"

	$files = Get-ChildItem -Path $dataDirectory -File | where {$_.Extension -in ".dat",".xml",".txt",""}
	try
    {	
		foreach ($file in $files) {
			write-host "Encrypting " $file.FullName
            			
            if (-not(Test-Path -Path $_folderPath -PathType Container))
			{
				New-Item -Path $_folderPath -ItemType Directory | Out-Null
				Write-Host "Created Directory $_folderPath"
			}
            

            $_pgpFile = $file.FullName + "_$timestamp.pgp"


            gpg --encrypt --recipient "cmbitconnectdigitalmappingvalidation@hsbc.co.in" `
            --output $_pgpFile $file.FullName
			
            if ($LASTEXITCODE -ne 0 -or -not (Test-Path $_pgpFile)) {
                Write-Error "Encryption failed for $($file.FullName). Skipping send."
                continue
            }

            Write-Host "Encryption successful: $_pgpFile"

            Write-Host "Archiving File Into Folder " $_folderPath
            Move-Item $file.FullName -Destination $_folderPath
            
            $_pgpfileName = [System.IO.Path]::GetFileName($_pgpFile)
            $_remotePath = $remotePath.TrimEnd('/')+"/" + $_pgpfileName


            Write-Host "Sending File $_pgpfileName to HSBC" 
            . "E:\Scripts\WinScpFTP\HSBC\HSBC-US.ps1"			
            $result = SendFile -LocalFile $_pgpFile -RemotePath $_remotePath -Environment $Environment
            
            # Move encrypted file after sending
            Move-Item $_pgpFile -Destination $_folderPath
         
        }
    }
    catch {
        Write-Host "An error occurred: $_"
    }
}

function main {
    try {
        Encrypt-Files
        Send-Files
    }
    catch {
        Write-Host "An error occurred: $_"
    }
}


main

Stop-Transcript -ErrorAction SilentlyContinue