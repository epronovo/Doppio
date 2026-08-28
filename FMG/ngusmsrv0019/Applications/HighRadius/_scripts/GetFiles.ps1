param (
    [Parameter(Mandatory=$true)]
    [ValidateSet("stage","production")]
    [string]$Environment
)

$basePath = Split-Path -Path $PSScriptRoot -Parent
$configPath = Join-Path $basePath "configuration\$Environment.json"
write-host $configPath
if (-not (Test-Path $configPath)) {
    Write-Error "Config file not found: $configPath"
    exit 1
}

$config = Get-Content $configPath | ConvertFrom-Json
$remotePath = $config.App_GetRemotePath
$processedPath = $config.App_ProcessedPath
$dataDirectory = $config.App_DataDirectory

Start-Transcript -Path "$PSScriptRoot\get_log.log" -Append

function Get_Files {

. "E:\Scripts\WinScpFTP\HighRadius\HighRadiusConnect.ps1"			
    GetFiles -RemotePath $remotePath -SavePath $dataDirectory -Environment $Environment

}

function SendToFileShare {
	Param
	(
		[string]$Erp,
        [string]$FileToProcess
	)
	
    Write-host "Send To FileShare : $Erp File $FileToProcess"

	switch ($Erp.ToUpper()){
		"M3" {
			
			$M3saveLocation = $config.App_GetM3SaveLocation
			
			Write-Host "Copy $FileToProcess to $M3saveLocation"
			Copy-Item -Path $FileToProcess -Destination $M3saveLocation -ErrorAction Stop
			
			$dateFolder = Get-Date -Format "yyyy-MM-dd"
			$processedDirectory = Join-Path ($processedPath) $dateFolder
			
			if (-not (Test-Path $processedDirectory)) {
				New-Item -ItemType Directory -Path $processedDirectory -Force | Out-Null
			}
			
			Move-Item -Path $FileToProcess -Destination $processedDirectory -Force -ErrorAction Stop
			
			break
		}
		default {
			Write-Error "Unknown Erp File System $Erp"
		}
		
	}
}

function SendToJBA {
	Param 
	(
		[string]$FileToProcess
	)
	
	Write-host "Send To JBA : $FileToProcess"
	$sendToJBA = $config.App_GetJBASaveLocation
	Copy-Item -Path $FileToProcess -Destination $sendToJBA
	
	$dateFolder = Get-Date -Format "yyyy-MM-dd"
	$processedDirectory = Join-Path ($processedPath) $dateFolder
	
	if (-not (Test-Path $processedDirectory)) {
		New-Item -ItemType Directory -Path $processedDirectory -Force | Out-Null
	}
	
	Move-Item -Path $FileToProcess -Destination $processedDirectory -Force -ErrorAction Stop
}


function Process_Files {

	$files = Get-ChildItem -Path $dataDirectory -File
	Write-Host "====== Processing Files ====="
	foreach ($file in $files) {
		
		$baseName = $file.BaseName
		$recordType = $baseName.split("_")
		
		write-host "Processing File $baseName "
		
		
		switch ($recordType[1].ToUpper()){
		
			{ $_ -eq "CHECK" -or $_ -eq "EDI" } {
				
				if ($recordType.Count -lt 3) {
					Write-Error "Unable to determine file type from filename: $($file.Name)"
					continue
				}
				
				$erp = $recordType[2]
				SendToFileShare -Erp $erp -FileToProcess $file.FullName
				break
			}
			
			"AURORA" {
				SendToJBA -FileToProcess $file.FullName
				break
			}			
			default {
				Write-Error "Unknown File Type"
			}
		
		}
	
	}
}


function main {
    try{
        Get_Files
		Process_Files
    }
    catch{
        Write-Error $_
        Stop-Transcript
        exit 1
    }
}


main

Stop-Transcript -ErrorAction SilentlyContinue
