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
$acknowledgmentsDirectory = $config.App_AcknowledgmentPath
$decryptedDirectory = $config.App_AcknowledgmentDecryptedPath
$processedRoot = $config.App_ProcessedPath
Write-Host $acknowledgmentsDirectory


Start-Transcript -Path "$PSScriptRoot\get_log.log" -Append

function Get-AcknowledgmentFiles {

    Write-Host "Downloading files from HSBC"

    . "E:\Scripts\WinScpFTP\HSBC\HSBC-US.ps1"			
    GetFiles -RemotePath "/Inbox/" -SavePath $acknowledgmentsDirectory -Environment $Environment

}


function Decrypt_Files {

	. "E:\Scripts\WinScpFTP\HSBC\HSBC-US.ps1"	
    $passPhrase = Get-SecretValue -PhraseLocation $config.App_KeyPassPhraseLocation -FileKey $config.File_Key

    $folderName = Get-Date -Format "yyyy-MM-dd"
    $_folderPath = Join-Path -Path $decryptedDirectory -ChildPath $folderName

    $files = Get-ChildItem -Path $acknowledgmentsDirectory -File | Where-Object {$_.Extension -in ".txt",".xml"}

    if ($files.Count -eq 0) {
        Write-Host "No files to decrypt"
        return
    }

    if (-not (Test-Path $_folderPath)) {
        New-Item -Path $_folderPath -ItemType Directory | Out-Null
        Write-Host "Created Directory $_folderPath"
    }

    foreach ($file in $files) {
        Write-Host "Decrypting $($file.FullName)"

        $_extension = [System.IO.Path]::GetExtension($file.FullName)
        $outputFile = Join-Path $file.DirectoryName ("decrypted_" + $file.BaseName + $_extension)

        gpg --quiet --batch --yes --passphrase $passPhrase `
            --pinentry-mode loopback `
            -o $outputFile `
            -d $file.FullName 2>$null

        if ($LASTEXITCODE -ne 0 -or -not (Test-Path $outputFile)) {
            Write-Error "Decryption failed for $($file.Name). Will retry next run."
            continue
        }
		
		$erpFolder = $null
        $originalFile = $null

		if ($file.Name -like "*JBA_*") {
			
            $erpFolder = "HysonJBA"

            if ($file.BaseName -match 'ACK[12]PSRV3\.PC\d+\.JBA_(?<type>[A-Z]+)_(?<id>[A-F0-9]+)\.') {

                $type = $matches['type']
                $id = $matches['id']

                $searchPattern = "JBA_${type}_${id}*.xml"

                Write-Host "Searching for JBA original file: $searchPattern"

                $originalFile = Get-ChildItem `
                    -Path $processedRoot `
                    -Recurse `
                    -File `
                    -Filter $searchPattern `
                    -ErrorAction SilentlyContinue |
                    Select-Object -First 1
            }
            
		}
		elseif ($file.Name -like "*RaymondM3_*") {
			
            $erpFolder = "ASRaymondM3"

            if ($file.BaseName -match 'ACK[12]PSRV3\.PC\d+\.(?<erp>RaymondM3)_(?<ref>\d+)_(?<seq>\d+)\.') {

                $erp = $matches['erp']
                $ref = $matches['ref']
                $seq = $matches['seq']

                $searchPattern = "${erp}_*_${ref}_${seq}.xml"

                Write-Host "Searching for RaymondM3 original file: $searchPattern"

                $originalFile = Get-ChildItem `
                    -Path $processedRoot `
                    -Recurse `
                    -File `
                    -Filter $searchPattern `
                    -ErrorAction SilentlyContinue |
                    Select-Object -First 1
            }
		}
        elseif ($file.Name -match '^decrypted_ACK[12]STDK') {

            Write-Host "Detected NACHA ACK file"

            #
            # Read ACK file contents
            #
            $lines = Get-Content $outputFile

            $companyId = $null

            #
            # Extract company id
            #
            foreach ($line in $lines) {

                if ($line -match 'PC000078445(\d{9})') {

                    $companyId = $matches[1]
                    break
                }
            }

            Write-Host "Detected Company ID: $companyId"

            $searchCompanyId = $null
            $searchPattern = $null

            switch ($companyId) {

                #
                # Hyson JBA NACHA
                #
                "797025162" {

                    $erpFolder = "HysonJBA"

                    # reversed + leading 5
                    # 797025162 -> 5261520797
                    $searchCompanyId = "5261520797"

                    $searchPattern = "JBA_ACH_*.DAT"

                    break
                }

                #
                # RaymondM3 NACHA
                #
                "797025154" {

                    $erpFolder = "ASRaymondM3"

                    # reversed + leading 5
                    # 797025154 -> 5451520797
                    $searchCompanyId = "5451520797"

                    $searchPattern = "RaymondM3_ACH_*.txt"

                    break
                }

                default {

                    Write-Warning "Unknown NACHA company id: $companyId"

                    $erpFolder = "Unknown"
                }
            }

            #
            # Search for original ACH file
            #
            if ($searchCompanyId -and $searchPattern) {

                Write-Host "Searching for original NACHA file using $searchCompanyId"

                $candidateFiles = Get-ChildItem `
                    -Path $processedRoot `
                    -Recurse `
                    -File `
                    -Filter $searchPattern `
                    -ErrorAction SilentlyContinue

                foreach ($candidate in $candidateFiles) {

                    $candidateLines = Get-Content `
                        $candidate.FullName `
                        -TotalCount 5

                    foreach ($candidateLine in $candidateLines) {

                        if ($candidateLine -match $searchCompanyId) {

                            $originalFile = $candidate

                            Write-Host "Found original NACHA file: $($candidate.FullName)"

                            break
                        }
                    }

                    if ($originalFile) {
                        break
                    }
                }
            }

        }
		else {
			Write-Warning "Unknown ERP type for file $($file.Name). Moving to fallback."
			$erpFolder = "Unknown"
		}
		
		# Build final destination path		
		$finalPath = Join-Path "E:\shared\hsbc" -ChildPath "$erpFolder\$Environment\$folderName"
		
		# Create directory if it doesn't exist
		if (-not (Test-Path $finalPath)) {
			New-Item -Path $finalPath -ItemType Directory -Force | Out-Null
			Write-Host "Created Directory $finalPath"

		}

        Write-Host "Decryption successful: $outputFile"

        
		#Move-Item $outputFile -Destination $finalPath -Force
        $destFile = Join-Path $finalPath (Split-Path $outputFile -Leaf)

        robocopy (Split-Path $outputFile) $finalPath (Split-Path $outputFile -Leaf) /MOV /COPY:DAT /R:1 /W:1 /NFL /NDL /NP | Out-Null
        
        if ($LASTEXITCODE -ge 8) {
            Write-Error "Robocopy failed for $outputFile with code $LASTEXITCODE"
        }
        
        # Move original payment file        
        if ($originalFile) {

            Write-Host "Moving original file: $($originalFile.FullName)"

           <# Move-Item `
                -Path $originalFile.FullName `
                -Destination $finalPath `
                -Force
            #>
            robocopy (Split-Path $originalFile.FullName) $finalPath (Split-Path $originalFile.FullName -Leaf) /MOV /COPY:DAT /R:1 /W:1 /NFL /NDL /NP | Out-Null
            
            if ($LASTEXITCODE -ge 8) {
                Write-Error "Robocopy failed for $originalFile with code $LASTEXITCODE"
            }
        }
        else {

            Write-Warning "Original file not found for $($file.Name)"
        }
		
        #Move-Item $file.FullName -Destination $_folderPath -Force
        robocopy (Split-Path $file.FullName) $_folderPath (Split-Path $file.FullName -Leaf) `
            /MOV /COPY:DAT /R:1 /W:1 /NFL /NDL /NP | Out-Null
        
        if ($LASTEXITCODE -ge 8) {
            Write-Error "Robocopy failed for $file with code $LASTEXITCODE"
        }

    }

}


function old-main {
    Write-Host "Sending File $_pgpfileName to HSBC" 
    . "E:\Scripts\WinScpFTP\HSBC\HSBC-US.ps1"			
    GetFiles -RemotePath "/Inbox/" -SavePath $acknowledgmentsDirectory -Environment $Environment



    $folderName = Get-Date -Format "yyyy-MM-dd"
    $_folderPath = Join-Path -Path $decryptedDirectory -ChildPath $folderName

    $files = Get-ChildItem -Path $acknowledgmentsDirectory | where {$_.Extension -in ".txt",".xml"}

    try
    {

    if ($files.Count -gt 0) {

        if (-not (Test-Path -Path $_folderPath -PathType Container)) {
            New-Item -Path $_folderPath -ItemType Directory | Out-Null
            Write-Host "Created Directory $_folderPath"
        }

            foreach ($file in $files) {
                write-host "Decrypting " $file.FullName

                if (-not(Test-Path -Path $_folderPath -PathType Container))
		        {
			        New-Item -Path $_folderPath -ItemType Directory | Out-Null
			        Write-Host "Created Directory $_folderPath"
		        }

                $_extension = [System.IO.Path]::GetExtension($file.FullName)
        
                $outputFile = Join-Path $file.DirectoryName ("decrypted_" + $file.BaseName + $_extension)
        
                gpg --quiet --batch --yes --passphrase $passPhrase `
                --pinentry-mode loopback `
                -o $outputFile `
                -d $file.FullName 2>$null

                if ($LASTEXITCODE -ne 0) {
                    throw "GPG decryption failed with exit code $LASTEXITCODE"
                }
                else 
                {

                    Move-Item $outputFile -Destination $_folderPath
                    Move-Item $file.FullName -Destination $_folderPath
                }
        
            }
        }
    }
    catch{
        Write-Error $_
        Stop-Transcript
        exit 1
    }
}

function main {
    try{
        Get-AcknowledgmentFiles
        Decrypt_Files
    }
    catch{
        Write-Error $_
        Stop-Transcript
        exit 1
    }
}

main

Stop-Transcript -ErrorAction SilentlyContinue
            