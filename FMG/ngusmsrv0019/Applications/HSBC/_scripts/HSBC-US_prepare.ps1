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
Start-Transcript -Path "$PSScriptRoot\prepare.log" -Append

$dataDirectory = $config.App_DataDirectory

foreach ($location in $config.App_ApOutFileLocations){

    Write-host "Processing: $location"

    if(Test-Path $location){
        $files = Get-ChildItem -Path $location

        foreach ($file in $files){
            $destFile = Join-Path $dataDirectory $file.Name

            try{
                Move-Item -Path $file.FullName -Destination $destFile -Force
                Write-Host "Moved: $($file.FullName) -> $destFile"
            }
            catch {
                Write-Host "ERROR moving file: $($file.FullName)"
                Write-Host $_
            }
        }
            
    }
    else
    {
        Write-Host "Path not accessible"
    }
}
Stop-Transcript -ErrorAction SilentlyContinue            