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

function main {
    try{
        Get_Files
		#Process_Files
    }
    catch{
        Write-Error $_
        Stop-Transcript
        exit 1
    }
}


main