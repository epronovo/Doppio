# Load WinSCP .NET assembly
Add-Type -Path "C:\Program Files (x86)\WinSCP\WinSCPnet.dll"

function Get-SftpFiles {    
    param (
        [string]$HostName,
        [string]$UserName,
        [string]$SshKeyPath,
        [string]$PassPhrase,
        [string]$SshKey,
        [string]$RemotePath,
        [string]$SavePath,
		[string]$Environment
    )
	
	if ($Environment -eq "production"){
		$sessionOptions = New-Object WinSCP.SessionOptions -Property @{
			Protocol              = [WinSCP.Protocol]::Sftp
			HostName              = $HostName
			UserName              = $UserName
			SshPrivateKeyPath     = $SshKeyPath
			PrivateKeyPassphrase  = $PassPhrase
			SshHostKeyFingerprint = $SshKey
			PortNumber = 10022
		}
	}
	else{
		$sessionOptions = New-Object WinSCP.SessionOptions -Property @{
			Protocol = [WinSCP.Protocol]::Sftp
			HostName = $HostName
			UserName = $UserName
			Password = $PassPhrase
			SshHostKeyFingerprint = $SshKey
		}
	}
	
    Write-Host "Get files from $RemotePath saving the files to $SavePath"
    
    try
    {
        Write-Host "Connecting to HighRadius"
        
        $session = New-Object WinSCP.Session
        $session.SessionLogPath = Join-Path -Path $PSScriptRoot -ChildPath "highradius-get.log"

        $session.Open($sessionOptions)
        if ($session.Opened) {
            Write-Host "Connected to HighRadius $RemotePath"

            <#$session.GetFiles($RemotePath, $SavePath).Check()#>
            
            $remoteFiles = $session.ListDirectory($RemotePath).Files | Where-Object { -not $_.IsDirectory }
            $saveFile = $SavePath.TrimEnd('/') + '/'

            $transferOptions = New-Object WinSCP.TransferOptions
            $transferOptions.TransferMode = [WinSCP.TransferMode]::Binary

            foreach ($file in $remoteFiles)
            {
                Write-Host "Preparing to download file: $($file.Name)"
            
                #$remoteFile = ($RemotePath.TrimEnd('/') + '/' + $file.Name)
                $remoteFile = $file.FullName

                Write-Host "Remote file : $remoteFile moving to Save Directory $saveFile"

                $result = $session.GetFiles($remoteFile,(Join-Path $SavePath $file.Name),$true,$transferOptions)
                $result.Check()

                Write-Host "Download complete!"
            }
            
        }
    }
    finally
    {
        $session.Dispose()
    }
}


function GetFiles {
    [CmdletBinding()]
    param (
        
        [string]$RemotePath,
        [string]$SavePath,
        [string]$Environment
    )

    $configPath = Join-Path -Path $PSScriptRoot -ChildPath "/config/$Environment/config.json"
    $config = Get-Content $configPath | ConvertFrom-Json	
	
    $_hostName = $config.App_HostName
	$_userName = $config.App_UserName
	$_sshKey = $config.App_SSHKey
    $_sshKeyPath = $config.App_KeyPath

    $_passPhrase = $config.App_KeyPassPhrase

     
    Get-SftpFiles -HostName $_hostName -UserName $_userName -PassPhrase $_passPhrase -SshKey $_sshKey -SshKeyPath $_sshKeyPath -RemotePath $RemotePath -SavePath $SavePath -Environment $Environment
}