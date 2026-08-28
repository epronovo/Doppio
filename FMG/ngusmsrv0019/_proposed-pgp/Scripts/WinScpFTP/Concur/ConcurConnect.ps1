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
        [string]$SavePath
    )

	$sessionOptions = New-Object WinSCP.SessionOptions -Property @{
        Protocol              = [WinSCP.Protocol]::Sftp
        HostName              = $HostName
        UserName              = $UserName
        SshPrivateKeyPath     = $SshKeyPath
        PrivateKeyPassphrase  = $PassPhrase
        SshHostKeyFingerprint = $SshKey
        PortNumber = 22
    }

	Write-Host "Get files from $RemotePath saving the files to $SavePath"

	try
	{
		Write-Host "Connecting to Concur"

		$session = New-Object WinSCP.Session
        $session.SessionLogPath = Join-Path -Path $PSScriptRoot -ChildPath "concur-get.log"

		$session.Open($sessionOptions)
		if($session.Opened) {
			Write-Host "Connected to Concur $RemotePath"

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

                $result = $session.GetFiles($remoteFile,(Join-Path $SavePath $file.Name),$false,$transferOptions)
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


function Send-SftpFile {
	param (
        [string]$HostName,
        [string]$UserName,
        [string]$SshKeyPath,
        [string]$PassPhrase,
        [string]$SshKey,
        [string]$LocalFile,
        [string]$RemotePath = "/"
    )

	$sessionOptions = New-Object WinSCP.SessionOptions -Property @{
        Protocol              = [WinSCP.Protocol]::Sftp
        HostName              = $HostName
        UserName              = $UserName
        SshPrivateKeyPath     = $SshKeyPath
        PrivateKeyPassphrase  = $PassPhrase
        SshHostKeyFingerprint = $SshKey
        PortNumber = 22
    }

	$session = New-Object WinSCP.Session

	try
	{
		Write-Host "Connecting to Concur"

		$session.SessionLogPath = Join-Path -Path $PSScriptRoot -ChildPath "concur-send.log"
		$session.Open($sessionOptions)

		if ($session.Opened) {
			Write-Host "Connected to Concur, uploading $LocalFile to $RemotePath"

			$transferOptions = New-Object WinSCP.TransferOptions
			$transferOptions.TransferMode = [WinSCP.TransferMode]::Binary
			$transferOptions.PreserveTimestamp = $false

			$transferResult = $session.PutFiles($LocalFile, $RemotePath, $false, $transferOptions)

			if ($transferResult.IsSuccess) {
				Write-Host "Sent Successfully"
				return $true
			}
			else {
				# Surface WHY, not just that it failed - HSBC's version throws this away.
				foreach ($failure in $transferResult.Failures) {
					Write-Error "Transfer failure: $($failure.Message)"
				}
				Write-Error "File transfer failed for $LocalFile"
				return $false
			}
		}
		else {
			Write-Error "Unable to open session"
			return $false
		}
	}
	catch {
		Write-Error "An unexpected error occurred: $($_.Exception.Message)"
		return $false
	}
	finally
	{
		$session.Dispose()
	}

}


function GetFiles {
	[CmdletBinding()]
	param (
		[string]$SavePath,
        [string]$Environment
	)



	$configPath = Join-Path -Path $PSScriptRoot -ChildPath "/config/$Environment/config.json"
    $config = Get-Content $configPath | ConvertFrom-Json

	$_hostName = $config.App_HostName
	$_userName = $config.App_UserName
	$_sshKey = $config.App_SSHKey
	$_passPhrase = $config.App_KeyPassPhrase
    $_sshKeyPath = $config.App_KeyPath
	$RemotePath = "/out/"

	Get-SftpFiles -HostName $_hostName -UserName $_userName -PassPhrase $_passPhrase -SshKey $_sshKey -SshKeyPath $_sshKeyPath -RemotePath $RemotePath -SavePath $SavePath

}


function SendFile {
	[CmdletBinding()]
	param (
		[string]$LocalFile,
		[string]$RemotePath,
		[string]$Environment
	)

	$configPath = Join-Path -Path $PSScriptRoot -ChildPath "/config/$Environment/config.json"
	$config = Get-Content $configPath | ConvertFrom-Json

	$_hostName = $config.App_HostName
	$_userName = $config.App_UserName
	$_sshKey = $config.App_SSHKey
	$_passPhrase = $config.App_KeyPassPhrase
	$_sshKeyPath = $config.App_KeyPath

	# Concur's inbound path. CONFIRM THIS WITH CONCUR before the first production
	# run - "/in/" is the convention, not something observed on this endpoint.
	if (-not $RemotePath) {
		$RemotePath = $config.App_SendRemotePath
	}
	if (-not $RemotePath) {
		Write-Error "No remote path configured. Set App_SendRemotePath in $configPath."
		return $false
	}

	$result = Send-SftpFile -HostName $_hostName -UserName $_userName -PassPhrase $_passPhrase -SshKey $_sshKey -SshKeyPath $_sshKeyPath -LocalFile $LocalFile -RemotePath $RemotePath
	return $result

}
