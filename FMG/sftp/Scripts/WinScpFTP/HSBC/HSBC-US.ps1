# Load WinSCP .NET assembly
Add-Type -Path "C:\Program Files (x86)\WinSCP\WinSCPnet.dll"


#$_keyPhrase = Get-KeyPhraseValue -SecretName "SFTP-HSBC-US"

# Set up session options
$sessionOptions = New-Object WinSCP.SessionOptions -Property @{
    Protocol = [WinSCP.Protocol]::Sftp
    HostName = "ecom-sftp.fguk-pprd2.hsbc.com"
    PortNumber = 10022
    UserName = "CT000032270_35644"
    SshHostKeyFingerprint = "ssh-rsa 4096 2M0mmjdU54IHuTlP1go72XnejXLxYF5HkrsZnsb6ElM"
    SshPrivateKeyPath = "E:\Keys_Encryptions\HSBC\SSH\hsbc_private_key.ppk"
    SshPrivateKeyPassPhrase = $_keyPhrase
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
        PortNumber = 10022
    }

    $session = New-Object WinSCP.Session

    try
    {
    write-host "Connecting..."
        # Connect
        $session.SessionLogPath =  Join-Path -Path $PSScriptRoot -ChildPath "hsbc-us-send.log"
        $session.Open($sessionOptions)
        
        if ($session.Opened) {            
            Write-Host "Connected to HSBC, uploading $LocalFile to $RemotePath"
            
            $transferOptions = New-Object WinSCP.TransferOptions
            $transferOptions.PreserveTimestamp = $false

			$transferResult = $session.PutFiles($LocalFile, $RemotePath, $false, $transferOptions)
			if($transferResult.IsSuccess){
				Write-Host "Sent Successfully"
				return $true
			}
			else {
				Write-Error "File transfer failed."
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
        PortNumber = 10022
    }

    Write-Host "Get files from $RemotePath saving the files to $SavePath"
    
    try
    {
        Write-Host "Connecting to HSBC"
        
        $session = New-Object WinSCP.Session
        $session.SessionLogPath = Join-Path -Path $PSScriptRoot -ChildPath "hsbc-us-get.log"

        $session.Open($sessionOptions)
        if ($session.Opened) {
            Write-Host "Connected to HSBC $RemotePath"

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

    $_passPhrase = Get-SecretValue -PhraseLocation $config.App_KeyPassPhraseLocation -FileKey $config.File_Key

     
    Get-SftpFiles -HostName $_hostName -UserName $_userName -PassPhrase $_passPhrase -SshKey $_sshKey -SshKeyPath $_sshKeyPath -RemotePath $RemotePath -SavePath $SavePath
}



function SendFile {
    param (
        [string]$LocalFile, 
        [string]$RemotePath = "/",
        [string]$Environment
    )
    $configPath = Join-Path -Path $PSScriptRoot -ChildPath "/config/$Environment/config.json"
    $config = Get-Content $configPath | ConvertFrom-Json

    $_hostName = $config.App_HostName
	$_userName = $config.App_UserName
	$_sshKey = $config.App_SSHKey
    $_sshKeyPath = $config.App_KeyPath

    $_passPhrase = Get-SecretValue -PhraseLocation $config.App_KeyPassPhraseLocation -FileKey $config.File_Key

    $result = Send-SftpFile -HostName $_hostName -UserName $_userName -PassPhrase $_passPhrase -SshKey $_sshKey -LocalFile $LocalFile -SshKeyPath $_sshKeyPath -RemotePath $RemotePath
	return $result
}

function Get-SecretValue {
    param (
        [string]$PhraseLocation,
        [byte[]]$FileKey
    )

    $key = $FileKey
    #$secure = Get-Content $PhraseLocation | ConvertTo-SecureString -Key $key

    #$ptr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
    #$_password = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($ptr)
    $_password = "Tomorrow-Today-99!Fox"
    return $_password
}