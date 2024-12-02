# Server configurations
$servers = @{
    RI = @{
        IP = '192.168.0.22'
        Prefix = 'RI'
    }
}

function Get-ExchangeCredential {
    param (
        [string]$SecretsFilePath
    )

    if (Test-Path $SecretsFilePath) {
        try {
            $Cred = Import-Clixml -Path $SecretsFilePath
            return $Cred
        } catch {
            Write-Warning "Stored credentials are invalid or expired. Please enter new credentials."
            Remove-Item -Path $SecretsFilePath -Force
            return Get-ExchangeCredential -SecretsFilePath $SecretsFilePath
        }
    } else {
        Write-Host "Please enter domain credentials in the format 'domain\username'" -ForegroundColor Yellow
        $Cred = Get-Credential
        $Cred | Export-Clixml -Path $SecretsFilePath
        return $Cred
    }
}

foreach ($server in $servers.GetEnumerator()) {
    Test-NetConnection $server.Value.IP -Port 5985

    # Configure WinRM
    Write-Host "Configuring WinRM for $($server.Key)..." -ForegroundColor Yellow
    winrm quickconfig -force
    Set-Item WSMan:\localhost\Client\TrustedHosts -Value $server.Value.IP -Force
    Set-Item WSMan:\localhost\Client\AllowUnencrypted -Value $true -Force
    Restart-Service WinRM

    # Credentials handling
    $SecretsFile = Join-Path -Path $PSScriptRoot -ChildPath "secrets.exchange.$($server.Value.IP).psd1"
    $UserCredential = Get-ExchangeCredential -SecretsFilePath $SecretsFile

    $sessionParams = @{
        ConfigurationName = 'Microsoft.Exchange'
        ConnectionUri     = "http://$($server.Value.IP)/PowerShell/"
        Authentication   = 'Basic'
        AllowRedirection = $true
        Credential       = $UserCredential
        SessionOption    = (New-PSSessionOption -SkipCACheck -SkipCNCheck -SkipRevocationCheck)
        ErrorAction      = 'Stop'
    }

    $importParams = @{
        DisableNameChecking = $true
        AllowClobber       = $true
        Prefix             = $server.Value.Prefix
    }

    try {
        Write-Host "Connecting to $($server.Key) Exchange server..." -ForegroundColor Cyan
        $servers[$server.Key].Session = New-PSSession @sessionParams
        Import-PSSession $servers[$server.Key].Session @importParams
        Write-Host "Successfully connected to $($server.Key)!" -ForegroundColor Green
    }
    catch {
        Write-Host "Connection to $($server.Key) failed: $($_.Exception.Message)" -ForegroundColor Red
    }
}