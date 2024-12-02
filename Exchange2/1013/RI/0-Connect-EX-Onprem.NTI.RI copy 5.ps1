$servers = @{
    RI = @{
        IP = '192.168.0.22'
        FQDN = 'NTI-RI-EX02.RI.nti.local'
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
    # Kerberos connection parameters
    $kerberosParams = @{
        ConfigurationName = 'Microsoft.Exchange'
        ConnectionUri     = "http://$($server.Value.FQDN)/PowerShell/"
        Authentication   = 'Kerberos'
        AllowRedirection = $true
        ErrorAction      = 'Stop'
    }
 
    # Basic auth parameters
    $basicParams = @{
        ConfigurationName = 'Microsoft.Exchange'
        ConnectionUri     = "http://$($server.Value.IP)/PowerShell/"
        Authentication   = 'Basic'
        AllowRedirection = $true
        SessionOption    = (New-PSSessionOption -SkipCACheck -SkipCNCheck -SkipRevocationCheck)
        ErrorAction      = 'Stop'
    }
 
    # Import parameters
    $importParams = @{
        DisableNameChecking = $true
        AllowClobber       = $true
        Prefix             = $server.Value.Prefix
    }
 
    # Try Kerberos first
    try {
        Write-Host "Attempting Kerberos authentication to $($server.Key)..." -ForegroundColor Cyan
        
        # Test DNS resolution
        if (Test-Connection -ComputerName $server.Value.FQDN -Count 1 -Quiet) {
            $SecretsFile = Join-Path -Path $PSScriptRoot -ChildPath "secrets.exchange.$($server.Value.FQDN).psd1"
            $UserCredential = Get-ExchangeCredential -SecretsFilePath $SecretsFile
            $kerberosParams.Credential = $UserCredential
            
            $servers[$server.Key].Session = New-PSSession @kerberosParams
            Import-PSSession $servers[$server.Key].Session @importParams
            # Write-Host "Successfully connected using Kerberos!" -ForegroundColor Green
            Write-Host "Successfully connected to $($server.Key) using Kerberos!" -ForegroundColor Green
            continue
        }
    }
    catch {
        Write-Host "Kerberos authentication failed: $($_.Exception.Message)" -ForegroundColor Yellow
        Write-Host "Falling back to Basic authentication..." -ForegroundColor Yellow
    }
 
    # Fall back to Basic auth
    try {
        # Configure WinRM for Basic auth
        winrm quickconfig -force
        Set-Item WSMan:\localhost\Client\TrustedHosts -Value $server.Value.IP -Force
        Set-Item WSMan:\localhost\Client\AllowUnencrypted -Value $true -Force
        Restart-Service WinRM
 
        $SecretsFile = Join-Path -Path $PSScriptRoot -ChildPath "secrets.exchange.$($server.Value.IP).psd1"
        $UserCredential = Get-ExchangeCredential -SecretsFilePath $SecretsFile
        $basicParams.Credential = $UserCredential
 
        $servers[$server.Key].Session = New-PSSession @basicParams
        Import-PSSession $servers[$server.Key].Session @importParams
        # Write-Host "Successfully connected using Basic auth!" -ForegroundColor Green
        Write-Host "Successfully connected to $($server.Key) using Basic auth!" -ForegroundColor Green
    }
    catch {
        Write-Host "All connection attempts failed for $($server.Key): $($_.Exception.Message)" -ForegroundColor Red
    }
 }