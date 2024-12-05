$ProgressPreference = 'SilentlyContinue'

$servers = @{
    OTT = @{
        IP = '10.20.48.112'
        FQDN = 'NTI-OTT-EXCH01.ott.nti.local'
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

 $sessionOption = New-PSSessionOption -IdleTimeout 2147483 -SkipCACheck -SkipCNCheck -SkipRevocationCheck
 
 foreach ($server in $servers.GetEnumerator()) {
    # Kerberos connection parameters
    $kerberosParams = @{
        ConfigurationName = 'Microsoft.Exchange'
        ConnectionUri     = "http://$($server.Value.FQDN)/PowerShell/"
        Authentication   = 'Kerberos'
        AllowRedirection = $true
        ErrorAction      = 'Stop'
        SessionOption    =  $sessionOption
    }
 
    # Basic auth parameters  
    $basicParams = @{
        ConfigurationName = 'Microsoft.Exchange'
        ConnectionUri     = "http://$($server.Value.IP)/PowerShell/"
        Authentication   = 'Basic'
        AllowRedirection = $true
        SessionOption    =  $sessionOption
        ErrorAction      = 'Stop'
    }
 
    # Import parameters
    $importParams = @{
        DisableNameChecking = $true
        AllowClobber       = $true
    }
 
    # Try Kerberos first
    try {
        Write-Host "Attempting Kerberos authentication to $($server.Key)..." -ForegroundColor Cyan
        
        if (-not (Test-Connection -ComputerName $server.Value.FQDN -Count 1 -Quiet)) {
            throw "DNS resolution failed for $($server.Value.FQDN)"
        }
        
        $SecretsFile = Join-Path -Path $PSScriptRoot -ChildPath "secrets.exchange.$($server.Value.FQDN).psd1"
        $UserCredential = Get-ExchangeCredential -SecretsFilePath $SecretsFile
        $kerberosParams.Credential = $UserCredential
        
        Write-Host "Creating new PS Session..." -ForegroundColor Yellow
        $servers[$server.Key].Session = New-PSSession @kerberosParams
        
        Write-Host "Importing PS Session..." -ForegroundColor Yellow 
        Import-PSSession $servers[$server.Key].Session @importParams
        
        Write-Host "Successfully connected to $($server.Key) using Kerberos!" -ForegroundColor Green
        continue
    }
    catch {
        Write-Host "Kerberos authentication failed with error: $_" -ForegroundColor Red
        Write-Host "Falling back to Basic authentication..." -ForegroundColor Yellow
    }
 
    # Fall back to Basic auth
    try {
        # Configure WinRM for Basic auth
        Write-Host "Configuring WinRM for Basic auth..." -ForegroundColor Yellow
        winrm quickconfig -force
        Set-Item WSMan:\localhost\Client\TrustedHosts -Value $server.Value.IP -Force
        Set-Item WSMan:\localhost\Client\AllowUnencrypted -Value $true -Force
        Restart-Service WinRM
 
        $SecretsFile = Join-Path -Path $PSScriptRoot -ChildPath "secrets.exchange.$($server.Value.IP).psd1"
        $UserCredential = Get-ExchangeCredential -SecretsFilePath $SecretsFile
        $basicParams.Credential = $UserCredential
 
        Write-Host "Creating new PS Session with Basic auth..." -ForegroundColor Yellow
        $servers[$server.Key].Session = New-PSSession @basicParams
        
        Write-Host "Importing PS Session..." -ForegroundColor Yellow
        Import-PSSession $servers[$server.Key].Session @importParams
        
        Write-Host "Successfully connected to $($server.Key) using Basic auth!" -ForegroundColor Green
    }
    catch {
        Write-Host "All connection attempts failed for $($server.Key): $($_.Exception.Message)" -ForegroundColor Red
    }
 }