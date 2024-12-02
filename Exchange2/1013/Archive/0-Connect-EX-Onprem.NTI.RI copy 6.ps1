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
    $session = $null

    # Try Kerberos first
    try {
        Write-Host "Attempting Kerberos authentication to $($server.Key)..." -ForegroundColor Cyan
        
        if (Test-Connection -ComputerName $server.Value.FQDN -Count 1 -Quiet) {
            $SecretsFile = Join-Path -Path $PSScriptRoot -ChildPath "secrets.exchange.$($server.Value.FQDN).psd1"
            $UserCredential = Get-ExchangeCredential -SecretsFilePath $SecretsFile
            
            $kerberosParams = @{
                ConfigurationName = 'Microsoft.Exchange'
                ConnectionUri     = "http://$($server.Value.FQDN)/PowerShell/"
                Authentication   = 'Kerberos'
                Credential       = $UserCredential
                AllowRedirection = $true
                ErrorAction      = 'Stop'
            }
            
            $session = New-PSSession @kerberosParams
            Write-Host "Kerberos connection established!" -ForegroundColor Green
        }
    }
    catch {
        Write-Host "Kerberos authentication failed: $($_.Exception.Message)" -ForegroundColor Yellow
        Write-Host "Falling back to Basic authentication..." -ForegroundColor Yellow
        
        try {
            winrm quickconfig -force
            Set-Item WSMan:\localhost\Client\TrustedHosts -Value $server.Value.IP -Force
            Set-Item WSMan:\localhost\Client\AllowUnencrypted -Value $true -Force
            Restart-Service WinRM

            $SecretsFile = Join-Path -Path $PSScriptRoot -ChildPath "secrets.exchange.$($server.Value.IP).psd1"
            $UserCredential = Get-ExchangeCredential -SecretsFilePath $SecretsFile

            $basicParams = @{
                ConfigurationName = 'Microsoft.Exchange'
                ConnectionUri     = "http://$($server.Value.IP)/PowerShell/"
                Authentication   = 'Basic'
                Credential       = $UserCredential
                AllowRedirection = $true
                SessionOption    = (New-PSSessionOption -SkipCACheck -SkipCNCheck -SkipRevocationCheck)
                ErrorAction      = 'Stop'
            }

            $session = New-PSSession @basicParams
            Write-Host "Basic auth connection established!" -ForegroundColor Green
        }
        catch {
            Write-Host "All connection attempts failed for $($server.Key): $($_.Exception.Message)" -ForegroundColor Red
            continue
        }
    }

    if ($session) {
        try {
            $importParams = @{
                Session           = $session
                DisableNameChecking = $true
                AllowClobber       = $true
                Prefix             = $server.Value.Prefix
            }
            
            Import-PSSession @importParams
            $servers[$server.Key].Session = $session
            Write-Host "Successfully imported Exchange session for $($server.Key)!" -ForegroundColor Green
        }
        catch {
            Write-Host "Failed to import session for $($server.Key): $($_.Exception.Message)" -ForegroundColor Red
            if ($session) { Remove-PSSession $session }
        }
    }
}