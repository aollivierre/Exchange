$ProgressPreference = 'SilentlyContinue'

$servers = @{
    IQ = @{
        IP = '192.168.14.230'
        FQDN = 'NTI-IQ-EX01.iq.nti.local'
    }
}

function Get-ExchangeCredential {
    param ([string]$SecretsFilePath)
 
    if (Test-Path $SecretsFilePath) {
        try {
            return Import-Clixml -Path $SecretsFilePath
        } catch {
            Write-Warning "Invalid credentials. Enter new credentials."
            Remove-Item -Path $SecretsFilePath -Force
            return Get-ExchangeCredential -SecretsFilePath $SecretsFilePath
        }
    } else {
        Write-Host "Enter domain credentials (domain\username)" -ForegroundColor Yellow
        $Cred = Get-Credential
        $Cred | Export-Clixml -Path $SecretsFilePath
        return $Cred
    }
}

$sessionParams = @{
    IdleTimeout = 2147483
    SkipCACheck = $true
    SkipCNCheck = $true
    SkipRevocationCheck = $true
    MaximumReceivedDataSizePerCommand = 1GB
    MaximumRedirection = 0
    NoCompression = $false
}
$sessionOption = New-PSSessionOption @sessionParams

$requiredCommands = @('Get-MailContact', 'Set-MailContact')

foreach ($server in $servers.GetEnumerator()) {
    $exchangeParams = @{
        ConfigurationName = 'Microsoft.Exchange'
        ConnectionUri = "http://$($server.Value.FQDN)/PowerShell/"
        Authentication = 'Kerberos'
        AllowRedirection = $true
        SessionOption = $sessionOption
        ErrorAction = 'Stop'
    }

    try {
        Write-Host "Connecting to $($server.Key)..." -ForegroundColor Cyan
        
        if (-not (Test-Connection -ComputerName $server.Value.FQDN -Count 1 -Quiet)) {
            throw "DNS resolution failed for $($server.Value.FQDN)"
        }
        
        $SecretsFile = Join-Path -Path $PSScriptRoot -ChildPath "secrets.exchange.$($server.Value.FQDN).psd1"
        $exchangeParams.Credential = Get-ExchangeCredential -SecretsFilePath $SecretsFile
        
        $servers[$server.Key].Session = New-PSSession @exchangeParams

        $importParams = @{
            Session = $servers[$server.Key].Session
            CommandName = $requiredCommands
            AllowClobber = $true
            DisableNameChecking = $true
        }
        Import-PSSession @importParams
        
        Write-Host "Connected to $($server.Key)!" -ForegroundColor Green
        continue
    }
    catch {
        Write-Error "Connection failed: $_"
    }
}