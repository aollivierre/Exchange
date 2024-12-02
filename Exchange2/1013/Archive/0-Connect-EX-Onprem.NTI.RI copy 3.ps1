# Define the IP address since DNS isn't resolving
$serverIP = '192.168.0.22'

# Test connectivity first
Write-Host "Testing connectivity..." -ForegroundColor Cyan
Test-NetConnection $serverIP -Port 5985

# Configure WinRM if needed
Write-Host "Configuring WinRM..." -ForegroundColor Yellow
winrm quickconfig -force
Set-Item WSMan:\localhost\Client\TrustedHosts -Value $serverIP -Force
Set-Item WSMan:\localhost\Client\AllowUnencrypted -Value $true -Force
Restart-Service WinRM

# Path to the credentials file
$SecretsFile = Join-Path -Path $PSScriptRoot -ChildPath "secrets.exchange.$serverIP.psd1"

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

# Retrieve the credentials
$UserCredential = Get-ExchangeCredential -SecretsFilePath $SecretsFile

# Try different connection URIs and authentication methods
$connectionAttempts = @(
    @{
        Uri = "http://$serverIP/PowerShell/"
        Auth = "Basic"
    },
    @{
        Uri = "http://$serverIP/PowerShell-LiveID"
        Auth = "Basic"
    },
    @{
        Uri = "http://$serverIP/PowerShell/"
        Auth = "Negotiate"
    },
    @{
        Uri = "http://$serverIP:5985/PowerShell/"
        Auth = "Basic"
    }
)

$sessionEstablished = $false

foreach ($attempt in $connectionAttempts) {
    try {
        Write-Host "Attempting connection with URI: $($attempt.Uri) and Authentication: $($attempt.Auth)" -ForegroundColor Cyan
        
        $Session = New-PSSession `
            -ConfigurationName Microsoft.Exchange `
            -ConnectionUri $attempt.Uri `
            -Authentication $attempt.Auth `
            -AllowRedirection `
            -Credential $UserCredential `
            -SessionOption (New-PSSessionOption -SkipCACheck -SkipCNCheck -SkipRevocationCheck) `
            -ErrorAction Stop

        # If we get here, the connection was successful
        Write-Host "Successfully connected!" -ForegroundColor Green
        Import-PSSession $Session -DisableNameChecking -AllowClobber
        $sessionEstablished = $true
        break
    }
    catch {
        Write-Host "Connection attempt failed: $($_.Exception.Message)" -ForegroundColor Red
        continue
    }
}

if (-not $sessionEstablished) {
    Write-Host "All connection attempts failed. Please check:" -ForegroundColor Red
    Write-Host "1. Your credentials (use domain\username format)" -ForegroundColor Yellow
    Write-Host "2. The Exchange Management Tools are installed" -ForegroundColor Yellow
    Write-Host "3. The WinRM service is running on both machines" -ForegroundColor Yellow
    Write-Host "4. Any firewall rules that might be blocking the connection" -ForegroundColor Yellow
    Write-Host "5. The Exchange server's PowerShell virtual directory settings" -ForegroundColor Yellow
}