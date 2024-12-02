# Define the IP address since DNS isn't resolving
$serverIP = '192.168.0.22'

# Test connectivity first
Write-Host "Testing connectivity..." -ForegroundColor Cyan
Test-NetConnection $serverIP -Port 5985

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
        $Cred = Get-Credential
        $Cred | Export-Clixml -Path $SecretsFilePath
        return $Cred
    }
}

# First, add the IP to TrustedHosts if not already there
$currentTrustedHosts = Get-Item WSMan:\localhost\Client\TrustedHosts
if ($currentTrustedHosts.Value -notlike "*$serverIP*") {
    Write-Host "Adding $serverIP to TrustedHosts..." -ForegroundColor Yellow
    Set-Item WSMan:\localhost\Client\TrustedHosts -Value $serverIP -Force
}

# Retrieve the credentials
$UserCredential = Get-ExchangeCredential -SecretsFilePath $SecretsFile

# Define the correct PowerShell endpoint URL
$ConnectionUri = "http://$serverIP/PowerShell-LiveID?PSVersion=5.1.20348.2849"

# Establish a new PowerShell session with the Exchange server
try {
    Write-Host "Attempting to establish PowerShell session..." -ForegroundColor Cyan
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange `
        -ConnectionUri $ConnectionUri `
        -Authentication Negotiate `
        -Credential $UserCredential `
        -ErrorAction Stop

    # Import the session
    Write-Host "Importing PowerShell session..." -ForegroundColor Cyan
    Import-PSSession $Session -DisableNameChecking -AllowClobber
    
    Write-Host "Successfully connected to Exchange server!" -ForegroundColor Green
} catch {
    Write-Host "Error connecting to Exchange server:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    
    # If first attempt fails, try with Basic authentication
    Write-Host "Attempting connection with Basic authentication..." -ForegroundColor Yellow
    try {
        $Session = New-PSSession -ConfigurationName Microsoft.Exchange `
            -ConnectionUri $ConnectionUri `
            -Authentication Basic `
            -AllowRedirection `
            -Credential $UserCredential `
            -ErrorAction Stop
        
        Import-PSSession $Session -DisableNameChecking -AllowClobber
        Write-Host "Successfully connected using Basic authentication!" -ForegroundColor Green
    } catch {
        Write-Host "All connection attempts failed. Please check your credentials and network connectivity." -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red
    }
}