# 1- Check if Exchange Online PS Module is installed. If not, install it
if (!(Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Write-Host "Installing Exchange Online Management Module..." -ForegroundColor Cyan
    Install-Module -Name ExchangeOnlineManagement -Force
    Write-Host "Installation of Exchange Online Management Module is complete!" -ForegroundColor Green
}

# 2- Import the module
Import-Module ExchangeOnlineManagement

# 3- Connect to Exchange Online
Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
# $UserCredential = Get-Credential
# Connect-ExchangeOnline -Credential $UserCredential -ShowBanner:$false
Connect-ExchangeOnline
Write-Host "Connection to Exchange Online is successful!" -ForegroundColor Green