# 1- Check if Azure AD PS Module is installed. If not, install it
if (!(Get-Module -ListAvailable -Name AzureAD)) {
    Write-Host "Installing Azure AD Module..." -ForegroundColor Cyan
    Install-Module -Name AzureAD -Force -AllowClobber
    Write-Host "Installation of Azure AD Module is complete!" -ForegroundColor Green
}

# 2- Import the module
Import-Module AzureAD

# 3- Connect to Azure AD
Write-Host "Connecting to Azure AD..." -ForegroundColor Cyan
# If you need to specify credentials:
# $UserCredential = Get-Credential
# Connect-AzureAD -Credential $UserCredential
Connect-AzureAD
Write-Host "Connection to Azure AD is successful!" -ForegroundColor Green
