# Check if the Exchange Online Management Module (EXO V2) is installed
$module = Get-Module -Name ExchangeOnlineManagement -ListAvailable

# If the module is not installed, install it from the PowerShell Gallery
if (-not $module) {
    # Install the module
    Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber -Scope CurrentUser
}

# Import the module
Import-Module ExchangeOnlineManagement

# Function to connect to Exchange Online
# function Connect-EXO {
#     # Check if a session is already active
#     $existingSession = Get-PSSession | Where-Object { $_.ConfigurationName -eq 'Microsoft.Exchange' -and $_.State -eq 'Opened' }
#     if ($existingSession) {
#         Write-Host "An Exchange Online session is already active." -ForegroundColor Cyan
#     } else {
#         # Connect to Exchange Online
#         $userCredential = Get-Credential
#         Connect-ExchangeOnline -Credential $userCredential -ShowProgress $true
        
#         Write-Host "Connected to Exchange Online." -ForegroundColor Green
#     }
# }

# # Connect to Exchange Online
# Connect-EXO

# Place your Exchange Online management commands below
# Example: List all mailboxes
# Get-Mailbox | Format-Table DisplayName, PrimarySmtpAddress

# Example command to run: Add mailbox permissions (Replace <MailboxIdentity> and <User> with actual values)
# Add-MailboxPermission -Identity "<MailboxIdentity>" -User "<User>" -AccessRights FullAccess -InheritanceType All -AutoMapping $false

Connect-ExchangeOnline


Add-MailboxPermission -Identity "SPLC@railcan.ca" -User "mbarfoot@railcan.ca" -AccessRights FullAccess -InheritanceType All -AutoMapping $true
Add-MailboxPermission -Identity "SPLC@railcan.ca" -User "bbowman@railcan.ca" -AccessRights FullAccess -InheritanceType All -AutoMapping $true
Add-MailboxPermission -Identity "SPLC@railcan.ca" -User "EDeBenetti@railcan.ca" -AccessRights FullAccess -InheritanceType All -AutoMapping $true