# Load the required assembly for UserPrincipal
Add-Type -AssemblyName "System.DirectoryServices.AccountManagement"

# Define variables
$userName = "NewLocalAdmin9"
$password = ConvertTo-SecureString "ENTER your Password here" -AsPlainText -Force
$localAdminGroup = "Administrators"

# Create a local account
New-LocalUser -Name $userName -Password $password -FullName "Local Administrator Account" -Description "Local admin account created by PowerShell script"

# Add the account to the local Administrators group
Add-LocalGroupMember -Group $localAdminGroup -Member $userName

# Set the account's password to never expire
Set-LocalUser -Name $userName -PasswordNeverExpires $true

# Disable the requirement for the user to change their password on their next login
$contextType = [System.DirectoryServices.AccountManagement.ContextType]::Machine
$principalContext = New-Object -TypeName System.DirectoryServices.AccountManagement.PrincipalContext -ArgumentList $contextType

$userPrincipal = [System.DirectoryServices.AccountManagement.UserPrincipal]::FindByIdentity($principalContext, $userName)
$userPrincipal.UserCannotChangePassword = $true
$userPrincipal.Save()
