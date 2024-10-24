# Load the required assembly for UserPrincipal
Add-Type -AssemblyName "System.DirectoryServices.AccountManagement"

# Import Active Directory module
Import-Module ActiveDirectory

# Define variables
$userName = "AOllivierre_Admin"
$password = ConvertTo-SecureString "ENTER your Password here" -AsPlainText -Force
$group = "Domain Admins"

# Create an Active Directory user account
New-ADUser -Name $userName -AccountPassword $password -Enabled $true -PasswordNeverExpires $true -CannotChangePassword $true -DisplayName "Local Administrator Account" -Description "Domain admin account created by PowerShell script"

# Add the user to the specified group
try {
    Add-ADGroupMember -Identity $group -Members $userName
} catch {
    Write-Host "Error: Failed to add user to the $group group."
}