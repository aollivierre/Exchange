# Import the Active Directory module
Import-Module ActiveDirectory

# Prompt for the User Principal Name
# $upn = Read-Host "Please enter the User Principal Name (UPN)"
$upn = "vkandhasamy"

# Retrieve the user from Active Directory
$user = Get-ADUser -Identity $upn -Properties *

# Display all properties in a grid view
$user | Out-GridView
$user