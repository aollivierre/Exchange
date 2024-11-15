Install-Module -Name AzureAD -Force -Scope AllUsers
Install-Module -Name AzureAD.Standard.Preview
Import-Module AzureAD
Import-Module AzureAD.Standard.Preview

connect-azuread


# Find contacts with the conflicting SMTP address
$conflictingProxyAddress = "SMTP:mkruze@baltecusa.com"
Get-AzureADContact -All $true | Where-Object { $_.ProxyAddresses -contains $conflictingProxyAddress } | Select-Object DisplayName, ObjectId, ProxyAddresses

# DisplayName ObjectId                             ProxyAddresses
# ----------- --------                             --------------
# Ed Stots    400af537-f20d-4349-8f3d-af20aa503d4b {x500:/o=ExchangeLabs/ou=Exchange Administrative Group (FYDIBOHF23SPDLT)/cn=Recipients/cn=16f70d031bae429e83f0480ef67f1507-EdStots, X500:/O=Delta Environmen..


# Object ID of the contact to update
# $objectIdToUpdate = "<ObjectID of Ed Stots>"
$objectIdToUpdate = "400af537-f20d-4349-8f3d-af20aa503d4b"

# Retrieve the current proxy addresses
$contact = Get-AzureADContact -ObjectId $objectIdToUpdate
Remove-AzureADContact -ObjectId $Contact.ObjectId