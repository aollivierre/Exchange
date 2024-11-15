# Import Active Directory Module
Import-Module ActiveDirectory

# Define the email address to search in proxyAddresses
$emailAddress = "dyoon@chfcanada.coop"

# Search for AD objects with the specified email address in their proxyAddresses
$adObjects = Get-ADObject -Filter 'proxyAddresses -like "*$emailAddress*"' -Properties proxyAddresses

# Check if any objects were found
if ($null -ne $adObjects) {
    Write-Host "Found AD object(s) with the specified email in proxyAddresses:" -ForegroundColor Green
    foreach ($obj in $adObjects) {
        Write-Host "Object Name: $($obj.Name)" -ForegroundColor Yellow
        Write-Host "Distinguished Name: $($obj.DistinguishedName)" -ForegroundColor Cyan
        # Display all proxyAddresses for the object
        foreach ($address in $obj.proxyAddresses) {
            Write-Host "Proxy Address: $address" -ForegroundColor White
        }
    }
} else {
    Write-Host "No AD objects found with the specified email in proxyAddresses." -ForegroundColor Red
}
