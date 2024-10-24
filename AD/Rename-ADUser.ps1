# Update the Name property of the user
$existingUser = Get-ADUser -Filter {UserPrincipalName -eq "JHeuser@arnpriorhealth.ca"}
if ($existingUser) {
    Rename-ADObject -Identity $existingUser.DistinguishedName -NewName "Jordan Heuser"
    Write-Host "Updated Name to Jordan Heuser." -ForegroundColor Green

    # Update the DisplayName property
    Set-ADUser -Identity $existingUser -DisplayName "Jordan Heuser"
    Write-Host "Updated DisplayName to Jordan Heuser." -ForegroundColor Green
} else {
    Write-Host "User not found." -ForegroundColor Red
}
