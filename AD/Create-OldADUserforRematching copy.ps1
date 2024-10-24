# Define the user details
$userDetails = @{
    GivenName           = "Khushpreet"
    Surname             = "Kaur"
    SamAccountName      = "KKaur"
    UserPrincipalName   = "KKaur@contoso.com"
    Name                = "Khushpreet Kaur"
    Path                = "OU=ADMH,OU=User Accounts,OU=Users,OU=ARH,DC=admh,DC=arnpriorhospital,DC=com"
    AccountPassword     = ConvertTo-SecureString "P@ssword123" -AsPlainText -Force
    Enabled             = $true
    PasswordNeverExpires = $true
    DisplayName         = "Khushpreet Kaur"
}

$emailDetails = @{
    Identity           = "KKaur"
    EmailAddress       = "KKaur@contoso.com"
    UserPrincipalName  = "KKaur@contoso.com"
    DisplayName        = "Khushpreet Kaur"
}

# Create the new user
New-ADUser @userDetails -PassThru

# Set additional user properties
Set-ADUser @emailDetails

Write-Host "User $($userDetails.DisplayName) with UPN $($userDetails.UserPrincipalName) has been created and configured." -ForegroundColor Green
