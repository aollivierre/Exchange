# Define the user details
$userDetails = @{
    GivenName           = "Jordan"
    Surname             = "Heuser"
    SamAccountName      = "JHeuser-old"
    UserPrincipalName   = "JHeuser-old@arnpriorhealth.ca"
    Name                = "Jordan Heuser-old"
    Path                = "OU=ADMH,OU=User Accounts,OU=Users,OU=ARH,DC=admh,DC=arnpriorhospital,DC=com"
    AccountPassword     = ConvertTo-SecureString "P@ssword123" -AsPlainText -Force
    Enabled             = $true
    PasswordNeverExpires = $true
    DisplayName         = "Jordan Heuser-old"
}

$emailDetails = @{
    Identity           = "JHeuser-old"
    EmailAddress       = "JHeuser-old@arnpriorhealth.ca"
    UserPrincipalName  = "JHeuser-old@arnpriorhealth.ca"
    DisplayName        = "Jordan Heuser-old"
}

# Create the new user
New-ADUser @userDetails -PassThru

# Set additional user properties
Set-ADUser @emailDetails

Write-Host "User $($userDetails.DisplayName) with UPN $($userDetails.UserPrincipalName) has been created and configured." -ForegroundColor Green
