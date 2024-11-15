# 1- Import AD module
Import-Module ActiveDirectory

# 2- Dynamically grab the domain name
$domainName = (Get-ADDomain).DNSRoot
Write-Host "Domain Name: $domainName" -ForegroundColor Green

# 3- Create a container in AD called PendingMigrationtoExchangeOnline
$ouPath = "OU=PendingMigrationtoExchangeOnline,DC=" + $domainName.Replace('.', ',DC=')
if (!(Get-ADOrganizationalUnit -Filter {DistinguishedName -eq $ouPath})) {
    New-ADOrganizationalUnit -Name "PendingMigrationtoExchangeOnline" -Path ("DC=" + $domainName.Replace('.', ',DC='))
    Write-Host "Created Organizational Unit: PendingMigrationtoExchangeOnline" -ForegroundColor Green
}

# 4- Read CSV file
$csvFilePath = "C:\Code\CB\Exchange\Glebe\Exports\Glebe-Migration_32Users_AD_Email_Alias.csv" # Replace with the actual CSV file path
$csvFile = Import-Csv -Path $csvFilePath

# 5- List of users and their current OU
$existingUsers = 0
$nonExistingUsers = 0

foreach ($item in $csvFile) {
    $userPrincipalName = $item.ALIAS
    $emailAddress = $item.'EMAILADDRESS'

    # Search for the user in AD
    $user = Get-ADUser -Filter {UserPrincipalName -eq $userPrincipalName -or mail -eq $emailAddress} -Properties DistinguishedName, mail

    if ($user) {
        # Move user to the new OU
        Move-ADObject -Identity $user.DistinguishedName -TargetPath $ouPath
        Write-Host "User $userPrincipalName/$emailAddress exists in Active Directory and has been moved to $ouPath." -ForegroundColor Green
        $existingUsers++
    } else {
        Write-Host "User $userPrincipalName/$emailAddress does not exist in Active Directory." -ForegroundColor Red
        $nonExistingUsers++
    }
}

# 6- Display the total number of users, existing users, and non-existing users
Write-Host "Total users in CSV: $($csvFile.Count)" -ForegroundColor Cyan
Write-Host "Total existing users in Active Directory and moved to new OU: $existingUsers" -ForegroundColor Green
Write-Host "Total non-existing users in Active Directory: $nonExistingUsers" -ForegroundColor Red

# 7- List all of the users in the new OU
$usersAfterMove = Get-ADUser -Filter * -SearchBase $ouPath
Write-Host "$(Get-Date) - [INFO] Users in 'PendingMigrationtoExchangeOnline' OU after move:" -ForegroundColor Cyan
$usersAfterMove | ForEach-Object { Write-Host $_.UserPrincipalName }

Write-Host "$(Get-Date) - [INFO] Total users in 'PendingMigrationtoExchangeOnline' OU: $($usersAfterMove.Count)" -ForegroundColor Green
