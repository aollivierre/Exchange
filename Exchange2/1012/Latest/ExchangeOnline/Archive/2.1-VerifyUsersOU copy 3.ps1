# 1- Import AD module
Import-Module ActiveDirectory

# 2- Dynamically grab the domain name
$domainName = (Get-ADDomain).DNSRoot
Write-Host "Domain Name: $domainName" -ForegroundColor Green

# 3- Read CSV file
$csvFilePath = "C:\Code\CB\Exchange\Glebe\Exports\Glebe-Migration_20Users_AD_Email_Alias.csv" # Replace with the actual CSV file path
$csvFile = Import-Csv -Path $csvFilePath

# 4- List of users and their current OU
$userList = @()
$existingUsers = 0
$nonExistingUsers = 0

foreach ($item in $csvFile) {
    $userPrincipalName = $item.ALIAS
    $emailAddress = $item.'EMAILADDRESS' 

    # Search for the user in AD
    $user = Get-ADUser -Filter {UserPrincipalName -eq $userPrincipalName -or mail -eq $emailAddress} -Properties DistinguishedName, mail, proxyAddresses

    if ($user) {
        $currentOU = ($user.DistinguishedName -split ',',2)[1]
        $proxyAddresses = $user.proxyAddresses -join ", "

        $userList += New-Object PSObject -Property @{
            'UserPrincipalName' = $userPrincipalName
            'Email' = $emailAddress
            'CurrentOU' = $currentOU
            'ProxyAddresses' = $proxyAddresses
        }
        Write-Host "User $userPrincipalName/$emailAddress exists in Active Directory." -ForegroundColor Green
        $existingUsers++
    } else {
        Write-Host "User $userPrincipalName/$emailAddress does not exist in Active Directory." -ForegroundColor Red
        $nonExistingUsers++
    }
}

# 5- Display the list of users and their current OU in GridView
$userList | Out-GridView

# 6- Display the total number of users, existing users, and non-existing users
Write-Host "Total users in CSV: $($csvFile.Count)" -ForegroundColor Cyan
Write-Host "Total existing users in Active Directory: $existingUsers" -ForegroundColor Green
Write-Host "Total non-existing users in Active Directory: $nonExistingUsers" -ForegroundColor Red
