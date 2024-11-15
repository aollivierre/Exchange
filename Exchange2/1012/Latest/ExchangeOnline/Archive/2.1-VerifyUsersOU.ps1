# 1- Import AD module
Import-Module ActiveDirectory

# 2- Dynamically grab the domain name
$domainName = (Get-ADDomain).DNSRoot
Write-Host "Domain Name: $domainName" -ForegroundColor Green

# 3- Read CSV file
$csvFilePath = "C:\Code\CB\Exchange\Glebe\Exports\Glebe_Exchange_Remote_Move_32pilotusers_Import_AD.csv" # Replace with the actual CSV file path
$csvFile = Import-Csv -Path $csvFilePath

# 4- List of users and their current OU
$userList = @()
$existingUsers = 0
$nonExistingUsers = 0

foreach ($item in $csvFile) {
    $userPrincipalName = $item.UserPrincipalName
    $user = Get-ADUser -Filter {UserPrincipalName -eq $userPrincipalName} -Properties DistinguishedName

    if ($user) {
        $currentOU = ($user.DistinguishedName -split ',',2)[1]
        $userList += New-Object PSObject -Property @{
            'UserPrincipalName' = $userPrincipalName
            'CurrentOU' = $currentOU
        }
        Write-Host "User $userPrincipalName exists in AD." -ForegroundColor Green
        $existingUsers++
    } else {
        Write-Host "User $userPrincipalName does not exist in AD." -ForegroundColor Red
        $nonExistingUsers++
    }
}

# 5- Display the list of users and their current OU in GridView
$userList | Out-GridView

# 6- Display the total number of users, existing users and non-existing users
Write-Host "Total users in CSV: $($csvFile.Count)" -ForegroundColor Cyan
Write-Host "Total existing users in AD: $existingUsers" -ForegroundColor Green
Write-Host "Total non-existing users in AD: $nonExistingUsers" -ForegroundColor Red
