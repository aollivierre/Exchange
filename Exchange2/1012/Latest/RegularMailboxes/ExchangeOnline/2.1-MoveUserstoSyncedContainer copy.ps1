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

# 4- Move all users listed in CSV to this container (because they were not synced and hence Exchange Online does not know about them and thus we need to sync them first to be able to migrate them from exchange on-prem to exchange online)
$csvFilePath = "C:\Code\CB\Exchange\Glebe\Exports\Glebe_Exchange_Remote_Move_32pilotusers_Import_AD.csv" # Replace with the actual CSV file path
$csvFile = Import-Csv -Path $csvFilePath

$csvFile | ForEach-Object {
    $userPrincipalName = $_.UserPrincipalName
    $user = Get-ADUser -Filter {UserPrincipalName -eq $userPrincipalName}
    if ($user) {
        Move-ADObject -Identity $user.DistinguishedName -TargetPath $ouPath
    }
}

# 5- List all of the users in the CSV before the move and all of the number of users in the container after the move color coded time stamped
Write-Host "$(Get-Date) - [INFO] Users in CSV file before move:" -ForegroundColor Cyan
$csvFile | ForEach-Object { Write-Host $_.UserPrincipalName }

$usersAfterMove = Get-ADUser -Filter * -SearchBase $ouPath
Write-Host "$(Get-Date) - [INFO] Users in 'PendingMigrationtoExchangeOnline' OU after move:" -ForegroundColor Cyan
$usersAfterMove | ForEach-Object { Write-Host $_.UserPrincipalName }

Write-Host "$(Get-Date) - [INFO] Total users in CSV: $($csvFile.Count)" -ForegroundColor Green
Write-Host "$(Get-Date) - [INFO] Total users in 'PendingMigrationtoExchangeOnline' OU: $($usersAfterMove.Count)" -ForegroundColor Green
