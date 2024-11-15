# Import the AzureAD module
# Import-Module AzureAD

# Connect to Azure AD
Connect-AzureAD

# Read the CSV file containing the list of users
$csvFile = Import-Csv -Path "C:\Code\CB\Exchange\Glebe\Exports\Glebe_Exchange_Remote_Move_5pilotusers_Import_UPNs.csv"

# Define the Group ID
$groupId = '62c10754-d84d-4049-9179-f4a4d9e0c0c4'

# Retrieve the members of the group
$groupMembers = Get-AzureADGroupMember -ObjectId $groupId | Select-Object ObjectId

$existingUsers = 0
$nonExistingUsers = 0
$groupMembersCount = 0

# For each user in the CSV file
foreach ($item in $csvFile) {
    $userPrincipalName = $item.ALIAS
    $emailAddress = $item.'EmailAddress'

    try {
        # Get the user from Azure AD
        $azureUser = Get-AzureADUser -SearchString $userPrincipalName

        # If the user exists, check if they are a member of the group
        if ($azureUser){
            if ($groupMembers.ObjectId -contains $azureUser.ObjectId) {
                Write-Host "User $userPrincipalName is a member of the group." -ForegroundColor Green
                $groupMembersCount++
            } else {
                Write-Host "User $userPrincipalName is not a member of the group." -ForegroundColor Red
            }
            $existingUsers++
        } else {
            $azureUserByEmail = Get-AzureADUser -SearchString $emailAddress
            if($azureUserByEmail){
                if ($groupMembers.ObjectId -contains $azureUserByEmail.ObjectId) {
                    Write-Host "User $emailAddress is a member of the group." -ForegroundColor Green
                    $groupMembersCount++
                } else {
                    Write-Host "User $emailAddress is not a member of the group." -ForegroundColor Red
                }
                $existingUsers++
            } else {
                Write-host "User $userPrincipalName/$emailAddress does not exist in Azure AD." -ForegroundColor Red
                $nonExistingUsers++
            }
        }
    }
    catch {
        # If the user does not exist, print a message to the console
        Write-host "Error occurred while searching for user $userPrincipalName/$emailAddress in Azure AD." -ForegroundColor Red
    }
}

# Calculate the number of users who are not in the group
$nonGroupMembersCount = $existingUsers - $groupMembersCount

# Print stats
Write-Host "Total users in CSV: $($csvFile.Count)" -ForegroundColor Cyan
Write-Host "Total existing users in Azure AD: $existingUsers" -ForegroundColor Green
Write-Host "Total non-existing users in Azure AD: $nonExistingUsers" -ForegroundColor Red
Write-Host "Total users in the group: $groupMembersCount" -ForegroundColor Green
Write-Host "Total users not in the group: $nonGroupMembersCount" -ForegroundColor Red

# Disconnect from Azure AD
# Disconnect-AzureAD
