# Import the AzureAD module
Import-Module AzureAD

# Connect to Azure AD
# $UserCredential = Get-Credential
# Connect-AzureAD -Credential $UserCredential
# Connect-AzureAD

# Read the CSV file containing the list of users
$csvFile = Import-Csv -Path "C:\Code\CB\Exchange\Glebe\Exports\Glebe_Exchange_Remote_Move_10pilotusers_Import_UPNs.csv"

# Define the Group ID
$groupId = '62c10754-d84d-4049-9179-f4a4d9e0c0c4'

$existingUsers = 0
$nonExistingUsers = 0
$groupMembersCount = 0

# Get the current members of the group
$currentGroupMembers = Get-AzureADGroupMember -ObjectId $groupId | Select-Object ObjectId

# For each user in the CSV file
foreach ($item in $csvFile) {
    $userPrincipalName = $item.ALIAS
    $emailAddress = $item.'EmailAddress'

    # Get the user from Azure AD
    $azureUser = Get-AzureADUser -SearchString $userPrincipalName

    if ($azureUser) {
        # If the user exists, add the user to the group if not already a member
        if ($currentGroupMembers.ObjectId -notcontains $azureUser.ObjectId) {
            Add-AzureADGroupMember -ObjectId $groupId -RefObjectId $azureUser.ObjectId
            Write-Output "User $userPrincipalName was added to the group."
            $groupMembersCount++
        } else {
            Write-Output "User $userPrincipalName is already a member of the group."
        }
        $existingUsers++
    } else {
        $azureUserByEmail = Get-AzureADUser -SearchString $emailAddress
        if($azureUserByEmail){
            if ($currentGroupMembers.ObjectId -notcontains $azureUserByEmail.ObjectId) {
                Add-AzureADGroupMember -ObjectId $groupId -RefObjectId $azureUserByEmail.ObjectId
                Write-Output "User $emailAddress was added to the group."
                $groupMembersCount++
            } else {
                Write-Output "User $emailAddress is already a member of the group."
            }
            $existingUsers++
        } else {
            # If the user does not exist, print a message to the console
            Write-Output "User $userPrincipalName/$emailAddress does not exist in Azure AD."
            $nonExistingUsers++
        }
    }
}

# Print stats
Write-Host "Total users in CSV: $($csvFile.Count)" -ForegroundColor Cyan
Write-Host "Total existing users in Azure AD: $existingUsers" -ForegroundColor Green
Write-Host "Total non-existing users in Azure AD: $nonExistingUsers" -ForegroundColor Red
Write-Host "Total users added to the group: $groupMembersCount" -ForegroundColor Green

# Disconnect from Azure AD
# Disconnect-AzureAD
