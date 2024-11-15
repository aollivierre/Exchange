# Import the AzureAD module
# Import-Module AzureAD

# Connect to Azure AD
# Connect-AzureAD

# Read the CSV file containing the list of users
$users = Import-Csv -Path "C:\Code\CB\Exchange\Glebe\Exports\Glebe_Exchange_Remote_Move_32pilotusers_Import.csv"

# Define the Group ID
$groupId = '62c10754-d84d-4049-9179-f4a4d9e0c0c4'

# Retrieve the members of the group
$groupMembers = Get-AzureADGroupMember -ObjectId $groupId | Select-Object ObjectId

# For each user in the CSV file
foreach ($user in $users) {
    $emailAddress = $user.EmailAddress

    # Get the user from Azure AD
    $azureUser = Get-AzureADUser -ObjectId $emailAddress

    if ($azureUser) {
        # If the user exists, check if they are a member of the group
        if ($groupMembers.ObjectId -contains $azureUser.ObjectId) {
            Write-host "User $emailAddress is a member of the group." -ForegroundColor Green
        } else {
            Write-host "User $emailAddress is not a member of the group." -ForegroundColor red
        }
    } else {
        # If the user does not exist, print a message to the console
        Write-host "User $emailAddress does not exist in Azure AD." -ForegroundColor red
    }
}

# Disconnect from Azure AD
# Disconnect-AzureAD
