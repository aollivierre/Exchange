# Import the AzureAD module
Import-Module AzureAD

# Connect to Azure AD
# $UserCredential = Get-Credential
# Connect-AzureAD -Credential $UserCredential
# Connect-AzureAD

# Read the CSV file containing the list of users
$users = Import-Csv -Path "C:\Code\CB\Exchange\Glebe\Exports\Glebe-Migration_32Users_AD_Email_Alias.csv"

# Define the Group ID
$groupId = '62c10754-d84d-4049-9179-f4a4d9e0c0c4'

# For each user in the CSV file
foreach ($user in $users) {
    $emailAddress = $user.EmailAddress

    # Get the user from Azure AD
    $azureUser = Get-AzureADUser -ObjectId $emailAddress

    if ($azureUser) {
        # If the user exists, add the user to the group
        Add-AzureADGroupMember -ObjectId $groupId -RefObjectId $azureUser.ObjectId
        Write-Output "User $emailAddress was added to the group."
    } else {
        # If the user does not exist, print a message to the console
        Write-Output "User $emailAddress does not exist in Azure AD."
    }
}

# Disconnect from Azure AD
# Disconnect-AzureAD
