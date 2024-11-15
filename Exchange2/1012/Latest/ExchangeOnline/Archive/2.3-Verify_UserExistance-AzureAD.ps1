# Import the AzureAD module
# Install-Module AzureAD
# Import-Module AzureAD

# Connect to Azure AD
# $UserCredential = Get-Credential
# Connect-AzureAD -Credential $UserCredential
# Connect-AzureAD

# Read the CSV file containing the list of users
$users = Import-Csv -Path "C:\Code\CB\Exchange\Glebe\Exports\Glebe-Migration_32Users_AD_Email_Alias.csv"

# For each user in the CSV file
foreach ($user in $users) {
    $emailAddress = $user.EmailAddress

    try {
        # Attempt to get the user from Azure AD
        $azureUser = Get-AzureADUser -ObjectId $emailAddress 

        # If the user exists, print a message to the console
        Write-host "User $emailAddress exists in Azure AD." -ForegroundColor Green
    } 
    catch {
        # If the user does not exist, print a message to the console
        Write-host "User $emailAddress does not exist in Azure AD." -ForegroundColor Red
    }
}

# Disconnect from Azure AD
# Disconnect-AzureAD