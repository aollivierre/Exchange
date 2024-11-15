# Import the AzureAD module
# Install-Module AzureAD
# Import-Module AzureAD

# Connect to Azure AD
# $UserCredential = Get-Credential
# Connect-AzureAD -Credential $UserCredential
# Connect-AzureAD

# Read the CSV file containing the list of users
$csvFile = Import-Csv -Path "C:\Code\CB\Exchange\LHC\Exports\LHC-Migration_21Users_AD_Email_Alias.csv"

$existingUsers = 0
$nonExistingUsers = 0

# For each user in the CSV file
foreach ($item in $csvFile) {
    $userPrincipalName = $item.ALIAS
    $emailAddress = $item.'EMAILADDRESS'

    try {
        # Attempt to get the user from Azure AD
        $azureUser = Get-AzureADUser -SearchString $userPrincipalName

        # If the user exists, print a message to the console
        if($azureUser){
            Write-Host "User $userPrincipalName exists in Azure AD." -ForegroundColor Green
            $existingUsers++
        } else {
            $azureUserByEmail = Get-AzureADUser -SearchString $emailAddress
            if($azureUserByEmail){
                Write-Host "User $emailAddress exists in Azure AD." -ForegroundColor Green
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

# Print stats
Write-Host "Total users in CSV: $($csvFile.Count)" -ForegroundColor Cyan
Write-Host "Total existing users in Azure AD: $existingUsers" -ForegroundColor Green
Write-Host "Total non-existing users in Azure AD: $nonExistingUsers" -ForegroundColor Red

# Disconnect from Azure AD
# Disconnect-AzureAD