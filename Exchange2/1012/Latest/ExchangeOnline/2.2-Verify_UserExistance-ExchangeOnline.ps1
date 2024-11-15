$ErrorActionPreference = 'SilentlyContinue'

# Import the EXOv2 module
# Import-Module ExchangeOnlineManagement

# Connect to Exchange Online
# $UserCredential = Get-Credential
# Connect-ExchangeOnline -Credential $UserCredential


#The following should show that users do not exist before migration which is expected
# Read the CSV file containing the list of users
$users = Import-Csv -Path "C:\Code\CB\Exchange\Glebe\Exports\Glebe_Exchange_Migration_Last_batch_7users_June_14_2023.csv"

# For each user in the CSV file
foreach ($user in $users) {
    $emailAddress = $user.EmailAddress  # Replace "Email" with the name of the column that contains the user's email address

    # Attempt to get the user from Exchange Online
    $exchangeUser = Get-ExoMailbox -Identity $emailAddress

    if ($exchangeUser) {
        # If the user exists, print a message to the console
        Write-host "User $emailAddress exists in Exchange Online." -ForegroundColor Green
    } else {
        # If the user does not exist, print a message to the console
        Write-host "User $emailAddress does not exist in Exchange Online." -ForegroundColor red
    }
}

# Disconnect from Exchange Online
# Disconnect-ExchangeOnline -Confirm:$false
