# Import the necessary module
Import-Module ActiveDirectory


# Get all users from the Exchange server
$users = Get-ADUser -Filter * -Properties EmailAddress

# Create an empty array to hold the results
$results = @()

# Loop through each user
foreach ($user in $users) {
    # Get the mailbox type
    $mailboxType = if (Get-RemoteMailbox $user.SamAccountName) {"Office 365"} else {"Local"}

    # Check if the email domain matches the required one
    $emailDomainCheck = $user.EmailAddress -like "*@glebecentre.mail.onmicrosoft.com"

    # Color code output for console
    $color = if ($emailDomainCheck) {"Green"} else {"Red"}

    # Write to console
    Write-Host ("User: " + $user.SamAccountName + ", Email Domain Check: " + $emailDomainCheck + ", Mailbox Type: " + $mailboxType) -ForegroundColor $color

    # Add result to the array
    $results += New-Object PSObject -Property @{
        User = $user.SamAccountName
        EmailDomainCheck = $emailDomainCheck
        MailboxType = $mailboxType
    }
}

# Output results to grid view
$results | Out-GridView

# Export results to a CSV file
# $results | Export-Csv -Path "ExchangeUserReport.csv" -NoTypeInformation