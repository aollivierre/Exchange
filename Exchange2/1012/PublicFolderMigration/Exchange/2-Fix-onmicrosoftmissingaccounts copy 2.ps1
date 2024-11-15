# Import user list from CSV
$users = Import-Csv -Path 'C:\Code\CB\Exchange\Exports\failed_email_check_users.csv'

foreach ($user in $users) {
    $userPrincipalName = $user.UserPrincipalName

    # Extract the username part from the UPN
    $username = $userPrincipalName.Split('@')[0]

    # Get mailbox
    $mailbox = Get-Mailbox -Identity $userPrincipalName -ErrorAction SilentlyContinue

    if ($mailbox) {
        # Check if the mailbox has the @glebecentre.mail.onmicrosoft.com SMTP address
        $hasGlebeCentreEmail = $mailbox.EmailAddresses -like "SMTP:*@glebecentre.mail.onmicrosoft.com"

        if (-not $hasGlebeCentreEmail) {
            # Add @glebecentre.mail.onmicrosoft.com SMTP address
            $newAddress = "SMTP:$username@glebecentre.mail.onmicrosoft.com"
            Set-Mailbox -Identity $userPrincipalName -EmailAddresses @{Add=$newAddress}
            Write-Host "Added $newAddress to $userPrincipalName's email addresses."
        }
    } else {
        Write-Warning "Could not find mailbox for user $userPrincipalName."
    }
}
