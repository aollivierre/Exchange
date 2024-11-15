# Import user list from CSV
$users = Import-Csv -Path 'C:\Code\CB\Exchange\Glebe\Exports\failed_email_check_users.csv'

foreach ($user in $users) {
    $userPrincipalName = $user.UserPrincipalName

    # Get mailbox
    $mailbox = Get-Mailbox -Identity $userPrincipalName -ErrorAction SilentlyContinue

    if ($mailbox) {
        # Check if the mailbox has the @glebecentre.mail.onmicrosoft.com SMTP address
        $hasGlebeCentreEmail = $mailbox.EmailAddresses -like "SMTP:*@glebecentre.mail.onmicrosoft.com"

        if (-not $hasGlebeCentreEmail) {
            # Add @glebecentre.mail.onmicrosoft.com SMTP address
            $newAddress = "SMTP:$userPrincipalName@glebecentre.mail.onmicrosoft.com"
            Set-Mailbox -Identity $userPrincipalName -EmailAddresses @{Add=$newAddress}
            Write-Host "Added $newAddress to $userPrincipalName's email addresses."
        }
    } else {
        Write-Warning "Could not find mailbox for user $userPrincipalName."
    }
}
