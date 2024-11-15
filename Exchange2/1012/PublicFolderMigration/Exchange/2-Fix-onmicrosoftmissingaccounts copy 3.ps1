$users = Import-Csv -Path 'C:\Code\CB\Exchange\Exports\failed_email_check_users.csv'

foreach ($user in $users) {
    $userPrincipalName = $user.UserPrincipalName
    $username = $userPrincipalName.Split('@')[0]

    $mailbox = Get-Mailbox -Identity $userPrincipalName -ErrorAction SilentlyContinue

    if ($mailbox) {
        $emailAddresses = $mailbox.EmailAddresses

        if ($emailAddresses -notcontains "smtp:$username@glebecentre.mail.onmicrosoft.com") {
            $emailAddresses += "smtp:$username@glebecentre.mail.onmicrosoft.com"

            Set-Mailbox -Identity $userPrincipalName -EmailAddresses $emailAddresses
            Write-Host "Added '@glebecentre.mail.onmicrosoft.com' to the email addresses of $username."
        } else {
            Write-Host "$username already has '@glebecentre.mail.onmicrosoft.com' in their email addresses."
        }
    } else {
        Write-Host "Could not find a mailbox for $userPrincipalName."
    }
}
