$ErrorActionPreference = 'SilentlyContinue'
$users = Get-User -ResultSize Unlimited

$results = foreach ($user in $users) {
    $mailboxType = $null
    $emailAddresses = $null
    $emailAddressCheck = $null
    $recipientTypeDetails = $null
    $isRemoteUser = $null -ne (Get-Recipient $user.UserPrincipalName | Where-Object {$_.RecipientTypeDetails -eq "RemoteUserMailbox"})
    if ($isRemoteUser) {
        try {
            $remoteMailbox = Get-RemoteMailbox -Identity $user.UserPrincipalName -ErrorAction SilentlyContinue
            if ($remoteMailbox.EmailAddresses) {
                $emailAddresses = $remoteMailbox.EmailAddresses | Where-Object { $_ -like "*SMTP:*" } | ForEach-Object { $_.ToString().TrimStart('SMTP:') }
                $emailAddressCheck = $emailAddresses -like "*@glebecentre.mail.onmicrosoft.com*"
                $mailboxType = 'Remote'
                $recipientTypeDetails = $remoteMailbox.RecipientTypeDetails
            }
        } catch {
            Write-Host "$(Get-Date) - [WARNING] Could not find remote mailbox for user $($user.UserPrincipalName)" -ForegroundColor Yellow
        }
    } else {
        try {
            $localMailbox = Get-Mailbox -Identity $user.UserPrincipalName -ErrorAction SilentlyContinue
            if ($localMailbox.EmailAddresses) {
                $emailAddresses = $localMailbox.EmailAddresses | Where-Object { $_ -like "*SMTP:*" } | ForEach-Object { $_.ToString().TrimStart('SMTP:') }
                $emailAddressCheck = $emailAddresses -like "*@glebecentre.mail.onmicrosoft.com*"
                $mailboxType = 'Local'
                $recipientTypeDetails = $localMailbox.RecipientTypeDetails
            }
        } catch {
            Write-Host "$(Get-Date) - [WARNING] Could not find local mailbox for user $($user.UserPrincipalName)" -ForegroundColor Yellow
        }
    }

    # Only process users with email addresses
    if ($emailAddresses) {
        $userProperties = @{
            UserPrincipalName    = $user.UserPrincipalName
            EmailAddresses       = $emailAddresses -join ', '
            EmailAddressCheck    = $emailAddressCheck
            MailboxType          = $mailboxType
            RecipientTypeDetails = $recipientTypeDetails
        }
        New-Object PsObject -Property $userProperties
    }
}

# Get all unique EmailAddressCheck values
$uniqueEmailChecks = $results.EmailAddressCheck | Select-Object -Unique

Write-Host "Unique EmailAddressCheck values:"
$uniqueEmailChecks | ForEach-Object { Write-Host $_ }
