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
                $mailboxType = 'Local'
                $recipientTypeDetails = $localMailbox.RecipientTypeDetails
            }
        } catch {
            Write-Host "$(Get-Date) - [WARNING] Could not find local mailbox for user $($user.UserPrincipalName)" -ForegroundColor Yellow
        }
    }

    # Set EmailAddressCheck based on email address presence
    if ($emailAddresses) {
        $emailAddressCheck = $emailAddresses -like "*@glebecentre.mail.onmicrosoft.com*"
    } else {
        $emailAddressCheck = $null  # or set it to any desired value for accounts without email addresses
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
# $uniqueEmailChecks = $results.EmailAddressCheck | Select-Object -Unique

Write-Host "$(Get-Date) - [INFO] Total Users in Exchange: $($users.Count)" -ForegroundColor Green
Write-Host "$(Get-Date) - [INFO] Total Users with email addresses: $($results.Count)" -ForegroundColor Green
Write-Host "$(Get-Date) - [INFO] Total Remote Users: $($results | Where-Object { $_.MailboxType -eq 'Remote' }).Count" -ForegroundColor Green
Write-Host "$(Get-Date) - [INFO] Total Local Users: $($results | Where-Object { $_.MailboxType -eq 'Local' }).Count" -ForegroundColor Green
Write-Host "$(Get-Date) - [INFO] Total Users Passed Email Check: $($results | Where-Object { $_.EmailAddressCheck }).Count" -ForegroundColor Green
Write-Host "$(Get-Date) - [INFO] Total Users Did Not Pass Email Check: $($results | Where-Object { -not $_.EmailAddressCheck }).Count" -ForegroundColor Green
Write-Host "$(Get-Date) - [INFO] Total Users with '@glebecentre.mail.onmicrosoft.com': $($results | Where-Object { $_.EmailAddressCheck -like '*@glebecentre.mail.onmicrosoft.com*' }).Count" -ForegroundColor Green
Write-Host "$(Get-Date) - [INFO] Total Users without '@glebecentre.mail.onmicrosoft.com': $($results | Where-Object { $_.EmailAddressCheck -notlike '*@glebecentre.mail.onmicrosoft.com*' }).Count" -ForegroundColor Green

# Write-Host "Unique EmailAddressCheck values:"
# $uniqueEmailChecks | ForEach-Object { Write-Host $_ }

# Group by mailbox type and count each type
$mailboxTypeStats = $results | Group-Object -Property RecipientTypeDetails
foreach ($mailboxType in $mailboxTypeStats) {
    Write-Host "$(Get-Date) - [INFO] Total $($mailboxType.Name) Mailboxes: $($mailboxType.Count)" -ForegroundColor Green
}

$results | Out-GridView
