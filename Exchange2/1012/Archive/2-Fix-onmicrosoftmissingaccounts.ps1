# $ErrorActionPreference = 'SilentlyContinue'
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

                if ($emailAddresses -notlike "*@glebecentre.mail.onmicrosoft.com*") {
                    $newAddress = $user.UserPrincipalName + "@glebecentre.mail.onmicrosoft.com"
                    Set-RemoteMailbox -Identity $user.UserPrincipalName -EmailAddresses @{Add=$newAddress}
                    Write-Host "$(Get-Date) - [INFO] Added $newAddress to $user.UserPrincipalName's remote mailbox." -ForegroundColor Green
                }
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

                if ($emailAddresses -notlike "*@glebecentre.mail.onmicrosoft.com*") {
                    $newAddress = $user.UserPrincipalName + "@glebecentre.mail.onmicrosoft.com"
                    Set-Mailbox -Identity $user.UserPrincipalName -EmailAddresses @{Add=$newAddress}
                    Write-Host "$(Get-Date) - [INFO] Added $newAddress to $user.UserPrincipalName's local mailbox." -ForegroundColor Green
                }
            }
        } catch {
            Write-Host "$(Get-Date) - [WARNING] Could not find local mailbox for user $($user.UserPrincipalName)" -ForegroundColor Yellow
        }
    }

    if ($emailAddresses) {
        $emailAddressCheck = $emailAddresses -like "*@glebecentre.mail.onmicrosoft.com*"
    } else {
        $emailAddressCheck = @() 
    }

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

$results = $results | Sort-Object -Property MailboxType, EmailAddressCheck -Descending
$totalUsers = $users.Count
$totalUsersWithEmailAddresses = $results.Count
$totalRemote = ($results | Where-Object { $_.MailboxType -eq "Remote" }).Count
$totalLocal = ($results | Where-Object { $_.MailboxType -eq "Local" }).Count
$totalPassedEmailCheck = ($results | Where-Object { $_.EmailAddressCheck.Count -eq 1 }).Count
$totalFailedEmailCheck = ($results | Where-Object { $_.EmailAddressCheck.Count -eq 0 -or $_.EmailAddressCheck -eq $false }).Count
$totalWithGlebeCentreEmail = ($results | Where-Object { $_.EmailAddressCheck -like "*@glebecentre.mail.onmicrosoft.com*" }).Count
$totalWithoutGlebeCentreEmail = ($results | Where-Object { $_.EmailAddressCheck.Count -eq 0 -or $_.EmailAddressCheck -eq $false }).Count

Write-Host "$(Get-Date) - [INFO] Total Users in Exchange: $totalUsers" -ForegroundColor Green
Write-Host "$(Get-Date) - [INFO] Total Users with email addresses: $totalUsersWithEmailAddresses" -ForegroundColor Green
Write-Host "$(Get-Date) - [INFO] Total Remote Users: $totalRemote" -ForegroundColor Green
Write-Host "$(Get-Date) - [INFO] Total Local mailboxes: $totalLocal" -ForegroundColor Green
Write-Host "$(Get-Date) - [INFO] Total Users Passed Email Check: $totalPassedEmailCheck" -ForegroundColor Green
Write-Host "$(Get-Date) - [INFO] Total Users Failed Email Check: $totalFailedEmailCheck" -ForegroundColor Green
Write-Host "$(Get-Date) - [INFO] Total Users with '@glebecentre.mail.onmicrosoft.com': $totalWithGlebeCentreEmail" -ForegroundColor Green
Write-Host "$(Get-Date) - [INFO] Total Users without '@glebecentre.mail.onmicrosoft.com': $totalWithoutGlebeCentreEmail" -ForegroundColor Green

$mailboxTypeStats = $results | Group-Object -Property RecipientTypeDetails
foreach ($mailboxType in $mailboxTypeStats) {
    Write-Host "$(Get-Date) - [INFO] Total $($mailboxType.Name) Mailboxes: $($mailboxType.Count)" -ForegroundColor Green
}

$results | Out-GridView
