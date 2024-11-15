$users = $null
$users = Get-User -ResultSize Unlimited

$results = $null
$results = foreach($user in $users){
    $mailboxType = $null
    $emailAddresses = $null
    $emailAddressCheck = $null
    $isRemoteUser = $null -ne (Get-Recipient $user.UserPrincipalName | Where-Object {$_.RecipientTypeDetails -eq "RemoteUserMailbox"})
    if($isRemoteUser){
        try{
            $remoteMailbox = Get-RemoteMailbox -Identity $user.UserPrincipalName -ErrorAction SilentlyContinue
            if($remoteMailbox.EmailAddresses){
                $emailAddresses = $remoteMailbox.EmailAddresses | Where-Object {$_ -like "*SMTP:*"} | ForEach-Object {$_.ToString().TrimStart('SMTP:')}
                $emailAddressCheck = $emailAddresses -like "*@glebecentre.mail.onmicrosoft.com*"
                $mailboxType = 'Remote'
            }
        }catch{
            Write-Host "$(Get-Date) - [WARNING] Could not find remote mailbox for user $($user.UserPrincipalName)" -ForegroundColor Yellow
        }
    }else{
        try{
            $localMailbox = Get-Mailbox -Identity $user.UserPrincipalName -ErrorAction SilentlyContinue
            if($localMailbox.EmailAddresses){
                $emailAddresses = $localMailbox.EmailAddresses | Where-Object {$_ -like "*SMTP:*"} | ForEach-Object {$_.ToString().TrimStart('SMTP:')}
                $emailAddressCheck = $emailAddresses -like "*@glebecentre.mail.onmicrosoft.com*"
                $mailboxType = 'Local'
            }
        }catch{
            Write-Host "$(Get-Date) - [WARNING] Could not find local mailbox for user $($user.UserPrincipalName)" -ForegroundColor Yellow
        }
    }
    
    # Only process users with email addresses
    if($emailAddresses){
        $userProperties = @{
            UserPrincipalName = $user.UserPrincipalName
            EmailAddresses = $emailAddresses -join ', '
            EmailAddressCheck = $emailAddressCheck
            MailboxType = $mailboxType
        }
        New-Object PsObject -Property $userProperties
    }
}

$totalUsers = $users.count
$results = $results | Sort-Object -Property MailboxType, EmailAddressCheck -Descending
$totalusersinresults = $results.count

$totalRemote = ($results | Where-Object {$_.MailboxType -eq "Remote"}).count
$totalLocal = ($results | Where-Object {$_.MailboxType -eq "Local"}).count
$totalEmailCheckPassed = ($results | Where-Object {$_.EmailAddressCheck -ne $null}).count
$totalEmailCheckFailed = ($results | Where-Object {$_.EmailAddressCheck -eq $false}).count

Write-Host "$(Get-Date) - [INFO] Total Users in Exchange: $totalUsers" -ForegroundColor Green
Write-Host "$(Get-Date) - [INFO] Total Users in results: $totalusersinresults" -ForegroundColor Green
Write-Host "$(Get-Date) - [INFO] Total Remote Users: $totalRemote" -ForegroundColor Green
Write-Host "$(Get-Date) - [INFO] Total Local Users: $totalLocal" -ForegroundColor Green
Write-Host "$(Get-Date) - [INFO] Total Users Passed Email Check: $totalEmailCheckPassed" -ForegroundColor Green
Write-Host "$(Get-Date) - [INFO] Total Users Did Not Pass Email Check: $totalEmailCheckFailed" -ForegroundColor Green

# $results | Export-Csv -Path 'MailboxDetails.csv' -NoTypeInformation

$results | Out-GridView