$users = Get-User -ResultSize Unlimited

$results = foreach($user in $users){
    $mailboxType = $null
    $emailAddresses = $null
    $emailAddressCheck = $null
    $isRemoteUser = $null -ne (Get-Recipient $user.UserPrincipalName | Where-Object {$_.RecipientTypeDetails -eq "RemoteUserMailbox"})
    if($isRemoteUser){
        try{
            $remoteMailbox = Get-RemoteMailbox -Identity $user.UserPrincipalName -ErrorAction SilentlyContinue
            $emailAddresses = $remoteMailbox.EmailAddresses | Where-Object {$_ -like "*SMTP:*"} | ForEach-Object {$_.ToString().TrimStart('SMTP:')}
            $emailAddressCheck = $emailAddresses -like "*@glebecentre.mail.onmicrosoft.com*"
            $mailboxType = 'Remote'
        }catch{
            Write-Host "$(Get-Date) - [WARNING] Could not find remote mailbox for user $($user.UserPrincipalName)" -ForegroundColor Yellow
        }
    }else{
        try{
            $localMailbox = Get-Mailbox -Identity $user.UserPrincipalName -ErrorAction SilentlyContinue
            $emailAddresses = $localMailbox.EmailAddresses | Where-Object {$_ -like "*SMTP:*"} | ForEach-Object {$_.ToString().TrimStart('SMTP:')}
            $emailAddressCheck = $emailAddresses -like "*@glebecentre.mail.onmicrosoft.com*"
            $mailboxType = 'Local'
        }catch{
            Write-Host "$(Get-Date) - [WARNING] Could not find local mailbox for user $($user.UserPrincipalName)" -ForegroundColor Yellow
        }
    }

    $userProperties = @{
        UserPrincipalName = $user.UserPrincipalName
        EmailAddresses = $emailAddresses -join ', '
        EmailAddressCheck = $emailAddressCheck
        MailboxType = $mailboxType
    }
    New-Object PsObject -Property $userProperties
}

$results = $results | Sort-Object -Property MailboxType, EmailAddressCheck -Descending

# $results | Export-Csv -Path 'MailboxDetails.csv' -NoTypeInformation
$results | Out-GridView
