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
            Write-Warning "Could not find remote mailbox for user $($user.UserPrincipalName)"
        }
    }else{
        try{
            $localMailbox = Get-Mailbox -Identity $user.UserPrincipalName -ErrorAction SilentlyContinue
            $emailAddresses = $localMailbox.EmailAddresses | Where-Object {$_ -like "*SMTP:*"} | ForEach-Object {$_.ToString().TrimStart('SMTP:')}
            $emailAddressCheck = $emailAddresses -like "*@glebecentre.mail.onmicrosoft.com*"
            $mailboxType = 'Local'
        }catch{
            Write-Warning "Could not find local mailbox for user $($user.UserPrincipalName)"
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

$results | Export-Csv -Path 'C:\Code\CB\Exchange\Glebe\MailboxDetails.csv' -NoTypeInformation
$results | Out-GridView
