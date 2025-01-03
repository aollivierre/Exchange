$users = Get-User -ResultSize Unlimited

$results = foreach($user in $users){
    $mailboxType = $null
    $emailAddresses = $null
    $isRemoteUser = $null -ne (Get-Recipient $user.UserPrincipalName | Where-Object {$_.RecipientTypeDetails -eq "RemoteUserMailbox"})
    if($isRemoteUser){
        try{
            $remoteMailbox = Get-RemoteMailbox -Identity $user.UserPrincipalName -ErrorAction SilentlyContinue
            $emailAddresses = $remoteMailbox.EmailAddresses | Where-Object {$_ -like "*SMTP:*"} | ForEach-Object {$_.ToString().TrimStart('SMTP:')}
            $mailboxType = 'Remote'
        }catch{
            Write-Warning "Could not find remote mailbox for user $($user.UserPrincipalName)"
        }
    }else{
        try{
            $localMailbox = Get-Mailbox -Identity $user.UserPrincipalName -ErrorAction SilentlyContinue
            $emailAddresses = $localMailbox.EmailAddresses | Where-Object {$_ -like "*SMTP:*"} | ForEach-Object {$_.ToString().TrimStart('SMTP:')}
            $mailboxType = 'Local'
        }catch{
            Write-Warning "Could not find local mailbox for user $($user.UserPrincipalName)"
        }
    }

    $userProperties = @{
        UserPrincipalName = $user.UserPrincipalName
        EmailAddresses = $emailAddresses -join ', '
        MailboxType = $mailboxType
    }
    New-Object PsObject -Property $userProperties
}

# $results | Export-Csv -Path 'MailboxDetails.csv' -NoTypeInformation
$results | Out-GridView
