$users = Get-User -ResultSize Unlimited

$results = foreach($user in $users){
    $mailboxType = $null
    try{
        $remoteMailbox = Get-RemoteMailbox -Identity $user.UserPrincipalName -ErrorAction Stop
        $mailboxType = 'Remote'
    }catch{
        try{
            $localMailbox = Get-Mailbox -Identity $user.UserPrincipalName -ErrorAction Stop
            $mailboxType = 'Local'
        }catch{
            Write-Warning "Could not find local or remote mailbox for user $($user.UserPrincipalName)"
            continue
        }
    }
    $userProperties = @{
        UserPrincipalName = $user.UserPrincipalName
        EmailAddressCheck = $user.EmailAddresses -like "*@glebecentre.mail.onmicrosoft.com*"
        MailboxType = $mailboxType
    }
    New-Object PsObject -Property $userProperties
}

# $results | Export-Csv -Path 'MailboxDetails.csv' -NoTypeInformation
$results | Out-GridView