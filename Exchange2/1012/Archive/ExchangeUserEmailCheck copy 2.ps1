$users = Get-User -ResultSize Unlimited

$results = foreach($user in $users){
    $mailboxType = $null
    # Check if the user has a Remote User Mailbox
    $isRemoteUser = $null -ne (Get-Recipient $user.UserPrincipalName | Where-Object {$_.RecipientTypeDetails -eq "RemoteUserMailbox"})
    if($isRemoteUser){
        try{
            $remoteMailbox = Get-RemoteMailbox -Identity $user.UserPrincipalName -ErrorAction SilentlyContinue
            $mailboxType = 'Remote'
        }catch{
            Write-Warning "Could not find remote mailbox for user $($user.UserPrincipalName)"
        }
    }else{
        try{
            $localMailbox = Get-Mailbox -Identity $user.UserPrincipalName -ErrorAction SilentlyContinue
            $mailboxType = 'Local'
        }catch{
            Write-Warning "Could not find local mailbox for user $($user.UserPrincipalName)"
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
