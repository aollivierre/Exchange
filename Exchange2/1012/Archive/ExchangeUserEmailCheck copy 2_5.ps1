# $ErrorActionPreference = 'SilentlyContinue'

$users = Get-User -ResultSize Unlimited

$results = foreach($user in $users){
    $mailboxType = $null
    $emailAddressCheck = $null

    # Check if the user has a Remote User Mailbox
    $isRemoteUser = $null -ne (Get-Recipient $user.UserPrincipalName | Where-Object {$_.RecipientTypeDetails -eq "RemoteUserMailbox"})
    
    if($isRemoteUser){
        try{
            $remoteMailbox = Get-RemoteMailbox -Identity $user.UserPrincipalName -ErrorAction SilentlyContinue
            $mailboxType = 'Remote'
            $emailAddressCheck = $remoteMailbox.EmailAddresses -contains "glebecentre.mail.onmicrosoft.com"
        }catch{
            Write-Warning "Could not find remote mailbox for user $($user.UserPrincipalName)"
        }
    }else{
        try{
            $localMailbox = Get-Mailbox -Identity $user.UserPrincipalName -ErrorAction SilentlyContinue
            $mailboxType = 'Local'
            $emailAddressCheck = $localMailbox.EmailAddresses -contains "glebecentre.mail.onmicrosoft.com"
        }catch{
            Write-Warning "Could not find local mailbox for user $($user.UserPrincipalName)"
        }
    }

    $userProperties = @{
        UserPrincipalName = $user.UserPrincipalName
        EmailAddressCheck = $emailAddressCheck
        MailboxType = $mailboxType
    }

    New-Object PsObject -Property $userProperties
}

# $results | Export-Csv -Path 'MailboxDetails.csv' -NoTypeInformation
$results | Out-GridView
