$users = Get-User -ResultSize Unlimited

$results = foreach($user in $users){
    $mailboxType = $null
    $remoteMailbox = $null
    $localMailbox = $null
    
    # Check if the user has a local mailbox
    try {
        $localMailbox = Get-Mailbox -Identity $user.UserPrincipalName -ErrorAction Stop
        $mailboxType = 'Local'
    }
    catch {
        # If the user doesn't have a local mailbox, check if they have a remote mailbox
        try {
            $remoteMailbox = Get-RemoteMailbox -Identity $user.UserPrincipalName -ErrorAction Stop
            $mailboxType = 'Remote'
        }
        catch {
            Write-Warning "Could not find a mailbox for user $($user.UserPrincipalName)"
        }
    }

    $userProperties = @{
        UserPrincipalName = $user.UserPrincipalName
        EmailAddressCheck = $user.EmailAddresses -like "*@glebecentre.mail.onmicrosoft.com*"
        MailboxType = $mailboxType
    }
    New-Object PsObject -Property $userProperties
}

$results | Export-Csv -Path 'MailboxDetails.csv' -NoTypeInformation
$results | Out-GridView
