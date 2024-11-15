$sharedMailboxes = Get-Mailbox -ResultSize Unlimited | Where-Object {$_.RecipientTypeDetails -eq "SharedMailbox"}

$results = foreach ($mailbox in $sharedMailboxes) {
    $delegates = Get-MailboxPermission $mailbox.DistinguishedName | Where-Object { ($_.IsInherited -eq $false) -and ($_.User -notlike 'NT AUTHORITY\SELF') -and ($_.Deny -eq $false) }

    $userProperties = @{
        SharedMailbox = $mailbox.UserPrincipalName
        Delegates = $delegates.User -join ', '
    }

    New-Object PsObject -Property $userProperties
}

$results | Out-GridView

$results | Export-Csv -Path 'C:\Code\CB\Exchange\Glebe\Exports\Glebe_6_Shared_Mailboxes_to_Migrate_.csv' -NoTypeInformation
