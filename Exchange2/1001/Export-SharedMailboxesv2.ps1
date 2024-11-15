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

$results | Export-Csv -Path 'C:\Code\Exchange\Exports\AGH_October_3rd_2023_12_34_PM_SharedMailbox_PowerShell_Export_V4.csv' -NoTypeInformation
