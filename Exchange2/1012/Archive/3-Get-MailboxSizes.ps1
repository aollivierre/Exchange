$mailboxes = Get-Mailbox -ResultSize Unlimited

$results = $null
$results = foreach ($mailbox in $mailboxes) {
    $stats = Get-MailboxStatistics -Identity $mailbox.UserPrincipalName

    $userProperties = @{
        UserPrincipalName = $mailbox.UserPrincipalName
        MailboxSize = $stats.TotalItemSize.Value.ToMB()  # Convert size to MB for easier reading
    }

    New-Object PsObject -Property $userProperties
}

$results = $results | Sort-Object -Property MailboxSize -Descending
$results | Out-GridView
