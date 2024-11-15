# Define a script block to gather mailbox data
$gatherMailboxData = {
    param($mailbox)
    # Connect-ExchangeOnline -UserPrincipalName "admin@yourdomain.com" -Password (ConvertTo-SecureString -String "YourPassword" -AsPlainText -Force)
    Connect-ExchangeOnline
    $stats = Get-MailboxStatistics $mailbox.UserPrincipalName
    $sizeGB = [math]::Round(([int64]$stats.TotalItemSize.Value.ToString().Split("(")[1].Split(" ")[0].Replace(",","") / 1GB), 2)
    return New-Object PSObject -Property @{
        UserPrincipalName = $mailbox.UserPrincipalName
        ArchiveEnabled = if($mailbox.ArchiveDatabase -ne $null) { "Yes" } else { "No" }
        MailboxSizeGB = $sizeGB
    }
}

# Retrieve all mailboxes and gather mailbox data in parallel
$mailboxes = Get-Mailbox -ResultSize Unlimited 
$Users = $mailboxes | ForEach-Object -Parallel $gatherMailboxData -ThrottleLimit 10