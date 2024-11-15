# Connect to Exchange Online
Connect-ExchangeOnline

# Define an empty array to store user data
$Users = @()

# Retrieve all mailboxes
$mailboxes = Get-Mailbox -ResultSize Unlimited 

# Loop over each mailbox to gather data
foreach ($mailbox in $mailboxes) {
    $stats = Get-MailboxStatistics $mailbox.UserPrincipalName
    $sizeGB = [math]::Round(([int64]$stats.TotalItemSize.Value.ToString().Split("(")[1].Split(" ")[0].Replace(",","") / 1GB), 2)
    $Users += New-Object PSObject -Property @{
        UserPrincipalName = $mailbox.UserPrincipalName
        ArchiveEnabled = if($null -ne $mailbox.ArchiveDatabase) { "Yes" } else { "No" }
        MailboxSizeGB = $sizeGB
    }
}

# Export to CSV and console
$Users | Export-Csv -Path "C:\Code\CB\Exchange\CPHA\Exports\ExchangeUsers.csv" -NoTypeInformation
Write-Output "Total users: $($Users.Count)"
Write-Output "Total users with more than 4.5GB: $($Users | Where-Object {$_.MailboxSizeGB -gt 4.5}).Count"
Write-Output "Total users with less than 4.5GB: $($Users | Where-Object {$_.MailboxSizeGB -lt 4.5}).Count"
Write-Output "Total users with no archiving enabled: $($Users | Where-Object {$_.ArchiveEnabled -eq 'No'}).Count"
