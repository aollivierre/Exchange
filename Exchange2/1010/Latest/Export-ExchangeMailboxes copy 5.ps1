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

# Sort users by size (descending)
$Users = $Users | Sort-Object MailboxSizeGB -Descending

# Users with more than 4.5GB
$UsersMoreThanThreshold = $Users | Where-Object {$_.MailboxSizeGB -gt 4.5}
Write-Output "Total users with more than 4.5GB: $($UsersMoreThanThreshold.Count)"

# Users with less than 4.5GB
$UsersLessThanThreshold = $Users | Where-Object {$_.MailboxSizeGB -le 4.5}
Write-Output "Total users with less than 4.5GB: $($UsersLessThanThreshold.Count)"

# Users with no archiving enabled
$UsersNoArchiving = $Users | Where-Object {$_.ArchiveEnabled -eq 'No'}
Write-Output "Total users with no archiving enabled: $($UsersNoArchiving.Count)"

# Export all users to CSV
$Users | Export-Csv -Path "C:\Code\CB\Exchange\CPHA\Exports\ExchangeUsers.csv" -NoTypeInformation

# Export groups to separate CSV files
$UsersMoreThanThreshold | Export-Csv -Path "C:\Code\CB\Exchange\CPHA\Exports\UsersMoreThanThreshold.csv" -NoTypeInformation
$UsersLessThanThreshold | Export-Csv -Path "C:\Code\CB\Exchange\CPHA\Exports\UsersLessThanThreshold.csv" -NoTypeInformation

# Display in a grid view
$UsersMoreThanThreshold | Out-GridView -Title 'User Mailboxes More Than 4.5GB'
$UsersLessThanThreshold | Out-GridView -Title 'User Mailboxes Less Than 4.5GB'
