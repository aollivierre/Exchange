Connect-ExchangeOnline

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

# Retrieve all mailboxes and gather mailbox data in parallel
# $Users = Get-Mailbox -ResultSize Unlimited | ForEach-Object -Parallel $gatherMailboxData -ThrottleLimit 10

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
$UsersMoreThanThreshold | Export-Csv -Path "C:\Code\CB\Exchange\CPHA\Exports\UsersMoreThanThreshold-parallel.csv" -NoTypeInformation
$UsersLessThanThreshold | Export-Csv -Path "C:\Code\CB\Exchange\CPHA\Exports\UsersLessThanThreshold-parallel.csv" -NoTypeInformation

# Display in a grid view
$UsersMoreThanThreshold | Out-GridView -Title 'User Mailboxes More Than 4.5GB'
$UsersLessThanThreshold | Out-GridView -Title 'User Mailboxes Less Than 4.5GB'