Install-Module -Name ExchangeOnlineManagement
Import-Module ExchangeOnlineManagement
Connect-ExchangeOnline

# Export the users, mailbox usage, and archiving status to a CSV file
$Users = Get-Mailbox -ResultSize Unlimited | Select-Object UserPrincipalName, ArchiveStatus, @{Name="MailboxSizeGB"; Expression={(Get-MailboxStatistics $_.UserPrincipalName).TotalItemSize.Value.ToGB()}}

# Filter the users based on the conditions and export them to the console and CSV
$Users | Export-Csv -Path "C:\Code\CB\Exchange\CPHA\Exports\ExchangeUsers.csv" -NoTypeInformation
Write-Output "Total users: $($Users.Count)"
Write-Output "Total users with more than 4.5GB: $($Users | Where-Object {$_.MailboxSizeGB -gt 4.5}).Count"
Write-Output "Total users with less than 4.5GB: $($Users | Where-Object {$_.MailboxSizeGB -lt 4.5}).Count"
Write-Output "Total users with no archiving enabled: $($Users | Where-Object {$_.ArchiveStatus -eq 'None'}).Count"