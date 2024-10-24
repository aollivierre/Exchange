$dbs = Get-MailboxDatabase
$dbs | ForEach-Object {Get-MailboxStatistics -Database $_.DistinguishedName} 