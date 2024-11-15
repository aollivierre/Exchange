# $mappings = Import-Csv <Folder-to-mailbox map path>
$mappings = Import-Csv "C:\Code\CB\Exchange\Glebe\Latest\PublicFolders\Exchange\CSV_Generated_From_OnPrem_EMS\map.csv"
$primaryMailboxName = ($mappings | Where-Object FolderPath -eq "\" ).TargetMailbox;
New-Mailbox -HoldForMigration:$true -PublicFolder -IsExcludedFromServingHierarchy:$false $primaryMailboxName
($mappings | Where-Object TargetMailbox -ne $primaryMailboxName).TargetMailbox | Sort-Object -unique | ForEach-Object { New-Mailbox -PublicFolder -IsExcludedFromServingHierarchy:$false $_ }