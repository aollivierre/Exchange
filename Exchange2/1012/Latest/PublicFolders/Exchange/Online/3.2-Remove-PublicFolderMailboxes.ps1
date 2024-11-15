# After the public folders are removed, run the following commands to remove all public folder mailboxes:

$hierarchyMailboxGuid = $(Get-OrganizationConfig).RootPublicFolderMailbox.HierarchyMailboxGuid
Get-Mailbox -PublicFolder | Where-Object {$_.ExchangeGuid -ne $hierarchyMailboxGuid} | Remove-Mailbox -PublicFolder -Confirm:$false -Force
Get-Mailbox -PublicFolder | Where-Object {$_.ExchangeGuid -eq $hierarchyMailboxGuid} | Remove-Mailbox -PublicFolder -Confirm:$false -Force
Get-Mailbox -PublicFolder -SoftDeletedMailbox | ForEach-Object {Remove-Mailbox -PublicFolder $_.PrimarySmtpAddress -PermanentlyDelete:$true -force -Confirm:$false}  
$soft=Get-Mailbox -PublicFolder -SoftDeletedMailbox; foreach ($mbx in $soft){if ($mbx.Name -like "*CNF:*" -or $mbx.identity -like "*CNF:*") {Remove-Mailbox -PublicFolder        $mbx.ExchangeGUID.GUID -RemoveCNFPublicFolderMailboxPermanently -Force -Confirm:$false}}


# Exception: Unable to index into an object of type
# "System.Collections.Generic.Dictionary`2[System.String,System.Collections.Generic.IEnumerable`1[System.String]]".
