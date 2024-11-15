# If you do have any public folders in Microsoft 365 or Office 365 or Exchange Online, run the following PowerShell command to remove them (after confirming that they are not needed). Make sure that you've saved any information within these public folders before deleting them, because all information will be permanently deleted when you remove the public folders.

Get-MailPublicFolder -ResultSize Unlimited | Where-Object {$_.EntryId -ne $null}| Disable-MailPublicFolder -Confirm:$false
Get-PublicFolder -GetChildren \ -ResultSize Unlimited | Remove-PublicFolder -Recurse -Confirm:$false



# Get-PublicFolder: |Microsoft.Exchange.Data.StoreObjects.ObjectNotFoundException|No active public folder mailboxes were found for organization
# glebecentre.onmicrosoft.com. This happens when no public folder mailboxes are provisioned or they are provisioned in 'HoldForMigration'
# mode. If you're not currently performing a migration, create a public folder mailbox.