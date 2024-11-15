# After all the mailboxes are synced and the migration batch has a status of "Synced", you can complete the batch.
# Note: Do NOT run this line until all the mailboxes are synced. It might take several hours or even days depending on the size and number of the mailboxes.

$newMigrationBatchName = '[May232023]-[Prod]-[6SharedMailboxes]'

Complete-MigrationBatch -Identity $newMigrationBatchName
Write-Host "$(Get-Date) - [INFO] Migration batch completed successfully!" -ForegroundColor Green