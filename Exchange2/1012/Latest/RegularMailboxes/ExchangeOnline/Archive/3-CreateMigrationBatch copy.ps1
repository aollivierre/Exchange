# 5- Output the total number of users in the CSV file
$csvFilePath = "C:\Code\CB\Exchange\Glebe\Exports\Glebe_1_Shared_Mailboxes_to_Migrate_RemoteMove.csv" # Replace with your CSV file path
$csvFile = Import-Csv -Path $csvFilePath
$totalUsers = $csvFile.Count
Write-Host "$(Get-Date) - [INFO] Total number of users in the CSV file: $totalUsers" -ForegroundColor Green

# 6- Create a New-MigrationBatch for a CSV file containing a number of users
$newMigrationBatchName = "[May232023]-[Prod]-[6SharedMailboxes]-[Batch2]" # Replace with your preferred batch name
Write-Host "$(Get-Date) - [INFO] Creating a new migration batch..." -ForegroundColor Cyan
New-MigrationBatch -Name $newMigrationBatchName -CSVData ([System.IO.File]::ReadAllBytes($csvFilePath)) -NotificationEmails "NovaAdmin-Abdullah@glebecentre.ca" # Replace with the admin's email

# Start the migration batch
Start-MigrationBatch -Identity $newMigrationBatchName
Write-Host "$(Get-Date) - [INFO] New migration batch started successfully!" -ForegroundColor Green

# After all the mailboxes are synced and the migration batch has a status of "Synced", you can complete the batch.
# Note: Do NOT run this line until all the mailboxes are synced. It might take several hours or even days depending on the size and number of the mailboxes.
# Complete-MigrationBatch -Identity $newMigrationBatchName
# Write-Host "$(Get-Date) - [INFO] Migration batch completed successfully!" -ForegroundColor Green
