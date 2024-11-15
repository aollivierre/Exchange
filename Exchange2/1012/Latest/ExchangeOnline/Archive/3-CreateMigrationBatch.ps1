# 5- Output the total number of users in the CSV file
$csvFilePath = "C:\Code\CB\Exchange\Glebe\Exports\Glebe_6_Shared_Mailboxes_to_Migrate_.csv" # Replace with the actual CSV file path
$csvFile = Import-Csv -Path $csvFilePath
$totalUsers = $csvFile.Count
Write-Host "Total number of users in the CSV file: $totalUsers" -ForegroundColor Green

# 6- Create a New-MigrationBatch for a CSV file containing a number of users
$newMigrationBatchName = "[May232023]-[Prod]-[6SharedMailboxes]-[Batch2]" # Replace with your preferred batch name
Write-Host "Creating a new migration batch..." -ForegroundColor Cyan
New-MigrationBatch -Name $newMigrationBatchName -CSVData ([System.IO.File]::ReadAllBytes($csvFilePath)) -NotificationEmails "NovaAdmin-Abdullah@glebecentre.ca" # Replace with the admin's email

Write-Host "New migration batch created successfully!" -ForegroundColor Green