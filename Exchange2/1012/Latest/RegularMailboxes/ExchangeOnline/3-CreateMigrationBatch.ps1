# 5- Output the total number of users in the CSV file
<<<<<<< Updated upstream:Exchange/Glebe/Latest/RegularMailboxes/ExchangeOnline/3-CreateMigrationBatch.ps1
$csvFilePath = "C:\Code\CB\Exchange\Glebe\Exports\Glebe_Exchange_Migration_Last_batch_7users_June_14_2023.csv" # Replace with your CSV file path
=======
$csvFilePath = "C:\Code\CB\Exchange\LHC\Exports\onpremremaining_1user_migration_batch1.csv" # Replace with your CSV file path
>>>>>>> Stashed changes:Exchange/Glebe/Latest/ExchangeOnline/3-CreateMigrationBatch.ps1
$csvFile = @(Import-Csv -Path $csvFilePath)
$totalUsers = $csvFile.Count
Write-Host "$(Get-Date) - [INFO] Total number of users in the CSV file: $totalUsers" -ForegroundColor Green


# Then, you can proceed with creating the migration batch
<<<<<<< Updated upstream:Exchange/Glebe/Latest/RegularMailboxes/ExchangeOnline/3-CreateMigrationBatch.ps1
$newMigrationBatchName = "[June172023]-[Prod]-[7Prod]-[Batch8]-[Phase6]" # Replace with your preferred batch name
=======
$newMigrationBatchName = "[Aug212023]-[Prod]-[1Prod]-[Batch3]-[Phase2]" # Replace with your preferred batch name
>>>>>>> Stashed changes:Exchange/Glebe/Latest/ExchangeOnline/3-CreateMigrationBatch.ps1

# Define parameters in a hashtable (splatting)
$params = @{
    Name = $newMigrationBatchName
    SourceEndpoint = "Hybrid Migration Endpoint - EWS (Default Web Site)"
    TargetDeliveryDomain = "lcwhc.mail.onmicrosoft.com"
    CSVData = [System.IO.File]::ReadAllBytes($csvFilePath)
<<<<<<< Updated upstream:Exchange/Glebe/Latest/RegularMailboxes/ExchangeOnline/3-CreateMigrationBatch.ps1
    NotificationEmails = "NovaAdmin-Abdullah@glebecentre.ca"
    CompleteAfter = "2023-06-17 11:59:00 PM"
=======
    NotificationEmails = "Admin-CCI@lhc.ca"
    CompleteAfter = "2023-08-23 11:59:00 PM"
>>>>>>> Stashed changes:Exchange/Glebe/Latest/ExchangeOnline/3-CreateMigrationBatch.ps1
}

# Create the new migration batch with the defined parameters
New-MigrationBatch @params

# Start the migration batch
Start-MigrationBatch -Identity $newMigrationBatchName
