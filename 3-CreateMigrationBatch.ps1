# 5- Output the total number of users in the CSV file
$csvFilePath = "C:\Code\CB\Exchange\Glebe\Exports\Glebe-Migration_32Users_AD_Email_Alias_Hybrid_Migration.csv" # Replace with your CSV file path
$csvFile = @(Import-Csv -Path $csvFilePath)
$totalUsers = $csvFile.Count
Write-Host "$(Get-Date) - [INFO] Total number of users in the CSV file: $totalUsers" -ForegroundColor Green


# Then, you can proceed with creating the migration batch
$newMigrationBatchName = "[June132023]-[Prod]-[32Prod]-[Batch7]-[Phase5]" # Replace with your preferred batch name

# Define parameters in a hashtable (splatting)
$params = @{
    Name = $newMigrationBatchName
    SourceEndpoint = "Hybrid Migration Endpoint - EWS (Default Web Site)"
    TargetDeliveryDomain = "glebecentre.mail.onmicrosoft.com"
    CSVData = [System.IO.File]::ReadAllBytes($csvFilePath)
    NotificationEmails = "NovaAdmin-Abdullah@glebecentre.ca"
    CompleteAfter = "2023-06-13 11:59:00 PM"
}

# Create the new migration batch with the defined parameters
New-MigrationBatch @params

# Start the migration batch
Start-MigrationBatch -Identity $newMigrationBatchName
