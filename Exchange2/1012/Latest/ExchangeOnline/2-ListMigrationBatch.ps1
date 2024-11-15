# List all current migration batches
Write-Host "Listing all current migration batches..." -ForegroundColor Cyan
$migrationBatches = Get-MigrationBatch
$batchStatistics = @()

foreach ($batch in $migrationBatches) {
    Write-Host "Batch Name: $($batch.Identity)" -ForegroundColor Cyan
    Write-Host "Batch Status: $($batch.Status)" -ForegroundColor Cyan
    Write-Host "Batch Start Time: $($batch.StartTime)" -ForegroundColor Cyan
    Write-Host "Batch Last Synced Time: $($batch.LastSyncedTime)" -ForegroundColor Cyan
    Write-Host "Batch Source Endpoint: $($batch.SourceEndpoint)" -ForegroundColor Cyan
    Write-Host "Batch Target Delivery Domain: $($batch.TargetDeliveryDomain)" -ForegroundColor Cyan
    Write-Host "Batch Notification Emails: $($batch.NotificationEmails -join ', ')" -ForegroundColor Cyan
    
    # Getting user statistics for the current batch
    Write-Host "User Statistics for the batch $($batch.Identity):" -ForegroundColor Cyan
    
    $migrationUsers = Get-MigrationUser -BatchId $batch.Identity
    foreach ($user in $migrationUsers) {
        $userStats = Get-MigrationUserStatistics -Identity $user.Identity
        $batchStatistics += New-Object PSObject -Property @{
            'BatchName' = $batch.Identity
            'User' = $userStats.Identity
            'UserStatus' = $userStats.Status
            'LastSuccessfulSyncTime' = $userStats.LastSuccessfulSyncTime
        }
    }
}

$batchStatistics | Out-GridView

$csvFilePath = "C:\Code\CB\Exchange\Glebe\Exports\migrationBatchStatistics7.csv"  # Modify this with the actual path
$batchStatistics | Export-Csv -Path $csvFilePath -NoTypeInformation

Write-Host "CSV file created at $csvFilePath" -ForegroundColor Green