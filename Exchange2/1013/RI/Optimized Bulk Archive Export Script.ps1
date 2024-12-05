# Configuration
$exportPath = "\\$env:COMPUTERNAME\C$\ExchangeArchives"
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$logFile = "$exportPath\ExportLog_$timestamp.csv"
$maxConcurrentExports = 2

# Setup export directory and permissions
if (-not (Test-Path $exportPath)) { $null = New-Item -ItemType Directory -Force -Path $exportPath }
$acl = Get-Acl $exportPath
$exchangeIdentities = @("Exchange Trusted Subsystem", "SYSTEM", "NETWORK SERVICE")
foreach ($identity in $exchangeIdentities) {
    $acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule($identity, "FullControl", "ContainerInherit,ObjectInherit", "None", "Allow")))
}
Set-Acl -Path $exportPath -AclObject $acl

# Initialize ArrayList for results
$results = [System.Collections.ArrayList]::new()

# Get all mailboxes with archives in a single query
$mailboxes = @(Get-Mailbox -ResultSize Unlimited -Filter {
    ArchiveDatabase -ne $null -and 
    ArchiveState -eq 'Local' -and 
    RecipientTypeDetails -ne 'DiscoveryMailbox'
} | Where-Object {
    $_.Identity -notlike "SystemMailbox*" -and
    $_.Identity -ne "Administrator"
})

Write-Host "Found $($mailboxes.Count) mailboxes with archives to process"

foreach ($mailbox in $mailboxes) {
    # Check current exports using a single query
    while (@(Get-MailboxExportRequest -Status InProgress).Count -ge $maxConcurrentExports) {
        Write-Host "Maximum concurrent exports reached. Waiting..."
        Start-Sleep -Seconds 30
    }

    $startTime = Get-Date
    # Single query for archive stats
    $archiveStats = Get-MailboxStatistics -Identity $mailbox.Identity -Archive
    $pstPath = "$exportPath\$($mailbox.Alias)_archive_$timestamp.pst"
    
    Write-Host "`nProcessing: $($mailbox.DisplayName)"
    Write-Host "Archive size: $($archiveStats.TotalItemSize)"
    Write-Host "Items: $($archiveStats.ItemCount)"

    try {
        # Clean up existing requests in a single operation
        Get-MailboxExportRequest -Name "Archive_Export_$($mailbox.Alias)*" | 
            Remove-MailboxExportRequest -Confirm:$false -ErrorAction SilentlyContinue

        $exportParams = @{
            Mailbox = $mailbox.Identity
            IsArchive = $true
            FilePath = $pstPath
            Name = "Archive_Export_$($mailbox.Alias)_$timestamp"
        }

        $null = New-MailboxExportRequest @exportParams
        
        do {
            Start-Sleep -Seconds 10
            # Single status query
            $status = Get-MailboxExportRequest -Name $exportParams.Name
            if ($status.Status -like "*Failed*") { throw $status.LastError }
            Write-Host "Status: $($status.Status) - $(if($status.PercentComplete){"$($status.PercentComplete)%"}else{"Calculating..."})"
        } while ($status.Status -notlike "*Completed*")

        $endTime = Get-Date
        $pstSize = if (Test-Path $pstPath) { (Get-Item $pstPath).Length } else { 0 }

        $result = [PSCustomObject]@{
            Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            User = $mailbox.DisplayName
            Email = $mailbox.PrimarySmtpAddress
            Status = "Success"
            OriginalSize = $archiveStats.TotalItemSize
            FinalPSTSize = "$([math]::Round($pstSize / 1MB, 2)) MB"
            ItemCount = $archiveStats.ItemCount
            Duration = "$([math]::Round(($endTime - $startTime).TotalMinutes, 2)) minutes"
            ExportPath = $pstPath
            Error = ""
        }
    }
    catch {
        $result = [PSCustomObject]@{
            Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            User = $mailbox.DisplayName
            Email = $mailbox.PrimarySmtpAddress
            Status = "Failed"
            OriginalSize = $archiveStats.TotalItemSize
            FinalPSTSize = "N/A"
            ItemCount = $archiveStats.ItemCount
            Duration = "$([math]::Round(((Get-Date) - $startTime).TotalMinutes, 2)) minutes"
            ExportPath = $pstPath
            Error = $_.ToString()
        }
        Write-Warning "Error processing $($mailbox.DisplayName): $_"
    }
    
    # Add to ArrayList and export in single operations
    $null = $results.Add($result)
    $results | Export-Csv -Path $logFile -NoTypeInformation -Force
}

$successful = @($results | Where-Object {$_.Status -eq "Success"}).Count
$failed = @($results | Where-Object {$_.Status -eq "Failed"}).Count

Write-Host "`nExport Summary:"
Write-Host "Total: $($results.Count) | Successful: $successful | Failed: $failed"
Write-Host "Log: $logFile"