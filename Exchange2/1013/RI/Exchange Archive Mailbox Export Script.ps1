# Script to export archive mailboxes to PST files and generate a report
# Requires Exchange Management Shell

# Configuration
$config = @{
    ExportPath = "\\nti-ri-vmhost2\F$\ArchiveExports"
    ReportPath = "\\nti-ri-vmhost2\F$\ArchiveExports\ExportReport.csv"
    LogPath = "\\nti-ri-vmhost2\F$\ArchiveExports\ExportLog.txt"
    Timestamp = (Get-Date -Format "yyyyMMdd_HHmmss")
}

# Create necessary folders
$null = New-Item -ItemType Directory -Force -Path $config.ExportPath
$null = New-Item -ItemType Directory -Force -Path "$($config.ExportPath)\Logs"

# Start logging
Start-Transcript -Path $config.LogPath -Append

Write-Host "Starting archive mailbox export process..."

# Get all mailboxes with archives - single server call
$getMailboxParams = @{
    ResultSize = 'Unlimited'
}
$mailboxes = @(Get-Mailbox @getMailboxParams | Where-Object {$_.ArchiveDatabase})

# Initialize ArrayList for better performance with large datasets
$exportResults = [System.Collections.ArrayList]::new()

foreach ($mailbox in $mailboxes) {
    $startTime = Get-Date
    $status = "Success"
    $errorMessage = ""
    $pstPath = ""
    
    Write-Host "Processing archive for user: $($mailbox.DisplayName)"
    
    try {
        # Create user-specific folder
        $userFolder = Join-Path $config.ExportPath $mailbox.Alias
        $null = New-Item -ItemType Directory -Force -Path $userFolder
        
        # Prepare export parameters
        $pstPath = Join-Path $userFolder "$($mailbox.Alias)_archive_$($config.Timestamp).pst"
        
        $exportParams = @{
            Mailbox = $mailbox.Identity
            IsArchive = $true
            FilePath = $pstPath
            Name = "Archive_Export_$($mailbox.Alias)_$($config.Timestamp)"
        }
        
        # Single server call for export request
        $null = New-MailboxExportRequest @exportParams
        
        # Wait for export to complete (with timeout)
        $timeout = 3600 # 1 hour timeout
        $timer = [Diagnostics.Stopwatch]::StartNew()
        
        do {
            Start-Sleep -Seconds 30
            $requestParams = @{
                Name = "Archive_Export_$($mailbox.Alias)_$($config.Timestamp)"
            }
            $request = Get-MailboxExportRequest @requestParams
            
            if ($timer.Elapsed.TotalSeconds -gt $timeout) {
                throw "Export timed out after 1 hour"
            }
        } while ($request.Status -notlike "*Completed*")
        
        # Clean up export request
        $removeParams = @{
            Identity = $request.Identity
            Confirm = $false
        }
        Remove-MailboxExportRequest @removeParams
        
    } catch {
        $status = "Failed"
        $errorMessage = $_.Exception.Message
        Write-Warning "Error processing $($mailbox.DisplayName): $errorMessage"
    }
    
    # Get archive statistics - single server call
    $statsParams = @{
        Identity = $mailbox.Identity
        Archive = $true
    }
    $archiveStats = Get-MailboxStatistics @statsParams
    
    # Create result object and add to ArrayList
    $resultObject = [PSCustomObject]@{
        Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        User = $mailbox.DisplayName
        Email = $mailbox.PrimarySmtpAddress
        Status = $status
        Error = $errorMessage
        Duration = "{0:hh\:mm\:ss}" -f ((Get-Date) - $startTime)
        ExportPath = $pstPath
        ArchiveSize = $archiveStats.TotalItemSize
    }
    
    $null = $exportResults.Add($resultObject)
}

# Export results to CSV in one operation
$exportResults | Export-Csv -Path $config.ReportPath -NoTypeInformation

# Generate summary statistics
$successful = @($exportResults | Where-Object {$_.Status -eq "Success"}).Count
$failed = @($exportResults | Where-Object {$_.Status -eq "Failed"}).Count
$totalSize = ($exportResults | Measure-Object -Property ArchiveSize -Sum).Sum

# Display summary
$summary = @"

Export Summary:
----------------
Total mailboxes processed: $($exportResults.Count)
Successful exports: $successful
Failed exports: $failed
Total archive size: $totalSize
Detailed report saved to: $($config.ReportPath)
Full log saved to: $($config.LogPath)
"@

Write-Host $summary

Stop-Transcript