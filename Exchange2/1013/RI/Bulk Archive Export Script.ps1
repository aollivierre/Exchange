# Configuration
$exportPath = "\\$env:COMPUTERNAME\C$\ExchangeArchives"
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$logFile = "$exportPath\ExportLog_$timestamp.csv"
$maxConcurrentExports = 2

# Setup export directory and permissions
if (-not (Test-Path $exportPath)) { $null = New-Item -ItemType Directory -Force -Path $exportPath }
$acl = Get-Acl $exportPath
@("Exchange Trusted Subsystem", "SYSTEM", "NETWORK SERVICE") | ForEach-Object {
    $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule($_, "FullControl", "ContainerInherit,ObjectInherit", "None", "Allow")
    $acl.AddAccessRule($accessRule)
}
Set-Acl -Path $exportPath -AclObject $acl

# Get all mailboxes with archives
$mailboxes = Get-Mailbox -ResultSize Unlimited | Where-Object {
    $_.ArchiveDatabase -ne $null -and 
    $_.ArchiveState -eq 'Local' -and 
    $_.RecipientTypeDetails -ne 'DiscoveryMailbox' -and
    $_.Identity -notlike "SystemMailbox*" -and
    $_.Identity -ne "Administrator"
}

Write-Host "Found $($mailboxes.Count) mailboxes with archives to process"

# Results tracking
$results = @()

foreach ($mailbox in $mailboxes) {
    # Check current export count
    $currentExports = Get-MailboxExportRequest | Where-Object {$_.Status -eq "InProgress"}
    while ($currentExports.Count -ge $maxConcurrentExports) {
        Write-Host "Maximum concurrent exports reached. Waiting..."
        Start-Sleep -Seconds 30
        $currentExports = Get-MailboxExportRequest | Where-Object {$_.Status -eq "InProgress"}
    }

    $startTime = Get-Date
    $archiveStats = Get-MailboxStatistics -Identity $mailbox.Identity -Archive
    $pstPath = "$exportPath\$($mailbox.Alias)_archive_$timestamp.pst"
    
    Write-Host "`nProcessing: $($mailbox.DisplayName)"
    Write-Host "Archive size: $($archiveStats.TotalItemSize)"
    Write-Host "Items: $($archiveStats.ItemCount)"

    try {
        # Clean up existing requests for this mailbox
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
            $status = Get-MailboxExportRequest -Name $exportParams.Name
            if ($status.Status -like "*Failed*") { throw $status.LastError }
            Write-Host "Status: $($status.Status) - $(if($status.PercentComplete){"$($status.PercentComplete)%"}else{"Calculating..."})"
        } while ($status.Status -notlike "*Completed*")

        $endTime = Get-Date
        $pstSize = if (Test-Path $pstPath) { (Get-Item $pstPath).Length } else { 0 }

        $results += [PSCustomObject]@{
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
        Write-Warning "Error processing $($mailbox.DisplayName): $_"
        $results += [PSCustomObject]@{
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
    }
    
    # Export current results after each mailbox
    $results | Export-Csv -Path $logFile -NoTypeInformation -Force
}

# Final Summary
Write-Host "`nExport Summary:"
Write-Host "Total mailboxes processed: $($results.Count)"
Write-Host "Successful exports: $($results | Where-Object {$_.Status -eq 'Success'}).Count"
Write-Host "Failed exports: $($results | Where-Object {$_.Status -eq 'Failed'}).Count"
Write-Host "Detailed log saved to: $logFile"