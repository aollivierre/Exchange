# Connect to Exchange if not already connected
# Assuming you're already connected as mentioned

# Get all mailboxes and their archive details
$results = Get-Mailbox -ResultSize Unlimited | ForEach-Object {
    $mailbox = $_
    $stats = $null
    $archiveStats = $null
    
    # Try to get mailbox statistics
    try {
        $stats = Get-MailboxStatistics $mailbox.Identity -ErrorAction SilentlyContinue
        if ($mailbox.ArchiveDatabase) {
            $archiveStats = Get-MailboxStatistics $mailbox.Identity -Archive -ErrorAction SilentlyContinue
        }
    } catch {
        Write-Warning "Could not get statistics for mailbox: $($mailbox.Identity)"
    }
    
    # Create custom object with desired properties
    [PSCustomObject]@{
        DisplayName = $mailbox.DisplayName
        PrimarySmtpAddress = $mailbox.PrimarySmtpAddress
        MailboxType = $mailbox.RecipientTypeDetails
        IsEnabled = $stats.IsValid
        ArchiveEnabled = [bool]$mailbox.ArchiveDatabase
        ArchiveState = $mailbox.ArchiveState
        ArchivePolicy = if ($mailbox.RetentionPolicy) { $mailbox.RetentionPolicy } else { "No Policy" }
        ArchiveQuota = if ($mailbox.ArchiveQuota) { $mailbox.ArchiveQuota } else { "No Quota" }
        ArchiveWarningQuota = if ($mailbox.ArchiveWarningQuota) { $mailbox.ArchiveWarningQuota } else { "No Warning Quota" }
        ArchiveSizeGB = if ($archiveStats) { 
            [math]::Round(($archiveStats.TotalItemSize.ToString().Split("(")[1].Split(" ")[0].Replace(",","")/1GB), 2)
        } else { 0 }
        ArchiveCreationDate = $mailbox.ArchiveCreationDate
        LastLogonTime = $stats.LastLogonTime
        ItemCount = $stats.ItemCount
        DatabaseName = $mailbox.Database
        ServerName = $mailbox.ServerName
    }
}

# Export results to CSV
$date = Get-Date -Format "yyyyMMdd-HHmmss"
$results | Export-CSV "ArchiveAudit_$date.csv" -NoTypeInformation
$results | Out-HtmlView -FilePath "ArchiveAudit_$date.HTML" -Title "NTI_OTT_onprem_exchange_$date"

# Display summary on screen
Write-Host "`nArchive Status Summary:"
Write-Host "Total Mailboxes: $($results.Count)"
Write-Host "Mailboxes with Archives Enabled: $(($results | Where-Object {$_.ArchiveEnabled}).Count)"
Write-Host "Total Archive Size (GB): $(($results | Measure-Object -Property ArchiveSizeGB -Sum).Sum)"
Write-Host "`nReport exported to: ArchiveAudit_$date.csv"