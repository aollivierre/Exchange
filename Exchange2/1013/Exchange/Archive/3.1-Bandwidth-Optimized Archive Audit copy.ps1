# Function to collect and process Exchange archive data
function Get-ArchiveReport {
    param(
        [string]$ExportPath = ".\ExchangeData",
        [switch]$CollectOnly,
        [switch]$ProcessOnly
    )

    if (!$ProcessOnly) {
        Write-Host "Collecting data from Exchange server..." -ForegroundColor Cyan
        
        # Create directory if it doesn't exist
        if (!(Test-Path $ExportPath)) {
            New-Item -ItemType Directory -Path $ExportPath | Out-Null
        }

        # Collect all required data in single calls and export to files
        Write-Host "Getting mailbox data..."
        $mailboxes = Get-Mailbox | Select-Object *
        $mailboxes | Export-Clixml "$ExportPath\mailboxes.xml"

        Write-Host "Getting mailbox statistics..."
        $stats = @()
        foreach ($mbx in $mailboxes) {
            Write-Host "  Processing stats for $($mbx.DisplayName)..." -ForegroundColor Gray
            try {
                $mbxStats = Get-MailboxStatistics $mbx.Identity -ErrorAction Stop
                $stats += $mbxStats
            }
            catch {
                Write-Warning "Could not get statistics for $($mbx.DisplayName): $_"
            }
        }
        $stats | Export-Clixml "$ExportPath\mailboxstats.xml"

        Write-Host "Getting archive statistics..."
        $archiveStats = @()
        foreach ($mbx in ($mailboxes | Where-Object {$_.ArchiveDatabase})) {
            Write-Host "  Processing archive stats for $($mbx.DisplayName)..." -ForegroundColor Gray
            try {
                $arcStats = Get-MailboxStatistics $mbx.Identity -Archive -ErrorAction Stop
                $archiveStats += $arcStats
            }
            catch {
                Write-Warning "Could not get archive statistics for $($mbx.DisplayName): $_"
            }
        }
        $archiveStats | Export-Clixml "$ExportPath\archivestats.xml"

        Write-Host "Data collection complete. Files saved to $ExportPath"
    }

    if (!$CollectOnly) {
        Write-Host "`nProcessing collected data..." -ForegroundColor Cyan

        # Import collected data
        $mailboxes = Import-Clixml "$ExportPath\mailboxes.xml"
        $stats = Import-Clixml "$ExportPath\mailboxstats.xml"
        $archiveStats = Import-Clixml "$ExportPath\archivestats.xml"

        # Process data locally
        $results = [System.Collections.Generic.List[PSObject]]::new()
        
        foreach ($mbx in $mailboxes) {
            $mbxStats = $stats | Where-Object { $_.MailboxGuid -eq $mbx.ExchangeGuid }
            $arcStats = $archiveStats | Where-Object { $_.MailboxGuid -eq $mbx.ExchangeGuid }

            # Function to safely convert size string to GB
            function Convert-ToGB {
                param($sizeString)
                if (!$sizeString) { return 0 }
                try {
                    if ($sizeString.ToString() -match "\(([0-9,]+) bytes\)") {
                        return [math]::Round([decimal]$matches[1].Replace(",","")/1GB, 4)
                    }
                    return 0
                } catch {
                    return 0
                }
            }

            $results.Add([PSCustomObject]@{
                DisplayName = $mbx.DisplayName
                PrimarySmtpAddress = $mbx.PrimarySmtpAddress
                MailboxType = $mbx.RecipientTypeDetails
                IsEnabled = $mbxStats.IsValid
                
                # Primary Mailbox Properties
                PrimaryMailboxDatabase = $mbx.Database
                PrimaryMailboxServer = $mbx.ServerName
                PrimaryMailboxItemCount = $mbxStats.ItemCount
                PrimaryMailboxSizeGB = Convert-ToGB $mbxStats.TotalItemSize
                PrimaryMailboxLastLogon = $mbxStats.LastLogonTime
                
                # Archive Properties
                ArchiveEnabled = [bool]$mbx.ArchiveDatabase
                ArchiveState = $mbx.ArchiveState
                ArchiveDatabase = $mbx.ArchiveDatabase
                ArchiveServer = if ($mbx.ArchiveDatabase) { 
                    (($mbx.ArchiveDatabase -split '\\')[0]) 
                } else { "N/A" }
                ArchivePolicy = if ($mbx.RetentionPolicy) { $mbx.RetentionPolicy } else { "No Policy" }
                ArchiveQuota = $mbx.ArchiveQuota
                ArchiveWarningQuota = $mbx.ArchiveWarningQuota
                ArchiveItemCount = if ($arcStats) { $arcStats.ItemCount } else { 0 }
                ArchiveSizeGB = Convert-ToGB $arcStats.TotalItemSize
                ArchiveCreationDate = $mbx.ArchiveCreationDate
            })
        }

        # Generate summary
        Write-Host "`nMailbox and Archive Summary:"
        Write-Host "----------------------------------------"
        Write-Host "Total Mailboxes: $($results.Count)"
        Write-Host "Mailboxes with Archives Enabled: $(($results | Where-Object {$_.ArchiveEnabled}).Count)"
        
        # Primary mailbox stats
        $totalPrimarySize = ($results | Measure-Object -Property PrimaryMailboxSizeGB -Sum).Sum
        $totalPrimaryItems = ($results | Measure-Object -Property PrimaryMailboxItemCount -Sum).Sum
        Write-Host "`nPrimary Mailbox Statistics:"
        Write-Host "Total Primary Size: $totalPrimarySize GB"
        Write-Host "Total Primary Items: $totalPrimaryItems"

        # Archive stats
        $totalArchiveSize = ($results | Measure-Object -Property ArchiveSizeGB -Sum).Sum
        $totalArchiveItems = ($results | Measure-Object -Property ArchiveItemCount -Sum).Sum
        Write-Host "`nArchive Statistics:"
        Write-Host "Total Archive Size: $totalArchiveSize GB"
        Write-Host "Total Archive Items: $totalArchiveItems"

        Write-Host "`nArchive Distribution by Database:"
        $results | 
            Where-Object {$_.ArchiveEnabled} | 
            Group-Object ArchiveDatabase | 
            ForEach-Object {
                $dbSize = ($_.Group | Measure-Object -Property ArchiveSizeGB -Sum).Sum
                $dbItems = ($_.Group | Measure-Object -Property ArchiveItemCount -Sum).Sum
                Write-Host "`nDatabase: $($_.Name)"
                Write-Host "  Size (GB): $dbSize"
                Write-Host "  Items: $dbItems"
                Write-Host "  Mailbox Count: $($_.Count)"
            }

        # Export results
        $date = Get-Date -Format "yyyyMMdd-HHmmss"
        $results | Export-CSV "ArchiveReport_$date.csv" -NoTypeInformation

        return $results
    }
}


# Example usage:
<#
# To collect data from Exchange server:
Get-ArchiveReport -CollectOnly -ExportPath "C:\ExchangeData"

# To process previously collected data:
Get-ArchiveReport -ProcessOnly -ExportPath "C:\ExchangeData"

# To do both in one go:
Get-ArchiveReport -ExportPath "C:\ExchangeData"
#>

Get-ArchiveReport -ExportPath "C:\ExchangeData"