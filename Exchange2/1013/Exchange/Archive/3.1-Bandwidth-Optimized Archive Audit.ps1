# First collect all needed data in one go to minimize remote calls
function Get-ArchiveReport {
    param(
        [string]$ExportPath = ".\ArchiveData",
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
        $mailboxes = Get-Mailbox -ResultSize Unlimited | Select-Object *
        $mailboxes | Export-Clixml "$ExportPath\mailboxes.xml"

        Write-Host "Getting mailbox statistics..."
        $stats = Get-MailboxStatistics -ResultSize Unlimited | Select-Object *
        $stats | Export-Clixml "$ExportPath\mailboxstats.xml"

        Write-Host "Getting archive statistics..."
        $archiveStats = Get-MailboxStatistics -ResultSize Unlimited -Archive -ErrorAction SilentlyContinue | 
            Select-Object *
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
                LastLogonTime = $mbxStats.LastLogonTime
            })
        }

        # Generate summary
        Write-Host "`nArchive Status Summary:"
        Write-Host "----------------------------------------"
        Write-Host "Total Mailboxes: $($results.Count)"
        Write-Host "Mailboxes with Archives Enabled: $(($results | Where-Object {$_.ArchiveEnabled}).Count)"
        Write-Host "Total Primary Mailbox Size (GB): $(($results | Measure-Object -Property PrimaryMailboxSizeGB -Sum).Sum)"
        Write-Host "Total Archive Size (GB): $(($results | Measure-Object -Property ArchiveSizeGB -Sum).Sum)"

        Write-Host "`nArchive Size by Database:"
        $results | 
            Where-Object {$_.ArchiveEnabled} | 
            Group-Object ArchiveDatabase | 
            ForEach-Object {
                $totalSize = ($_.Group | Measure-Object -Property ArchiveSizeGB -Sum).Sum
                Write-Host "Database: $($_.Name)"
                Write-Host "  Total Size (GB): $totalSize"
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