function Get-ArchiveReport {
    param(
        [string]$ExportPath = ".\ExchangeData",
        [switch]$CollectOnly,
        [switch]$ProcessOnly,
        [switch]$QuietMode
    )

    function Write-StatusMessage {
        param([string]$Message)
        if (!$QuietMode) {
            Write-Host $Message -ForegroundColor Cyan
        }
    }

    function Convert-ToGB {
        param($sizeString)
        if (!$sizeString) { return 0 }
        try {
            if ($sizeString.ToString() -match "\(([0-9,]+) bytes\)") {
                return [math]::Round([decimal]$matches[1].Replace(",", "") / 1GB, 4)
            }
            return 0
        }
        catch {
            return 0
        }
    }

    if (!$ProcessOnly) {
        Write-StatusMessage "Collecting data from Exchange server..."
        
        if (!(Test-Path $ExportPath)) {
            New-Item -ItemType Directory -Path $ExportPath | Out-Null
        }

        # Better Exchange Server FQDN detection
        try {
            # Try to get FQDN from the first mailbox database
            $db = Get-MailboxDatabase | Select-Object -First 1
            if ($db) {
                $exchServer = $db.Server
                $exchangeFQDN = if ($exchServer -match '\.') {
                    $exchServer
                }
                else {
                    $(try { [System.Net.Dns]::GetHostEntry($exchServer).HostName } catch { $exchServer })
                }
            }
            else {
                throw "No mailbox database found"
            }
        }
        catch {
            Write-Warning "Could not determine Exchange server FQDN: $_"
            $exchangeFQDN = "UnknownServer"
        }

        Write-StatusMessage "Connected to Exchange Server: $exchangeFQDN"

        # Collect and process data in a single pass
        Write-StatusMessage "Collecting mailbox and statistics data..."
        $results = Get-Mailbox -ResultSize Unlimited | ForEach-Object {
            $mailbox = $_
            $stats = $null
            $archiveStats = $null
            $lastLogon = $null
            
            try {
                $stats = Get-MailboxStatistics $mailbox.Identity -ErrorAction Stop
                $lastLogon = $stats.LastLogonTime
            }
            catch {
                if ($_.ToString() -match "hasn't logged on") {
                    Write-Warning "User hasn't logged on yet: $($mailbox.DisplayName)"
                    $lastLogon = "Never"
                }
                else {
                    Write-Warning "Could not get statistics for mailbox: $($mailbox.DisplayName)"
                }
            }

            try {
                if ($mailbox.ArchiveDatabase) {
                    $archiveStats = Get-MailboxStatistics $mailbox.Identity -Archive -ErrorAction Stop
                }
            }
            catch {
                Write-Warning "Could not get archive statistics for mailbox: $($mailbox.DisplayName)"
            }
            
            [PSCustomObject]@{
                DisplayName = $mailbox.DisplayName
                PrimarySmtpAddress = $mailbox.PrimarySmtpAddress
                LastLogonTime = $lastLogon
                PrimaryMailboxItemCount = $stats.ItemCount
                PrimaryMailboxSizeGB = Convert-ToGB $stats.TotalItemSize
                ArchiveItemCount = if ($archiveStats) { $archiveStats.ItemCount } else { 0 }
                ArchiveSizeGB = Convert-ToGB $archiveStats.TotalItemSize
                MailboxType = $mailbox.RecipientTypeDetails
                IsEnabled = if ($stats) { $stats.IsValid } else { "Unknown" }
                PrimaryMailboxDatabase = $mailbox.Database
                PrimaryMailboxServer = $mailbox.ServerName
                ArchiveEnabled = [bool]$mailbox.ArchiveDatabase
                ArchiveState = $mailbox.ArchiveState
                ArchiveDatabase = $mailbox.ArchiveDatabase
                ArchiveServer = $exchangeFQDN
                ArchivePolicy = if ($mailbox.RetentionPolicy) { $mailbox.RetentionPolicy } else { "No Policy" }
                ArchiveQuota = $mailbox.ArchiveQuota
                ArchiveWarningQuota = $mailbox.ArchiveWarningQuota
                ArchiveCreationDate = $mailbox.ArchiveCreationDate
            }
        }

        # Export the results
        $results | Export-Clixml "$ExportPath\results.xml"
        $exchangeFQDN | Out-File "$ExportPath\server.txt"
        
        Write-StatusMessage "Data collection complete. Files saved to $ExportPath"
    }

    if (!$CollectOnly) {
        Write-StatusMessage "`nProcessing collected data..."

        $results = Import-Clixml "$ExportPath\results.xml"
        $exchangeFQDN = Get-Content "$ExportPath\server.txt"

        if (!$QuietMode) {
            Write-Host "`nMailbox and Archive Summary:"
            Write-Host "----------------------------------------"
            Write-Host "Total Mailboxes: $($results.Count)"
            Write-Host "Never Logged On: $(($results | Where-Object {$_.LastLogonTime -eq 'Never'}).Count)"
            Write-Host "Mailboxes with Archives Enabled: $(($results | Where-Object {$_.ArchiveEnabled}).Count)"
            
            $primaryStats = $results | Measure-Object -Property PrimaryMailboxSizeGB, PrimaryMailboxItemCount -Sum
            $archiveStats = $results | Measure-Object -Property ArchiveSizeGB, ArchiveItemCount -Sum
            
            Write-Host "`nPrimary Mailbox Statistics:"
            Write-Host "Total Primary Size: $($primaryStats[0].Sum) GB"
            Write-Host "Total Primary Items: $($primaryStats[1].Sum)"

            Write-Host "`nArchive Statistics:"
            Write-Host "Total Archive Size: $($archiveStats[0].Sum) GB"
            Write-Host "Total Archive Items: $($archiveStats[1].Sum)"

            Write-Host "`nArchive Size by Database:"
            Write-Host "----------------------------------------"
            $results | 
                Where-Object {$_.ArchiveEnabled} | 
                Group-Object ArchiveDatabase | 
                ForEach-Object {
                    $totalSize = ($_.Group | Measure-Object -Property ArchiveSizeGB -Sum).Sum
                    $totalItems = ($_.Group | Measure-Object -Property ArchiveItemCount -Sum).Sum
                    Write-Host "Database: $($_.Name)"
                    Write-Host "  Total Size (GB): $totalSize"
                    Write-Host "  Total Items: $totalItems"
                    Write-Host "  Mailbox Count: $($_.Count)"
                }
        }

        $date = Get-Date -Format "yyyyMMdd-HHmmss"
        $baseFileName = "ArchiveReport_{0}_{1}" -f $exchangeFQDN, $date

        $results | Export-CSV "$ExportPath\$baseFileName.csv" -NoTypeInformation
        $results | Out-HtmlView -FilePath "$ExportPath\$baseFileName.HTML" -Title "Exchange_Archive_Report_$exchangeFQDN`_$date"

        Write-StatusMessage "`nExport complete:"
        Write-StatusMessage "CSV: $ExportPath\$baseFileName.csv"
        Write-StatusMessage "HTML: $ExportPath\$baseFileName.html"

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