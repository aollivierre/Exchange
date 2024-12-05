function Get-ArchiveReport {
    param(
        [string]$ExportPath = ".\ExchangeData",
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

    Write-StatusMessage "Collecting data from Exchange server..."
        
    if (!(Test-Path $ExportPath)) {
        New-Item -ItemType Directory -Path $ExportPath | Out-Null
    }

    # Get Exchange Server FQDN from database - single call
    try {
        $exchangeFQDN = (Get-MailboxDatabase -Status | Select-Object -First 1 -ExpandProperty Server)
    }
    catch {
        Write-Warning "Could not determine Exchange server FQDN: $_"
        $exchangeFQDN = $env:COMPUTERNAME
    }

    Write-StatusMessage "Connected to Exchange Server: $exchangeFQDN"

    # Get all mailboxes with minimal required properties - single call
    Write-StatusMessage "Getting mailbox data..."
    $mailboxes = Get-Mailbox -ResultSize Unlimited | Select-Object @(
        'DisplayName',
        'PrimarySmtpAddress',
        'RecipientTypeDetails',
        'Database',
        'ServerName',
        'ExchangeGuid',
        'ArchiveDatabase',
        'ArchiveState',
        'RetentionPolicy',
        'ArchiveQuota',
        'ArchiveWarningQuota',
        'ArchiveCreationDate'
    )

    Write-StatusMessage "Processing mailbox statistics..."
    
    # Pre-create lookup for archive-enabled mailboxes to optimize filtering
    $archiveMailboxes = $mailboxes | Where-Object { $_.ArchiveDatabase } | 
    Select-Object -ExpandProperty PrimarySmtpAddress

    # Preallocate the results array with the exact size we need
    $results = [System.Collections.Generic.List[PSCustomObject]]::new($mailboxes.Count)
    
    # Process each mailbox
    foreach ($mailbox in $mailboxes) {
        try {
            # Get primary stats
            $stats = Get-MailboxStatistics $mailbox.PrimarySmtpAddress -ErrorAction Stop
            $lastLogon = $stats.LastLogonTime
            $neverLoggedOn = $false
            
            # Get archive stats only if needed
            $archiveStats = if ($mailbox.ArchiveDatabase) {
                Get-MailboxStatistics $mailbox.PrimarySmtpAddress -Archive -ErrorAction Stop
            }

            $results.Add([PSCustomObject]@{
                    DisplayName             = $mailbox.DisplayName
                    PrimarySmtpAddress      = $mailbox.PrimarySmtpAddress
                    LastLogonTime           = $lastLogon
                    PrimaryMailboxItemCount = $stats.ItemCount
                    PrimaryMailboxSizeGB    = Convert-ToGB $stats.TotalItemSize
                    ArchiveItemCount        = if ($archiveStats) { $archiveStats.ItemCount } else { 0 }
                    ArchiveSizeGB           = Convert-ToGB $archiveStats.TotalItemSize
                    MailboxType             = $mailbox.RecipientTypeDetails
                    IsEnabled               = $stats.IsValid
                    PrimaryMailboxDatabase  = $mailbox.Database
                    PrimaryMailboxServer    = $mailbox.ServerName
                    ArchiveEnabled          = [bool]$mailbox.ArchiveDatabase
                    ArchiveState            = $mailbox.ArchiveState
                    ArchiveDatabase         = $mailbox.ArchiveDatabase
                    ArchiveServer           = $exchangeFQDN
                    ArchivePolicy           = if ($mailbox.RetentionPolicy) { $mailbox.RetentionPolicy } else { "No Policy" }
                    ArchiveQuota            = $mailbox.ArchiveQuota
                    ArchiveWarningQuota     = $mailbox.ArchiveWarningQuota
                    ArchiveCreationDate     = $mailbox.ArchiveCreationDate
                    NeverLoggedOn           = $false
                })
        }
        catch {
            if ($_.Exception.Message -match "hasn't logged on") {
                Write-Warning "User hasn't logged on yet: $($mailbox.DisplayName)"
                
                $results.Add([PSCustomObject]@{
                        DisplayName             = $mailbox.DisplayName
                        PrimarySmtpAddress      = $mailbox.PrimarySmtpAddress
                        LastLogonTime           = "Never"
                        PrimaryMailboxItemCount = 0
                        PrimaryMailboxSizeGB    = 0
                        ArchiveItemCount        = 0
                        ArchiveSizeGB           = 0
                        MailboxType             = $mailbox.RecipientTypeDetails
                        IsEnabled               = "Never Logged On"
                        PrimaryMailboxDatabase  = $mailbox.Database
                        PrimaryMailboxServer    = $mailbox.ServerName
                        ArchiveEnabled          = [bool]$mailbox.ArchiveDatabase
                        ArchiveState            = $mailbox.ArchiveState
                        ArchiveDatabase         = $mailbox.ArchiveDatabase
                        ArchiveServer           = $exchangeFQDN
                        ArchivePolicy           = if ($mailbox.RetentionPolicy) { $mailbox.RetentionPolicy } else { "No Policy" }
                        ArchiveQuota            = $mailbox.ArchiveQuota
                        ArchiveWarningQuota     = $mailbox.ArchiveWarningQuota
                        ArchiveCreationDate     = $mailbox.ArchiveCreationDate
                        NeverLoggedOn           = $true
                    })
            }
            else {
                Write-Warning "Error processing mailbox $($mailbox.DisplayName): $_"
            }
        }
    }

    $finalResults = $results.ToArray()

    if (!$QuietMode) {
        # Calculate statistics using efficient methods
        $archiveEnabled = @($finalResults | Where-Object { $_.ArchiveEnabled }).Count
        $neverLoggedOn = @($finalResults | Where-Object { $_.NeverLoggedOn }).Count
        $primaryStats = $finalResults | Measure-Object -Property PrimaryMailboxSizeGB, PrimaryMailboxItemCount -Sum
        $archiveStats = $finalResults | Measure-Object -Property ArchiveSizeGB, ArchiveItemCount -Sum

        Write-Host "`nMailbox and Archive Summary:"
        Write-Host "----------------------------------------"
        Write-Host "Total Mailboxes: $($finalResults.Count)"
        Write-Host "Never Logged On: $neverLoggedOn"
        Write-Host "Mailboxes with Archives Enabled: $archiveEnabled"
        
        Write-Host "`nPrimary Mailbox Statistics:"
        Write-Host "Total Primary Size: $($primaryStats[0].Sum) GB"
        Write-Host "Total Primary Items: $($primaryStats[1].Sum)"

        Write-Host "`nArchive Statistics:"
        Write-Host "Total Archive Size: $($archiveStats[0].Sum) GB"
        Write-Host "Total Archive Items: $($archiveStats[1].Sum)"

        Write-Host "`nLargest Archives:"
        Write-Host "----------------------------------------"
        $finalResults | 
        Where-Object { $_.ArchiveSizeGB -gt 0 } |
        Sort-Object ArchiveSizeGB -Descending |
        Select-Object -First 5 |
        ForEach-Object {
            Write-Host "$($_.DisplayName) - $($_.ArchiveSizeGB) GB ($($_.ArchiveItemCount) items)"
        }

        Write-Host "`nArchive Size by Database:"
        Write-Host "----------------------------------------"
        $finalResults | 
        Where-Object { $_.ArchiveEnabled } | 
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

    $finalResults | Export-CSV "$ExportPath\$baseFileName.csv" -NoTypeInformation
    $finalResults | Out-HtmlView -FilePath "$ExportPath\$baseFileName.HTML" -Title "Exchange_Archive_Report_$exchangeFQDN`_$date"

    Write-StatusMessage "`nExport complete:"
    Write-StatusMessage "CSV: $ExportPath\$baseFileName.csv"
    Write-StatusMessage "HTML: $ExportPath\$baseFileName.html"

    return $finalResults
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