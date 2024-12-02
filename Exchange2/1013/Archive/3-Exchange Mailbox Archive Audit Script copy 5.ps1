# First, verify we have an active Exchange session
$exchangeSession = Get-PSSession | Where-Object {
    $_.ConfigurationName -eq "Microsoft.Exchange" -and $_.State -eq "Opened"
} | Select-Object -First 1

if (-not $exchangeSession) {
    Write-Error "No active Exchange session found. Please run the connection script first."
    return
}

Write-Host "Using existing Exchange session: $($exchangeSession.Id)" -ForegroundColor Green

# Get all mailboxes and process them locally to avoid multiple remote calls
$mailboxes = Invoke-Command -Session $exchangeSession -ScriptBlock {
    Get-Mailbox -ResultSize Unlimited
}

# Function to safely convert size string to GB with more decimal places
function Convert-ToGB {
    param($sizeString)
    if (!$sizeString) { return 0 }
    try {
        if ($sizeString.ToString() -match "\(([0-9,]+) bytes\)") {
            return [math]::Round([decimal]$matches[1].Replace(",","")/1GB, 4)
        }
        return 0
    } catch {
        Write-Warning "Could not convert size for: $sizeString"
        return 0
    }
}

# Process each mailbox
$results = $mailboxes | ForEach-Object {
    $mailbox = $_
    Write-Host "Processing mailbox: $($mailbox.DisplayName)" -ForegroundColor Cyan
    
    # Get statistics in a single remote call
    $stats = $archiveStats = $null
    $statsData = Invoke-Command -Session $exchangeSession -ScriptBlock {
        param($identity)
        $s = Get-MailboxStatistics $identity -ErrorAction SilentlyContinue
        $a = if ((Get-Mailbox $identity).ArchiveDatabase) {
            Get-MailboxStatistics $identity -Archive -ErrorAction SilentlyContinue
        }
        return @{
            Stats = $s
            ArchiveStats = $a
        }
    } -ArgumentList $mailbox.Identity
    
    $stats = $statsData.Stats
    $archiveStats = $statsData.ArchiveStats

    [PSCustomObject]@{
        DisplayName = $mailbox.DisplayName
        PrimarySmtpAddress = $mailbox.PrimarySmtpAddress
        MailboxType = $mailbox.RecipientTypeDetails
        IsEnabled = $stats.IsValid
        
        # Primary Mailbox Properties
        PrimaryMailboxDatabase = $mailbox.Database
        PrimaryMailboxServer = $mailbox.ServerName
        PrimaryMailboxItemCount = $stats.ItemCount
        PrimaryMailboxSizeGB = Convert-ToGB $stats.TotalItemSize
        
        # Archive Properties
        ArchiveEnabled = [bool]$mailbox.ArchiveDatabase
        ArchiveState = $mailbox.ArchiveState
        ArchiveDatabase = $mailbox.ArchiveDatabase
        ArchiveServer = if ($mailbox.ArchiveDatabase) { 
            (Invoke-Command -Session $exchangeSession -ScriptBlock {
                param($db)
                (Get-MailboxDatabase $db).Server
            } -ArgumentList $mailbox.ArchiveDatabase)
        } else { "N/A" }
        ArchivePolicy = if ($mailbox.RetentionPolicy) { $mailbox.RetentionPolicy } else { "No Policy" }
        ArchiveQuota = if ($mailbox.ArchiveQuota) { $mailbox.ArchiveQuota } else { "No Quota" }
        ArchiveWarningQuota = if ($mailbox.ArchiveWarningQuota) { $mailbox.ArchiveWarningQuota } else { "No Warning Quota" }
        ArchiveItemCount = if ($archiveStats) { $archiveStats.ItemCount } else { 0 }
        ArchiveSizeGB = Convert-ToGB $archiveStats.TotalItemSize
        ArchiveCreationDate = $mailbox.ArchiveCreationDate
        LastLogonTime = $stats.LastLogonTime
    }
}

# Rest of your export and summary code remains the same...

# Execute the script block in the existing session
$results = Invoke-Command -Session $exchangeSession -ScriptBlock $scriptBlock

# Rest of your script remains the same (exports and summaries)
# Get the server's FQDN
$serverFQDN = [System.Net.Dns]::GetHostEntry([string]$env:computername).HostName

# Create filename with date and FQDN
$date = Get-Date -Format "yyyyMMdd-HHmmss"
$baseFileName = "ArchiveAudit_{0}_{1}" -f $serverFQDN, $date

# Export results to CSV and HTML
$results | Export-CSV "$baseFileName.csv" -NoTypeInformation
$results | Out-HtmlView -FilePath "$baseFileName.HTML" -Title "NTI_RI_onprem_exchange_$serverFQDN`_$date"

Write-Host "`nReports exported to:"
Write-Host "CSV: $baseFileName.csv"
Write-Host "HTML: $baseFileName.html"


# Display comprehensive summary on screen
Write-Host "`nArchive Status Summary:"
Write-Host "----------------------------------------"
Write-Host "Overall Statistics:"
Write-Host "Total Mailboxes: $($results.Count)"
Write-Host "Mailboxes with Archives Enabled: $(($results | Where-Object {$_.ArchiveEnabled}).Count)"
Write-Host "Total Primary Mailbox Size (GB): $(($results | Measure-Object -Property PrimaryMailboxSizeGB -Sum).Sum)"
Write-Host "Total Archive Size (GB): $(($results | Measure-Object -Property ArchiveSizeGB -Sum).Sum)"
Write-Host "Total Primary Items: $(($results | Measure-Object -Property PrimaryMailboxItemCount -Sum).Sum)"
Write-Host "Total Archived Items: $(($results | Measure-Object -Property ArchiveItemCount -Sum).Sum)"

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

Write-Host "`nArchive Size by Server:"
Write-Host "----------------------------------------"
$results | 
    Where-Object {$_.ArchiveEnabled} | 
    Group-Object ArchiveServer | 
    ForEach-Object {
        $totalSize = ($_.Group | Measure-Object -Property ArchiveSizeGB -Sum).Sum
        $totalItems = ($_.Group | Measure-Object -Property ArchiveItemCount -Sum).Sum
        Write-Host "Server: $($_.Name)"
        Write-Host "  Total Size (GB): $totalSize"
        Write-Host "  Total Items: $totalItems"
        Write-Host "  Mailbox Count: $($_.Count)"
    }

# Write-Host "`nReport exported to: ArchiveAudit_$date.csv"