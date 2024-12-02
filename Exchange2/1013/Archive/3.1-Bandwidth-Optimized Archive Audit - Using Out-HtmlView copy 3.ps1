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

        # Get Exchange Server name from a mailbox database
        try {
            $firstMbx = Get-Mailbox -Identity "Administrator" -ErrorAction Stop
            $exchServer = $firstMbx.ServerName
            if ($exchServer -match '\.') {
                $exchangeFQDN = $exchServer
            }
            else {
                # Try to get FQDN if only hostname was returned
                $exchFQDN = [System.Net.Dns]::GetHostEntry($exchServer).HostName
                if ($exchFQDN) {
                    $exchangeFQDN = $exchFQDN
                }
                else {
                    $exchangeFQDN = $exchServer
                }
            }
        }
        catch {
            Write-Warning "Could not determine Exchange server FQDN: $_"
            $exchangeFQDN = "UnknownServer"
        }

        Write-Host "Connected to Exchange Server: $exchangeFQDN"

        # Initialize arrays with capacity for better performance
        Write-Host "Getting mailbox count..."
        $mbxCount = @(Get-Mailbox).Count
        # Or even simpler, since we're using modern PowerShell:
        # $stats = [System.Collections.Generic.List[PSObject]]::new($mbxCount)
        $archiveStats = [System.Collections.Generic.List[PSObject]]::new($mbxCount)

        # Collect all required data in single calls and export to files
        Write-Host "Getting mailbox data..."
        $mailboxes = @(Get-Mailbox | Select-Object *)
        $mailboxes | Export-Clixml "$ExportPath\mailboxes.xml"

        # Write-Host "Getting mailbox statistics..."
        # $progressCount = 0
        # foreach ($mbx in $mailboxes) {
        #     $progressCount++
        #     Write-Progress -Activity "Collecting Mailbox Statistics" -Status "$($mbx.DisplayName)" -PercentComplete (($progressCount / $mbxCount) * 100)
            
        #     try {
        #         $mbxStats = Get-MailboxStatistics $mbx.Identity -ErrorAction Stop
        #         [void]$stats.Add($mbxStats)
        #     }
        #     catch {
        #         Write-Warning "Could not get statistics for $($mbx.DisplayName): $_"
        #     }
        # }



        # Try bulk collection first, fall back to individual if it fails
        try {
            Write-Host "Attempting bulk statistics collection..."
            $stats = [System.Collections.Generic.List[PSObject]]::new()
            $allStats = Get-MailboxStatistics
            $stats.AddRange($allStats)
        }
        catch {
            Write-Warning "Bulk collection failed, falling back to individual collection: $_"
            $stats = [System.Collections.Generic.List[PSObject]]::new($mbxCount)
            $progressCount = 0
    
            foreach ($mbx in $mailboxes) {
                $progressCount++
                Write-Progress -Activity "Collecting Mailbox Statistics" -Status "$($mbx.DisplayName)" -PercentComplete (($progressCount / $mbxCount) * 100)
        
                try {
                    $mbxStats = Get-MailboxStatistics $mbx.Identity -ErrorAction Stop
                    [void]$stats.Add($mbxStats)
                }
                catch {
                    Write-Warning "Could not get statistics for $($mbx.DisplayName): $_"
                }
            }
        }

        Write-Progress -Activity "Collecting Mailbox Statistics" -Completed
        $stats | Export-Clixml "$ExportPath\mailboxstats.xml"

        Write-Host "Getting archive statistics..."
        $archiveMbx = @($mailboxes | Where-Object { $_.ArchiveDatabase })
        $progressCount = 0
        foreach ($mbx in $archiveMbx) {
            $progressCount++
            Write-Progress -Activity "Collecting Archive Statistics" -Status "$($mbx.DisplayName)" -PercentComplete (($progressCount / $archiveMbx.Count) * 100)
            
            try {
                $arcStats = Get-MailboxStatistics $mbx.Identity -Archive -ErrorAction Stop
                [void]$archiveStats.Add($arcStats)
            }
            catch {
                Write-Warning "Could not get archive statistics for $($mbx.DisplayName): $_"
            }
        }
        Write-Progress -Activity "Collecting Archive Statistics" -Completed
        $archiveStats | Export-Clixml "$ExportPath\archivestats.xml"

        Write-Host "Data collection complete. Files saved to $ExportPath"
        
        # Save the Exchange FQDN for processing phase
        $exchangeFQDN | Out-File "$ExportPath\server.txt"
    }

    if (!$CollectOnly) {
        Write-Host "`nProcessing collected data..." -ForegroundColor Cyan

        # Import collected data
        $mailboxes = Import-Clixml "$ExportPath\mailboxes.xml"
        $stats = Import-Clixml "$ExportPath\mailboxstats.xml"
        $archiveStats = Import-Clixml "$ExportPath\archivestats.xml"
        $exchangeFQDN = Get-Content "$ExportPath\server.txt"

        # Process data locally
        $results = [System.Collections.Generic.List[PSObject]]::new($mailboxes.Count)
        
        # Function to safely convert size string to GB
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

        foreach ($mbx in $mailboxes) {
            $mbxStats = $stats | Where-Object { $_.MailboxGuid -eq $mbx.ExchangeGuid }
            $arcStats = $archiveStats | Where-Object { $_.MailboxGuid -eq $mbx.ExchangeGuid }

            $results.Add([PSCustomObject]@{
                    DisplayName             = $mbx.DisplayName
                    PrimarySmtpAddress      = $mbx.PrimarySmtpAddress
                    MailboxType             = $mbx.RecipientTypeDetails
                    IsEnabled               = $mbxStats.IsValid
                    PrimaryMailboxDatabase  = $mbx.Database
                    PrimaryMailboxServer    = $mbx.ServerName
                    PrimaryMailboxItemCount = $mbxStats.ItemCount
                    PrimaryMailboxSizeGB    = Convert-ToGB $mbxStats.TotalItemSize
                    PrimaryMailboxLastLogon = $mbxStats.LastLogonTime
                    ArchiveEnabled          = [bool]$mbx.ArchiveDatabase
                    ArchiveState            = $mbx.ArchiveState
                    ArchiveDatabase         = $mbx.ArchiveDatabase
                    ArchiveServer           = $exchangeFQDN
                    ArchivePolicy           = if ($mbx.RetentionPolicy) { $mbx.RetentionPolicy } else { "No Policy" }
                    ArchiveQuota            = $mbx.ArchiveQuota
                    ArchiveWarningQuota     = $mbx.ArchiveWarningQuota
                    ArchiveItemCount        = if ($arcStats) { $arcStats.ItemCount } else { 0 }
                    ArchiveSizeGB           = Convert-ToGB $arcStats.TotalItemSize
                    ArchiveCreationDate     = $mbx.ArchiveCreationDate
                })
        }

        # Generate summary
        Write-Host "`nMailbox and Archive Summary:"
        Write-Host "----------------------------------------"
        Write-Host "Total Mailboxes: $($results.Count)"
        Write-Host "Mailboxes with Archives Enabled: $(($results | Where-Object {$_.ArchiveEnabled}).Count)"
        
        Write-Host "`nPrimary Mailbox Statistics:"
        Write-Host "Total Primary Size: $(($results | Measure-Object -Property PrimaryMailboxSizeGB -Sum).Sum) GB"
        Write-Host "Total Primary Items: $(($results | Measure-Object -Property PrimaryMailboxItemCount -Sum).Sum)"

        Write-Host "`nArchive Statistics:"
        Write-Host "Total Archive Size: $(($results | Measure-Object -Property ArchiveSizeGB -Sum).Sum) GB"
        Write-Host "Total Archive Items: $(($results | Measure-Object -Property ArchiveItemCount -Sum).Sum)"

        # Create filename with date and FQDN
        $date = Get-Date -Format "yyyyMMdd-HHmmss"
        $baseFileName = "ArchiveReport_{0}_{1}" -f $exchangeFQDN, $date

        # Export results
        $results | Export-CSV "$ExportPath\$baseFileName.csv" -NoTypeInformation
        $results | Out-HtmlView -FilePath "$ExportPath\$baseFileName.HTML" -Title "Exchange_Archive_Report_$exchangeFQDN`_$date"

        Write-Host "`nExport complete:"
        Write-Host "CSV: $ExportPath\$baseFileName.csv"
        Write-Host "HTML: $ExportPath\$baseFileName.html"

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