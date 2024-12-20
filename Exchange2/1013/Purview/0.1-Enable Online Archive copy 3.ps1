# Import required module for HTML reporting
Import-Module PSWriteHTML -ErrorAction Stop

function Write-OperationLog {
    param(
        [Parameter(Mandatory)]
        [string]$Message,
        
        [ValidateSet('Info', 'Warning', 'Error')]
        [string]$Level = 'Info'
    )
    
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $logMessage = "[$timestamp] [$Level] $Message"
    
    switch ($Level) {
        'Info' { Write-Host $logMessage -ForegroundColor Green }
        'Warning' { Write-Host $logMessage -ForegroundColor Yellow }
        'Error' { Write-Host $logMessage -ForegroundColor Red }
    }
}

function Connect-ExchangeOnlineWithCheck {
    [CmdletBinding()]
    param()
    
    Write-OperationLog "Checking Exchange Online connection..."
    
    $exchangeSession = Get-PSSession | Where-Object { 
        $_.ConfigurationName -eq 'Microsoft.Exchange' -and $_.State -eq 'Opened' 
    }
    
    if (-not $exchangeSession) {
        Write-OperationLog "No active Exchange Online session found. Connecting..." -Level Warning
        Connect-ExchangeOnline
        Write-OperationLog "Successfully connected to Exchange Online" -Level Info
    }
    else {
        Write-OperationLog "Active Exchange Online session found" -Level Info
    }
}

function Get-MailboxSizeReport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string[]]$Mailboxes,
        
        [Parameter()]
        [string]$OutputDir = ".\MailboxReports"
    )

    # Ensure output directory exists
    $outputDir = Join-Path $PSScriptRoot $OutputDir
    if (-not (Test-Path $outputDir)) {
        New-Item -ItemType Directory -Path $outputDir | Out-Null
    }

    # Connect to Exchange Online
    Connect-ExchangeOnlineWithCheck

    # Initialize results array
    $results = [System.Collections.Generic.List[object]]::new()

    foreach ($mailbox in $Mailboxes) {
        try {
            Write-OperationLog "Processing mailbox: $mailbox"
            
            # Get mailbox statistics
            $primaryStats = Get-MailboxStatistics -Identity $mailbox -ErrorAction Stop
            $archiveStats = Get-MailboxStatistics -Identity $mailbox -Archive -ErrorAction SilentlyContinue
            $mailboxInfo = Get-Mailbox -Identity $mailbox -ErrorAction Stop

            # Calculate sizes in GB and round to 2 decimal places
            $primarySizeGB = [math]::Round(($primaryStats.TotalItemSize.ToString().Split("(")[1].Split(" ")[0].Replace(",","")/1GB), 2)
            $primaryItemCount = $primaryStats.ItemCount
            
            # Check archive status from mailbox properties
            $archiveStatus = if ($mailboxInfo.ArchiveStatus -eq 'Active') { 
                "Enabled" 
            } elseif ($mailboxInfo.ArchiveDatabase) { 
                "Provisioning" 
            } else { 
                "Not Enabled" 
            }

            # Get archive info if it exists
            if ($archiveStats) {
                if ($archiveStats.TotalItemSize) {
                    $archiveSizeGB = [math]::Round(($archiveStats.TotalItemSize.ToString().Split("(")[1].Split(" ")[0].Replace(",","")/1GB), 2)
                } else {
                    $archiveSizeGB = 0
                }
                $archiveItemCount = $archiveStats.ItemCount
            } else {
                $archiveSizeGB = 0
                $archiveItemCount = 0
            }

            $resultObject = [PSCustomObject]@{
                Mailbox = $mailbox
                PrimaryCreationDate = $mailboxInfo.WhenCreated
                PrimarySizeGB = $primarySizeGB
                PrimaryItemCount = $primaryItemCount
                PrimaryLastLogon = $primaryStats.LastLogonTime
                ArchiveStatus = $archiveStatus
                ArchiveCreationDate = if ($mailboxInfo.ArchiveDatabase) { $mailboxInfo.WhenMailboxCreated } else { $null }
                ArchiveSizeGB = $archiveSizeGB
                ArchiveItemCount = $archiveItemCount
                TotalSizeGB = $primarySizeGB + $archiveSizeGB
                Status = "Success"
                Error = "-"
            }

            $results.Add($resultObject)
            Write-OperationLog "Successfully processed: $mailbox"
        }
        catch {
            Write-OperationLog "Error processing $mailbox : $_" -Level Error
            
            $resultObject = [PSCustomObject]@{
                Mailbox = $mailbox
                PrimaryCreationDate = $null
                PrimarySizeGB = 0
                PrimaryItemCount = 0
                PrimaryLastLogon = $null
                ArchiveStatus = "Error"
                ArchiveCreationDate = $null
                ArchiveSizeGB = 0
                ArchiveItemCount = 0
                TotalSizeGB = 0
                Status = "Failed"
                Error = $_.Exception.Message
            }
            
            $results.Add($resultObject)
        }
    }

    # Generate reports
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $htmlPath = Join-Path $outputDir "MailboxSizeReport_$timestamp.html"
    $csvPath = Join-Path $outputDir "MailboxSizeReport_$timestamp.csv"

    # Export to CSV
    $results | Export-Csv -Path $csvPath -NoTypeInformation

    # Calculate metadata
    $metadata = @{
        GeneratedBy = $env:USERNAME
        GeneratedOn = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        TotalMailboxes = $results.Count
        SuccessCount = ($results | Where-Object Status -eq "Success").Count
        FailureCount = ($results | Where-Object Status -eq "Failed").Count
        TotalPrimarySize = [math]::Round(($results | Measure-Object -Property PrimarySizeGB -Sum).Sum, 2)
        TotalArchiveSize = [math]::Round(($results | Measure-Object -Property ArchiveSizeGB -Sum).Sum, 2)
    }

    # Generate HTML Report
    New-HTML -Title "Mailbox Size Report" -FilePath $htmlPath -ShowHTML {
        New-HTMLSection -HeaderText "Report Summary" {
            New-HTMLPanel {
                New-HTMLText -Text @"
                <h3>Report Details</h3>
                <ul>
                    <li>Generated By: $($metadata.GeneratedBy)</li>
                    <li>Generated On: $($metadata.GeneratedOn)</li>
                    <li>Total Mailboxes Processed: $($metadata.TotalMailboxes)</li>
                    <li>Successful Operations: $($metadata.SuccessCount)</li>
                    <li>Failed Operations: $($metadata.FailureCount)</li>
                    <li>Total Primary Mailbox Size: $($metadata.TotalPrimarySize) GB</li>
                    <li>Total Archive Size: $($metadata.TotalArchiveSize) GB</li>
                </ul>
"@
            }
        }
        
        New-HTMLSection -HeaderText "Mailbox Size Details" {
            New-HTMLTable -DataTable $results -ScrollX -Buttons @('copyHtml5', 'excelHtml5', 'csvHtml5') -SearchBuilder {
                New-TableCondition -Name 'Status' -ComparisonType string -Operator eq -Value 'Failed' -BackgroundColor Salmon -Color Black
                New-TableCondition -Name 'Status' -ComparisonType string -Operator eq -Value 'Success' -BackgroundColor LightGreen -Color Black
                New-TableCondition -Name 'ArchiveStatus' -ComparisonType string -Operator eq -Value 'Enabled' -BackgroundColor LightBlue -Color Black
            }
        }
    }

    Write-OperationLog "Results exported to: $outputDir"
    return $results
}

# Example usage:
$mailboxes = @(
    "cc@tunngavik.com",
    "epo@tunngavik.com",
    "hrclerk@tunngavik.com",
    "Invoices@tunngavik.com",
    "payroll@tunngavik.com",
    "relief@tunngavik.com",
    "ri-support@tunngavik.com"
)

Get-MailboxSizeReport -Mailboxes $mailboxes -OutputDir ".\MailboxSizeReports"