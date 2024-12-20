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

function Enable-MailboxArchive {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$MailboxIdentity
    )

    try {
        $mailbox = Get-Mailbox -Identity $MailboxIdentity -ErrorAction Stop
        
        if ($mailbox.ArchiveStatus -eq "Active") {
            Write-OperationLog "Archive already enabled for $MailboxIdentity" -Level Warning
            return [PSCustomObject]@{
                Mailbox = $MailboxIdentity
                Status = "Already Enabled"
                ArchiveStatus = $mailbox.ArchiveStatus
                Error = $null
            }
        }

        Enable-Mailbox -Identity $MailboxIdentity -Archive -ErrorAction Stop
        Start-Sleep -Seconds 2  # Brief pause to allow for backend processing
        
        $updatedMailbox = Get-Mailbox -Identity $MailboxIdentity
        $status = if ($updatedMailbox.ArchiveStatus -eq "Active") { "Enabled" } else { "Failed" }
        
        Write-OperationLog "Archive $status for $MailboxIdentity" -Level Info
        
        return [PSCustomObject]@{
            Mailbox = $MailboxIdentity
            Status = $status
            ArchiveStatus = $updatedMailbox.ArchiveStatus
            Error = $null
        }
    }
    catch {
        Write-OperationLog "Failed to enable archive for $MailboxIdentity : $_" -Level Error
        return [PSCustomObject]@{
            Mailbox = $MailboxIdentity
            Status = "Failed"
            ArchiveStatus = "Error"
            Error = $_.Exception.Message
        }
    }
}

function Export-ArchiveReport {
    param (
        [Parameter(Mandatory)]
        [array]$Results,
        [Parameter(Mandatory)]
        [string]$OutputPath
    )

    $reportParams = @{
        FilePath = $OutputPath
        ShowHTML = $true
        Title = "Online Archive Enablement Report"
    }

    New-HTML @reportParams {
        New-HTMLSection -HeaderText "Archive Enablement Results" {
            New-HTMLTable -DataTable $Results {
                New-TableCondition -Name 'Status' -ComparisonType string -Operator eq -Value 'Failed' -BackgroundColor '#ffcdd2' -Color '#000000'
                New-TableCondition -Name 'Status' -ComparisonType string -Operator eq -Value 'Enabled' -BackgroundColor '#c8e6c9' -Color '#000000'
            }
        }
    }
}

# Main execution
$mailboxes = @(
    "cc@tunngavik.com",
    "epo@tunngavik.com",
    "hrclerk@tunngavik.com",
    "Invoices@tunngavik.com",
    "payroll@tunngavik.com",
    "relief@tunngavik.com",
    "ri-support@tunngavik.com"
)

$outputDir = Join-Path $PSScriptRoot "ArchiveReports_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
if (-not (Test-Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir | Out-Null
}

# Connect to Exchange Online
Connect-ExchangeOnlineWithCheck

# Process mailboxes
$results = [System.Collections.Generic.List[object]]::new()

foreach ($mailbox in $mailboxes) {
    $result = Enable-MailboxArchive -MailboxIdentity $mailbox
    $results.Add($result)
}

# Export results
$csvPath = Join-Path $outputDir "archive_results.csv"
$htmlPath = Join-Path $outputDir "archive_report.html"

$results | Export-Csv -Path $csvPath -NoTypeInformation
Export-ArchiveReport -Results $results -OutputPath $htmlPath

Write-OperationLog "Results exported to: $outputDir"