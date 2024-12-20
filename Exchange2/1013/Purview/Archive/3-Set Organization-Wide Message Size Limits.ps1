function Set-OrganizationMessageSizeLimits {
    param (
        [Parameter()]
        [ValidateRange(1,150)]
        [int]$MaxSizeMB = 150,
        
        [Parameter()]
        [switch]$IncludeExistingMailboxes,
        
        [Parameter()]
        [string]$LogPath = ".\MessageSizeLimits_Log"
    )
    
    # Create log directory
    $null = New-Item -ItemType Directory -Force -Path $LogPath
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $logFile = Join-Path $LogPath "SizeLimits_Changes_$timestamp.html"
    
    # Initialize results collection
    $results = [System.Collections.ArrayList]::new()
    
    try {
        # Step 1: Set mailbox plan defaults
        Write-Host "Setting mailbox plan defaults..." -ForegroundColor Cyan
        $plans = Get-MailboxPlan
        foreach ($plan in $plans) {
            $null = $results.Add([PSCustomObject]@{
                Component = "Mailbox Plan: $($plan.Name)"
                MaxSendSize = "$($MaxSizeMB)MB"
                MaxReceiveSize = "$($MaxSizeMB)MB"
                Status = "Processing"
            })
            
            Set-MailboxPlan $plan.Name -MaxSendSize "$($MaxSizeMB)MB" -MaxReceiveSize "$($MaxSizeMB)MB"
            
            $results[-1].Status = "Updated"
        }
        
        # Step 2: Update existing mailboxes if requested
        if ($IncludeExistingMailboxes) {
            Write-Host "Updating existing mailboxes..." -ForegroundColor Cyan
            $mailboxes = Get-Mailbox -ResultSize Unlimited
            $totalMailboxes = $mailboxes.Count
            $current = 0
            
            foreach ($mbx in $mailboxes) {
                $current++
                Write-Progress -Activity "Updating mailboxes" -Status "$current of $totalMailboxes" -PercentComplete (($current / $totalMailboxes) * 100)
                
                $null = $results.Add([PSCustomObject]@{
                    Component = "Mailbox: $($mbx.PrimarySmtpAddress)"
                    MaxSendSize = "$($MaxSizeMB)MB"
                    MaxReceiveSize = "$($MaxSizeMB)MB"
                    Status = "Processing"
                })
                
                try {
                    Set-Mailbox $mbx.Identity -MaxSendSize "$($MaxSizeMB)MB" -MaxReceiveSize "$($MaxSizeMB)MB"
                    $results[-1].Status = "Updated"
                }
                catch {
                    $results[-1].Status = "Failed: $($_.Exception.Message)"
                }
            }
        }
        
        # Generate HTML report
        New-HTML -Title "Message Size Limits Configuration Report" -FilePath $logFile -ShowHTML {
            New-HTMLSection -HeaderText "Configuration Summary" {
                New-HTMLPanel {
                    New-HTMLText -Text @"
                    <h3>Overview</h3>
                    <ul>
                        <li>Maximum Size Set: $($MaxSizeMB) MB</li>
                        <li>Include Existing Mailboxes: $($IncludeExistingMailboxes)</li>
                        <li>Execution Time: $(Get-Date)</li>
                    </ul>
"@
                }
            }
            
            New-HTMLSection -HeaderText "Configuration Results" {
                New-HTMLTable -DataTable $results -ScrollX
            }
        }
        
        Write-Host "`nConfiguration completed successfully!" -ForegroundColor Green
        Write-Host "Report generated at: $logFile" -ForegroundColor Cyan
        
        return $results
        
    }
    catch {
        Write-Error "Error during configuration: $_"
        throw
    }
}

# Example usage
$params = @{
    MaxSizeMB = 150  # Maximum allowed in Exchange Online
    IncludeExistingMailboxes = $true  # Set to update existing mailboxes
    LogPath = ".\MessageSizeLimits_Log"
}

Set-OrganizationMessageSizeLimits @params