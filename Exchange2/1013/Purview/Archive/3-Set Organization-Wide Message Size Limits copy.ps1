function Set-OrganizationMessageSizeLimits {
    [CmdletBinding()]
    param (
        [Parameter()]
        [ValidateRange(1,150)]
        [int]$MaxSizeMB = 150,
        
        [Parameter()]
        [switch]$IncludeExistingMailboxes,
        
        [Parameter()]
        [string]$LogPath = ".\MessageSizeLimits_Log"
    )
    
    # Create log directory silently
    $null = New-Item -ItemType Directory -Force -Path $LogPath
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $logFile = Join-Path $LogPath "SizeLimits_Changes_$timestamp.html"
    
    # Initialize results array with capacity
    $results = [System.Collections.Generic.List[PSCustomObject]]::new(1000)
    
    try {
        # Step 1: Set mailbox plan defaults (bulk operation)
        $plans = Get-MailboxPlan
        foreach ($plan in $plans) {
            $results.Add([PSCustomObject]@{
                Component = "Mailbox Plan: $($plan.Name)"
                MaxSendSize = "$($MaxSizeMB)MB"
                MaxReceiveSize = "$($MaxSizeMB)MB"
                Status = "Updated"
            })
            
            $null = Set-MailboxPlan $plan.Name -MaxSendSize "$($MaxSizeMB)MB" -MaxReceiveSize "$($MaxSizeMB)MB"
        }
        
        # Step 2: Update existing mailboxes if requested
        if ($IncludeExistingMailboxes) {
            # Get all mailboxes in one query
            $mailboxes = Get-Mailbox -ResultSize Unlimited
            $total = $mailboxes.Count
            $batchSize = 100
            $processed = 0
            
            # Process in batches for better performance
            for ($i = 0; $i -lt $total; $i += $batchSize) {
                $batch = $mailboxes | Select-Object -Skip $i -First $batchSize
                
                # Update progress every 100 mailboxes only
                $processed += $batchSize
                if ($processed % 100 -eq 0) {
                    Write-Progress -Activity "Updating mailboxes" -PercentComplete (($processed/$total)*100) -Status "$processed of $total"
                }
                
                foreach ($mbx in $batch) {
                    try {
                        $null = Set-Mailbox $mbx.Identity -MaxSendSize "$($MaxSizeMB)MB" -MaxReceiveSize "$($MaxSizeMB)MB" -ErrorAction Stop
                        $results.Add([PSCustomObject]@{
                            Component = "Mailbox: $($mbx.PrimarySmtpAddress)"
                            MaxSendSize = "$($MaxSizeMB)MB"
                            MaxReceiveSize = "$($MaxSizeMB)MB"
                            Status = "Updated"
                        })
                    }
                    catch {
                        $results.Add([PSCustomObject]@{
                            Component = "Mailbox: $($mbx.PrimarySmtpAddress)"
                            MaxSendSize = "$($MaxSizeMB)MB"
                            MaxReceiveSize = "$($MaxSizeMB)MB"
                            Status = "Failed: $($_.Exception.Message)"
                        })
                    }
                }
            }
            Write-Progress -Activity "Updating mailboxes" -Completed
        }
        
        # Generate HTML report (one-time operation)
        New-HTML -Title "Message Size Limits Configuration Report" -FilePath $logFile -ShowHTML {
            New-HTMLSection -HeaderText "Configuration Summary" {
                New-HTMLPanel {
                    New-HTMLText -Text @"
                    <h3>Overview</h3>
                    <ul>
                        <li>Maximum Size Set: $($MaxSizeMB) MB</li>
                        <li>Include Existing Mailboxes: $($IncludeExistingMailboxes)</li>
                        <li>Total Updates: $($results.Count)</li>
                        <li>Execution Time: $(Get-Date)</li>
                    </ul>
"@
                }
            }
            
            New-HTMLSection -HeaderText "Configuration Results" {
                New-HTMLTable -DataTable $results -ScrollX
            }
        }
        
        return $results
    }
    catch {
        Write-Error "Error during configuration: $_"
        throw
    }
}

# Example usage
$params = @{
    MaxSizeMB = 150
    IncludeExistingMailboxes = $true
    LogPath = ".\MessageSizeLimits_Log"
}

$results = Set-OrganizationMessageSizeLimits @params