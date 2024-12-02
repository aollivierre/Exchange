# Function to disable archiving while preserving archive mailbox
function Disable-MailboxArchiving {
    param(
        [Parameter(Mandatory=$false)]
        [string[]]$Mailboxes,
        [switch]$AllMailboxes,
        [switch]$WhatIf
    )

    if (!$AllMailboxes -and !$Mailboxes) {
        Write-Host "You must either specify mailboxes or use -AllMailboxes switch" -ForegroundColor Yellow
        return
    }

    # Initialize results array with capacity
    $results = [System.Collections.Generic.List[PSObject]]::new()
    
    try {
        # Get mailboxes based on parameter
        $mailboxesToProcess = if ($AllMailboxes) {
            Get-Mailbox -ResultSize Unlimited
        } else {
            Get-Mailbox -Identity $Mailboxes
        }

        # Show preview of mailboxes to be processed
        Write-Host "`nMailboxes to be processed:" -ForegroundColor Cyan
        $mailboxPreview = $mailboxesToProcess | Select-Object DisplayName, PrimarySmtpAddress, ArchiveDatabase, RetentionPolicy |
            Format-Table -AutoSize
        $mailboxPreview

        # Count statistics
        $totalMailboxes = ($mailboxesToProcess | Measure-Object).Count
        $archiveEnabledCount = ($mailboxesToProcess | Where-Object {$_.ArchiveDatabase} | Measure-Object).Count

        Write-Host "Summary:" -ForegroundColor Cyan
        Write-Host "Total Mailboxes: $totalMailboxes"
        Write-Host "Mailboxes with Archives: $archiveEnabledCount"
        
        # Ask for confirmation
        $confirmation = Read-Host "`nDo you want to proceed with disabling archive policies for these mailboxes? (Yes/No)"
        if ($confirmation -notmatch '^[Yy]') {
            Write-Host "Operation cancelled by user" -ForegroundColor Yellow
            return
        }

        # Process mailboxes
        $i = 0
        foreach ($mbx in $mailboxesToProcess) {
            $i++
            Write-Progress -Activity "Processing Mailboxes" -Status "$i of $totalMailboxes" -PercentComplete (($i / $totalMailboxes) * 100)
            
            Write-Host "`nProcessing ($i/$totalMailboxes): $($mbx.DisplayName) - $($mbx.PrimarySmtpAddress)" -ForegroundColor Cyan
            Write-Host "Current Archive Database: $($mbx.ArchiveDatabase)"
            Write-Host "Current Retention Policy: $($mbx.RetentionPolicy)"

            # Create result object
            $result = [PSCustomObject]@{
                Mailbox = $mbx.DisplayName
                Email = $mbx.PrimarySmtpAddress
                ArchiveDatabase = $mbx.ArchiveDatabase
                OriginalRetentionPolicy = $mbx.RetentionPolicy
                NewRetentionPolicy = $null
                Status = $null
                TimeStamp = (Get-Date)
                Notes = $null
            }

            # If there's an archive database, the archive is enabled
            if ($mbx.ArchiveDatabase) {
                if (!$WhatIf) {
                    Write-Host "Disabling archiving policy..." -ForegroundColor Yellow
                    try {
                        # Remove retention policy (stops new items from being archived)
                        Set-Mailbox $mbx.Identity -RetentionPolicy $null -WarningAction SilentlyContinue
                        
                        # Get updated mailbox info
                        $updatedMbx = Get-Mailbox $mbx.Identity
                        
                        $result.NewRetentionPolicy = $updatedMbx.RetentionPolicy
                        $result.Status = "Archive Policies Disabled"
                        $result.Notes = "Archive mailbox preserved but new archiving stopped"
                        
                        Write-Host "Successfully disabled archiving policy while preserving archive mailbox" -ForegroundColor Green
                    }
                    catch {
                        $result.Status = "Error"
                        $result.Notes = $_.Exception.Message
                        Write-Warning "Error processing $($mbx.DisplayName): $_"
                    }
                } else {
                    Write-Host "WhatIf: Would disable archiving policy" -ForegroundColor Cyan
                    $result.Status = "WhatIf - Would Disable"
                    $result.Notes = "WhatIf mode - no changes made"
                }
            } else {
                Write-Host "Skipping - No archive mailbox found" -ForegroundColor Yellow
                $result.Status = "Skipped"
                $result.Notes = "No archive mailbox exists"
            }

            # Add to results list
            $results.Add($result)
        }

        # Export results to CSV
        $date = Get-Date -Format "yyyyMMdd-HHmmss"
        $fileName = "ArchiveDisabled_$date.csv"
        $results | Export-CSV $fileName -NoTypeInformation
        
        # Display summary
        Write-Host "`nOperation Complete!" -ForegroundColor Green
        Write-Host "Results exported to: $fileName"
        Write-Host "`nSummary of operations:"
        $results | Group-Object Status | Select-Object @{
            Name='Status'; Expression={$_.Name}
        }, @{
            Name='Count'; Expression={$_.Count}
        } | Format-Table -AutoSize

        Write-Host "`nDetailed results:"
        $results | Format-Table Mailbox, Email, Status, Notes -AutoSize
    }
    catch {
        Write-Error "An error occurred: $_"
    }
    finally {
        Write-Progress -Activity "Processing Mailboxes" -Completed
    }
    
    return $results
}

# Example usage:
<#
# For specific mailboxes:
Disable-MailboxArchiving -Mailboxes "user1@domain.com", "user2@domain.com" -WhatIf

# For all mailboxes:
Disable-MailboxArchiving -AllMailboxes -WhatIf

# To actually make changes, remove -WhatIf:
Disable-MailboxArchiving -AllMailboxes
#>


# Disable-MailboxArchiving -AllMailboxes -WhatIf
Disable-MailboxArchiving -AllMailboxes