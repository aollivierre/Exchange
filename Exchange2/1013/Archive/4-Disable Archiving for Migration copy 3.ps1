# Function to disable archiving while preserving archive mailbox
function Disable-MailboxArchiving {
    param(
        [Parameter(Mandatory=$true)]
        [string[]]$Mailboxes,
        [switch]$WhatIf
    )

    # Store the results for reporting
    $results = @()

    foreach ($mailbox in $Mailboxes) {
        try {
            # Get current mailbox info
            $mbx = Get-Mailbox $mailbox -ErrorAction Stop
            
            $initialState = [PSCustomObject]@{
                Mailbox = $mbx.DisplayName
                Email = $mbx.PrimarySmtpAddress
                ArchiveDatabase = $mbx.ArchiveDatabase
                ArchiveState = $mbx.ArchiveState
                OriginalRetentionPolicy = $mbx.RetentionPolicy
            }

            Write-Host "`nProcessing: $($mbx.DisplayName) - $($mbx.PrimarySmtpAddress)"
            Write-Host "Current Archive Database: $($mbx.ArchiveDatabase)"
            Write-Host "Current Retention Policy: $($mbx.RetentionPolicy)"

            # If there's an archive database, the archive is enabled
            if ($mbx.ArchiveDatabase) {
                if (!$WhatIf) {
                    Write-Host "Disabling archiving policy..."
                    
                    # Remove retention policy (stops new items from being archived)
                    Set-Mailbox $mailbox -RetentionPolicy $null -WarningAction SilentlyContinue
                    
                    # Get updated mailbox info
                    $updatedMbx = Get-Mailbox $mailbox
                    
                    $results += [PSCustomObject]@{
                        Mailbox = $mbx.DisplayName
                        Email = $mbx.PrimarySmtpAddress
                        ArchiveDatabase = $mbx.ArchiveDatabase
                        OriginalRetentionPolicy = $initialState.OriginalRetentionPolicy
                        NewRetentionPolicy = $updatedMbx.RetentionPolicy
                        Status = "Archive Policies Disabled"
                        TimeStamp = (Get-Date)
                        Notes = "Archive mailbox preserved but new archiving stopped"
                    }

                    Write-Host "Successfully disabled archiving policy while preserving archive mailbox"
                } else {
                    Write-Host "WhatIf: Would disable archiving policy for $($mbx.DisplayName) while preserving archive mailbox"
                    $results += [PSCustomObject]@{
                        Mailbox = $mbx.DisplayName
                        Email = $mbx.PrimarySmtpAddress
                        ArchiveDatabase = $mbx.ArchiveDatabase
                        OriginalRetentionPolicy = $initialState.OriginalRetentionPolicy
                        Status = "WhatIf - Would Disable"
                        TimeStamp = (Get-Date)
                        Notes = "WhatIf mode - no changes made"
                    }
                }
            } else {
                Write-Host "Skipping - No archive mailbox found"
                $results += [PSCustomObject]@{
                    Mailbox = $mbx.DisplayName
                    Email = $mbx.PrimarySmtpAddress
                    ArchiveDatabase = "None"
                    OriginalRetentionPolicy = $initialState.OriginalRetentionPolicy
                    Status = "Skipped"
                    TimeStamp = (Get-Date)
                    Notes = "No archive mailbox exists"
                }
            }
        } catch {
            Write-Warning "Error processing $mailbox : $_"
            $results += [PSCustomObject]@{
                Mailbox = $mailbox
                Email = $mailbox
                Status = "Error"
                TimeStamp = (Get-Date)
                Notes = $_.Exception.Message
            }
        }
    }

    # Export results for documentation
    $date = Get-Date -Format "yyyyMMdd-HHmmss"
    $fileName = "ArchiveDisabled_$date.csv"
    $results | Export-CSV $fileName -NoTypeInformation
    
    Write-Host "`nResults exported to: $fileName"
    Write-Host "`nSummary of operations:"
    $results | Format-Table Mailbox, Email, Status, Notes -AutoSize
    
    return $results
}

Disable-MailboxArchiving -Mailboxes "ACampbell@tunngavik.com"