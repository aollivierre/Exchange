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
            
            # Add debug information
            Write-Host "Debug Info for $($mbx.DisplayName):"
            Write-Host "Archive Database: $($mbx.ArchiveDatabase)"
            Write-Host "Archive State: $($mbx.ArchiveState)"
            Write-Host "Archive Status: $($mbx.ArchiveStatus)"
            Write-Host "Archive GUID: $($mbx.ArchiveGuid)"
            
            $initialState = [PSCustomObject]@{
                Mailbox = $mbx.DisplayName
                Email = $mbx.PrimarySmtpAddress
                ArchiveDatabase = $mbx.ArchiveDatabase
                ArchiveState = $mbx.ArchiveState
                ArchiveGuid = $mbx.ArchiveGuid
                OriginalRetentionPolicy = $mbx.RetentionPolicy
            }

            # Check if archive exists using ArchiveDatabase property
            if ($mbx.ArchiveDatabase) {
                Write-Host "Processing $($mbx.DisplayName)..."
                
                if (!$WhatIf) {
                    # Remove retention policy (stops new items from being archived)
                    Write-Host "Removing retention policy..."
                    Set-Mailbox $mailbox -RetentionPolicy $null -WarningAction SilentlyContinue
                    
                    # Disable auto-expanding archive if it's enabled
                    Write-Host "Disabling auto-expanding archive..."
                    Set-Mailbox $mailbox -AutoExpandingArchive $false -WarningAction SilentlyContinue
                    
                    # Get updated mailbox info
                    $updatedMbx = Get-Mailbox $mailbox
                    
                    $results += [PSCustomObject]@{
                        Mailbox = $mbx.DisplayName
                        Email = $mbx.PrimarySmtpAddress
                        OriginalArchiveState = $initialState.ArchiveState
                        OriginalRetentionPolicy = $initialState.OriginalRetentionPolicy
                        NewRetentionPolicy = $updatedMbx.RetentionPolicy
                        ArchiveDatabase = $mbx.ArchiveDatabase
                        Status = "Archiving Disabled"
                        TimeStamp = (Get-Date)
                    }
                } else {
                    Write-Host "WhatIf: Would disable archiving for $($mbx.DisplayName)"
                }
            } else {
                Write-Host "Skipping $($mbx.DisplayName) - No archive database found"
                $results += [PSCustomObject]@{
                    Mailbox = $mbx.DisplayName
                    Email = $mbx.PrimarySmtpAddress
                    OriginalArchiveState = $initialState.ArchiveState
                    OriginalRetentionPolicy = $initialState.OriginalRetentionPolicy
                    ArchiveDatabase = "None"
                    Status = "Skipped - No Archive Database"
                    TimeStamp = (Get-Date)
                }
            }
        } catch {
            Write-Warning "Error processing $mailbox : $_"
            $results += [PSCustomObject]@{
                Mailbox = $mailbox
                Status = "Error: $_"
                TimeStamp = (Get-Date)
            }
        }
    }

    # Export results for documentation
    $date = Get-Date -Format "yyyyMMdd-HHmmss"
    $results | Export-CSV "ArchiveDisabled_$date.csv" -NoTypeInformation
    
    # Display results on screen
    $results | Format-Table -AutoSize
    
    return $results
}

# Example usage:
<#
# For a single mailbox:
Disable-MailboxArchiving -Mailboxes "user@domain.com" -WhatIf

# For multiple mailboxes:
$mailboxesToProcess = @(
    "user1@domain.com",
    "user2@domain.com"
)
Disable-MailboxArchiving -Mailboxes $mailboxesToProcess

# For mailboxes from a CSV:
$mailboxesToProcess = Import-CSV "mailboxes.csv" | Select-Object -ExpandProperty EmailAddress
Disable-MailboxArchiving -Mailboxes $mailboxesToProcess
#>

# For a single mailbox:
Disable-MailboxArchiving -Mailboxes "ACampbell@tunngavik.com" -WhatIf