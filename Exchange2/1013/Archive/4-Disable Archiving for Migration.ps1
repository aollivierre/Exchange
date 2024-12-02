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
                OriginalArchiveState = $mbx.ArchiveStatus
                OriginalRetentionPolicy = $mbx.RetentionPolicy
            }

            # If archiving is enabled, disable it
            if ($mbx.ArchiveStatus -eq "Active") {
                Write-Host "Processing $($mbx.DisplayName)..."
                
                if (!$WhatIf) {
                    # Remove retention policy (stops new items from being archived)
                    Set-Mailbox $mailbox -RetentionPolicy $null -WarningAction SilentlyContinue
                    
                    # Disable auto-expanding archive if it's enabled
                    Set-Mailbox $mailbox -AutoExpandingArchive $false -WarningAction SilentlyContinue
                    
                    # Get updated mailbox info
                    $updatedMbx = Get-Mailbox $mailbox
                    
                    $results += [PSCustomObject]@{
                        Mailbox = $mbx.DisplayName
                        Email = $mbx.PrimarySmtpAddress
                        OriginalArchiveState = $initialState.OriginalArchiveState
                        OriginalRetentionPolicy = $initialState.OriginalRetentionPolicy
                        NewRetentionPolicy = $updatedMbx.RetentionPolicy
                        Status = "Archiving Disabled"
                        TimeStamp = (Get-Date)
                    }
                } else {
                    Write-Host "WhatIf: Would disable archiving for $($mbx.DisplayName)"
                }
            } else {
                Write-Host "Skipping $($mbx.DisplayName) - Archiving not active"
                $results += [PSCustomObject]@{
                    Mailbox = $mbx.DisplayName
                    Email = $mbx.PrimarySmtpAddress
                    OriginalArchiveState = $initialState.OriginalArchiveState
                    OriginalRetentionPolicy = $initialState.OriginalRetentionPolicy
                    Status = "Skipped - Not Active"
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