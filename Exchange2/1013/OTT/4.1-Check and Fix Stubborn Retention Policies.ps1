# First let's identify all mailboxes that still have the "6M Archive" policy
$stubbornMailboxes = Get-Mailbox -ResultSize Unlimited | Where-Object {$_.RetentionPolicy -eq "6M Archive"}

Write-Host "`nMailboxes still having '6M Archive' policy:" -ForegroundColor Yellow
$stubbornMailboxes | Format-Table DisplayName, PrimarySmtpAddress, RetentionPolicy, ArchiveDatabase -AutoSize

# Function to forcefully remove retention policy
function Remove-StubornRetentionPolicy {
    param(
        [Parameter(Mandatory=$true)]
        [string[]]$Mailboxes,
        [switch]$WhatIf
    )

    $results = [System.Collections.Generic.List[PSObject]]::new()

    foreach ($mailbox in $Mailboxes) {
        try {
            $mbx = Get-Mailbox $mailbox -ErrorAction Stop
            
            Write-Host "`nProcessing: $($mbx.DisplayName)" -ForegroundColor Cyan
            Write-Host "Current Retention Policy: $($mbx.RetentionPolicy)"
            
            # Try to force remove the retention policy
            if (!$WhatIf) {
                Write-Host "Attempting to force remove retention policy..." -ForegroundColor Yellow
                
                # First attempt - standard removal
                Set-Mailbox $mbx.Identity -RetentionPolicy $null -Force -ErrorAction SilentlyContinue
                
                # Verify the change
                $updatedMbx = Get-Mailbox $mbx.Identity
                
                $result = [PSCustomObject]@{
                    Mailbox = $mbx.DisplayName
                    Email = $mbx.PrimarySmtpAddress
                    OriginalRetentionPolicy = $mbx.RetentionPolicy
                    NewRetentionPolicy = $updatedMbx.RetentionPolicy
                    Status = if ($updatedMbx.RetentionPolicy -eq $null) { "Success" } else { "Failed" }
                    Notes = "Attempted force removal of retention policy"
                }
                
                Write-Host "New Retention Policy: $($updatedMbx.RetentionPolicy)"
            } else {
                Write-Host "WhatIf: Would attempt to force remove retention policy" -ForegroundColor Cyan
                $result = [PSCustomObject]@{
                    Mailbox = $mbx.DisplayName
                    Email = $mbx.PrimarySmtpAddress
                    OriginalRetentionPolicy = $mbx.RetentionPolicy
                    Status = "WhatIf"
                    Notes = "WhatIf mode - no changes made"
                }
            }
            
            $results.Add($result)
        }
        catch {
            Write-Warning "Error processing $mailbox : $_"
            $results.Add([PSCustomObject]@{
                Mailbox = $mailbox
                Status = "Error"
                Notes = $_.Exception.Message
            })
        }
    }

    # Display results
    Write-Host "`nOperation Results:" -ForegroundColor Cyan
    $results | Format-Table -AutoSize

    return $results
}

# Example usage:
if ($stubbornMailboxes) {
    Write-Host "`nWould you like to attempt to force remove the retention policy from these mailboxes? (Yes/No)"
    $confirmation = Read-Host
    
    if ($confirmation -match '^[Yy]') {
        # First try with WhatIf
        Write-Host "`nTesting with WhatIf first..." -ForegroundColor Cyan
        Remove-StubornRetentionPolicy -Mailboxes $stubbornMailboxes.PrimarySmtpAddress -WhatIf
        
        Write-Host "`nWould you like to proceed with the actual changes? (Yes/No)"
        $confirmation = Read-Host
        
        if ($confirmation -match '^[Yy]') {
            Remove-StubornRetentionPolicy -Mailboxes $stubbornMailboxes.PrimarySmtpAddress
        }
    }
}
else {
    Write-Host "No mailboxes found with '6M Archive' policy." -ForegroundColor Green
}