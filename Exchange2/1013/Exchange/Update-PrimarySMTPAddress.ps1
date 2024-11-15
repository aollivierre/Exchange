# Function to update primary SMTP address
function Update-PrimarySmtpAddress {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$Identity,
        
        [Parameter(Mandatory = $true)]
        [string]$NewPrimarySmtpAddress,

        [Parameter(Mandatory = $false)]
        [string]$TenantName = "tunngavik"  # Default tenant name
    )
    
    try {
        # First get the mailbox to check current configuration
        $mailbox = Get-RemoteMailbox -Identity $Identity
        if (-not $mailbox) {
            throw "Mailbox not found: $Identity"
        }

        Write-Host "Current configuration for $Identity" -ForegroundColor Cyan
        Write-Host "Primary SMTP: $($mailbox.PrimarySmtpAddress)" -ForegroundColor Gray
        Write-Host "Remote Routing: $($mailbox.RemoteRoutingAddress)" -ForegroundColor Gray
        
        # Disable email address policy if enabled
        if ($mailbox.EmailAddressPolicyEnabled) {
            Write-Host "Disabling email address policy for $Identity..." -ForegroundColor Yellow
            Set-RemoteMailbox -Identity $Identity -EmailAddressPolicyEnabled $false -Confirm:$false
        }
        
        # Prepare the correct routing address
        $name = $Identity.ToLower()
        $remoteRoutingAddress = "$name@$TenantName.mail.onmicrosoft.com"
        
        # Create new address collection
        [System.Collections.Generic.List[string]]$newAddresses = @()
        $newAddresses.Add("SMTP:$NewPrimarySmtpAddress")  # Primary SMTP
        $newAddresses.Add("smtp:$remoteRoutingAddress")   # Hybrid routing address
        
        Write-Host "Setting new email addresses for $Identity..." -ForegroundColor Yellow
        Set-RemoteMailbox -Identity $Identity -EmailAddresses $newAddresses -Confirm:$false
        
        Write-Host "Setting remote routing address..." -ForegroundColor Yellow
        Set-RemoteMailbox -Identity $Identity -RemoteRoutingAddress $remoteRoutingAddress -Confirm:$false
        
        # Verify changes
        $updatedMailbox = Get-RemoteMailbox -Identity $Identity
        
        Write-Host "`nUpdated configuration for $Identity" -ForegroundColor Green
        Write-Host "Email Address Policy: $($updatedMailbox.EmailAddressPolicyEnabled)" -ForegroundColor Yellow
        Write-Host "Primary SMTP: $($updatedMailbox.PrimarySmtpAddress)" -ForegroundColor Yellow
        Write-Host "Remote Routing: $($updatedMailbox.RemoteRoutingAddress)" -ForegroundColor Yellow
        Write-Host "All Email Addresses:" -ForegroundColor Yellow
        $updatedMailbox.EmailAddresses | ForEach-Object {
            Write-Host "  $_" -ForegroundColor Gray
        }
    }
    catch {
        Write-Error "Error updating $Identity : $_"
        Write-Host "Full Error Details:" -ForegroundColor Red
        $_.Exception | Format-List -Force
    }
}

# Main execution
try {
    # Update SMTP mailbox
    Write-Host "`n=== Updating SMTP mailbox ===" -ForegroundColor Cyan
    Update-PrimarySmtpAddress -Identity "SMTP" -NewPrimarySmtpAddress "SMTP@tunngavik.com"

    # Update DMARC mailbox
    Write-Host "`n=== Updating DMARC mailbox ===" -ForegroundColor Cyan
    Update-PrimarySmtpAddress -Identity "DMARC" -NewPrimarySmtpAddress "DMARC@tunngavik.com"

    # Final verification
    Write-Host "`n=== Final Configuration Verification ===" -ForegroundColor Cyan
    @("SMTP", "DMARC") | ForEach-Object {
        $mailbox = Get-RemoteMailbox -Identity $_
        Write-Host "`nMailbox: $_" -ForegroundColor Green
        Write-Host "Status Summary:" -ForegroundColor Yellow
        Write-Host "- Email Address Policy Enabled: $($mailbox.EmailAddressPolicyEnabled)"
        Write-Host "- Primary SMTP Address: $($mailbox.PrimarySmtpAddress)"
        Write-Host "- Remote Routing Address: $($mailbox.RemoteRoutingAddress)"
        Write-Host "- Email Addresses:"
        $mailbox.EmailAddresses | ForEach-Object {
            Write-Host "  $_"
        }
    }
}
catch {
    Write-Error "Script execution failed: $_"
    Write-Host "Full Error Details:" -ForegroundColor Red
    $_.Exception | Format-List -Force
}