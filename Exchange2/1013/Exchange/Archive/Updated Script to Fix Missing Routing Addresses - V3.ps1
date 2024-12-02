function Fix-MailboxRoutingAddresses {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$TenantDomain
    )

    try {
        # Validate tenant domain
        if (-not $TenantDomain.EndsWith('.mail.onmicrosoft.com')) {
            $TenantDomain = $TenantDomain + '.mail.onmicrosoft.com'
        }

        # Get mailboxes missing routing address
        $mailboxes = Get-Mailbox -ResultSize Unlimited | Where-Object {
            -not ($_.EmailAddresses | Where-Object { $_ -match "smtp:.+\.mail\.onmicrosoft\.com$" })
        }

        if ($mailboxes.Count -eq 0) {
            Write-Host "No mailboxes found missing routing addresses." -ForegroundColor Green
            return
        }

        # Display affected mailboxes
        Write-Host "`nMailboxes missing routing address:" -ForegroundColor Yellow
        $mailboxes | Format-Table DisplayName, PrimarySmtpAddress, Alias -AutoSize

        # Confirm before proceeding
        $confirm = Read-Host "`nWill add routing address using domain '$TenantDomain'. Proceed? (Y/N)"
        if ($confirm -ne 'Y') {
            Write-Host "Operation cancelled." -ForegroundColor Yellow
            return
        }

        foreach ($mailbox in $mailboxes) {
            try {
                Write-Host "`nProcessing $($mailbox.DisplayName)..." -ForegroundColor Cyan
                
                # Construct new routing address
                $username = ($mailbox.PrimarySmtpAddress -split '@')[0]
                $newAddress = "smtp:$username@$TenantDomain"
                
                # Create the email addresses string
                $emailAddressesString = $mailbox.EmailAddresses | ForEach-Object { $_.ToString() }
                $emailAddressesString += $newAddress
                
                # Create the update command
                # $updateCommand = "Set-Mailbox -Identity '$($mailbox.Alias)' -EmailAddresses '$($emailAddressesString -join ","')'"
                
                # Execute the command
                Write-Host "Executing update..." -ForegroundColor Gray
                Invoke-Expression $updateCommand
                
                Write-Host "Successfully added routing address for $($mailbox.DisplayName)" -ForegroundColor Green
                Write-Host "Added: $newAddress" -ForegroundColor Gray
            }
            catch {
                Write-Host "Error processing $($mailbox.DisplayName): $_" -ForegroundColor Red
                Write-Host "Failed Command: $updateCommand" -ForegroundColor Red
                Write-Host "Full Error:" -ForegroundColor Red
                $_.Exception | Format-List -Force
            }
        }

        Write-Host "`nOperation completed. Running verification..." -ForegroundColor Yellow
        
        # Verify changes
        $remainingIssues = Get-Mailbox -ResultSize Unlimited | Where-Object {
            -not ($_.EmailAddresses | Where-Object { $_ -match "smtp:.+\.mail\.onmicrosoft\.com$" })
        }

        if ($remainingIssues.Count -eq 0) {
            Write-Host "All mailboxes now have routing addresses!" -ForegroundColor Green
        }
        else {
            Write-Host "Some mailboxes still missing routing addresses:" -ForegroundColor Red
            $remainingIssues | Format-Table DisplayName, PrimarySmtpAddress, Alias -AutoSize
        }
    }
    catch {
        Write-Error "Error: $_"
        Write-Host "Full Error Details:" -ForegroundColor Red
        $_.Exception | Format-List -Force
    }
}

# Example usage:
Fix-MailboxRoutingAddresses -TenantDomain "tunngavik.mail.onmicrosoft.com"