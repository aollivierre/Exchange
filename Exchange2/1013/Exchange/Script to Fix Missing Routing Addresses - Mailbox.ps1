function Fix-MailboxRoutingAddresses {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
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
        $mailboxes | Format-Table DisplayName, PrimarySmtpAddress -AutoSize

        # Confirm before proceeding
        $confirm = Read-Host "`nWill add routing address using domain '$TenantDomain'. Proceed? (Y/N)"
        if ($confirm -ne 'Y') {
            Write-Host "Operation cancelled." -ForegroundColor Yellow
            return
        }

        foreach ($mailbox in $mailboxes) {
            try {
                # Construct new routing address
                $username = ($mailbox.PrimarySmtpAddress -split '@')[0]
                $newAddress = "smtp:$username@$TenantDomain"

                # Backup current addresses
                $currentAddresses = $mailbox.EmailAddresses

                # Add new routing address
                Write-Host "`nProcessing $($mailbox.DisplayName)..." -ForegroundColor Cyan
                Set-Mailbox -Identity $mailbox.Identity -EmailAddresses @{Add=$newAddress}

                Write-Host "Successfully added routing address for $($mailbox.DisplayName)" -ForegroundColor Green
            }
            catch {
                Write-Host "Error processing $($mailbox.DisplayName): $_" -ForegroundColor Red
            }
        }

        Write-Host "`nOperation completed. Please verify the changes." -ForegroundColor Green
    }
    catch {
        Write-Error "Error: $_"
    }
}

# Example usage:
Fix-MailboxRoutingAddresses -TenantDomain "tunngavik.mail.onmicrosoft.com"