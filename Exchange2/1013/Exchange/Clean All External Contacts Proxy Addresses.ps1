function Remove-AllExternalTunngavikAddresses {
    [CmdletBinding()]
    param()

    try {
        $ErrorActionPreference = 'Stop'
        $ConfirmPreference = 'None'
        $WarningPreference = 'SilentlyContinue'

        $contacts = Get-MailContact -ResultSize Unlimited | Where-Object {
            -not ($_.PrimarySmtpAddress -match "tunngavik\.com")
        }

        Write-Host "Found $($contacts.Count) external contacts to clean" -ForegroundColor Yellow

        foreach ($contact in $contacts) {
            $newAddresses = $contact.EmailAddresses | Where-Object { 
                $addr = $_.ToString()
                ($addr -notmatch "tunngavik\.com" -and $addr -notmatch "tunngavik\.mail\.onmicrosoft\.com") -or $addr -cmatch "^SMTP:"
            }
            
            $params = @{
                Identity = $contact.Identity
                EmailAddresses = $newAddresses
                EmailAddressPolicyEnabled = $false
                Confirm = $false
                Force = $true
            }
            
            $null = Set-MailContact @params

            Write-Host "Cleaned proxy addresses for $($contact.PrimarySmtpAddress)" -ForegroundColor Green
        }
    }
    catch {
        Write-Error "Error in bulk cleanup: $_"
    }
}

Remove-AllExternalTunngavikAddresses