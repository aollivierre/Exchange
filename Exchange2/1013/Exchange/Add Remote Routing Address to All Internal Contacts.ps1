function Add-BulkRemoteRoutingAddress {
    [CmdletBinding()]
    param()

    try {
        $ErrorActionPreference = 'Stop'
        $ConfirmPreference = 'None'
        $WarningPreference = 'SilentlyContinue'

        Write-Host "Starting bulk update process..." -ForegroundColor Cyan
        
        # Get all contacts missing routing address
        $contacts = Get-MailContact -ResultSize Unlimited | Where-Object {
            ($_.EmailAddresses | ForEach-Object { $_.ToString() }) -match "tunngavik\.com" -and
            -not (($_.EmailAddresses | ForEach-Object { $_.ToString() }) -match "smtp:.*?@tunngavik\.mail\.onmicrosoft\.com$")
        }

        Write-Host "Found $($contacts.Count) contacts to update" -ForegroundColor Yellow

        foreach ($contact in $contacts) {
            $username = $contact.PrimarySmtpAddress.Split('@')[0]
            $routingAddress = "smtp:$username@tunngavik.mail.onmicrosoft.com"

            $null = Set-MailContact -Identity $contact.Identity `
                                  -EmailAddresses @{add="$routingAddress"} `
                                  -Confirm:$false `
                                  -Force

            Write-Host "Updated $($contact.PrimarySmtpAddress)" -ForegroundColor Green
        }
    }
    catch {
        Write-Error "Error: $_"
    }
}

Add-BulkRemoteRoutingAddress