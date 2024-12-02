function Add-RemoteRoutingAddress {
    [CmdletBinding(SupportsShouldProcess)]
    param (
        [Parameter(Mandatory)]
        [string]$EmailAddress
    )

    try {
        $contact = Get-MailContact -Identity $EmailAddress -ErrorAction Stop
        
        # Create the remote routing address
        $username = $contact.PrimarySmtpAddress.Split('@')[0]
        $routingAddress = "smtp:$username@tunngavik.mail.onmicrosoft.com"
        
        # Check if address already exists
        $hasRoutingAddress = $contact.EmailAddresses | Where-Object { $_.ToString() -eq $routingAddress }
        
        if ($hasRoutingAddress) {
            Write-Host "Remote routing address already exists for $EmailAddress" -ForegroundColor Yellow
            return
        }

        if ($PSCmdlet.ShouldProcess($EmailAddress, "Add remote routing address")) {
            $newAddresses = $contact.EmailAddresses + $routingAddress
            Set-MailContact -Identity $EmailAddress -EmailAddresses $newAddresses
            Write-Host "Successfully added remote routing address to $EmailAddress" -ForegroundColor Green
        }
    }
    catch {
        Write-Error "Error adding remote routing address: $_"
    }
}

# Test with pilot user
# Add-RemoteRoutingAddress -EmailAddress "A0TestCB@cambridgebay.tunngavik.com" -WhatIf

Add-RemoteRoutingAddress -EmailAddress "A0TestCB@cambridgebay.tunngavik.com"