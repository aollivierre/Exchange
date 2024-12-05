function Remove-TunngavikProxyAddress {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$EmailAddress
    )

    try {
        $contact = Get-MailContact -Identity $EmailAddress
        
        $newAddresses = $contact.EmailAddresses | Where-Object { 
            $addr = $_.ToString()
            ($addr -notmatch "tunngavik\.com" -and $addr -notmatch "tunngavik\.mail\.onmicrosoft\.com") -or $addr -cmatch "^SMTP:"
        }
        
        Set-MailContact -Identity $contact.Identity -EmailAddresses $newAddresses -EmailAddressPolicyEnabled $false -Confirm:$false
        Write-Host "Successfully cleaned proxy addresses for $EmailAddress" -ForegroundColor Green
        
        $updatedContact = Get-MailContact -Identity $EmailAddress
        Write-Host "Current email addresses:" -ForegroundColor Yellow
        $updatedContact.EmailAddresses | ForEach-Object { Write-Host $_ }
    }
    catch {
        Write-Error "Error cleaning proxy addresses: $_"
    }
}

# Test with pilot contact
Remove-TunngavikProxyAddress -EmailAddress "daisylahure@gmail.com"