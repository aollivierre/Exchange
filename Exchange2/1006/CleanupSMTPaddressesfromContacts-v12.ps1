# Timestamp format
$timestampFormat = "yyyy-MM-dd HH:mm:ss"

# Retrieve the list of accepted domains
$acceptedDomains = Get-AcceptedDomain | ForEach-Object { $_.DomainName.ToString().ToLower() }

# Fetch the mail contact details for Lisa Berting
$lisaBerting = Get-MailContact -Identity "Lisa Berting" -ErrorAction SilentlyContinue

if ($null -ne $lisaBerting) {
    Write-Host "$(Get-Date -Format $timestampFormat) - Analyzing proxy addresses for $($lisaBerting.Name):" -ForegroundColor Magenta
    $lisaBerting.EmailAddresses | ForEach-Object { Write-Host "Original Address: $_" -ForegroundColor Gray }

    # Filter out SMTP proxy addresses that match any of the accepted domains
    $filteredProxyAddresses = $lisaBerting.EmailAddresses | Where-Object {
        $address = $_.ToString()
        if ($address -cmatch '^SMTP:') {
            # Always keep the primary SMTP address
            $true
        } elseif ($address -cmatch '^smtp:') {
            # Extract the domain part of the SMTP address
            $domain = $address -replace '^smtp:[^@]+@', ''
            Write-Host "Extracted domain from SMTP: $domain" -ForegroundColor Yellow
            # Keep if the domain is not in the accepted domains list
            -not ($acceptedDomains -contains $domain.ToLower())
        } else {
            # Keep non-SMTP addresses
            $true
        }
    }

    # Update the mail contact with the filtered list of proxyAddresses
    Set-MailContact -Identity $lisaBerting.Identity -EmailAddresses $filteredProxyAddresses -ErrorAction SilentlyContinue

    Write-Host "$(Get-Date -Format $timestampFormat) - Update applied successfully for $($lisaBerting.Name)." -ForegroundColor Green

    # Display filtered proxy addresses
    Write-Host "$(Get-Date -Format $timestampFormat) - Filtered proxy addresses for $($lisaBerting.Name):" -ForegroundColor Green
    $filteredProxyAddresses | ForEach-Object { Write-Host "Kept Address: $_" -ForegroundColor Gray }
} else {
    Write-Host "$(Get-Date -Format $timestampFormat) - Lisa Berting not found" -ForegroundColor Red
}
