# Timestamp format
$timestampFormat = "yyyy-MM-dd HH:mm:ss"

# Retrieve the list of accepted domains from Exchange
$acceptedDomains = Get-AcceptedDomain | ForEach-Object { $_.DomainName.ToString().ToLower() }

# Function to update proxy addresses in AD for a contact
function Update-ADContactProxyAddresses {
    param (
        [Parameter(Mandatory = $true)]
        $ADContact
    )

    # Retrieve current proxy addresses
    $proxyAddresses = $ADContact.proxyAddresses

    Write-Host "$(Get-Date -Format $timestampFormat) - Analyzing proxy addresses for $($ADContact.Name):" -ForegroundColor Magenta
    $proxyAddresses | ForEach-Object { Write-Host "Original Address: $_" -ForegroundColor Gray }

    # Initialize a list to hold the filtered proxy addresses
    $filteredProxyAddresses = @()

    foreach ($address in $proxyAddresses) {
        Write-Host "Processing Address: $address" -ForegroundColor Yellow
        $domain = $address -replace '^[smtp|SMTP]+:[^@]+@', ''
        if ($address -cmatch '^SMTP:') {
            Write-Host "Keeping primary SMTP address: $address" -ForegroundColor Green
            $filteredProxyAddresses += $address
        } elseif ($address -cmatch '^smtp:' -and -not ($acceptedDomains -contains $domain.ToLower())) {
            Write-Host "Keeping address as its domain is not an accepted domain: $address" -ForegroundColor Green
            $filteredProxyAddresses += $address
        } elseif ($address -cnotmatch '^smtp:') {
            Write-Host "Keeping non-SMTP address: $address" -ForegroundColor Green
            $filteredProxyAddresses += $address
        } else {
            Write-Host "Removing address as its domain is an accepted domain: $address" -ForegroundColor Red
        }
    }

    # Update the AD contact with the filtered list of proxyAddresses
    Set-ADObject -Identity $ADContact.DistinguishedName -Replace @{proxyAddresses = $filteredProxyAddresses}
    Write-Host "$(Get-Date -Format $timestampFormat) - Update applied successfully for $($ADContact.Name)." -ForegroundColor Green

    # Display filtered proxy addresses
    Write-Host "$(Get-Date -Format $timestampFormat) - Filtered proxy addresses for $($ADContact.Name):" -ForegroundColor Green
    $filteredProxyAddresses | ForEach-Object { Write-Host "Kept Address: $_" -ForegroundColor Gray }
}

# Fetch the AD contact for Lisa Berting using an LDAP filter to target only contact objects
$lisaBertingContact = Get-ADObject -LDAPFilter "(objectClass=contact)" -Properties proxyAddresses, Name | Where-Object { $_.Name -eq "Lisa Berting" }

if ($null -ne $lisaBertingContact) {
    Update-ADContactProxyAddresses -ADContact $lisaBertingContact
} else {
    Write-Host "$(Get-Date -Format $timestampFormat) - Lisa Berting contact not found in AD" -ForegroundColor Red
}
