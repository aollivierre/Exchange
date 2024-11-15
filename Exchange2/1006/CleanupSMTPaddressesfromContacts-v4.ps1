# Retrieve the list of accepted domains
$acceptedDomains = Get-AcceptedDomain | ForEach-Object { $_.DomainName.ToString().ToLower() }

# Get all mail contacts
$mailContacts = Get-MailContact -ResultSize Unlimited

# Iterate over each mail contact
foreach ($contact in $mailContacts) {
    # Retrieve current proxy addresses and convert them to lowercase for comparison
    $currentProxyAddresses = $contact.EmailAddresses | ForEach-Object { $_.ToString().ToLower() }

    # Filter out proxy addresses that match any of the accepted domains
    $filteredProxyAddresses = $currentProxyAddresses | Where-Object {
        $address = $_
        $remove = $false
        foreach ($domain in $acceptedDomains) {
            if ($address -like "*@$domain") {
                $remove = $true
                break
            }
        }
        -not $remove
    }

    # Update the mail contact with the filtered list of proxyAddresses
    # Convert back to original format (e.g., "smtp:" to "SMTP:") if needed
    $filteredProxyAddressesForUpdate = $filteredProxyAddresses | ForEach-Object {
        if ($_ -cmatch '^smtp:') {
            $_ -replace '^smtp:', 'smtp:'
        } else {
            $_
        }
    }
    Set-MailContact -Identity $contact.Identity -EmailAddresses $filteredProxyAddressesForUpdate
}

Write-Host "Completed removing proxy addresses with accepted domains for all mail contacts."
