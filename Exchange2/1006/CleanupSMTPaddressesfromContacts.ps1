# Define the name of the AD contact
$contactName = "Lisa Berting"

# Get the AD contact object by its name
$contact = Get-ADObject -Filter { Name -eq $contactName } -Properties proxyAddresses

# Check if the contact exists
if ($contact -ne $null) {
    # Filter the proxyAddresses to keep only X400, X500, and uppercase SMTP addresses
    $filteredProxyAddresses = $contact.proxyAddresses | Where-Object {
        ($_ -like "X400:*") -or
        ($_ -like "X500:*") -or
        ($_ -like "SMTP:*")
    }

    # Update the AD contact with the filtered list of proxyAddresses
    Set-ADObject -Identity $contact.DistinguishedName -Replace @{proxyAddresses = $filteredProxyAddresses}

    Write-Host "Proxy addresses for '$contactName' have been updated."
} else {
    Write-Host "Contact '$contactName' not found."
}
