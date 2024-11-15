# Timestamp format
$timestampFormat = "yyyy-MM-dd HH:mm:ss"

# Get all mail contacts
$mailContacts = Get-MailContact -ResultSize Unlimited

# Count before changes
$countBefore = $mailContacts.Count
Write-Host "$(Get-Date -Format $timestampFormat) - Total mail contacts before changes: $countBefore" -ForegroundColor Cyan

# Iterate over each mail contact to update proxy addresses
$results = foreach ($contact in $mailContacts) {
    # Retrieve current proxy addresses
    $proxyAddresses = $contact.EmailAddresses | ForEach-Object { $_.ToString() }

    # Filter the proxyAddresses to keep only X400, X500, and uppercase SMTP addresses
    $filteredProxyAddresses = $proxyAddresses | Where-Object {
        ($_ -like "X400:*") -or
        ($_ -like "X500:*") -or
        ($_ -cmatch "^SMTP:")
    }

    # Update the mail contact with the filtered list of proxyAddresses
    Set-MailContact -Identity $contact.Identity -EmailAddresses $filteredProxyAddresses -ErrorAction SilentlyContinue

    # Output for CSV and grid view
    [PSCustomObject]@{
        Name             = $contact.Name
        Identity         = $contact.Identity
        FilteredAddresses = $filteredProxyAddresses -join "; "
    }
}

# Count after changes
$mailContactsAfter = Get-MailContact -ResultSize Unlimited
$countAfter = $mailContactsAfter.Count
Write-Host "$(Get-Date -Format $timestampFormat) - Total mail contacts after changes: $countAfter" -ForegroundColor Green

# Output results to CSV
$results | Export-Csv -Path "C:\code\exchange\file\mail_contacts_updated.csv" -NoTypeInformation

# Display results in Out-GridView
$results | Out-GridView -Title "Updated Mail Contacts"

