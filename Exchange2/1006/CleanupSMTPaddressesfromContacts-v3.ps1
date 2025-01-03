# Timestamp format
$timestampFormat = "yyyy-MM-dd HH:mm:ss"

# Function to update proxy addresses and return result object
function Update-MailContactProxyAddresses {
    param (
        [Parameter(Mandatory = $true)]
        $MailContact
    )

    # Retrieve current proxy addresses
    $proxyAddresses = $MailContact.EmailAddresses | ForEach-Object { $_.ToString() }

    # Filter the proxyAddresses to keep only X400, X500, and uppercase SMTP addresses
    $filteredProxyAddresses = $proxyAddresses | Where-Object {
        ($_ -like "X400:*") -or
        ($_ -like "X500:*") -or
        ($_ -cmatch "^SMTP:")
    }

    # Update the mail contact with the filtered list of proxyAddresses
    Set-MailContact -Identity $MailContact.Identity -EmailAddresses $filteredProxyAddresses -ErrorAction SilentlyContinue

    # Return updated contact info
    return [PSCustomObject]@{
        Name             = $MailContact.Name
        Identity         = $MailContact.Identity
        FilteredAddresses = $filteredProxyAddresses -join "; "
    }
}

# Start with Lisa Berting
$lisaBerting = Get-MailContact -Identity "Lisa Berting" -ErrorAction SilentlyContinue
if ($null -ne $lisaBerting) {
    $lisaResult = Update-MailContactProxyAddresses -MailContact $lisaBerting
    Write-Host "$(Get-Date -Format $timestampFormat) - Processed Lisa Berting" -ForegroundColor Cyan
    $results = @($lisaResult)
} else {
    Write-Host "$(Get-Date -Format $timestampFormat) - Lisa Berting not found" -ForegroundColor Red
    $results = @()
}

# Get all mail contacts excluding Lisa Berting
# $mailContacts = Get-MailContact -ResultSize Unlimited | Where-Object { $_.Identity -ne $lisaBerting.Identity }

# Iterate over each remaining mail contact to update proxy addresses
# foreach ($contact in $mailContacts) {
#     $result = Update-MailContactProxyAddresses -MailContact $contact
#     $results += $result
# }

# Output results to CSV
$results | Export-Csv -Path "C:\code\exchange\mail_contacts_updated.csv" -NoTypeInformation

# Display results in Out-GridView
$results | Out-GridView -Title "Updated Mail Contacts"
