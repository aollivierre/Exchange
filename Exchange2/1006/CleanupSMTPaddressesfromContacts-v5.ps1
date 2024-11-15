# Timestamp format
$timestampFormat = "yyyy-MM-dd HH:mm:ss"

# Retrieve the list of accepted domains
$acceptedDomains = Get-AcceptedDomain | ForEach-Object { $_.DomainName.ToString().ToLower() }

# Function to update proxy addresses and return result object
function Update-MailContactProxyAddresses {
    param (
        [Parameter(Mandatory = $true)]
        $MailContact
    )

    # Retrieve current proxy addresses
    $proxyAddresses = $MailContact.EmailAddresses | ForEach-Object { $_.ToString().ToLower() }

    # Filter out proxy addresses that match any of the accepted domains
    $filteredProxyAddresses = $proxyAddresses | Where-Object {
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
    Set-MailContact -Identity $MailContact.Identity -EmailAddresses $filteredProxyAddresses -ErrorAction SilentlyContinue

    # Return updated contact info
    return [PSCustomObject]@{
        Name             = $MailContact.Name
        Identity         = $MailContact.Identity
        FilteredAddresses = ($filteredProxyAddresses -join "; ")
    }
}

# Start with Lisa Berting
$lisaBerting = Get-MailContact -Identity "Lisa Berting" -ErrorAction SilentlyContinue
if ($null -ne $lisaBerting) {
    $lisaResult = Update-MailContactProxyAddresses -MailContact $lisaBerting
    Write-Host "$(Get-Date -Format $timestampFormat) - Processed Lisa Berting with updated proxy addresses." -ForegroundColor Cyan
    $results = @($lisaResult)
} else {
    Write-Host "$(Get-Date -Format $timestampFormat) - Lisa Berting not found" -ForegroundColor Red
    $results = @()
}

# Output results to CSV
$results | Export-Csv -Path "C:\code\exchange\mail_contacts_updated.csv" -NoTypeInformation

# Display results in Out-GridView
$results | Out-GridView -Title "Updated Mail Contacts"
