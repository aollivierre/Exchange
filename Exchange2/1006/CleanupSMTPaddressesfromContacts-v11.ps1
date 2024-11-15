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
    $proxyAddresses = $MailContact.EmailAddresses | ForEach-Object { $_.ToString() }

    # Display current proxy addresses
    Write-Host "$(Get-Date -Format $timestampFormat) - Analyzing proxy addresses for $($MailContact.Name):" -ForegroundColor Magenta
    $proxyAddresses | ForEach-Object { Write-Host "Original Address: $_" -ForegroundColor Gray }

    # Initialize a list to hold the filtered proxy addresses
    $filteredProxyAddresses = @()

    foreach ($address in $proxyAddresses) {
        Write-Host "Processing Address: $address" -ForegroundColor Yellow
        if ($address -cmatch '^SMTP:') {
            Write-Host "Keeping primary SMTP address: $address" -ForegroundColor Green
            $filteredProxyAddresses += $address
        } elseif ($address -cmatch '^smtp:') {
            # Extract the domain part of the SMTP address
            $domain = $address -replace '^smtp:[^@]+@', ''
            Write-Host "Extracted domain from SMTP: $domain" -ForegroundColor Yellow
            # Check if the domain is in the accepted domains list
            if (-not ($acceptedDomains -contains $domain.ToLower())) {
                Write-Host "Keeping address as its domain is not an accepted domain: $address" -ForegroundColor Green
                $filteredProxyAddresses += $address
            } else {
                Write-Host "Removing address as its domain is an accepted domain: $address" -ForegroundColor Red
            }
        } else {
            Write-Host "Keeping non-SMTP address: $address" -ForegroundColor Green
            $filteredProxyAddresses += $address
        }
    }

    # Update the mail contact with the filtered list of proxyAddresses
  
    try {
        # Update the mail contact with the filtered list of proxyAddresses
        Set-MailContact -Identity $MailContact.Identity -EmailAddresses $filteredProxyAddresses -ErrorAction Stop -Verbose
        Write-Host "Update applied successfully for $($MailContact.Name)." -ForegroundColor Green
    } catch {
        Write-Host "Failed to update $($MailContact.Name): $_" -ForegroundColor Red
    }

    # Display filtered proxy addresses
    Write-Host "$(Get-Date -Format $timestampFormat) - Filtered proxy addresses for $($MailContact.Name):" -ForegroundColor Green
    $filteredProxyAddresses | ForEach-Object { Write-Host "Kept Address: $_" -ForegroundColor Gray }

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
    Write-Host "$(Get-Date -Format $timestampFormat) - Finished processing Lisa Berting." -ForegroundColor Cyan
    $results = @($lisaResult)
} else {
    Write-Host "$(Get-Date -Format $timestampFormat) - Lisa Berting not found" -ForegroundColor Red
    $results = @()
}

# Output results to CSV
$results | Export-Csv -Path "C:\code\exchange\mail_contacts_updated.csv" -NoTypeInformation

# Display results in Out-GridView
$results | Out-GridView -Title "Updated Mail Contacts"
