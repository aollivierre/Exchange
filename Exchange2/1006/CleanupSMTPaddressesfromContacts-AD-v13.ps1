# Timestamp format
$timestampFormat = "yyyy-MM-dd HH:mm:ss"

# Retrieve the list of accepted domains from Exchange
$acceptedDomains = Get-AcceptedDomain | ForEach-Object { $_.DomainName.ToString().ToLower() }

# Function to update proxy addresses in AD
function Update-ADContactProxyAddresses {
    param (
        [Parameter(Mandatory = $true)]
        $ADObject
    )

    # Retrieve current proxy addresses
    $proxyAddresses = $ADObject.proxyAddresses

    Write-Host "$(Get-Date -Format $timestampFormat) - Analyzing proxy addresses for $($ADObject.Name):" -ForegroundColor Magenta
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

    # Update the AD object with the filtered list of proxyAddresses
    Set-ADObject -Identity $ADObject.DistinguishedName -Replace @{proxyAddresses = $filteredProxyAddresses}
    Write-Host "$(Get-Date -Format $timestampFormat) - Update applied successfully for $($ADObject.Name)." -ForegroundColor Green

    # Display filtered proxy addresses
    Write-Host "$(Get-Date -Format $timestampFormat) - Filtered proxy addresses for $($ADObject.Name):" -ForegroundColor Green
    $filteredProxyAddresses | ForEach-Object { Write-Host "Kept Address: $_" -ForegroundColor Gray }
}

# Fetch the AD object for Lisa Berting
$lisaBertingAD = Get-ADUser -Filter {Name -eq "Lisa Berting"} -Properties proxyAddresses

if ($null -ne $lisaBertingAD) {
    Update-ADContactProxyAddresses -ADObject $lisaBertingAD
} else {
    Write-Host "$(Get-Date -Format $timestampFormat) - Lisa Berting not found in AD" -ForegroundColor Red
}
