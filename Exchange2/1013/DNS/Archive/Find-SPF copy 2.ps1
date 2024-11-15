# Function to recursively retrieve and expand SPF records
Function Get-SPFRecordRecursive {
    param (
        [string]$domain,
        [System.Collections.Generic.HashSet[string]]$ProcessedDomains = $(New-Object System.Collections.Generic.HashSet[string])
    )

    # Avoid processing the same domain multiple times to prevent infinite loops
    if ($ProcessedDomains.Contains($domain)) {
        return @()
    }

    $ProcessedDomains.Add($domain)

    try {
        $txtRecords = Resolve-DnsName -Name $domain -Type TXT -ErrorAction Stop
    }
    catch {
        Write-Host "Failed to resolve TXT records for $domain"
        return @()
    }

    $spfRecord = $txtRecords | Where-Object { $_.Strings -match '^v=spf1' } | Select-Object -First 1

    if ($null -eq $spfRecord) {
        Write-Host "No SPF record found for $domain"
        return @()
    }

    # Combine multi-string SPF records into one
    $spfString = $spfRecord.Strings -join ' '

    # Remove 'v=spf1' and split into individual mechanisms
    $mechanisms = $spfString -replace '^v=spf1\s*', '' -split '\s+'

    $expandedMechanisms = @()

    foreach ($mechanism in $mechanisms) {
        if ($mechanism -like 'include:*') {
            $includedDomain = $mechanism -replace 'include:', ''
            Write-Host "Processing include: $includedDomain"
            $includedMechanisms = Get-SPFRecordRecursive -domain $includedDomain -ProcessedDomains $ProcessedDomains
            $expandedMechanisms += $includedMechanisms
        }
        else {
            $expandedMechanisms += $mechanism
        }
    }

    return $expandedMechanisms
}

# Define the list of domains
$domains = @(
    "Iqaluit.tunngavik.com",
    "Rankininlet.tunngavik.com",
    "cambridgebay.tunngavik.com",
    "ottawa.tunngavik.com",
    "tunngavik.com"
)

# Initialize an array to hold the SPF records
$spfRecords = @()

foreach ($domain in $domains) {
    Write-Host "Processing domain: $domain"

    $processedDomains = New-Object System.Collections.Generic.HashSet[string]
    $expandedMechanisms = Get-SPFRecordRecursive -domain $domain -ProcessedDomains $processedDomains

    if ($expandedMechanisms.Count -gt 0) {
        $spfRecords += [PSCustomObject]@{
            Domain    = $domain
            SPFRecord = 'v=spf1 ' + ($expandedMechanisms -join ' ')
        }
    }
    else {
        $spfRecords += [PSCustomObject]@{
            Domain    = $domain
            SPFRecord = "No SPF record found"
        }
    }
}

# # Export the results to an HTML file
# $spfRecords |
#     ConvertTo-Html -Property Domain, SPFRecord -Title "SPF Records" |
#     Out-File -FilePath "SPFRecords.html"

# # Open the HTML file in the default web browser
# Start-Process -FilePath "SPFRecords.html"



# Export the results using Out-HTMLView
$spfRecords | Out-HTMLView -Title "SPF Records" -FilePath "SPFRecords.html"