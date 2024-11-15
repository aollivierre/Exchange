# Define the list of domains
$domains = @(
    "Iqaluit.tunngavik.com",
    "Rankininlet.tunngavik.com",
    "Cambridgebay.tunngavik.com",
    "Ottawa.tunngavik.com",
    "Tunngavik.com"
)

# Initialize an array to store results
$results = @()

# Function to get the SPF record using Resolve-DnsName
function Get-SPFRecord {
    param (
        [string]$domain
    )

    try {
        # Use Resolve-DnsName to get the TXT records of the domain
        $dnsRecords = Resolve-DnsName -Name $domain -Type TXT -ErrorAction Stop

        # Filter out the SPF record from the TXT records
        $spfRecord = $dnsRecords | Where-Object { $_.Strings -match "v=spf1" }

        if ($spfRecord) {
            return $spfRecord.Strings
        } else {
            return "SPF record not found using Resolve-DnsName for $domain"
        }
    } catch {
        return "Error retrieving SPF record using Resolve-DnsName for $domain $_"
    }
}

# Function to get the SPF record using an external API
function Get-SPFRecordFromAPI {
    param (
        [string]$domain
    )

    try {
        $response = Invoke-RestMethod -Uri "https://dns.google/resolve?name=$domain&type=TXT" -Method Get
        $spfRecord = $response.Answer | Where-Object { $_.data -match "v=spf1" }

        if ($spfRecord) {
            return $spfRecord.data
        } else {
            return "SPF record not found using external API for $domain"
        }
    } catch {
        return "Error retrieving SPF record using external API for $domain $_"
    }
}

# Function to get MX records using Resolve-DnsName
function Get-MXRecord {
    param (
        [string]$domain
    )

    try {
        # Use Resolve-DnsName to get the MX records of the domain
        $mxRecords = Resolve-DnsName -Name $domain -Type MX -ErrorAction Stop

        if ($mxRecords) {
            return $mxRecords | Select-Object -ExpandProperty Exchange
        } else {
            return "MX record not found using Resolve-DnsName for $domain"
        }
    } catch {
        return "Error retrieving MX record using Resolve-DnsName for $domain $_"
    }
}

# Function to get MX record using an external API
function Get-MXRecordFromAPI {
    param (
        [string]$domain
    )

    try {
        $response = Invoke-RestMethod -Uri "https://dns.google/resolve?name=$domain&type=MX" -Method Get
        $mxRecord = $response.Answer | ForEach-Object { $_.exchange }

        if ($mxRecord) {
            return $mxRecord
        } else {
            return "MX record not found using external API for $domain"
        }
    } catch {
        return "Error retrieving MX record using external API for $domain $_"
    }
}

# Loop through each domain and get its SPF and MX records using both methods
foreach ($domain in $domains) {
    $spfRecordResolve = Get-SPFRecord -domain $domain
    $spfRecordAPI = Get-SPFRecordFromAPI -domain $domain
    $mxRecordResolve = Get-MXRecord -domain $domain
    $mxRecordAPI = Get-MXRecordFromAPI -domain $domain

    # Store results in a hashtable
    $result = [pscustomobject]@{
        Domain            = $domain
        SPF_Record_Resolve = $spfRecordResolve
        SPF_Record_API    = $spfRecordAPI
        MX_Record_Resolve = $mxRecordResolve
        MX_Record_API     = $mxRecordAPI
    }

    # Add to results array
    $results += $result
}

# Export results to HTML and display using Out-HTMLView
$results | Out-HTMLView -Title "SPF & MX Record Results"
