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
            return "SPF record not found using Resolve-DnsName for $domain possibly due to being checked form Internal DNS (run from an external network)"
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

# Loop through each domain and get its SPF record using both methods
foreach ($domain in $domains) {
    $spfRecordResolve = Get-SPFRecord -domain $domain
    $spfRecordAPI = Get-SPFRecordFromAPI -domain $domain

    # Store results in a hashtable
    $result = [pscustomobject]@{
        Domain           = $domain
        SPF_Record_Resolve = $spfRecordResolve
        SPF_Record_API   = $spfRecordAPI
    }

    # Add to results array
    $results += $result
}

# Export results to HTML and display using Out-HTMLView
$results | Out-HTMLView -Title "SPF Record Results"
