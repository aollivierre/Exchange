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

# Function to get the Autodiscover record using Resolve-DnsName
function Get-AutodiscoverRecord {
    param (
        [string]$domain
    )

    try {
        # Use Resolve-DnsName to get the CNAME or A records for Autodiscover
        $dnsRecords = Resolve-DnsName -Name "autodiscover.$domain" -ErrorAction Stop

        if ($dnsRecords) {
            # Join multiple records into a single string if necessary
            return ($dnsRecords | Select-Object -ExpandProperty NameHost) -join ", "
        } else {
            return "Autodiscover record not found using Resolve-DnsName for $domain"
        }
    } catch {
        return "Error retrieving Autodiscover record using Resolve-DnsName for $domain $($_)"
    }
}

# Function to get the Autodiscover record using an external API
function Get-AutodiscoverRecordFromAPI {
    param (
        [string]$domain
    )

    try {
        $response = Invoke-RestMethod -Uri "https://dns.google/resolve?name=autodiscover.$domain&type=CNAME" -Method Get
        $autodiscoverRecord = $response.Answer | ForEach-Object { $_.data }

        if ($autodiscoverRecord) {
            # Join multiple records into a single string if necessary
            return $autodiscoverRecord -join ", "
        } else {
            return "Autodiscover record not found using external API for $domain"
        }
    } catch {
        return "Error retrieving Autodiscover record using external API for $domain $($_)"
    }
}

# Loop through each domain and get its Autodiscover record using both methods
foreach ($domain in $domains) {
    $autodiscoverRecordResolve = Get-AutodiscoverRecord -domain $domain
    $autodiscoverRecordAPI = Get-AutodiscoverRecordFromAPI -domain $domain

    # Store results in a hashtable
    $result = [pscustomobject]@{
        Domain                 = $domain
        Autodiscover_Resolve   = $autodiscoverRecordResolve
        Autodiscover_API       = $autodiscoverRecordAPI
    }

    # Add to results array
    $results += $result
}

# Export results to HTML and display using Out-HTMLView
$results | Out-HTMLView -Title "Autodiscover DNS Record Results"
