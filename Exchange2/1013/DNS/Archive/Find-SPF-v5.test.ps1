# Define the list of domains
$domains = @(
    "Iqaluit.tunngavik.com",
    "Rankininlet.tunngavik.com",
    "Cambridgebay.tunngavik.com",
    "Ottawa.tunngavik.com",
    "Tunngavik.com"
)

# Function to get the SPF record of a domain
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
            return "SPF record not found for $domain"
        }
    } catch {
        return "Error retrieving SPF record for $domain $_"
    }
}

# Loop through each domain and get its SPF record
foreach ($domain in $domains) {
    $spfRecord = Get-SPFRecord -domain $domain
    Write-Host "Domain: $domain"
    Write-Host "SPF Record: $spfRecord"
    Write-Host "------------------------"
}
