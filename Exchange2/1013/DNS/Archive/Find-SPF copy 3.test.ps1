# Define the list of domains
$domains = @(
    "Iqaluit.tunngavik.com",
    "Rankininlet.tunngavik.com",
    "Cambridgebay.tunngavik.com",
    "Ottawa.tunngavik.com",
    "Tunngavik.com"
)

# Function to parse SPF record elements
function Parse-SPFRecord {
    param (
        [string]$spfRecord
    )

    $spfParts = $spfRecord -split "\s+"
    
    foreach ($part in $spfParts) {
        if ($part -match "^v=spf1") {
            Write-Host "v=spf1 - SPF record version"
        }
        elseif ($part -match "^ip4:(.+)") {
            Write-Host "+ ip4 $($matches[1]) - Match if IP is in the given range."
        }
        elseif ($part -match "^include:(.+)") {
            Write-Host "+ include $($matches[1]) - The specified domain is searched for an 'allow'."
        }
        elseif ($part -eq "~all") {
            Write-Host "~ all - SoftFail (Always matches at the end of your record)"
        }
        elseif ($part -eq "-all") {
            Write-Host "- all - Fail (Only the listed hosts are allowed to send mail)"
        }
    }
}

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
    Write-Host "Domain: $domain"
    $spfRecord = Get-SPFRecord -domain $domain
    if ($spfRecord -is [string]) {
        Write-Host "SPF Record: $spfRecord"
    } else {
        Parse-SPFRecord -spfRecord $spfRecord
    }
    Write-Host "------------------------"
}
