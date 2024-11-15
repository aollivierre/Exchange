# Define the list of domains
$domains = @(
    "Iqaluit.tunngavik.com",
    "Rankininlet.tunngavik.com",
    "cambridgebay.tunngavik.com",
    "ottawa.tunngavik.com",
    "tunngavik.com"
)

foreach ($domain in $domains) {
    try {
        # Resolve the domain, handling CNAMEs
        $resolvedDomain = (Resolve-DnsName -Name $domain -Type CNAME -ErrorAction SilentlyContinue).NameHost.TrimEnd('.') -or $domain

        # Retrieve TXT records
        $txtRecords = (Resolve-DnsName -Name $resolvedDomain -Type TXT -ErrorAction Stop).Strings

        # Find SPF record
        $spfRecord = $txtRecords | Where-Object { $_ -like "v=spf1*" } | Select-Object -First 1

        if ($spfRecord) {
            Write-Host "SPF record for $domain`n$spfRecord"
        }
        else {
            Write-Host "No SPF record found for $domain"
        }
    }
    catch {
        Write-Host "An error occurred while retrieving SPF record for $domain $_"
    }
    Write-Host "------------------------"
}
