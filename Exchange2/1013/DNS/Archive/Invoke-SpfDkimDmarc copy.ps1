#the following does not get the full SPF record and it does not get the SPF record for domains that do actually have an SPF record

# Install required modules (if not already installed)
# Install-Module -Name DomainHealthChecker -Scope CurrentUser -Force
# Install-Module -Name PSWriteHTML -Scope CurrentUser -Force

# Import modules
Import-Module -Name DomainHealthChecker
Import-Module -Name PSWriteHTML

# Define your list of domains
$domains = @(
    "Iqaluit.tunngavik.com",
    "Rankininlet.tunngavik.com",
    "cambridgebay.tunngavik.com",
    "ottawa.tunngavik.com",
    "tunngavik.com"
)

# Initialize an array to hold the results
$results = @()

# Loop over each domain and invoke the function
foreach ($domain in $domains) {
    $result = Invoke-SpfDkimDmarc -Name $domain -IncludeDNSSEC
    $results += $result
}

# Export the results to an HTML file and open it in the default browser
$results | Out-HTMLView -Title "Domain Health Check Results" -FilePath "DomainHealthCheckResults.html"