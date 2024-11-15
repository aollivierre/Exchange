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

# Invoke the function using the pipeline and collect results
$results = $domains | Invoke-SpfDkimDmarc -IncludeDNSSEC

# Export the results to an HTML file and open it in the default browser
$results | Out-HTMLView -Title "Domain Health Check Results" -FilePath "DomainHealthCheckResults.html"
