# Install NetSPF module if not already installed
if (-not (Get-Module -ListAvailable -Name NetSPF)) {
    Install-Module -Name NetSPF -Scope CurrentUser -Force
}

Import-Module NetSPF

# Install PSWriteHTML module if not already installed
if (-not (Get-Module -ListAvailable -Name PSWriteHTML)) {
    Install-Module -Name PSWriteHTML -Scope CurrentUser -Force
}

Import-Module PSWriteHTML

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
    try {
        # Parse the SPF record and expand includes
        $spf = Get-SPFRecord -Domain $domain -ResolveIncludes -ErrorAction Stop

        if ($spf) {
            # Join all mechanisms into a single string
            $spfExpanded = ($spf.Mechanisms | ForEach-Object { $_.ToString() }) -join " "

            # Add the expanded SPF record to the results
            $spfRecords += [PSCustomObject]@{
                Domain        = $domain
                SPFRecord     = $spfExpanded
            }
        } else {
            $spfRecords += [PSCustomObject]@{
                Domain        = $domain
                SPFRecord     = "No SPF record found"
            }
        }
    }
    catch {
        # Handle errors (e.g., domain not found)
        $spfRecords += [PSCustomObject]@{
            Domain        = $domain
            SPFRecord     = "Domain not found or error occurred"
        }
    }
}

# Export the results using Out-HTMLView
$spfRecords | Out-HTMLView -Title "SPF Records" -FilePath "SPFRecords.html"
