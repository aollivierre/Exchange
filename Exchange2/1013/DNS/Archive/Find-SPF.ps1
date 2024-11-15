# Ensure PSWriteHTML module is installed
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
        # Retrieve all TXT records for the domain
        $txtRecords = Resolve-DnsName -Name $domain -Type TXT -ErrorAction Stop

        # Initialize a flag to check if SPF record is found
        $spfFound = $false

        # Loop through each TXT record to find the SPF record
        foreach ($record in $txtRecords) {
            foreach ($txt in $record.Strings) {
                if ($txt -match "^v=spf1") {
                    # Add the SPF record to the results array
                    $spfRecords += [PSCustomObject]@{
                        Domain    = $domain
                        SPFRecord = $txt
                    }
                    $spfFound = $true
                }
            }
        }

        # If no SPF record is found, note it in the results
        if (-not $spfFound) {
            $spfRecords += [PSCustomObject]@{
                Domain    = $domain
                SPFRecord = "No SPF record found"
            }
        }
    }
    catch {
        # Handle errors (e.g., domain not found)
        $spfRecords += [PSCustomObject]@{
            Domain    = $domain
            SPFRecord = "Domain not found or error occurred"
        }
    }
}

# Export the results using Out-HTMLView
$spfRecords | Out-HTMLView -Title "SPF Records" -FilePath "SPFRecords.html"
