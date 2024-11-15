# Install PSWriteHTML module if not already installed
if (-not (Get-Module -ListAvailable -Name PSWriteHTML)) {
    Install-Module -Name PSWriteHTML -Scope CurrentUser -Force
}

# Import the PSWriteHTML module
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
        # Initialize variable for resolved domain
        $resolvedDomain = $domain

        # Check if the domain has a CNAME record
        $cnameRecord = Resolve-DnsName -Name $domain -Type CNAME -ErrorAction SilentlyContinue
        if ($cnameRecord) {
            $resolvedDomain = $cnameRecord.NameHost.TrimEnd('.')
        }

        # Retrieve all TXT records for the resolved domain
        $txtRecords = Resolve-DnsName -Name $resolvedDomain -Type TXT -ErrorAction Stop

        # Initialize a variable to store the SPF record
        $spfRecord = "No SPF record found"

        # Loop through each TXT record to find the SPF record
        foreach ($record in $txtRecords) {
            foreach ($txt in $record.Strings) {
                if ($txt -like "v=spf1*") {
                    $spfRecord = $txt
                    break
                }
            }
            if ($spfRecord -ne "No SPF record found") {
                break
            }
        }
    }
    catch {
        $spfRecord = "Error retrieving SPF record"
    }

    # Add the result to the array
    $spfRecords += [PSCustomObject]@{
        Domain    = $domain
        SPFRecord = $spfRecord
    }

    # Output to console
    Write-Host "Domain: $domain" -ForegroundColor Cyan
    Write-Host "SPF Record: $spfRecord" -ForegroundColor Green
    Write-Host "------------------------" -ForegroundColor DarkGray
}

# Export the results to an HTML file using Out-HTMLView
$spfRecords | Out-HTMLView -Title "SPF Records" -FilePath "SPFRecords.html"
