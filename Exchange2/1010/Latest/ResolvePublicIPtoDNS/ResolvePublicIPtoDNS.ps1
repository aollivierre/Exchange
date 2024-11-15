<#
.SYNOPSIS
    A short one-line action-based description, e.g. 'Tests if a function is valid'
.DESCRIPTION
    A longer description of the function, its purpose, common use cases, etc.
.NOTES
    Information or caveats about the function e.g. 'This function is not supported in Linux'
.LINK
    Specify a URI to a help page, this will show when Get-Help -Online is used.
.EXAMPLE
    Test-MyTestFunction -Verbose
    Explanation of the function or its result. You can include multiple examples with additional .EXAMPLE lines


72.1.218.181 -> mail.pharmacists.ca
173.243.133.62 -> gw3062.fortimail.com
167.89.98.146 -> o1678998x146.outbound-mail.sendgrid.net
mail.pharmacists.ca -> A record IP: 72.1.218.181
cpha-mail.pharmacists.ca -> A record IP: 72.1.218.181
#>



# List of IP addresses and domains
$addresses = @(
    "72.1.218.181",
    "173.243.133.62",
    "167.89.98.146",
    "mail.pharmacists.ca",   # This is a domain name, but we can also resolve it using Resolve-DnsName to get its IP.
    "cpha-mail.pharmacists.ca" # Same as above
)

# Loop through each address and try to resolve the hostname (for IPs) or IP (for domains)
foreach ($address in $addresses) {
    try {
        $result = Resolve-DnsName -Name $address -ErrorAction Stop
        # Check if the result is for an IP or a domain
        if ($address -match "\b\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\b") { # If it's an IP
            Write-Output "$address -> $($result.NameHost)"
        } else { # If it's a domain
            Write-Output "$address -> $($result.QueryType) record IP: $($result.IPAddress)"
        }
    } catch {
        Write-Output "$address -> Unable to resolve"
    }
}
