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


2a01:0111:f403:7053:0000:0000:0000:0000 -> Unable to resolve
216.122.168.228 -> hostingmx.thewebexpert.ca
209.216.150.55 -> server.sigma3inc.com
147.185.114.170 -> 170-114-185-147.rdns.colocationamerica.com
99.209.86.42 -> Unable to resolve
198.2.139.252 -> mail252.atl221.rsgsv.net
198.2.175.120 -> mail120.suw151.rsgsv.net
173.243.133.62 -> gw3062.fortimail.com
216.191.127.138 -> mx1.ssochamplain.ca

#>





# List of IP addresses
$ipAddresses = @(
    "2a01:0111:f403:7053:0000:0000:0000:0000", # Although you've given a CIDR notation (/96), we just need the IP address for resolution.
    "216.122.168.228",
    "209.216.150.55",
    "147.185.114.170",
    "99.209.86.42",
    "198.2.139.252",
    "198.2.175.120",
    "173.243.133.62",
    "216.191.127.138"
)

# Loop through each IP and try to resolve the hostname
foreach ($ip in $ipAddresses) {
    try {
        $result = Resolve-DnsName -Name $ip -ErrorAction Stop
        Write-Output "$ip -> $($result.NameHost)"
    } catch {
        Write-Output "$ip -> Unable to resolve"
    }
}
