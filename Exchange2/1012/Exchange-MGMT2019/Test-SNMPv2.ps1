# PowerShell Script for SNMP Troubleshooting

param (
    [string]$serverIP = "192.168.2.47",
    [string]$community = "NovaRead"
)

# Checking Network Connection
Write-Host "Checking Network Connectivity..."
if (Test-Connection -ComputerName $serverIP -Count 2 -Quiet) {
    Write-Host "Network Connectivity: OK"
} else {
    Write-Host "Network Connectivity: FAILED"
    exit
}

# Checking SNMP Port Accessibility
Write-Host "Checking SNMP Port Accessibility..."
try {
    $tcpclient = New-Object system.Net.Sockets.TcpClient
    $tcpclient.Connect($serverIP, 161)
    Write-Host "SNMP Port Accessibility: OK"
} catch {
    Write-Host "SNMP Port Accessibility: FAILED"
    exit
}

# Running SNMP Walk
Write-Host "Running SNMP Walk..."
try {
    $result = & 'c:\code\SnmpWalk.exe' -r:$serverIP -t:10 -c:$community -os:.1.3.6.1.2.1.1
    if ($result -match "Failed") {
        Write-Host "SNMP Walk: FAILED"
    } else {
        Write-Host "SNMP Walk: OK"
    }
} catch {
    Write-Host "Error running SNMP Walk: $_"
}

# Additional manual steps
Write-Host "Please manually check the following:"
Write-Host "- SNMP Service is running on the server $serverIP"
Write-Host "- SNMP Configuration on the server $serverIP"
Write-Host "- Community String $community is correct"
Write-Host "- SNMP Logs on the server $serverIP"
Write-Host "- Firewall settings on both machines"
