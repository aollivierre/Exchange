<#
.SYNOPSIS
    This script tests the connectivity and Print Server role installation status of a list of servers.
.DESCRIPTION
    This script performs the following tests on a list of servers:
    - Ping test
    - RDP port test
    - WSMan test
    - Print Server role installation status check
    The results are output to a grid view.
.PARAMETER ServerName
    The name of the server to test.
.EXAMPLE
    .\Test-PrintServersv2-ICMP-RDP-WSMAN-v2.ps1
    This example runs the script and tests the servers listed in the $servers array.
.NOTES
    Author: Unknown
    Date: Unknown
    Version: 3.0
#>


# Domain variable to be set for global use
$domainName = "RAILCAN.ca" # Update this to your domain


# Import the Active Directory module
Import-Module ActiveDirectory

function TestPing {
    param($ServerName)
    $pingResult = Test-Connection -ComputerName $ServerName -Count 2 -Quiet
    if ($pingResult) {
        Write-Host "Ping to $ServerName Successful" -ForegroundColor Green
    } else {
        Write-Host "Ping to $ServerName Failed" -ForegroundColor Red
    }
    return $pingResult
}

function TestRDP {
    param($ServerName)
    $rdpTest = $false
    try {
        $tcpConnection = New-Object System.Net.Sockets.TcpClient($ServerName, 3389)
        $tcpConnection.ReceiveTimeout = 5000
        $tcpConnection.SendTimeout = 5000
        $rdpTest = $true
        $tcpConnection.Close()
        Write-Host "RDP to $ServerName Successful" -ForegroundColor Green
    } catch {
        Write-Host "RDP to $ServerName Failed" -ForegroundColor Red
    }
    return $rdpTest
}

function TestWSMan {
    param($ServerName)
    $wsManTest = $false
    try {
        $wsManTestResult = Test-WSMan -ComputerName $ServerName
        if ($wsManTestResult) {
            $wsManTest = $true
            Write-Host "WSMan to $ServerName Successful" -ForegroundColor Green
        }
    } catch {
        Write-Host "WSMan to $ServerName Failed" -ForegroundColor Red
    }
    return $wsManTest
}

function CheckPrintServerRole {
    param($ServerName)
    $FQDN = "$ServerName.railcan.ca"
    try {
        $scriptBlock = {
            $feature = Get-WindowsFeature -Name Print-Server
            return $feature.Installed
        }
        $result = Invoke-Command -ComputerName $FQDN -ScriptBlock $scriptBlock
        if ($result) {
            Write-Host "Print Server role on $FQDN Installed" -ForegroundColor Green
        } else {
            Write-Host "Print Server role on $FQDN Not installed" -ForegroundColor Yellow
        }
        return $result
    } catch {
        Write-Host "Could not check Print Server role on $FQDN $_" -ForegroundColor Red
        return $false
    }
}




function CheckRDSRole {
    param($ServerName)
    $FQDN = "$ServerName.railcan.ca"
    try {
        $scriptBlock = {
            # You can change 'RDS-RD-Server' to 'Remote-Desktop-Services' based on your specific needs
            $feature = Get-WindowsFeature -Name 'RDS-RD-Server'
            return $feature.Installed
        }
        $result = Invoke-Command -ComputerName $FQDN -ScriptBlock $scriptBlock
        if ($result) {
            Write-Host "RDS role on $FQDN is installed" -ForegroundColor Green
        } else {
            Write-Host "RDS role on $FQDN is not installed" -ForegroundColor Yellow
        }
        return $result
    } catch {
        Write-Host "Could not check RDS role on $FQDN $_" -ForegroundColor Red
        return $false
    }
}



function CheckNPSServerRole {
    param($ServerName)
    $FQDN = "$ServerName.$domainName"
    try {
        $scriptBlock = {
            $feature = Get-WindowsFeature -Name NPAS
            return $feature.Installed
        }
        $result = Invoke-Command -ComputerName $FQDN -ScriptBlock $scriptBlock
        if ($result) {
            Write-Host "NPS Server role on $FQDN Installed" -ForegroundColor Green
        } else {
            Write-Host "NPS Server role on $FQDN Not installed" -ForegroundColor Yellow
        }
        return $result
    } catch {
        Write-Host "Could not check NPS Server role on $FQDN $_" -ForegroundColor Red
        return $false
    }
}


function GetOSVersion {
    param($ServerName)
    try {
        $osInfo = Invoke-Command -ComputerName $ServerName -ScriptBlock {
            Get-CimInstance -ClassName Win32_OperatingSystem | Select-Object Version, Caption
        }
        Write-Host "OS Version for $ServerName retrieved: $($osInfo.Caption)" -ForegroundColor Green
        return $osInfo
    } catch {
        Write-Host "Failed to get OS Version for $ServerName $_" -ForegroundColor Red
        return $null
    }
}

function CheckServiceStatus {
    param($ServerName)
    try {
        $serviceInfo = Invoke-Command -ComputerName $ServerName -ScriptBlock {
            $service = Get-WmiObject -Class Win32_Service -Filter "PathName LIKE '%VMagicPPII.exe%'"
            if ($service) {
                return @{
                    Exists = $true
                    Running = $service.State -eq 'Running'
                    DisplayName = $service.DisplayName
                }
            } else {
                return @{
                    Exists = $false
                    Running = $false
                    DisplayName = $null
                }
            }
        }
        $status = if ($serviceInfo.Running) { "Running" } else { "Stopped" }
        Write-Host "Service $($serviceInfo.DisplayName) on $ServerName is $status." -ForegroundColor Green
        return $serviceInfo
    } catch {
        Write-Host "Failed to check service status for $ServerName $_" -ForegroundColor Red
        return $null
    }
}



function TestClusterMembership {
    param($ServerName)
    try {
        $clusterInfo = Invoke-Command -ComputerName $ServerName -ScriptBlock {
            try {
                Get-ClusterNode -ErrorAction Stop
            } catch {
                $clusterRegistryKey = 'HKLM:\Cluster'
                Test-Path $clusterRegistryKey
            }
        }
        if ($clusterInfo) {
            Write-Host "Server $ServerName is part of a cluster." -ForegroundColor Green
            return $true
        } else {
            Write-Host "Server $ServerName is not part of a cluster." -ForegroundColor Yellow
            return $false
        }
    } catch {
        Write-Host "Failed to check cluster membership for $ServerName $_" -ForegroundColor Red
        return $null
    }
}





# Define the servers

# Fetch servers from Active Directory
$servers = Get-ADComputer -Filter {OperatingSystem -like "*Server*"} -Property Name | Select-Object -ExpandProperty Name

# Create an empty array to hold results
$results = @()

foreach ($server in $servers) {
    # Perform ping test
    $pingTest = TestPing -ServerName $server

    # Skip the server if ping test fails
    if (-not $pingTest) {
        Write-Host "Ping test failed for $server. Skipping server." -ForegroundColor Yellow
        continue
    }

    # Perform other tests
    $rdpTest = TestRDP -ServerName $server
    $wsManTest = TestWSMan -ServerName $server
    # $printServerRole = CheckPrintServerRole -ServerName $server
    $RDServerRole = CheckRDSRole -ServerName $server
    $osVersionInfo = GetOSVersion -ServerName $server
    $npsServerRole = CheckNPSServerRole -ServerName $server
    # $serviceStatusInfo = CheckServiceStatus -ServerName $server
    # $isClusterMember = TestClusterMembership -ServerName $server

    # Get IP Address
    $ipAddress = "Unknown"
    try {
        $ipAddress = (Resolve-DnsName $server).IPAddress
        Write-Host "IP Address for $server $ipAddress" -ForegroundColor Green
    } catch {
        Write-Host "Could not resolve IP address for $server $_" -ForegroundColor Red
    }

    # Store the result
    $result = New-Object PSObject -Property @{
        Server          = $server
        IPAddress       = $ipAddress
        PingTest        = $pingTest
        RDPPortTest     = $rdpTest
        WSManTest       = $wsManTest
        NPSServerRole   = $npsServerRole
        # PrintServerRole = $printServerRole
        RDSServerRole   = $RDServerRole
        OSVersion       = $osVersionInfo.Caption
        # ServiceStatus   = if ($serviceStatusInfo.Running) { "Running" } else { "Stopped" }
        # ServiceExists   = $serviceStatusInfo.Exists
        # ClusterMember   = $isClusterMember
    }

    # Add the result to the array
    $results += $result
}


# Output the results to a grid view
$results | Select-Object Server, IPAddress, PingTest, RDPPortTest, WSManTest, RDSServerRole, OSVersion, NPSServerRole | Out-GridView



# Generate a timestamp for the file name
$timestamp = Get-Date -Format "yyyyMMddHHmmss"

# Specify the file path with the timestamp
# $csvFilePath = "C:\Code\CB\RDS\UnivCan\Test-RDS-Servers_$timestamp.csv"
$csvFilePath = "C:\Code\Test-NPS-Servers_$timestamp.csv"

# Export the sorted results to the CSV file
$results | Select-Object Server, IPAddress, PingTest, RDPPortTest, WSManTest, RDSServerRole, OSVersion | Export-Csv -Path $csvFilePath -NoTypeInformation


# Display a message indicating the file path of the exported CSV
Write-Host "Exported results to $csvFilePath" -ForegroundColor Green