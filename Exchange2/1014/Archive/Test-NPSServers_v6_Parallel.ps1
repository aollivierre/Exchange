#requires -Version 7.0

<#
.SYNOPSIS
    This script tests the connectivity and NPS Server role installation status of a list of servers.
.DESCRIPTION
    This script performs the following tests on a list of servers:
    - Ping test
    - RDP port test
    - WSMan test
    - NPS Server role installation status check
    The results are output to a grid view.
.NOTES
    Author: Unknown
    Date: Unknown
    Version: 5.0 - Parallel
    Modified to check NPS role and remove specific domain references
#>

# Domain variable to be set for global use
$domainName = "RAILCAN.ca" # Update this to your domain

# Import the Active Directory module
Import-Module ActiveDirectory

# function TestPing {
#     param($ServerName)
#     $pingResult = Test-Connection -ComputerName $ServerName -Count 2 -Quiet
#     if ($pingResult) {
#         Write-Host "Ping to $ServerName Successful" -ForegroundColor Green
#     } else {
#         Write-Host "Ping to $ServerName Failed" -ForegroundColor Red
#     }
#     return $pingResult
# }

# function TestRDP {
#     param($ServerName)
#     $rdpTest = $false
#     try {
#         $tcpConnection = New-Object System.Net.Sockets.TcpClient($ServerName, 3389)
#         $tcpConnection.ReceiveTimeout = 5000
#         $tcpConnection.SendTimeout = 5000
#         $rdpTest = $true
#         $tcpConnection.Close()
#         Write-Host "RDP to $ServerName Successful" -ForegroundColor Green
#     } catch {
#         Write-Host "RDP to $ServerName Failed" -ForegroundColor Red
#     }
#     return $rdpTest
# }

# function TestWSMan {
#     param($ServerName)
#     $wsManTest = $false
#     try {
#         $wsManTestResult = Test-WSMan -ComputerName $ServerName
#         if ($wsManTestResult) {
#             $wsManTest = $true
#             Write-Host "WSMan to $ServerName Successful" -ForegroundColor Green
#         }
#     } catch {
#         Write-Host "WSMan to $ServerName Failed" -ForegroundColor Red
#     }
#     return $wsManTest
# }

# function CheckNPSServerRole {
#     param($ServerName)
#     $FQDN = "$ServerName.$domainName"
#     try {
#         $scriptBlock = {
#             $feature = Get-WindowsFeature -Name NPAS
#             return $feature.Installed
#         }
#         $result = Invoke-Command -ComputerName $FQDN -ScriptBlock $scriptBlock
#         if ($result) {
#             Write-Host "NPS Server role on $FQDN Installed" -ForegroundColor Green
#         } else {
#             Write-Host "NPS Server role on $FQDN Not installed" -ForegroundColor Yellow
#         }
#         return $result
#     } catch {
#         Write-Host "Could not check NPS Server role on $FQDN $_" -ForegroundColor Red
#         return $false
#     }
# }

# function GetOSVersion {
#     param($ServerName)
#     try {
#         $osInfo = Invoke-Command -ComputerName $ServerName -ScriptBlock {

#             Get-CimInstance -ClassName Win32_OperatingSystem | Select-Object Version, Caption
#         }
#         Write-Host "OS Version for $ServerName retrieved: $($osInfo.Caption)" -ForegroundColor Green
#         return $osInfo
#     } catch {
#         Write-Host "Failed to get OS Version for $ServerName $_" -ForegroundColor Red
#         return $null
#     }
# }

# function CheckServiceStatus {
#     param($ServerName)
#     try {
#         $serviceInfo = Invoke-Command -ComputerName $ServerName -ScriptBlock {
#             $service = Get-WmiObject -Class Win32_Service -Filter "PathName LIKE '%VMagicPPII.exe%'"
#             if ($service) {
#                 return @{
#                     Exists = $true
#                     Running = $service.State -eq 'Running'
#                     DisplayName = $service.DisplayName
#                 }
#             } else {
#                 return @{
#                     Exists = $false
#                     Running = $false
#                     DisplayName = $null
#                 }
#             }
#         }
#         $status = if ($serviceInfo.Running) { "Running" } else { "Stopped" }
#         Write-Host "Service $($serviceInfo.DisplayName) on $ServerName is $status." -ForegroundColor Green
#         return $serviceInfo
#     } catch {
#         Write-Host "Failed to check service status for $ServerName $_" -ForegroundColor Red
#         return $null
#     }
# }

# function TestClusterMembership {
#     param($ServerName)
#     try {
#         $clusterInfo = Invoke-Command -ComputerName $ServerName -ScriptBlock {
#             try {
#                 Get-ClusterNode -ErrorAction Stop
#             } catch {
#                 $clusterRegistryKey = 'HKLM:\Cluster'
#                 Test-Path $clusterRegistryKey
#             }
#         }
#         if ($clusterInfo) {
#             Write-Host "Server $ServerName is part of a cluster." -ForegroundColor Green
#             return $true
#         } else {
#             Write-Host "Server $ServerName is not part of a cluster." -ForegroundColor Yellow
#             return $false
#         }
#     } catch {
#         Write-Host "Failed to check cluster membership for $ServerName $_" -ForegroundColor Red
#         return $null
#     }
# }

# Define and populate the $servers array
$servers = Get-ADComputer -Filter {OperatingSystem -like "*Server*"} -Property Name | Select-Object -ExpandProperty Name

# Initialize an empty concurrent bag for results
$results = [System.Collections.Concurrent.ConcurrentBag[PSObject]]::new()

# Run checks in parallel
$servers | ForEach-Object -Parallel {


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
    
    # function CheckServiceStatus {
    #     param($ServerName)
    #     try {
    #         $serviceInfo = Invoke-Command -ComputerName $ServerName -ScriptBlock {
    #             $service = Get-WmiObject -Class Win32_Service -Filter "PathName LIKE '%VMagicPPII.exe%'"
    #             if ($service) {
    #                 return @{
    #                     Exists = $true
    #                     Running = $service.State -eq 'Running'
    #                     DisplayName = $service.DisplayName
    #                 }
    #             } else {
    #                 return @{
    #                     Exists = $false
    #                     Running = $false
    #                     DisplayName = $null
    #                 }
    #             }
    #         }
    #         $status = if ($serviceInfo.Running) { "Running" } else { "Stopped" }
    #         Write-Host "Service $($serviceInfo.DisplayName) on $ServerName is $status." -ForegroundColor Green
    #         return $serviceInfo
    #     } catch {
    #         Write-Host "Failed to check service status for $ServerName $_" -ForegroundColor Red
    #         return $null
    #     }
    # }
    
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

    $server = $_
    $pingTest = TestPing -ServerName $server
    if (-not $pingTest) {
        Write-Host "Ping test failed for $server. Skipping server." -ForegroundColor Yellow
        return
    }
    $rdpTest = TestRDP -ServerName $server
    $wsManTest = TestWSMan -ServerName $server
    $npsServerRole = CheckNPSServerRole -ServerName $server
    $osVersionInfo = GetOSVersion -ServerName $server
    # $serviceStatusInfo = CheckServiceStatus -ServerName $server
    $isClusterMember = TestClusterMembership -ServerName $server
    $ipAddress = "Unknown"
    try {
        $ipAddress = (Resolve-DnsName $server).IPAddress
    } catch {
        $ipAddress = "Could not resolve"
    }
    $result = [pscustomobject]@{
        Server          = $server
        IPAddress       = $ipAddress
        PingTest        = $pingTest
        RDPPortTest     = $rdpTest
        WSManTest       = $wsManTest
        NPSServerRole   = $npsServerRole
        OSVersion       = $osVersionInfo.Caption
        # ServiceStatus   = if ($serviceStatusInfo.Running) { "Running" } else { "Stopped" }
        # ServiceExists   = $serviceStatusInfo.Exists
        ClusterMember   = $isClusterMember
    }
    $ExecutionContext.InvokeCommand.InvokeScript({ param($resultArray, $item) $resultArray.Add($item) }, $using:results, $result)
} -ThrottleLimit 10

# Convert the concurrent bag to a regular array
$resultsArray = $results.ToArray()

# Output or process $resultsArray as needed
$resultsArray | Select-Object Server, IPAddress, PingTest, RDPPortTest, WSManTest, NPSServerRole, OSVersion, ServiceStatus, ServiceExists, ClusterMember | Out-GridView

# Generate a timestamp for the file name
$timestamp = Get-Date -Format "yyyyMMddHHmmss"

# Specify the file path with the timestamp
$csvFilePath = "C:\code\exports\Test-Servers_$timestamp.csv"

# Export the sorted results to the CSV file
$resultsArray | Export-Csv -Path $csvFilePath -NoTypeInformation

# Display a message indicating the file path of the exported CSV
Write-Host "Exported results to $csvFilePath" -ForegroundColor Green