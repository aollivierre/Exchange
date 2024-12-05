# Get all domain controllers from Active Directory
$dcs = Get-ADDomainController -Filter * | Select-Object Name, IPv4Address, Domain, Site

# Create array to store results
$results = foreach ($dc in $dcs) {
    # Test basic connectivity
    $pingTest = Test-Connection -ComputerName $dc.Name -Count 1 -Quiet
    $portTests = @{}
    
    # Test critical ports if ping succeeds
    if ($pingTest) {
        $ports = @(
            @{Port = 389; Service = "LDAP"},
            @{Port = 88; Service = "Kerberos"},
            @{Port = 445; Service = "SMB"},
            @{Port = 53; Service = "DNS"},
            @{Port = 135; Service = "RPC"}
        )
        
        foreach ($portInfo in $ports) {
            $tcp = New-Object System.Net.Sockets.TcpClient
            try {
                $result = $tcp.BeginConnect($dc.IPv4Address, $portInfo.Port, $null, $null)
                $success = $result.AsyncWaitHandle.WaitOne(1000)
                $portTests[$portInfo.Service] = $success
            }
            catch {
                $portTests[$portInfo.Service] = $false
            }
            finally {
                $tcp.Close()
            }
        }
    }
    
    # Create custom object with results
    [PSCustomObject]@{
        'DC Name' = $dc.Name
        'IP Address' = $dc.IPv4Address
        'Domain' = $dc.Domain
        'Site' = $dc.Site
        'Ping' = if ($pingTest) {"✅"} else {"❌"}
        'LDAP' = if ($portTests['LDAP']) {"✅"} else {"❌"}
        'Kerberos' = if ($portTests['Kerberos']) {"✅"} else {"❌"}
        'SMB' = if ($portTests['SMB']) {"✅"} else {"❌"}
        'DNS' = if ($portTests['DNS']) {"✅"} else {"❌"}
        'RPC' = if ($portTests['RPC']) {"✅"} else {"❌"}
    }
}

# Display results in a formatted table
$results | Format-Table -AutoSize