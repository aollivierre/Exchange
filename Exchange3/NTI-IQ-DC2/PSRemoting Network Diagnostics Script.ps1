# PSRemoting Network Diagnostics Script
param(
    [Parameter(Mandatory=$true)]
    [string]$RemoteComputer,
    
    [Parameter(Mandatory=$true)]
    [string]$Domain,
    
    [Parameter(Mandatory=$false)]
    [string]$Username
)

function Write-TestResult {
    param(
        [string]$TestName,
        [bool]$Success,
        [string]$Details
    )
    Write-Host "`n=== $TestName ===" -ForegroundColor Cyan
    if ($Success) {
        Write-Host "Status: PASS" -ForegroundColor Green
    } else {
        Write-Host "Status: FAIL" -ForegroundColor Red
    }
    Write-Host "Details: $Details" -ForegroundColor White
}

function Test-NetworkConnectivity {
    param([string]$ComputerName)
    
    Write-Host "`n=== Network Connectivity Tests ===" -ForegroundColor Yellow
    
    # DNS Resolution
    try {
        $dnsResult = Resolve-DnsName -Name $ComputerName -ErrorAction Stop
        Write-TestResult -TestName "DNS Resolution" -Success $true -Details "Resolved to $($dnsResult.IPAddress)"
        $ip = $dnsResult.IPAddress
    } catch {
        Write-TestResult -TestName "DNS Resolution" -Success $false -Details $_.Exception.Message
        return $false
    }
    
    # ICMP Test
    $pingResult = Test-Connection -ComputerName $ComputerName -Count 4 -Quiet
    Write-TestResult -TestName "ICMP Ping" -Success $pingResult -Details "Ping test to $ComputerName"
    
    # Latency Test
    $latencyResults = Test-Connection -ComputerName $ComputerName -Count 4 | 
        Select-Object -ExpandProperty ResponseTime
    $avgLatency = ($latencyResults | Measure-Object -Average).Average
    Write-TestResult -TestName "Network Latency" -Success ($avgLatency -lt 1000) -Details "Average latency: ${avgLatency}ms"
    
    # Traceroute
    Write-Host "`nTraceroute to $ComputerName" -ForegroundColor Yellow
    Test-NetConnection -ComputerName $ComputerName -TraceRoute | 
        Select-Object -ExpandProperty TraceRoute

    return $true
}

function Test-WinRMPorts {
    param([string]$ComputerName)
    
    Write-Host "`n=== WinRM Port Tests ===" -ForegroundColor Yellow
    
    $ports = @(5985, 5986)
    foreach ($port in $ports) {
        $result = Test-NetConnection -ComputerName $ComputerName -Port $port -WarningAction SilentlyContinue
        Write-TestResult -TestName "Port $port" -Success $result.TcpTestSucceeded -Details "TCP connection to port $port"
    }
}

function Test-WinRMConfiguration {
    Write-Host "`n=== Local WinRM Configuration ===" -ForegroundColor Yellow
    
    # Check WinRM Service
    $winrmService = Get-Service WinRM
    Write-TestResult -TestName "WinRM Service" -Success ($winrmService.Status -eq 'Running') -Details "Service Status: $($winrmService.Status)"
    
    # Check WinRM Configuration
    try {
        $config = winrm get winrm/config
        Write-Host "`nWinRM Configuration:" -ForegroundColor Yellow
        Write-Host $config
    } catch {
        Write-Host "Error getting WinRM configuration: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    # Check Trusted Hosts
    $trustedHosts = Get-Item WSMan:\localhost\Client\TrustedHosts
    Write-TestResult -TestName "Trusted Hosts" -Success $true -Details "Current value: $($trustedHosts.Value)"
}

function Test-RemoteWinRM {
    param(
        [string]$ComputerName,
        [string]$DomainName,
        [string]$Username
    )
    
    Write-Host "`n=== Remote WinRM Tests ===" -ForegroundColor Yellow
    
    # Test WS-Man
    try {
        $wsmanResult = Test-WSMan -ComputerName $ComputerName -ErrorAction Stop
        Write-TestResult -TestName "WS-Man Basic" -Success $true -Details "Successfully connected to remote WinRM"
    } catch {
        Write-TestResult -TestName "WS-Man Basic" -Success $false -Details $_.Exception.Message
    }
    
    # Test Authentication
    if ($Username) {
        $cred = Get-Credential -Message "Enter credentials for $Username" -UserName $Username
        try {
            $session = New-PSSession -ComputerName $ComputerName -Credential $cred -ErrorAction Stop
            Write-TestResult -TestName "Authentication" -Success $true -Details "Successfully authenticated"
            Remove-PSSession $session
        } catch {
            Write-TestResult -TestName "Authentication" -Success $false -Details $_.Exception.Message
        }
    }
}

function Test-FirewallRules {
    Write-Host "`n=== Firewall Rules ===" -ForegroundColor Yellow
    
    $rules = Get-NetFirewallRule -DisplayGroup "Windows Remote Management" |
        Where-Object Enabled -eq 'True' |
        Select-Object Name, DisplayName, Enabled, Direction, Action
    
    if ($rules) {
        Write-Host "Active WinRM Firewall Rules:" -ForegroundColor Yellow
        $rules | Format-Table -AutoSize
    } else {
        Write-Host "No active WinRM firewall rules found!" -ForegroundColor Red
    }
}

# Main execution
Clear-Host
Write-Host "Starting PSRemoting Network Diagnostics" -ForegroundColor Green
Write-Host "Target Computer: $RemoteComputer" -ForegroundColor Green
Write-Host "Domain: $Domain" -ForegroundColor Green
Write-Host "Testing as User: $(if($Username){$Username}else{'Current User'})" -ForegroundColor Green
Write-Host "Local Computer: $env:COMPUTERNAME" -ForegroundColor Green
Write-Host "Local IP(s): $((Get-NetIPAddress -AddressFamily IPv4).IPAddress -join ', ')" -ForegroundColor Green
Write-Host "Timestamp: $(Get-Date)" -ForegroundColor Green

$networkOK = Test-NetworkConnectivity -ComputerName $RemoteComputer
if ($networkOK) {
    Test-WinRMPorts -ComputerName $RemoteComputer
    Test-WinRMConfiguration
    Test-FirewallRules
    Test-RemoteWinRM -ComputerName $RemoteComputer -DomainName $Domain -Username $Username
}

Write-Host "`nDiagnostics Complete" -ForegroundColor Green
Write-Host "Next Steps:" -ForegroundColor Yellow
Write-Host "1. Review any FAIL results above" -ForegroundColor Yellow
Write-Host "2. Check network firewall logs for blocked traffic" -ForegroundColor Yellow
Write-Host "3. Use Wireshark to capture traffic on ports 5985/5986" -ForegroundColor Yellow
Write-Host "4. Verify domain trust relationships" -ForegroundColor Yellow