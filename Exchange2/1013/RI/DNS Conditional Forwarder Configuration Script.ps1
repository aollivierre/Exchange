# Function to set up and verify DNS conditional forwarder
function Set-DomainDNSForwarder {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ForwardDomain = "RI.NTI.LOCAL",
        
        [Parameter(Mandatory=$true)]
        [string[]]$MasterServers,
        
        [string]$DNSServer = $env:COMPUTERNAME
    )
    
    Write-Host "Setting up conditional forwarder for $ForwardDomain..." -ForegroundColor Green
    
    try {
        # Remove existing forwarder if it exists
        Remove-DnsServerConditionalForwarderZone -Name $ForwardDomain -Force -ErrorAction SilentlyContinue
        
        # Add new conditional forwarder
        Add-DnsServerConditionalForwarderZone `
            -Name $ForwardDomain `
            -MasterServers $MasterServers `
            -ReplicationScope "Forest" `
            -PassThru
        
        # Verify configuration
        Write-Host "`nVerifying DNS configuration:" -ForegroundColor Yellow
        
        # Check conditional forwarder
        $forwarder = Get-DnsServerConditionalForwarderZone -Name $ForwardDomain
        Write-Host "`nConditional Forwarder Details:" -ForegroundColor Cyan
        $forwarder | Format-List Name, ZoneType, MasterServers, ReplicationScope
        
        # Test DNS resolution
        Write-Host "`nTesting DNS resolution:" -ForegroundColor Cyan
        Resolve-DnsName -Name $ForwardDomain -Type NS -Server $DNSServer
        
        # Test SRV record resolution
        Write-Host "`nTesting SRV record resolution:" -ForegroundColor Cyan
        Resolve-DnsName -Name "_ldap._tcp.$ForwardDomain" -Type SRV -Server $DNSServer
    }
    catch {
        Write-Host "Error: $_" -ForegroundColor Red
        Write-Host "Stack: $($_.ScriptStackTrace)" -ForegroundColor Red
    }
}

# Example usage (you'll need to replace IP_ADDRESS with actual DC IP):
Set-DomainDNSForwarder -ForwardDomain "RI.NTI.LOCAL" -MasterServers "192.168.0.11"