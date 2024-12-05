# Function to add conditional forwarder using dnscmd
function Add-SimpleConditionalForwarder {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ForwardDomain = "RI.NTI.LOCAL",
        
        [Parameter(Mandatory=$true)]
        [string[]]$ServerIPs
    )
    
    Write-Host "Adding conditional forwarder for $ForwardDomain to servers: $($ServerIPs -join ', ')" -ForegroundColor Green
    
    try {
        # Remove existing forwarder first
        $remove = dnscmd $env:COMPUTERNAME /ZoneDelete $ForwardDomain /f
        Write-Host "Cleaned up existing forwarder (if any)" -ForegroundColor Yellow
        
        # Add new forwarder
        $serverList = $ServerIPs -join " "
        $add = dnscmd $env:COMPUTERNAME /ZoneAdd $ForwardDomain /Forwarder $serverList
        
        if ($LASTEXITCODE -eq 0) {
            Write-Host "`nSuccessfully added conditional forwarder" -ForegroundColor Green
            
            # Verify the setup
            Write-Host "`nTesting DNS resolution:" -ForegroundColor Cyan
            nslookup -type=NS $ForwardDomain
            
            Write-Host "`nTesting DC location:" -ForegroundColor Cyan
            nltest /dsgetdc:$ForwardDomain
        }
        else {
            Write-Host "Failed to add conditional forwarder. Exit code: $LASTEXITCODE" -ForegroundColor Red
        }
    }
    catch {
        Write-Host "Error occurred: $_" -ForegroundColor Red
    }
}

# Example usage (uncomment and replace IP with actual DC IP):
Add-SimpleConditionalForwarder -ForwardDomain "RI.NTI.LOCAL" -ServerIPs @("192.168.0.11")