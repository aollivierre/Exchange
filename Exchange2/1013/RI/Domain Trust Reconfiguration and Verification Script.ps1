# First remove the existing incorrect trust
Remove-ADTrust -Identity "RI.nti.local" -Confirm:$false

# Create new forest trust with proper settings
$creds = Get-Credential -Message "Enter Enterprise Admin credentials for RI.nti.local"
New-ADTrust -Name "RI.nti.local" `
            -TrustType Forest `
            -TrustDirection Bidirectional `
            -ForestTransitive $true `
            -SourceName "OTT.NTI.LOCAL" `
            -TargetName "RI.nti.local" `
            -Credential $creds

# Verification commands
function Test-DomainTrust {
    Write-Host "`nVerifying trust configuration:" -ForegroundColor Green
    
    # Check trust properties
    $trust = Get-ADTrust -Filter {Name -eq "RI.nti.local"}
    Write-Host "`nTrust Properties:" -ForegroundColor Yellow
    $trust | Format-List Name, TrustType, ForestTransitive, DisallowTransivity
    
    # Verify trust status using netdom
    Write-Host "`nNetdom Trust Status:" -ForegroundColor Yellow
    netdom verify RI.nti.local /domain:OTT.NTI.LOCAL
    
    # Test forest-wide authentication
    Write-Host "`nTesting Forest Authentication:" -ForegroundColor Yellow
    nltest /server:$env:COMPUTERNAME /sc_verify:RI.nti.local
    
    # Verify Kerberos tickets
    Write-Host "`nVerifying Kerberos Tickets:" -ForegroundColor Yellow
    klist
    
    # Test RPC connectivity
    Write-Host "`nTesting RPC Connectivity:" -ForegroundColor Yellow
    nltest /server:$env:COMPUTERNAME /trusted_domains
}

# Run verification
Test-DomainTrust