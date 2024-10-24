$ldapPort = 389
$ldapsPort = 636

# Check for existing LDAP rule
$ldapRule = Get-NetFirewallRule -DisplayName "Allow LDAP Port" -ErrorAction SilentlyContinue

if ($null -eq $ldapRule) {
    # Create a firewall rule for the LDAP port
    New-NetFirewallRule -DisplayName "Allow LDAP Port" -Direction Inbound -LocalPort $ldapPort -Protocol TCP -Action Allow -Enabled True
    Write-Host "$(Get-Date) - Created a new rule to allow LDAP port ($ldapPort)." -ForegroundColor Green
} else {
    Write-Host "$(Get-Date) - An existing rule for LDAP port ($ldapPort) is already in place." -ForegroundColor Yellow
}

# Check for existing LDAPS rule
$ldapsRule = Get-NetFirewallRule -DisplayName "Allow LDAPS Port" -ErrorAction SilentlyContinue

if ($null -eq $ldapsRule) {
    # Create a firewall rule for the LDAPS port
    New-NetFirewallRule -DisplayName "Allow LDAPS Port" -Direction Inbound -LocalPort $ldapsPort -Protocol TCP -Action Allow -Enabled True
    Write-Host "$(Get-Date) - Created a new rule to allow LDAPS port ($ldapsPort)." -ForegroundColor Green
} else {
    Write-Host "$(Get-Date) - An existing rule for LDAPS port ($ldapsPort) is already in place." -ForegroundColor Yellow
}
