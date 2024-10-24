# Import Active Directory module
Import-Module ActiveDirectory

# Get the closest domain controller
# $domainController = (Get-ADDomainController -Discover -NextClosestSite).HostName

#Using the Public IP of the domain controller (not the private IP)
# $domainController = "172.174.77.131"
# $domainController = "69.42.180.68"
# $domainController = "10.93.0.20"
$domainController = "AGH-DC01.AGH.com"

$ldapPort = 389
$ldapsPort = 636

# Check LDAP port
$ldapConnectionResult = Test-NetConnection -ComputerName $domainController -Port $ldapPort

if ($ldapConnectionResult.TcpTestSucceeded) {
    Write-Host "LDAP port ($ldapPort) is open on the domain controller ($domainController)."
} else {
    Write-Host "LDAP port ($ldapPort) is NOT open on the domain controller ($domainController)."
}

# Check LDAPS port
$ldapsConnectionResult = Test-NetConnection -ComputerName $domainController -Port $ldapsPort

if ($ldapsConnectionResult.TcpTestSucceeded) {
    Write-Host "LDAPS port ($ldapsPort) is open on the domain controller ($domainController)."
} else {
    Write-Host "LDAPS port ($ldapsPort) is NOT open on the domain controller ($domainController)."
}
