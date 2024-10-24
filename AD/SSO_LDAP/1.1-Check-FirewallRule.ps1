<#
.SYNOPSIS
    A short one-line action-based description, e.g. 'Tests if a function is valid'
.DESCRIPTION
    A longer description of the function, its purpose, common use cases, etc.
.NOTES
    Information or caveats about the function e.g. 'This function is not supported in Linux'
.LINK
    Specify a URI to a help page, this will show when Get-Help -Online is used.
.EXAMPLE
    Test-MyTestFunction -Verbose
    Explanation of the function or its result. You can include multiple examples with additional .EXAMPLE lines


    Name                DisplayName                                               DisplayGroup                     Protocol LocalPort RemotePort RemoteAddress Enabled Profile Direction Action
----                -----------                                               ------------                     -------- --------- ---------- ------------- ------- ------- --------- ------
ADDS-LDAPSEC-TCP-In Active Directory Domain Controller - Secure LDAP (TCP-In) Active Directory Domain Services TCP      636       Any        Any              True     Any   Inbound  Allow
ADDS-LDAP-TCP-In    Active Directory Domain Controller - LDAP (TCP-In)        Active Directory Domain Services TCP      389       Any        Any              True     Any   Inbound  Allow
ADDS-LDAP-UDP-In    Active Directory Domain Controller - LDAP (UDP-In)        Active Directory Domain Services UDP      389       Any        Any              True     Any   Inbound  Allow


    #>




$ldapPort = 389
$ldapsPort = 636

Get-NetFirewallRule -Direction Inbound |
Where-Object {($_ | Get-NetFirewallPortFilter).LocalPort -eq $ldapPort -or ($PSItem | Get-NetFirewallPortFilter).LocalPort -eq $ldapsPort} |
Format-Table -Property Name,
DisplayName,
DisplayGroup,
@{Name='Protocol';Expression={($_ | Get-NetFirewallPortFilter).Protocol}},
@{Name='LocalPort';Expression={($_ | Get-NetFirewallPortFilter).LocalPort}},
@{Name='RemotePort';Expression={($_ | Get-NetFirewallPortFilter).RemotePort}},
@{Name='RemoteAddress';Expression={($_ | Get-NetFirewallAddressFilter).RemoteAddress}},
Enabled,
Profile,
Direction,
Action