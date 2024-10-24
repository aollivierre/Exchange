$SMTPPort = 25
$SMTPsPort = 587

Get-NetFirewallRule -Direction Inbound |
Where-Object { ($_ | Get-NetFirewallPortFilter).LocalPort -eq $SMTPPort -or ($PSItem | Get-NetFirewallPortFilter).LocalPort -eq $SMTPsPort } |
Format-Table -Property Name,
DisplayName,
DisplayGroup,
@{Name = 'Protocol'; Expression = { ($_ | Get-NetFirewallPortFilter).Protocol } },
@{Name = 'LocalPort'; Expression = { ($_ | Get-NetFirewallPortFilter).LocalPort } },
@{Name = 'RemotePort'; Expression = { ($_ | Get-NetFirewallPortFilter).RemotePort } },
@{Name = 'RemoteAddress'; Expression = { ($_ | Get-NetFirewallAddressFilter).RemoteAddress } },
Enabled,
Profile,
Direction,
Action




Get-NetFirewallRule -Direction Outbound |
Where-Object { ($_ | Get-NetFirewallPortFilter).LocalPort -eq $SMTPPort -or ($PSItem | Get-NetFirewallPortFilter).LocalPort -eq $SMTPsPort } |
Format-Table -Property Name,
DisplayName,
DisplayGroup,
@{Name = 'Protocol'; Expression = { ($_ | Get-NetFirewallPortFilter).Protocol } },
@{Name = 'LocalPort'; Expression = { ($_ | Get-NetFirewallPortFilter).LocalPort } },
@{Name = 'RemotePort'; Expression = { ($_ | Get-NetFirewallPortFilter).RemotePort } },
@{Name = 'RemoteAddress'; Expression = { ($_ | Get-NetFirewallAddressFilter).RemoteAddress } },
Enabled,
Profile,
Direction,
Action