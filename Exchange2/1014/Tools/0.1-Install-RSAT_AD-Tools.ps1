# Install Remote Server Administration Tools (RSAT) and import ActiveDirectory module
if (!(Get-WindowsCapability -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0 -Online | Where-Object { $_.State -eq "Installed" })) {
    Add-WindowsCapability -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0 -Online
}

Import-Module ActiveDirectory