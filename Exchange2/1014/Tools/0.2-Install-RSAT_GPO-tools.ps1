# Check if Group Policy Management Tools are installed
if (!(Get-WindowsCapability -Name Rsat.GroupPolicy.Management.Tools~~~~0.0.1.0 -Online | Where-Object { $_.State -eq "Installed" })) {
    # Install Group Policy Management Tools
    Add-WindowsCapability -Name Rsat.GroupPolicy.Management.Tools~~~~0.0.1.0 -Online
}