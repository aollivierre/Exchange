# Install the Remote Server Administration Tools and import the ActiveDirectory module
# Install-WindowsFeature -Name RSAT-AD-PowerShell
# Import-Module ActiveDirectory

# Retrieve all disabled users from the domain
$disabledUsers = Get-ADUser -Filter 'Enabled -eq $False' -Properties LastLogonDate, LastLogon, DistinguishedName, Enabled

# Set the dynamic export path and create it if it does not exist
$exportPath = "C:\Code\AD\Exports\$(Get-Date -Format 'yyyy-MM-dd')\"
if (!(Test-Path $exportPath)) {
    New-Item -ItemType Directory -Path $exportPath | Out-Null
}

# Export disabled users to a CSV file
$csvExportPath = "${exportPath}disabled_users.csv"
$disabledUsers | Sort-Object LastLogonDate | Select-Object Name, LastLogonDate, @{Name='LastLogon';Expression={[DateTime]::FromFileTime($_.LastLogon)}}, DistinguishedName, Enabled | Export-Csv $csvExportPath -NoTypeInformation

# Display the data in a grid view
$disabledUsers | Sort-Object LastLogonDate | Select-Object Name, LastLogonDate, @{Name='LastLogon';Expression={[DateTime]::FromFileTime($_.LastLogon)}}, DistinguishedName, Enabled | Out-GridView

# Output the count of disabled users to the console
Write-Host "Total number of disabled users in the domain: $($disabledUsers.Count)" -ForegroundColor Green