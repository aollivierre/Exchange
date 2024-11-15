# Install the Remote Server Administration Tools and import the ActiveDirectory module
# Install-WindowsFeature -Name RSAT-AD-PowerShell
# Import-Module ActiveDirectory

# Retrieve all enabled users from the domain
$users = Get-ADUser -Filter {Enabled -eq $True} -Properties LastLogonDate, LastLogon, DistinguishedName, Enabled

# Set the dynamic export path and create it if it does not exist
$exportPath = "C:\Code\AD\Exports\$(Get-Date -Format 'yyyy-MM-dd')\"
if (!(Test-Path $exportPath)) {
    New-Item -ItemType Directory -Path $exportPath | Out-Null
}

# Export users to a CSV file
$csvExportPath = "${exportPath}all_enabled_users.csv"
$users | Sort-Object LastLogonDate | Select-Object Name, LastLogonDate, @{Name='LastLogon';Expression={[DateTime]::FromFileTime($_.LastLogon)}}, DistinguishedName, Enabled | Export-Csv $csvExportPath -NoTypeInformation

# Display the data in a grid view
$users | Sort-Object LastLogonDate | Select-Object Name, LastLogonDate, @{Name='LastLogon';Expression={[DateTime]::FromFileTime($_.LastLogon)}}, DistinguishedName, Enabled | Out-GridView

# Output the count of users to the console
Write-Host "Total number of enabled users in the domain: $($users.Count)" -ForegroundColor Green