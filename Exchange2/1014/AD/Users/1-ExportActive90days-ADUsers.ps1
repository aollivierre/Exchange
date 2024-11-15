
#The following script generated time stamps for last logon dates that were not so accurate I eneded up exporting the AD Tidy report. Perhpas in the future I can modify this script to use the last logon date attribute instead of the last logon date time stamp.


# Install the Remote Server Administration Tools and import the ActiveDirectory module
# Install-WindowsFeature -Name RSAT-AD-PowerShell
# Import-Module ActiveDirectory

# Set variables for the search base and inactive duration
$daysInactive = 90

# Calculate the cutoff date
$currentDate = Get-Date
$cutoffDate = $currentDate.AddDays(-$daysInactive)

# Retrieve enabled users from the domain who have been active in the last 90 days
$users = Get-ADUser -Filter {Enabled -eq $True -and LastLogonDate -gt $cutoffDate} -Properties LastLogonDate, LastLogon, DistinguishedName, Enabled | Where-Object { $_.Enabled -eq $True }

# Set the dynamic export path and create it if it does not exist
$exportPath = "C:\Code\AD\Exports\$(Get-Date -Format 'yyyy-MM-dd')\"
if (!(Test-Path $exportPath)) {
    New-Item -ItemType Directory -Path $exportPath | Out-Null
}

# Export active users to a CSV file
$users | Sort-Object LastLogonDate | Select-Object Name, LastLogonDate, @{Name='LastLogon';Expression={[DateTime]::FromFileTime($_.LastLogon)}}, DistinguishedName, Enabled | Export-Csv "${exportPath}active_90days_users.csv" -NoTypeInformation

# Output the count of users in the domain and the count of active users to the console
Write-Host "Total number of enabled users in the domain: $($users.Count)" -ForegroundColor Green
Write-Host "Total number of users who have been active in the last $daysInactive days: $($users.Count)" -ForegroundColor Blue