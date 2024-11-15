# Install the Remote Server Administration Tools and import the ActiveDirectory module
# Install-WindowsFeature -Name RSAT-AD-PowerShell
# Import-Module ActiveDirectory

# Set variables for the search base and stale durations
$daysInactive = 90

# Get the domain controller and workstations OU
# $domain = Get-ADDomain
# $workstationsOU = Get-ADOrganizationalUnit -Filter 'Name -eq "Workstations"' -SearchBase $domain.DistinguishedName
# $workstationsOU = Get-ADOrganizationalUnit -LDAPFilter '(name=Workstations)' -SearchBase 'OU=Workstations,DC=GLEBE,DC=LOCAL'


# Calculate the cutoff date
$currentDate = Get-Date
$cutoffDate = $currentDate.AddDays(-$daysInactive)

# Retrieve computers from the Workstations OU
# $computers = Get-ADComputer -Filter * -SearchBase $workstationsOU.DistinguishedName -Properties LastLogonDate, LastLogon, OperatingSystem, DistinguishedName | Where-Object { $_.Name -ne "AZUREADSSOACC" }
$computers = Get-ADComputer -Filter * -Properties LastLogonDate, LastLogon, OperatingSystem, DistinguishedName | Where-Object { $_.Name -ne "AZUREADSSOACC" }

# Filter computers based on LastLogon and LastLogonDate
$inactiveComputers = $computers | Where-Object { (-not $_.LastLogonDate -or $_.LastLogonDate -lt $cutoffDate) -and (-not $_.LastLogon -or [DateTime]::FromFileTime($_.LastLogon) -lt $cutoffDate) }

# Set the dynamic export path and create it if it does not exist
$exportPath = "C:\Code\AD\Exports\$(Get-Date -Format 'yyyy-MM-dd')\"
if (!(Test-Path $exportPath)) {
    New-Item -ItemType Directory -Path $exportPath | Out-Null
}

# Export inactive computers to a CSV file
$inactiveComputers | Sort-Object LastLogonDate | Select-Object Name, LastLogonDate, @{Name='LastLogon';Expression={[DateTime]::FromFileTime($_.LastLogon)}}, OperatingSystem, DistinguishedName | Export-Csv "${exportPath}premove_inactive_computers.csv" -NoTypeInformation
# $inactiveComputers | Sort-Object LastLogonDate | Select-Object Name| Export-Csv "${exportPath}inactive_computers2.csv" -NoTypeInformation

# Output the count of computers in the Workstations OU and the count of inactive computers to the console
Write-Host "Total number of computers in AD: $($computers.Count)" -ForegroundColor Yellow
Write-Host "Total number of computers that have not logged in for more than $daysInactive days: $($inactiveComputers.Count)" -ForegroundColor Red