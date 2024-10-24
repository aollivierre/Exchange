# Import Active Directory module
Import-Module ActiveDirectory

# Query all AD users, selecting necessary properties, including LastLogon and proxyAddresses
$users = Get-ADUser -Filter * -Properties proxyAddresses, LastLogonDate, Enabled, LastLogon, DistinguishedName

# Process each user to determine the most recent logon and prepare the output
$processedUsers = foreach ($user in $users) {
    # Filter out x400 and x500 addresses
    $validProxyAddresses = $user.proxyAddresses | Where-Object { $_ -notmatch "^(x400|x500):" }

    # Convert LastLogon to DateTime format, providing a default if LastLogon is $null
    $lastLogonDateTime = if ($user.LastLogon) { 
        [DateTime]::FromFileTime($user.LastLogon)
    } else { 
        [DateTime]::FromFileTime(0) # Default value, represents 1601-01-01 00:00:00
    }

    # Determine the most recent of LastLogonDate and LastLogonDateTime
    $mostRecentLogon = @($user.LastLogonDate, $lastLogonDateTime | Where-Object { $_ } | Sort-Object -Descending)[0]

    # Output a custom object for each user
    [PSCustomObject]@{
        CommonName = $user.Name
        ProxyAddresses = $validProxyAddresses -join "; "
        Enabled = $user.Enabled
        LastLogonDate = $user.LastLogonDate
        LastLogon = $lastLogonDateTime
        MostRecentLogon = $mostRecentLogon
        DistinguishedName = $user.DistinguishedName
    }
}

# Identify duplicates based on the proxyAddresses attribute
$duplicates = $processedUsers | Where-Object { $_.ProxyAddresses -ne $null } | Group-Object -Property ProxyAddresses | Where-Object { $_.Count -gt 1 }

# Prepare data for export, GridView, and console output with highlighting
$outputData = foreach ($dup in $duplicates) {
    $sortedGroup = $dup.Group | Sort-Object -Property MostRecentLogon -Descending
    $mostRecentUser = $sortedGroup[0]
    foreach ($user in $sortedGroup) {
        if ($user -eq $mostRecentUser) {
            # Highlight the most recent user in green
            Write-Host "Most Recent Logon (Green): CN: $($user.CommonName) - ProxyAddresses: $($user.ProxyAddresses) - Most Recent Logon: $($user.MostRecentLogon) - DN: $($user.DistinguishedName)" -ForegroundColor Green
        } else {
            # Other duplicates are shown in yellow
            Write-Host "Other Duplicates: CN: $($user.CommonName) - ProxyAddresses: $($user.ProxyAddresses) - Most Recent Logon: $($user.MostRecentLogon) - DN: $($user.DistinguishedName)" -ForegroundColor Yellow
        }
        # Return the user for export and GridView display
        $user
    }
}

# Define the export path with timestamp
$timestamp = Get-Date -Format "yyyyMMddHHmmss"
$exportFolder = Join-Path -Path $PSScriptRoot -ChildPath "exports"
$exportPath = Join-Path -Path $exportFolder -ChildPath "ADUsersWithDuplicateProxyAddresses_$timestamp.csv"

# Ensure the exports folder exists
if (-not (Test-Path -Path $exportFolder)) {
    New-Item -ItemType Directory -Path $exportFolder | Out-Null
}

# Export the data to CSV
$outputData | Export-Csv -Path $exportPath -NoTypeInformation

# Display in GridView
$outputData | Out-GridView -Title "AD Users with Duplicate ProxyAddresses Attribute and Recent Logon"

# Calculate additional totals for output
$TotalUsers = $processedUsers.Count
$TotalDuplicates = $duplicates.Count
$TotalDuplicateUsers = ($duplicates | ForEach-Object { $_.Group }).Count
$TotalEnabledDuplicates = ($duplicates | ForEach-Object { $_.Group }).Where({ $_.Enabled -eq $true }).Count
$TotalDisabledDuplicates = ($duplicates | ForEach-Object { $_.Group }).Where({ $_.Enabled -eq $false }).Count

# Display totals in the console
Write-Host "Total Users Processed: $TotalUsers" -ForegroundColor Cyan
Write-Host "Total Duplicate Groups: $TotalDuplicates" -ForegroundColor Cyan
Write-Host "Total Duplicate Users: $TotalDuplicateUsers" -ForegroundColor Cyan
Write-Host "Total Enabled Duplicate Users: $TotalEnabledDuplicates" -ForegroundColor Green
Write-Host "Total Disabled Duplicate Users: $TotalDisabledDuplicates" -ForegroundColor Red
