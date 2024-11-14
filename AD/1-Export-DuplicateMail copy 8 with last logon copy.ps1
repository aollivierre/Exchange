# Import Active Directory module
Import-Module ActiveDirectory

# Query all AD users, selecting necessary properties, including LastLogon
$users = Get-ADUser -Filter * -Properties mail, LastLogonDate, Enabled, LastLogon, DistinguishedName

# Process each user to determine the most recent logon and prepare the output
$processedUsers = foreach ($user in $users) {
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
        Mail = $user.Mail
        Enabled = $user.Enabled
        LastLogonDate = $user.LastLogonDate
        LastLogon = $lastLogonDateTime
        MostRecentLogon = $mostRecentLogon
    }
}

# Identify duplicates based on the mail attribute
$duplicates = $processedUsers | Where-Object { $_.Mail -ne $null } | Group-Object -Property Mail | Where-Object { $_.Count -gt 1 }

# Prepare data for export, GridView, and console output with highlighting
$outputData = foreach ($dup in $duplicates) {
    $sortedGroup = $dup.Group | Sort-Object -Property MostRecentLogon -Descending
    $mostRecentUser = $sortedGroup[0]
    foreach ($user in $sortedGroup) {
        if ($user -eq $mostRecentUser) {
            # Highlight the most recent user in green
            Write-Host "Most Recent Logon (Green): CN: $($user.CommonName) - Mail: $($user.Mail) - Most Recent Logon: $($user.MostRecentLogon)" -ForegroundColor Green
        } else {
            # Other duplicates are shown in yellow
            Write-Host "Other Duplicates: CN: $($user.CommonName) - Mail: $($user.Mail) - Most Recent Logon: $($user.MostRecentLogon)" -ForegroundColor Yellow
        }
        # Return the user for export and GridView display
        $user
    }
}

# Export and GridView parts remain unchanged...

# Define the export path with timestamp
$timestamp = Get-Date -Format "yyyyMMddHHmmss"
$exportFolder = Join-Path -Path $PSScriptRoot -ChildPath "exports"
$exportPath = Join-Path -Path $exportFolder -ChildPath "ADUsersWithDuplicateMail_$timestamp.csv"

# Ensure the exports folder exists
if (-not (Test-Path -Path $exportFolder)) {
    New-Item -ItemType Directory -Path $exportFolder | Out-Null
}

# Export the data to CSV
$outputData | Export-Csv -Path $exportPath -NoTypeInformation

# Display in GridView
$outputData | Out-GridView -Title "AD Users with Duplicate Mail Attribute and Recent Logon"


# Display totals in the console
$TotalUsers = $processedUsers.Count
$TotalDuplicates = $duplicates.Count
# Write-Host "Total Users Processed: $TotalUsers" -ForegroundColor Cyan
# Write-Host "Total Duplicate Groups: $TotalDuplicates" -ForegroundColor Cyan


# Continue from the previous script...

# Calculate additional totals for output
$TotalDuplicateUsers = ($duplicates | ForEach-Object { $_.Group }).Count
$TotalEnabledDuplicates = ($duplicates | ForEach-Object { $_.Group }).Where({ $_.Enabled -eq $true }).Count
$TotalDisabledDuplicates = ($duplicates | ForEach-Object { $_.Group }).Where({ $_.Enabled -eq $false }).Count

# Display totals in the console
Write-Host "Total Users Processed: $TotalUsers" -ForegroundColor Cyan
Write-Host "Total Duplicate Groups: $TotalDuplicates" -ForegroundColor Cyan
Write-Host "Total Duplicate Users: $TotalDuplicateUsers" -ForegroundColor Cyan
Write-Host "Total Enabled Duplicate Users: $TotalEnabledDuplicates" -ForegroundColor Green
Write-Host "Total Disabled Duplicate Users: $TotalDisabledDuplicates" -ForegroundColor Red