# Import the necessary modules
Import-Module ActiveDirectory

# Get the mailboxes
$mailboxes = Get-Mailbox

# Query all AD users, selecting necessary properties, including LastLogon
$users = Get-ADUser -Filter * -Properties SamAccountName, LastLogonDate, LastLogon, Enabled

# Initialize an array to hold the processed mailbox details
$mailboxDetails = @()

# Iterate through each mailbox and retrieve the required information
foreach ($mailbox in $mailboxes) {
    # Find the corresponding AD user by SamAccountName (matching the Alias of the mailbox)
    $adUser = $users | Where-Object { $_.SamAccountName -eq $mailbox.Alias }

    if ($adUser) {
        # Convert LastLogon to DateTime format, providing a default if LastLogon is $null
        $lastLogonDateTime = if ($adUser.LastLogon) { 
            [DateTime]::FromFileTime($adUser.LastLogon)
        } else { 
            [DateTime]::FromFileTime(0) # Default value, represents 1601-01-01 00:00:00
        }

        # Determine the most recent of LastLogonDate and LastLogonDateTime
        $mostRecentLogon = @($adUser.LastLogonDate, $lastLogonDateTime | Where-Object { $_ } | Sort-Object -Descending)[0]

        # Get the account status
        $isDisabled = !$adUser.Enabled
    } else {
        # If the AD user is not found, default values
        $mostRecentLogon = "N/A"
        $isDisabled = "N/A"
    }

    # Create a custom object to store the results and add it to the array
    $mailboxDetails += [PSCustomObject]@{
        Name               = $mailbox.Name
        Alias              = $mailbox.Alias
        LastLogonDate      = if ($adUser) { $adUser.LastLogonDate } else { "N/A" }
        LastLogon          = if ($adUser) { $lastLogonDateTime } else { "N/A" }
        MostRecentLogon    = $mostRecentLogon
        IsDisabled         = $isDisabled
    }
}

# Display the results in a table
$mailboxDetails | Format-Table -AutoSize
$mailboxDetails | Out-GridView -Title 'Mailbox Logon Dates'
$mailboxDetails | Out-HTMLView -Title 'Mailbox Logon Dates'

# Calculate and display totals in the console
$TotalMailboxes = $mailboxDetails.Count
$TotalEnabledUsers = ($mailboxDetails | Where-Object { $_.IsDisabled -eq $false }).Count
$TotalDisabledUsers = ($mailboxDetails | Where-Object { $_.IsDisabled -eq $true }).Count

Write-Host "Total Mailboxes Processed: $TotalMailboxes" -ForegroundColor Cyan
Write-Host "Total Enabled Users: $TotalEnabledUsers" -ForegroundColor Green
Write-Host "Total Disabled Users: $TotalDisabledUsers" -ForegroundColor Red
