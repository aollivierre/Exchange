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
            $null
        }

        # Determine the most recent of LastLogonDate and LastLogonDateTime
        $mostRecentLogon = @($adUser.LastLogonDate, $lastLogonDateTime) | Where-Object { $_ } | Sort-Object -Descending | Select-Object -First 1

        # Get the account status
        $isDisabled = -not $adUser.Enabled
    } else {
        # If the AD user is not found, default values
        $mostRecentLogon = "N/A"
        $isDisabled = "N/A"
    }

    # Get mailbox statistics
    $mailboxStats = Get-MailboxStatistics -Identity $mailbox.Identity

    # Get total item size
    $totalItemSize = $mailboxStats.TotalItemSize.ToString()

    # Check if archive is enabled
    $isArchiveEnabled = $mailbox.ArchiveEnabled

    if ($isArchiveEnabled) {
        try {
            # Get archive mailbox statistics
            $archiveStats = Get-MailboxStatistics -Identity $mailbox.Identity -Archive

            # Get archive total item size
            $archiveTotalItemSize = $archiveStats.TotalItemSize.ToString()
        }
        catch {
            $archiveTotalItemSize = 'Error retrieving archive size'
        }
    } else {
        $archiveTotalItemSize = 'N/A'
    }

    # Get mailbox type
    $mailboxType = $mailbox.RecipientTypeDetails

    # Get mailbox quotas
    $issueWarningQuota = $mailbox.IssueWarningQuota.ToString()
    $prohibitSendQuota = $mailbox.ProhibitSendQuota.ToString()
    $prohibitSendReceiveQuota = $mailbox.ProhibitSendReceiveQuota.ToString()

    # Get mailbox database
    $mailboxDatabase = $mailbox.Database.ToString()

    # Get email addresses
    $emailAddresses = ($mailbox.EmailAddresses | Where-Object { $_ -like 'SMTP:*' }) -join '; '

    # Create a custom object to store the results and add it to the array
    $mailboxDetails += [PSCustomObject]@{
        Name                    = $mailbox.Name
        Alias                   = $mailbox.Alias
        LastLogonDate           = if ($adUser) { $adUser.LastLogonDate } else { "N/A" }
        LastLogon               = if ($adUser) { $lastLogonDateTime } else { "N/A" }
        MostRecentLogon         = $mostRecentLogon
        IsDisabled              = $isDisabled
        MailboxSize             = $totalItemSize
        IsArchiveEnabled        = $isArchiveEnabled
        ArchiveSize             = $archiveTotalItemSize
        MailboxType             = $mailboxType
        IssueWarningQuota       = $issueWarningQuota
        ProhibitSendQuota       = $prohibitSendQuota
        ProhibitSendReceiveQuota= $prohibitSendReceiveQuota
        MailboxDatabase         = $mailboxDatabase
        EmailAddresses          = $emailAddresses
    }
}

# Display the results in a table
$mailboxDetails | Format-Table -AutoSize

# Optionally, output to GridView and HTML if desired
#$mailboxDetails | Out-GridView -Title 'Mailbox Details'
$mailboxDetails | Out-HTMLView -Title 'Mailbox Details'

# Calculate and display totals in the console
$TotalMailboxes = $mailboxDetails.Count
$TotalEnabledUsers = ($mailboxDetails | Where-Object { $_.IsDisabled -eq $false }).Count
$TotalDisabledUsers = ($mailboxDetails | Where-Object { $_.IsDisabled -eq $true }).Count

Write-Host "Total Mailboxes Processed: $TotalMailboxes" -ForegroundColor Cyan
Write-Host "Total Enabled Users: $TotalEnabledUsers" -ForegroundColor Green
Write-Host "Total Disabled Users: $TotalDisabledUsers" -ForegroundColor Red

# Export the results to a CSV file for further analysis
$mailboxDetails | Export-Csv -Path "MailboxDetails.csv" -NoTypeInformation
