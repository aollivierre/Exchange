# Section 1: Import Modules and Connect to Services

# Import the necessary modules
Import-Module ExchangeOnlineManagement
Import-Module Microsoft.Graph.Users

# Connect to Microsoft Graph with necessary permissions
Connect-MgGraph -Scopes "User.Read.All","AuditLog.Read.All"

# Connect to Exchange Online
Connect-ExchangeOnline

# Section 2: Get Mailboxes from Exchange Online

# Get all mailboxes from Exchange Online
$mailboxes = Get-EXOMailbox

# Section 3: Get Users from Entra (Azure AD) with Sign-In Activity

# Get all users with SignInActivity property
$users = Get-MgUser -All -Property "DisplayName,UserPrincipalName,SignInActivity,AccountEnabled"

# Convert to an array for easier processing
$users = $users.Value

# Section 4: Initialize List for Mailbox Details

$mailboxDetails = [System.Collections.Generic.List[PSObject]]::new()

# Section 5: Iterate Through Each Mailbox and Retrieve Information

foreach ($mailbox in $mailboxes) {
    # Find the corresponding Entra user by UserPrincipalName
    $azureADUser = $users | Where-Object { $_.UserPrincipalName -eq $mailbox.UserPrincipalName }

    # Initialize variables
    $lastSignInDate = $null
    $accountEnabled = $null

    if ($azureADUser) {
        # Get account status
        $accountEnabled = $azureADUser.AccountEnabled

        # Get last sign-in date
        if ($azureADUser.SignInActivity) {
            $lastSignInDate = $azureADUser.SignInActivity.LastSignInDateTime
        } else {
            $lastSignInDate = "Never signed in"
        }
    } else {
        # If the Entra user is not found, default values
        $lastSignInDate = "User not found in Entra"
        $accountEnabled = "N/A"
    }

    # Get mailbox statistics
    $mailboxStats = Get-EXOMailboxStatistics -Identity $mailbox.Identity

    # Get total item size
    $totalItemSize = $mailboxStats.TotalItemSize.ToString()

    # Check if archive is enabled
    $isArchiveEnabled = $mailbox.ArchiveStatus -eq 'Active'

    if ($isArchiveEnabled) {
        try {
            # Get archive mailbox statistics
            $archiveStats = Get-EXOMailboxStatistics -Identity $mailbox.Identity -Archive

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
    $issueWarningQuota        = $mailbox.IssueWarningQuota.ToString()
    $prohibitSendQuota        = $mailbox.ProhibitSendQuota.ToString()
    $prohibitSendReceiveQuota = $mailbox.ProhibitSendReceiveQuota.ToString()

    # Get email addresses
    $emailAddresses = ($mailbox.EmailAddresses | Where-Object { $_ -like 'SMTP:*' }) -join '; '

    # Create a custom object to store the results
    $mailboxDetail = [PSCustomObject]@{
        Name                     = $mailbox.DisplayName
        UserPrincipalName        = $mailbox.UserPrincipalName
        LastSignInDate           = $lastSignInDate
        AccountEnabled           = $accountEnabled
        MailboxSize              = $totalItemSize
        IsArchiveEnabled         = $isArchiveEnabled
        ArchiveSize              = $archiveTotalItemSize
        MailboxType              = $mailboxType
        IssueWarningQuota        = $issueWarningQuota
        ProhibitSendQuota        = $prohibitSendQuota
        ProhibitSendReceiveQuota = $prohibitSendReceiveQuota
        EmailAddresses           = $emailAddresses
    }

    # Add the mailbox detail to the collection
    $mailboxDetails.Add($mailboxDetail)
}

# Section 6: Display the Results

$mailboxDetails | Format-Table -AutoSize

# Section 7: Optionally Output to GridView or HTML

# $mailboxDetails | Out-GridView -Title 'Mailbox Details'
$mailboxDetails | Out-HTMLView -Title 'Mailbox Details' # Requires Out-HTMLView module

# Section 8: Calculate and Display Totals

$TotalMailboxes     = $mailboxDetails.Count
$TotalEnabledUsers  = ($mailboxDetails | Where-Object { $_.AccountEnabled -eq $true }).Count
$TotalDisabledUsers = ($mailboxDetails | Where-Object { $_.AccountEnabled -eq $false }).Count

Write-Host "Total Mailboxes Processed: $TotalMailboxes" -ForegroundColor Cyan
Write-Host "Total Enabled Users: $TotalEnabledUsers" -ForegroundColor Green
Write-Host "Total Disabled Users: $TotalDisabledUsers" -ForegroundColor Red

# Section 9: Export the Results to a CSV File

$mailboxDetails | Export-Csv -Path "MailboxDetails.csv" -NoTypeInformation
