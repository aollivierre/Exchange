# Section 1: Import Modules and Connect to Services

# Import the necessary modules
Import-Module ActiveDirectory

# Import Exchange Online module
Import-Module ExchangeOnlineManagement

# Import Microsoft Graph module for Azure AD interactions
Import-Module Microsoft.Graph.Users

# Connect to Exchange Online
Connect-ExchangeOnline

# Connect to Microsoft Graph with necessary permissions
Connect-MgGraph -Scopes "User.Read.All", "User.ReadWrite.All", "Directory.Read.All"

# Section 2: Get On-Premises Mailboxes and Users

# Get the on-premises mailboxes
$onPremMailboxes = Get-Mailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited

# Get all on-premises AD users, selecting necessary properties
$onPremUsers = Get-ADUser -Filter * -Properties SamAccountName, LastLogonDate, LastLogon, Enabled

# Section 3: Get Azure AD Users

# Get all Azure AD users
$azureADUsers = Get-MgUser -All -Property "DisplayName", "UserPrincipalName", "AccountEnabled", "SignInActivity"

# Convert to an array for easier processing
$azureADUsers = $azureADUsers | ForEach-Object { $_ }

# Section 4: Get Exchange Online Mailboxes

# Get all Exchange Online mailboxes
$exchangeOnlineMailboxes = Get-EXOMailbox -ResultSize Unlimited

# Section 5: Initialize List for Combined Mailbox Details

$combinedMailboxDetails = [System.Collections.Generic.List[PSObject]]::new()

# Section 6: Process Each On-Premises Mailbox

foreach ($mailbox in $onPremMailboxes) {
    # Find the corresponding on-premises AD user by SamAccountName (matching the Alias of the mailbox)
    $adUser = $onPremUsers | Where-Object { $_.SamAccountName -eq $mailbox.Alias }

    # Initialize variables
    $lastLogonDateTime = $null
    $mostRecentLogon = $null
    $isDisabled = $null

    if ($adUser) {
        # Convert LastLogon to DateTime format if it's not null or zero
        if ($adUser.LastLogon -and $adUser.LastLogon -ne 0) {
            $lastLogonDateTime = [DateTime]::FromFileTime($adUser.LastLogon)
        }

        # Determine the most recent of LastLogonDate and LastLogonDateTime
        $logonDates = @()
        if ($adUser.LastLogonDate) {
            $logonDates += $adUser.LastLogonDate
        }
        if ($lastLogonDateTime) {
            $logonDates += $lastLogonDateTime
        }

        if ($logonDates.Count -gt 0) {
            $mostRecentLogon = $logonDates | Sort-Object -Descending | Select-Object -First 1
        } else {
            $mostRecentLogon = "Never logged in"
        }

        # Get the account status
        $isDisabled = -not $adUser.Enabled
    } else {
        # If the AD user is not found, default values
        $mostRecentLogon = "N/A"
        $isDisabled = "N/A"
    }

    # Get on-premises mailbox statistics
    $mailboxStats = Get-MailboxStatistics -Identity $mailbox.Identity

    # Get total item size
    $totalItemSize = $mailboxStats.TotalItemSize.ToString()

    # Check if archive is enabled accurately
    $isArchiveEnabled = $mailbox.ArchiveGuid -ne [Guid]::Empty

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
    $issueWarningQuota        = $mailbox.IssueWarningQuota.ToString()
    $prohibitSendQuota        = $mailbox.ProhibitSendQuota.ToString()
    $prohibitSendReceiveQuota = $mailbox.ProhibitSendReceiveQuota.ToString()

    # Get mailbox database
    $mailboxDatabase = $mailbox.Database.ToString()

    # Get email addresses
    $emailAddresses = ($mailbox.EmailAddresses | Where-Object { $_ -like 'SMTP:*' }) -join '; '

    # Section 7: Check if User Exists in Azure AD

    # Use the primary SMTP address to find the user in Azure AD
    $userPrincipalName = $mailbox.PrimarySmtpAddress

    $azureADUser = $azureADUsers | Where-Object {
        $_.UserPrincipalName -eq $userPrincipalName -or
        $_.Mail -eq $userPrincipalName -or
        $_.OtherMails -contains $userPrincipalName
    }

    if ($azureADUser) {
        $existsInAzureAD = $true
        $azureAccountEnabled = $azureADUser.AccountEnabled
        $lastSignInDate = if ($azureADUser.SignInActivity) {
            $azureADUser.SignInActivity.LastSignInDateTime
        } else {
            "Never signed in"
        }
    } else {
        $existsInAzureAD = $false
        $azureAccountEnabled = $null
        $lastSignInDate = "User not found in Azure AD"
    }

    # Section 8: Check if Mailbox Exists in Exchange Online

    $exchangeOnlineMailbox = $exchangeOnlineMailboxes | Where-Object {
        $_.UserPrincipalName -eq $userPrincipalName -or
        $_.PrimarySmtpAddress -eq $userPrincipalName
    }

    if ($exchangeOnlineMailbox) {
        $existsInExchangeOnline = $true
        # Get Exchange Online mailbox size
        $exchangeOnlineMailboxStats = Get-EXOMailboxStatistics -Identity $exchangeOnlineMailbox.Identity
        $exchangeOnlineMailboxSize = $exchangeOnlineMailboxStats.TotalItemSize.ToString()
    } else {
        $existsInExchangeOnline = $false
        $exchangeOnlineMailboxSize = "N/A"
    }

    # Section 9: Create Combined Mailbox Detail Object

    $combinedMailboxDetail = [PSCustomObject]@{
        Name                         = $mailbox.Name
        Alias                        = $mailbox.Alias
        EmailAddresses               = $emailAddresses
        OnPrem_LastLogonDate         = if ($adUser.LastLogonDate) { $adUser.LastLogonDate } else { "Never logged in" }
        OnPrem_LastLogon             = if ($lastLogonDateTime) { $lastLogonDateTime } else { "Never logged in" }
        OnPrem_MostRecentLogon       = $mostRecentLogon
        OnPrem_IsDisabled            = $isDisabled
        OnPrem_MailboxSize           = $totalItemSize
        OnPrem_IsArchiveEnabled      = $isArchiveEnabled
        OnPrem_ArchiveSize           = $archiveTotalItemSize
        OnPrem_MailboxType           = $mailboxType
        OnPrem_IssueWarningQuota     = $issueWarningQuota
        OnPrem_ProhibitSendQuota     = $prohibitSendQuota
        OnPrem_ProhibitSendReceiveQuota = $prohibitSendReceiveQuota
        OnPrem_MailboxDatabase       = $mailboxDatabase
        AzureAD_Exists               = $existsInAzureAD
        AzureAD_AccountEnabled       = $azureAccountEnabled
        AzureAD_LastSignInDate       = $lastSignInDate
        ExchangeOnline_Exists        = $existsInExchangeOnline
        ExchangeOnline_MailboxSize   = $exchangeOnlineMailboxSize
    }

    # Add the combined detail to the collection
    $combinedMailboxDetails.Add($combinedMailboxDetail)
}

# Section 10: Display the Results

$combinedMailboxDetails | Format-Table -AutoSize

# Section 11: Optionally Output to GridView or HTML

# $combinedMailboxDetails | Out-GridView -Title 'Combined Mailbox Details'
$combinedMailboxDetails | Out-HTMLView -Title 'Combined Mailbox Details' # Requires Out-HTMLView module

# Section 12: Calculate and Display Totals

$TotalMailboxes                 = $combinedMailboxDetails.Count
$TotalEnabledOnPremUsers        = ($combinedMailboxDetails | Where-Object { $_.OnPrem_IsDisabled -eq $false }).Count
$TotalDisabledOnPremUsers       = ($combinedMailboxDetails | Where-Object { $_.OnPrem_IsDisabled -eq $true }).Count
$TotalUsersInAzureAD            = ($combinedMailboxDetails | Where-Object { $_.AzureAD_Exists -eq $true }).Count
$TotalUsersInExchangeOnline     = ($combinedMailboxDetails | Where-Object { $_.ExchangeOnline_Exists -eq $true }).Count

Write-Host "Total On-Premises Mailboxes Processed: $TotalMailboxes" -ForegroundColor Cyan
Write-Host "Total Enabled On-Premises Users: $TotalEnabledOnPremUsers" -ForegroundColor Green
Write-Host "Total Disabled On-Premises Users: $TotalDisabledOnPremUsers" -ForegroundColor Red
Write-Host "Total Users Existing in Azure AD: $TotalUsersInAzureAD" -ForegroundColor Cyan
Write-Host "Total Users Existing in Exchange Online: $TotalUsersInExchangeOnline" -ForegroundColor Cyan

# Section 13: Export the Results to a CSV File

$combinedMailboxDetails | Export-Csv -Path "CombinedMailboxDetails.csv" -NoTypeInformation
