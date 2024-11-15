# Section 1: Import Modules and Connect to Services

# Import the necessary modules
Import-Module ActiveDirectory

# Import the Azure AD module (Legacy)
Import-Module AzureAD

# Import the Microsoft Graph module for Entra
Import-Module Microsoft.Graph.Users

# Import the Exchange Online module with a prefix to avoid cmdlet conflicts
Import-Module ExchangeOnlineManagement

# Connect to Azure AD (Legacy)
Connect-AzureAD

# Connect to Microsoft Graph with necessary permissions for Entra data
Connect-MgGraph -Scopes "User.Read.All", "Directory.Read.All", "AuditLog.Read.All"

# Connect to Exchange Online with a cmdlet prefix 'EXOL' to avoid conflicts
Connect-ExchangeOnline -Prefix EXOL

# Section 2: Get On-Premises Mailboxes and Users

# Get the on-premises mailboxes (requires Exchange Management Tools)
$onPremMailboxes = Get-Mailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited

# Get all on-premises AD users, selecting necessary properties
$onPremUsers = Get-ADUser -Filter * -Properties SamAccountName, LastLogonDate, LastLogon, Enabled

# Section 3: Get Azure AD Users (Legacy)

# Get all Azure AD users using the AzureAD module
$azureADUsers = Get-AzureADUser -All $true

# Section 4: Get Entra Users via Microsoft Graph

# Get all Entra (Graph) users
$graphUsers = Get-MgUser -All -Property "DisplayName", "UserPrincipalName", "AccountEnabled", "SignInActivity"

# Convert to an array for easier processing
$graphUsers = $graphUsers | ForEach-Object { $_ }

# Section 5: Get Exchange Online Mailboxes

# Get all Exchange Online mailboxes using the cmdlets with the 'EXOL' prefix
$exchangeOnlineMailboxes = Get-EXOLMailbox -ResultSize Unlimited

# Section 6: Initialize List for Combined Mailbox Details

$combinedMailboxDetails = [System.Collections.Generic.List[PSObject]]::new()

# Section 7: Process Each On-Premises Mailbox

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

    # Section 8: Check if User Exists in Azure AD (Legacy)

    # Use the primary SMTP address to find the user in Azure AD
    $primaryEmail = $mailbox.PrimarySmtpAddress

    # Attempt to find the user by UserPrincipalName
    $azureADUser = Get-AzureADUser -Filter "UserPrincipalName eq '$primaryEmail'" -ErrorAction SilentlyContinue

    if (-not $azureADUser) {
        # Attempt to find the user by Mail property
        $azureADUser = Get-AzureADUser -Filter "Mail eq '$primaryEmail'" -ErrorAction SilentlyContinue
    }

    if ($azureADUser) {
        $existsInAzureAD = $true
        $azureAccountEnabled = $azureADUser.AccountEnabled
        # Note: Last sign-in date is not available via AzureAD module
        $azureAD_LastSignInDate = "Not available via AzureAD module"
    } else {
        $existsInAzureAD = $false
        $azureAccountEnabled = $null
        $azureAD_LastSignInDate = "User not found in Azure AD"
    }

    # Section 9: Check if User Exists in Entra (Graph)

    # Attempt to find the user in Entra via Microsoft Graph
    $graphUser = $graphUsers | Where-Object {
        $_.UserPrincipalName -eq $primaryEmail -or
        $_.Mail -eq $primaryEmail -or
        $_.OtherMails -contains $primaryEmail
    }

    if ($graphUser) {
        $existsInEntra = $true
        $entraAccountEnabled = $graphUser.AccountEnabled
        if ($graphUser.SignInActivity) {
            $entra_LastSignInDate = $graphUser.SignInActivity.LastSignInDateTime
        } else {
            $entra_LastSignInDate = "Never signed in"
        }
    } else {
        $existsInEntra = $false
        $entraAccountEnabled = $null
        $entra_LastSignInDate = "User not found in Entra"
    }

    # Section 10: Check if Mailbox Exists in Exchange Online

    $exchangeOnlineMailbox = $exchangeOnlineMailboxes | Where-Object {
        $_.UserPrincipalName -eq $primaryEmail -or
        $_.PrimarySmtpAddress -eq $primaryEmail
    }

    if ($exchangeOnlineMailbox) {
        $existsInExchangeOnline = $true
        # Get Exchange Online mailbox size using the cmdlets with the 'EXOL' prefix
        $exchangeOnlineMailboxStats = Get-EXOLMailboxStatistics -Identity $exchangeOnlineMailbox.Identity
        $exchangeOnlineMailboxSize = $exchangeOnlineMailboxStats.TotalItemSize.ToString()
    } else {
        $existsInExchangeOnline = $false
        $exchangeOnlineMailboxSize = "N/A"
    }

    # Section 11: Create Combined Mailbox Detail Object

    $combinedMailboxDetail = [PSCustomObject]@{
        Name                             = $mailbox.Name
        Alias                            = $mailbox.Alias
        EmailAddresses                   = $emailAddresses
        OnPrem_LastLogonDate             = if ($adUser.LastLogonDate) { $adUser.LastLogonDate } else { "Never logged in" }
        OnPrem_LastLogon                 = if ($lastLogonDateTime) { $lastLogonDateTime } else { "Never logged in" }
        OnPrem_MostRecentLogon           = $mostRecentLogon
        OnPrem_IsDisabled                = $isDisabled
        OnPrem_MailboxSize               = $totalItemSize
        OnPrem_IsArchiveEnabled          = $isArchiveEnabled
        OnPrem_ArchiveSize               = $archiveTotalItemSize
        OnPrem_MailboxType               = $mailboxType
        OnPrem_IssueWarningQuota         = $issueWarningQuota
        OnPrem_ProhibitSendQuota         = $prohibitSendQuota
        OnPrem_ProhibitSendReceiveQuota  = $prohibitSendReceiveQuota
        OnPrem_MailboxDatabase           = $mailboxDatabase
        'Azure AD (Legacy)_Exists'       = $existsInAzureAD
        'Azure AD (Legacy)_AccountEnabled' = $azureAccountEnabled
        'Azure AD (Legacy)_LastSignInDate' = $azureAD_LastSignInDate
        'Entra (Graph)_Exists'           = $existsInEntra
        'Entra (Graph)_AccountEnabled'   = $entraAccountEnabled
        'Entra (Graph)_LastSignInDate'   = $entra_LastSignInDate
        ExchangeOnline_Exists            = $existsInExchangeOnline
        ExchangeOnline_MailboxSize       = $exchangeOnlineMailboxSize
    }

    # Add the combined detail to the collection
    $combinedMailboxDetails.Add($combinedMailboxDetail)
}

# Section 12: Display the Results

$combinedMailboxDetails | Format-Table -AutoSize

# Section 13: Optionally Output to GridView or HTML

# $combinedMailboxDetails | Out-GridView -Title 'Combined Mailbox Details'
$combinedMailboxDetails | Out-HTMLView -Title 'Combined Mailbox Details' # Requires Out-HTMLView module

# Section 14: Calculate and Display Totals

$TotalMailboxes                  = $combinedMailboxDetails.Count
$TotalEnabledOnPremUsers         = ($combinedMailboxDetails | Where-Object { $_.OnPrem_IsDisabled -eq $false }).Count
$TotalDisabledOnPremUsers        = ($combinedMailboxDetails | Where-Object { $_.OnPrem_IsDisabled -eq $true }).Count
$TotalUsersInAzureAD             = ($combinedMailboxDetails | Where-Object { $_.'Azure AD (Legacy)_Exists' -eq $true }).Count
$TotalUsersInEntra               = ($combinedMailboxDetails | Where-Object { $_.'Entra (Graph)_Exists' -eq $true }).Count
$TotalUsersInExchangeOnline      = ($combinedMailboxDetails | Where-Object { $_.ExchangeOnline_Exists -eq $true }).Count

Write-Host "Total On-Premises Mailboxes Processed: $TotalMailboxes" -ForegroundColor Cyan
Write-Host "Total Enabled On-Premises Users: $TotalEnabledOnPremUsers" -ForegroundColor Green
Write-Host "Total Disabled On-Premises Users: $TotalDisabledOnPremUsers" -ForegroundColor Red
Write-Host "Total Users Existing in Azure AD (Legacy): $TotalUsersInAzureAD" -ForegroundColor Cyan
Write-Host "Total Users Existing in Entra (Graph): $TotalUsersInEntra" -ForegroundColor Cyan
Write-Host "Total Users Existing in Exchange Online: $TotalUsersInExchangeOnline" -ForegroundColor Cyan

# Section 15: Export the Results to a CSV File

$combinedMailboxDetails | Export-Csv -Path "CombinedMailboxDetails.csv" -NoTypeInformation
