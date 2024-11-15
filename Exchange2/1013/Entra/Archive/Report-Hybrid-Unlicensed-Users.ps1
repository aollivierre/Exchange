# Define required Graph modules
$requiredGraphModules = @(
    "Microsoft.Graph.Authentication",
    "Microsoft.Graph.Users",
    "Microsoft.Graph.Users.Actions",
    "Microsoft.Graph.Identity.DirectoryManagement",
    "Microsoft.Graph.Sites", # For OneDrive information
    "ExchangeOnlineManagement"  # For Exchange Online information
)

# Check and install required modules in AllUsers scope
Write-Host "Checking and installing required modules..."
foreach ($module in $requiredGraphModules) {
    if (-not (Get-Module -ListAvailable -Name $module)) {
        try {
            Write-Host "Installing $module module in AllUsers scope..."
            Install-Module -Name $module -Scope AllUsers -Force -AllowClobber
        }
        catch {
            Write-Error "Failed to install $module`: $_"
            exit
        }
    }
}

# Check and install AzureAD module
if (-not (Get-Module -ListAvailable -Name AzureAD)) {
    try {
        Write-Host "Installing AzureAD module in AllUsers scope..."
        Install-Module -Name AzureAD -Scope AllUsers -Force -AllowClobber
    }
    catch {
        Write-Error "Failed to install AzureAD module: $_"
        exit
    }
}

# # Import modules
# Write-Host "Importing modules..."
# Import-Module -Name AzureAD
# foreach ($module in $requiredGraphModules) {
#     Import-Module -Name $module
# }

# Function for friendly license names
function Get-FriendlyLicenseName {
    param (
        [string[]]$SkuPartNumbers
    )
    
    $licenseMap = @{
        'ENTERPRISEPREMIUM'                  = 'Office 365 E5'
        'ENTERPRISEPACK'                     = 'Office 365 E3'
        'FLOW_FREE'                          = 'Power Automate Free'
        'MCOPSTNC'                           = 'Microsoft Teams Domestic Calling'
        'Microsoft_Teams_Exploratory_Dept'   = 'Teams Exploratory'
        'POWER_BI_STANDARD'                  = 'Power BI Free'
        'POWER_BI_PRO'                       = 'Power BI Pro'
        'EXCHANGESTANDARD'                   = 'Exchange Online Plan 1'
        'EXCHANGEENTERPRISE'                 = 'Exchange Online Plan 2'
        'MCOSTANDARD'                        = 'Teams Plan 2'
        'Microsoft_365_E5_Developer'         = 'Microsoft 365 E5 Developer'
        'PROJECTPREMIUM'                     = 'Project Plan 5'
        'PROJECTPROFESSIONAL'                = 'Project Plan 3'
        'VISIOONLINE_PLAN1'                  = 'Visio Plan 1'
        'VISIOPRO'                           = 'Visio Plan 2'
        'SPE_E5'                             = 'Microsoft 365 E5'
        'SPE_E3'                             = 'Microsoft 365 E3'
        'M365_F1'                            = 'Microsoft 365 F1'
        'TEAMS_FREE'                         = 'Microsoft Teams Free'
        'MCOEV'                              = 'Microsoft Teams Phone System'
        'PHONESYSTEM_VIRTUALUSER'            = 'Microsoft Teams Phone Resource Account'
        'Microsoft_Teams_Audio_Conferencing' = 'Microsoft Teams Audio Conferencing'
    }

    $friendlyNames = foreach ($sku in $SkuPartNumbers) {
        $sku = $sku.Trim()
        if ($licenseMap.ContainsKey($sku)) {
            $licenseMap[$sku]
        }
        else {
            $sku  # Return original if no mapping found
        }
    }

    return ($friendlyNames -join '; ')
}




# Function for office location extraction
function Get-OfficeLocation {
    param (
        [string]$OnPremDomain
    )
    
    if (-not $OnPremDomain) { return '' }
    
    try {
        $firstPart = $OnPremDomain.Split('.')[0]
        return $firstPart.ToUpper()
    }
    catch {
        Write-Warning "Error extracting office location from domain $OnPremDomain"
        return ''
    }
}

# Safe path creation function
function Get-SafePath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$BasePath,
        
        [Parameter(Mandatory = $true)]
        [string]$FileName
    )
    
    try {
        if (-not (Test-Path -Path $BasePath)) {
            New-Item -ItemType Directory -Path $BasePath -Force | Out-Null
        }
        
        $fullPath = Join-Path -Path $BasePath -ChildPath $FileName
        return $fullPath
    }
    catch {
        Write-Error "Error creating path: $_"
        return $null
    }
}


# Function to get mailbox information
function Get-UserMailboxInfo {
    param (
        [Parameter(Mandatory = $true)]
        [string]$UserPrincipalName,
        
        [Parameter(Mandatory = $false)]
        [string]$ImmutableId
    )
    
    try {
        # Get the mailbox directly first
        $mailbox = Get-EXOMailbox -Identity $UserPrincipalName -ErrorAction Stop
        
        if ($mailbox) {
            $stats = Get-EXOMailboxStatistics -Identity $mailbox.PrimarySmtpAddress -ErrorAction Stop
            $archiveStats = if ($mailbox.ArchiveStatus -eq 'Active') {
                Get-EXOMailboxStatistics -Identity $mailbox.PrimarySmtpAddress -Archive -ErrorAction SilentlyContinue
            }
            else { $null }

            # Don't try to match ImmutableId if it's not provided (cloud-only users)
            $isMatch = if ($ImmutableId) {
                # Only try to convert if ExchangeGuid exists
                if ($mailbox.ExchangeGuid) {
                    $mailboxImmutableId = [Convert]::ToBase64String([Guid]::New($mailbox.ExchangeGuid.ToString()).ToByteArray())
                    $mailboxImmutableId -eq $ImmutableId
                }
                else {
                    $false
                }
            }
            else {
                # If no ImmutableId provided, consider it a match if primary SMTP matches
                $mailbox.PrimarySmtpAddress -eq $UserPrincipalName
            }

            if ($isMatch -or $mailbox.PrimarySmtpAddress -eq $UserPrincipalName) {
                return [PSCustomObject]@{
                    HasMailbox         = $true
                    MailboxType        = $mailbox.RecipientTypeDetails
                    PrimarySmtpAddress = $mailbox.PrimarySmtpAddress
                    MailboxSize        = if ($stats) { $stats.TotalItemSize } else { "N/A" }
                    ItemCount          = if ($stats) { $stats.ItemCount } else { 0 }
                    ArchiveStatus      = $mailbox.ArchiveStatus
                    ArchiveSize        = if ($archiveStats) { $archiveStats.TotalItemSize } else { "N/A" }
                    ArchiveItemCount   = if ($archiveStats) { $archiveStats.ItemCount } else { 0 }
                    DatabaseName       = $mailbox.Database
                    RetentionPolicy    = $mailbox.RetentionPolicy
                    ExchangeGuid       = $mailbox.ExchangeGuid
                    ImmutableId        = if ($mailbox.ExchangeGuid) { 
                        [Convert]::ToBase64String([Guid]::New($mailbox.ExchangeGuid.ToString()).ToByteArray()) 
                    }
                    else { $null }
                }
            }
        }
    }
    catch {
        Write-Verbose "No direct mailbox found for $UserPrincipalName. Error: $_"
        # Try searching by proxy addresses if direct lookup fails
        try {
            $mailbox = Get-EXORecipient -Filter "EmailAddresses -eq '$UserPrincipalName'" -ErrorAction Stop
            if ($mailbox) {
                $stats = Get-EXOMailboxStatistics -Identity $mailbox.PrimarySmtpAddress -ErrorAction SilentlyContinue
                return [PSCustomObject]@{
                    HasMailbox         = $true
                    MailboxType        = $mailbox.RecipientTypeDetails
                    PrimarySmtpAddress = $mailbox.PrimarySmtpAddress
                    MailboxSize        = if ($stats) { $stats.TotalItemSize } else { "N/A" }
                    ItemCount          = if ($stats) { $stats.ItemCount } else { 0 }
                    ArchiveStatus      = "N/A"
                    ArchiveSize        = "N/A"
                    ArchiveItemCount   = 0
                    DatabaseName       = "N/A"
                    RetentionPolicy    = "N/A"
                    ExchangeGuid       = $null
                    ImmutableId        = $null
                }
            }
        }
        catch {
            Write-Verbose "No mailbox found by proxy address search for $UserPrincipalName. Error: $_"
        }
    }

    # Return default object if no mailbox found
    return [PSCustomObject]@{
        HasMailbox         = $false
        MailboxType        = "None"
        PrimarySmtpAddress = "N/A"
        MailboxSize        = "N/A"
        ItemCount          = 0
        ArchiveStatus      = "None"
        ArchiveSize        = "N/A"
        ArchiveItemCount   = 0
        DatabaseName       = "N/A"
        RetentionPolicy    = "N/A"
        ExchangeGuid       = $null
        ImmutableId        = $null
    }
}




# Function to get OneDrive information
function Get-UserOneDriveInfo {
    param (
        [string]$UserPrincipalName
    )
    
    try {
        # First try to get the user's drive
        $drives = Get-MgUser -UserId $UserPrincipalName -Property Drive -ErrorAction Stop | 
        Select-Object -ExpandProperty Drive
        
        if ($drives) {
            return [PSCustomObject]@{
                HasOneDrive       = $true
                OneDriveUrl       = $drives.WebUrl
                OneDriveQuota     = $drives.Quota.Total
                OneDriveUsed      = $drives.Quota.Used
                OneDriveRemaining = $drives.Quota.Remaining
                OneDriveState     = $drives.Quota.State
            }
        }
        else {
            return [PSCustomObject]@{
                HasOneDrive       = $false
                OneDriveUrl       = "N/A"
                OneDriveQuota     = 0
                OneDriveUsed      = 0
                OneDriveRemaining = 0
                OneDriveState     = "None"
            }
        }
    }
    catch {
        Write-Warning "Error getting OneDrive info for $UserPrincipalName : $_"
        return [PSCustomObject]@{
            HasOneDrive       = $false
            OneDriveUrl       = "N/A"
            OneDriveQuota     = 0
            OneDriveUsed      = 0
            OneDriveRemaining = 0
            OneDriveState     = "None"
        }
    }
}






# Create and verify export directory
$exportPath = "C:\code\exports\NTI"
try {
    if (-not (Test-Path -Path $exportPath)) {
        Write-Host "Creating export directory: $exportPath"
        New-Item -ItemType Directory -Path $exportPath -Force | Out-Null
    }

    if (-not (Test-Path -Path $exportPath)) {
        throw "Failed to create or access export directory"
    }
}
catch {
    Write-Error "Error with export path: $_"
    exit
}

# Verify $exportPath is set and accessible
if (-not $exportPath) {
    Write-Error "Export path is null or empty"
    exit
}

Write-Host "Export path verified: $exportPath"

# Get current timestamp for file naming
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"

# Connect to services
try {
    Write-Host "Connecting to Microsoft Graph..."
    Connect-MgGraph -Scopes "User.Read.All", "Directory.Read.All", "Sites.Read.All"
    
    Write-Host "Connecting to Azure AD..."
    Connect-AzureAD

    Write-Host "Connecting to Exchange Online..."
    Connect-ExchangeOnline -ShowBanner:$false
}
catch {
    Write-Error "Failed to connect to one or more services: $_"
    exit
}

# Initialize Generic Lists for better performance
$graphResults = [System.Collections.Generic.List[PSCustomObject]]::new()
$azureADResults = [System.Collections.Generic.List[PSCustomObject]]::new()




# Get users from Microsoft Graph (Modern/Entra)
Write-Host "Gathering data from Microsoft Graph..."
try {
    $graphUsers = Get-MgUser -All -Property @(
        'DisplayName',
        'UserPrincipalName',
        'OnPremisesDistinguishedName',
        'OnPremisesDomainName',
        'UserType',
        'AccountEnabled',
        'Id',
        'CreatedDateTime',
        'SignInSessionsValidFromDateTime',
        'OnPremisesSyncEnabled',
        'OnPremisesImmutableId'
    ) | Where-Object { 
        $_.UserType -eq "Member" -and 
        -not ($_.UserPrincipalName -like "*#EXT#*")
    }

    Write-Host "Found $($graphUsers.Count) users in Microsoft Graph"

    foreach ($user in $graphUsers) {
        try {
            $licenses = Get-MgUserLicenseDetail -UserId $user.Id
            $licenseNames = Get-FriendlyLicenseName -SkuPartNumbers $licenses.SkuPartNumber
            $isLicensed = $licenses.Count -gt 0
            $officeLocation = Get-OfficeLocation -OnPremDomain $user.OnPremisesDomainName
            
            # Get Exchange and OneDrive information
            # $mailboxInfo = Get-UserMailboxInfo -UserPrincipalName $user.UserPrincipalName


            # In the main foreach loop where we process users, update to:
            $mailboxInfo = Get-UserMailboxInfo -UserPrincipalName $user.UserPrincipalName -ImmutableId $user.OnPremisesImmutableId
            if ($null -eq $mailboxInfo) {
                Write-Warning "Failed to get mailbox info for $($user.UserPrincipalName)"
                $mailboxInfo = [PSCustomObject]@{
                    HasMailbox         = $false
                    MailboxType        = "None"
                    PrimarySmtpAddress = "N/A"
                    MailboxSize        = "N/A"
                    ItemCount          = 0
                    ArchiveStatus      = "None"
                    ArchiveSize        = "N/A"
                    ArchiveItemCount   = 0
                    DatabaseName       = "N/A"
                    RetentionPolicy    = "N/A"
                    ExchangeGuid       = $null
                    ImmutableId        = $null
                }
            }


            $oneDriveInfo = Get-UserOneDriveInfo -UserPrincipalName $user.UserPrincipalName

            $graphResults.Add([PSCustomObject]@{
                    Source             = "Modern (Graph)"
                    DisplayName        = $user.DisplayName
                    UserPrincipalName  = $user.UserPrincipalName
                    OnPremisesDN       = $user.OnPremisesDistinguishedName
                    OnPremDomain       = $user.OnPremisesDomainName
                    OfficeLocation     = $officeLocation
                    ImmutableId        = $user.OnPremisesImmutableId
                    UserType           = $user.UserType
                    AccountEnabled     = $user.AccountEnabled
                    AssignedLicenses   = $licenseNames
                    IsLicensed         = $isLicensed
                    IsDirSynced        = $user.OnPremisesSyncEnabled
                    ObjectId           = $user.Id
                    CreatedDateTime    = $user.CreatedDateTime
                    LastSignInDateTime = $user.SignInSessionsValidFromDateTime

                    # Exchange Online properties
                    HasMailbox         = $mailboxInfo.HasMailbox
                    MailboxType        = $mailboxInfo.MailboxType
                    MailboxSize        = $mailboxInfo.MailboxSize
                    MailboxItemCount   = $mailboxInfo.ItemCount
                    ArchiveStatus      = $mailboxInfo.ArchiveStatus
                    ArchiveSize        = $mailboxInfo.ArchiveSize
                    ArchiveItemCount   = $mailboxInfo.ArchiveItemCount
                    MailboxDatabase    = $mailboxInfo.DatabaseName
                    RetentionPolicy    = $mailboxInfo.RetentionPolicy

                    # OneDrive properties
                    HasOneDrive        = $oneDriveInfo.HasOneDrive
                    OneDriveUrl        = $oneDriveInfo.OneDriveUrl
                    OneDriveQuota      = $oneDriveInfo.OneDriveQuota
                    OneDriveUsed       = $oneDriveInfo.OneDriveUsed
                    OneDriveRemaining  = $oneDriveInfo.OneDriveRemaining
                    OneDriveState      = $oneDriveInfo.OneDriveState
                })
        }
        catch {
            Write-Warning "Error processing user $($user.UserPrincipalName): $_"
        }
    }
}
catch {
    Write-Error "Error querying Microsoft Graph: $_"
}

# Get users from Azure AD (Legacy)
Write-Host "Gathering data from Azure AD..."
try {
    $azureADUsers = Get-AzureADUser -All $true | Where-Object { 
        $_.UserType -eq "Member" -and 
        -not ($_.UserPrincipalName -like "*#EXT#*")
    }

    Write-Host "Found $($azureADUsers.Count) users in Azure AD"

    $allSkus = Get-AzureADSubscribedSku

    foreach ($user in $azureADUsers) {
        $licenses = $user.AssignedLicenses | ForEach-Object {
            $skuId = $_.SkuId
            ($allSkus | Where-Object { $_.SkuId -eq $skuId }).SkuPartNumber
        }
        $licenseNames = Get-FriendlyLicenseName -SkuPartNumbers $licenses
        $isLicensed = $user.AssignedLicenses.Count -gt 0
        $officeLocation = Get-OfficeLocation -OnPremDomain $user.OnPremisesDomainName

        $azureADResults.Add([PSCustomObject]@{
                Source             = "Legacy (Azure AD)"
                DisplayName        = $user.DisplayName
                UserPrincipalName  = $user.UserPrincipalName
                OnPremisesDN       = $user.OnPremisesDistinguishedName
                OnPremDomain       = $user.OnPremisesDomainName
                OfficeLocation     = $officeLocation
                ImmutableId        = $user.ImmutableId
                UserType           = $user.UserType
                AccountEnabled     = $user.AccountEnabled
                AssignedLicenses   = $licenseNames
                IsLicensed         = $isLicensed
                IsDirSynced        = $user.DirSyncEnabled
                ObjectId           = $user.ObjectId
                CreatedDateTime    = $user.ExtensionProperty.createdDateTime
                LastSignInDateTime = $user.RefreshTokensValidFromDateTime
            })
    }
}
catch {
    Write-Error "Error querying Azure AD: $_"
}




# Merge results for side-by-side comparison
$mergedResults = [System.Collections.Generic.List[PSCustomObject]]::new()

foreach ($graphUser in $graphResults) {
    $azureADUser = $azureADResults | Where-Object { $_.UserPrincipalName -eq $graphUser.UserPrincipalName } | Select-Object -First 1
    
    if ($azureADUser) {
        $mergedResults.Add([PSCustomObject]@{
                UserPrincipalName       = $graphUser.UserPrincipalName
                DisplayName             = $graphUser.DisplayName
                OfficeLocation          = $graphUser.OfficeLocation
                ImmutableId             = $graphUser.ImmutableId
            
                # Graph (Modern) data
                Graph_AccountEnabled    = $graphUser.AccountEnabled
                Graph_LicenseStatus     = $graphUser.IsLicensed
                Graph_Licenses          = $graphUser.AssignedLicenses
                Graph_DirSyncEnabled    = $graphUser.IsDirSynced
                Graph_ObjectId          = $graphUser.ObjectId
                Graph_ImmutableId       = $graphUser.ImmutableId
                Graph_CreatedDateTime   = $graphUser.CreatedDateTime
                Graph_LastSignIn        = $graphUser.LastSignInDateTime
            
                # Azure AD (Legacy) data
                AzureAD_AccountEnabled  = $azureADUser.AccountEnabled
                AzureAD_LicenseStatus   = $azureADUser.IsLicensed
                AzureAD_Licenses        = $azureADUser.AssignedLicenses
                AzureAD_DirSyncEnabled  = $azureADUser.IsDirSynced
                AzureAD_ObjectId        = $azureADUser.ObjectId
                AzureAD_ImmutableId     = $azureADUser.ImmutableId
                AzureAD_CreatedDateTime = $azureADUser.CreatedDateTime
                AzureAD_LastSignIn      = $azureADUser.LastSignInDateTime
            
                # Common properties
                OnPremisesDN            = $graphUser.OnPremisesDN
                OnPremDomain            = $graphUser.OnPremDomain
            
                # Exchange Online properties
                HasMailbox              = $graphUser.HasMailbox
                MailboxType             = $graphUser.MailboxType
                MailboxSize             = $graphUser.MailboxSize
                MailboxItemCount        = $graphUser.MailboxItemCount
                ArchiveStatus           = $graphUser.ArchiveStatus
                ArchiveSize             = $graphUser.ArchiveSize
                ArchiveItemCount        = $graphUser.ArchiveItemCount
                MailboxDatabase         = $graphUser.MailboxDatabase
                RetentionPolicy         = $graphUser.RetentionPolicy

                # OneDrive properties
                HasOneDrive             = $graphUser.HasOneDrive
                OneDriveUrl             = $graphUser.OneDriveUrl
                OneDriveQuota           = $graphUser.OneDriveQuota
                OneDriveUsed            = $graphUser.OneDriveUsed
                OneDriveRemaining       = $graphUser.OneDriveRemaining
                OneDriveState           = $graphUser.OneDriveState
            
                # Comparison flags
                LicenseMismatch         = $graphUser.IsLicensed -ne $azureADUser.IsLicensed
                StatusMismatch          = $graphUser.AccountEnabled -ne $azureADUser.AccountEnabled
                DirSyncMismatch         = $graphUser.IsDirSynced -ne $azureADUser.IsDirSynced
                ImmutableIdMismatch     = $graphUser.ImmutableId -ne $azureADUser.ImmutableId
            })
    }
}

# Calculate all statistics
$mismatchStats = [PSCustomObject]@{
    TotalUsers            = $mergedResults.Count
    LicenseMismatches     = ($mergedResults | Where-Object { $_.LicenseMismatch -eq $true }).Count
    StatusMismatches      = ($mergedResults | Where-Object { $_.StatusMismatch -eq $true }).Count
    DirSyncMismatches     = ($mergedResults | Where-Object { $_.DirSyncMismatch -eq $true }).Count
    ImmutableIdMismatches = ($mergedResults | Where-Object { $_.ImmutableIdMismatch -eq $true }).Count
}

$cloudServiceStats = [PSCustomObject]@{
    UsersWithMailbox  = ($mergedResults | Where-Object { $_.HasMailbox -eq $true }).Count
    SharedMailboxes   = ($mergedResults | Where-Object { $_.MailboxType -eq 'SharedMailbox' }).Count
    UserMailboxes     = ($mergedResults | Where-Object { $_.MailboxType -eq 'UserMailbox' }).Count
    UsersWithArchive  = ($mergedResults | Where-Object { $_.ArchiveStatus -eq 'Active' }).Count
    UsersWithOneDrive = ($mergedResults | Where-Object { $_.HasOneDrive -eq $true }).Count
}

# Create comprehensive statistics report
$statsReport = @"
==========================================
User Account Comparison Report
Generated on: $(Get-Date)
==========================================

Comparison Statistics:
---------------------
Total Users: $($mismatchStats.TotalUsers)
License Mismatches: $($mismatchStats.LicenseMismatches)
Status Mismatches: $($mismatchStats.StatusMismatches)
DirSync Mismatches: $($mismatchStats.DirSyncMismatches)
ImmutableId Mismatches: $($mismatchStats.ImmutableIdMismatches)

Cloud Services Statistics:
------------------------
Users with Mailbox: $($cloudServiceStats.UsersWithMailbox)
- User Mailboxes: $($cloudServiceStats.UserMailboxes)
- Shared Mailboxes: $($cloudServiceStats.SharedMailboxes)
Users with Archive: $($cloudServiceStats.UsersWithArchive)
Users with OneDrive: $($cloudServiceStats.UsersWithOneDrive)
"@

# Display statistics in console
Write-Host $statsReport



# Create and verify file paths
$statsPath = Get-SafePath -BasePath $exportPath -FileName "UserStats_$timestamp.txt"
if (-not $statsPath) {
    Write-Error "Failed to create stats file path"
    exit
}

$csvPath = Get-SafePath -BasePath $exportPath -FileName "Combined_UserReport_$timestamp.csv"
if (-not $csvPath) {
    Write-Error "Failed to create CSV file path"
    exit
}

$htmlPath = Get-SafePath -BasePath $exportPath -FileName "Combined_UserReport_$timestamp.html"
if (-not $htmlPath) {
    Write-Error "Failed to create HTML file path"
    exit
}

# Export statistics to text file
$statsReport | Out-File -FilePath $statsPath

# Export to CSV
$mergedResults | Export-Csv -Path $csvPath -NoTypeInformation

# Create HTML report using Out-HTMLView
Write-Host "Generating HTML report..."
$mergedResults | Sort-Object UserPrincipalName | 
Select-Object @{Name = 'User Principal Name'; Expression = { $_.UserPrincipalName } },
@{Name = 'Display Name'; Expression = { $_.DisplayName } },
@{Name = 'Office Location'; Expression = { $_.OfficeLocation } },
@{Name = 'Has Mailbox'; Expression = { $_.HasMailbox } },
@{Name = 'Mailbox Type'; Expression = { $_.MailboxType } },
@{Name = 'Mailbox Size'; Expression = { $_.MailboxSize } },
@{Name = 'Archive Status'; Expression = { $_.ArchiveStatus } },
@{Name = 'Archive Size'; Expression = { $_.ArchiveSize } },
@{Name = 'Has OneDrive'; Expression = { $_.HasOneDrive } },
@{Name = 'OneDrive Size'; Expression = { if ($_.HasOneDrive) { [math]::Round($_.OneDriveUsed / 1GB, 2).ToString() + " GB" } else { "N/A" } } },
@{Name = 'Graph Licensed'; Expression = { $_.Graph_LicenseStatus } },
@{Name = 'Graph Licenses'; Expression = { $_.Graph_Licenses } },
@{Name = 'Account Enabled'; Expression = { $_.Graph_AccountEnabled } },
@{Name = 'DirSync Enabled'; Expression = { $_.Graph_DirSyncEnabled } },
@{Name = 'OnPrem DN'; Expression = { $_.OnPremisesDN } },
@{Name = 'OnPrem Domain'; Expression = { $_.OnPremDomain } },
@{Name = 'ImmutableId'; Expression = { $_.ImmutableId } },
@{Name = 'Status Mismatch'; Expression = { $_.StatusMismatch } },
@{Name = 'License Mismatch'; Expression = { $_.LicenseMismatch } },
@{Name = 'DirSync Mismatch'; Expression = { $_.DirSyncMismatch } } |
Out-HTMLView -Title "Entra ID User Report with Cloud Services" -FilePath $htmlPath

# Display completion message
Write-Host "`nReport generation completed!"
Write-Host "CSV report saved to: $csvPath"
Write-Host "Statistics report saved to: $statsPath"
Write-Host "HTML report has been generated and displayed"

# # Disconnect from all services
# Write-Host "Disconnecting from services..."
# Disconnect-MgGraph
# Disconnect-AzureAD
# Disconnect-ExchangeOnline -Confirm:$false