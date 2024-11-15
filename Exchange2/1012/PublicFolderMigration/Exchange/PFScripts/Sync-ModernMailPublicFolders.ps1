# .SYNOPSIS
#    Syncs modern mail-enabled public folder objects from the local Exchange deployment into O365. It uses the local Exchange deployment
#    as master to determine what changes need to be applied to O365. The script will create, update or delete mail-enabled public
#    folder objects on O365 Active Directory when appropriate.
#
# .DESCRIPTION
#    The script must be executed from an Exchange 2013 or later Management Shell window providing access to mail public folders in
#    the local Exchange deployment. Then, using the credentials provided, the script will create a session against Exchange Online,
#    which will be used to manipulate O365 Active Directory objects remotely.
#
#    Copyright (c) 2014 Microsoft Corporation. All rights reserved.
#
#    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK
#    OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
#
# .PARAMETER Credential
#    Exchange Online user name and password. Don't use this param if MFA is enabled.
#
# .PARAMETER CsvSummaryFile
#    The file path where sync operations and errors will be logged in a CSV format.
#
# .PARAMETER ConnectionUri
#    The Exchange Online remote PowerShell connection uri. If you are an Office 365 operated by 21Vianet customer in China, use "https://partner.outlook.cn/PowerShell".
#
# .PARAMETER Confirm
#    The Confirm switch causes the script to pause processing and requires you to acknowledge what the script will do before processing continues. You don't have to specify
#    a value with the Confirm switch.
#
# .PARAMETER FixInconsistencies
#    Fixes any inconsistencies such as orphaned, duplicate or disconnected mail public folders
#
# .PARAMETER Force
#    Force the script execution and bypass validation warnings.
#
# .PARAMETER WhatIf
#    The WhatIf switch instructs the script to simulate the actions that it would take on the object. By using the WhatIf switch, you can view what changes would occur
#    without having to apply any of those changes. You don't have to specify a value with the WhatIf switch.
#
# .EXAMPLE
#    .\Sync-ModernMailPublicFolders.ps1 -CsvSummaryFile:sync_summary.csv
#    
#    This example shows how to sync mail-public folders from your local deployment to Exchange Online. Note that the script outputs a CSV file listing all operations executed, and possibly errors encountered, during sync.
#
# .EXAMPLE
#    .\Sync-ModernMailPublicFolders.ps1 -CsvSummaryFile:sync_summary.csv -ConnectionUri:"https://partner.outlook.cn/PowerShell"
#    
#    This example shows how to use a different URI to connect to Exchange Online and sync modern mail-public folders from your local deployment.
#
param(
    [Parameter(Mandatory=$false)]
    [PSCredential] $Credential,

    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [string] $CsvSummaryFile,
    
    [Parameter(Mandatory=$false)]
    [ValidateNotNullOrEmpty()]
    [string] $ConnectionUri = "https://outlook.office365.com/powerShell-liveID",

    [Parameter(Mandatory=$false)]
    [bool] $Confirm = $true,

    [Parameter(Mandatory=$false)]
    [switch] $FixInconsistencies = $false,

    [Parameter(Mandatory=$false)]
    [switch] $Force = $false,

    [Parameter(Mandatory=$false)]
    [switch] $WhatIf = $false
)

# Writes a dated information message to console
function WriteInfoMessage()
{
    param ($message)
    Write-Host "[$($(Get-Date).ToString())]" $message;
}

# Writes a dated warning message to console
function WriteWarningMessage()
{
    param ($message)
    Write-Warning ("[{0}] {1}" -f (Get-Date),$message);
}

# Writes a verbose message to console
function WriteVerboseMessage()
{
    param ($message)
    Write-Host "[VERBOSE] $message" -ForegroundColor Green -BackgroundColor Black;
}

# Writes an error importing a mail public folder to the CSV summary
function WriteErrorSummary()
{
    param ($folder, $operation, $errorMessage, $commandtext)

    WriteOperationSummary $folder.Guid $operation $errorMessage $commandtext;
    $script:errorsEncountered++;
}

# Writes the operation executed and its result to the output CSV
function WriteOperationSummary()
{
    param ($folder, $operation, $result, $commandtext)

    $columns = @(
        (Get-Date).ToString(),
        $folder.Guid,
        $operation,
        (EscapeCsvColumn $result),
        (EscapeCsvColumn $commandtext)
    );

    Add-Content $CsvSummaryFile -Value ("{0},{1},{2},{3},{4}" -f $columns);
}

#Escapes a column value based on RFC 4180 (http://tools.ietf.org/html/rfc4180)
function EscapeCsvColumn()
{
    param ([string]$text)

    if ($text -eq $null)
    {
        return $text;
    }

    $hasSpecial = $false;
    for ($i=0; $i -lt $text.Length; $i++)
    {
        $c = $text[$i];
        if ($c -eq $script:csvEscapeChar -or
            $c -eq $script:csvFieldDelimiter -or
            $script:csvSpecialChars -contains $c)
        {
            $hasSpecial = $true;
            break;
        }
    }

    if (-not $hasSpecial)
    {
        return $text;
    }
    
    $ch = $script:csvEscapeChar.ToString([System.Globalization.CultureInfo]::InvariantCulture);
    return $ch + $text.Replace($ch, $ch + $ch) + $ch;
}

# Writes the current progress
function WriteProgress()
{
    param($statusFormat, $statusProcessed, $statusTotal)
    Write-Progress -Activity $LocalizedStrings.ProgressBarActivity `
        -Status ($statusFormat -f $statusProcessed,$statusTotal) `
        -PercentComplete (100 * ($script:itemsProcessed + $statusProcessed)/$script:totalItems);
}

# Create a tenant PSSession against Exchange Online with modern auth.
function InitializeExchangeOnlineRemoteSession()
{    
    WriteInfoMessage $LocalizedStrings.CreatingRemoteSession;

    $oldWarningPreference = $WarningPreference;
    $oldVerbosePreference = $VerbosePreference;

    try
    {
        Import-Module ExchangeOnlineManagement -ErrorAction SilentlyContinue;
        if (Get-Module ExchangeOnlineManagement)
        {
            $sessionOption = (New-PSSessionOption -SkipCACheck);
    	    Connect-ExchangeOnline -Credential $Credential -ConnectionURI $ConnectionUri -PSSessionOption $sessionOption -Prefix "Remote" -ErrorAction SilentlyContinue;
            $script:isConnectedToExchangeOnline = $true;
        }
        else
        {
            WriteWarningMessage $LocalizedStrings.EXOV2ModuleNotInstalled;
            Exit;
        }
    }
    finally
    {
        if ($script:isConnectedToExchangeOnline)
        {
            $WarningPreference = $oldWarningPreference;
            $VerbosePreference = $oldVerbosePreference;
        }
    }
    WriteInfoMessage $LocalizedStrings.RemoteSessionCreatedSuccessfully;
}

# Invokes New-SyncMailPublicFolder to create a new MEPF object on AD
function NewMailEnabledPublicFolder()
{
    param ($localFolder)

    if ($localFolder.PrimarySmtpAddress.ToString() -eq "")
    {
        $errorMsg = ($LocalizedStrings.FailedToCreateMailPublicFolderEmptyPrimarySmtpAddress -f $localFolder.Guid);
        Write-Error $errorMsg;
        WriteErrorSummary $localFolder $LocalizedStrings.CreateOperationName $errorMsg "";
        return;
    }

    # preserve the ability to reply via Outlook's nickname cache post-migration
    $emailAddressesArray = $localFolder.EmailAddresses.ToStringArray() + ("x500:" + $localFolder.LegacyExchangeDN);
           
    $newParams = @{};
    AddNewOrSetCommonParameters $localFolder $emailAddressesArray $newParams;

    [string]$commandText = (FormatCommand $script:NewSyncMailPublicFolderCommand $newParams);

    if ($script:verbose)
    {
        WriteVerboseMessage $commandText;
    }

    try
    {
        $result = &$script:NewSyncMailPublicFolderCommand @newParams;
        WriteOperationSummary $localFolder $LocalizedStrings.CreateOperationName $LocalizedStrings.CsvSuccessResult $commandText;

        if (-not $WhatIf)
        {
            $script:ObjectsCreated++;
        }
    }
    catch
    {
        WriteErrorSummary $localFolder $LocalizedStrings.CreateOperationName $error[0].Exception.Message $commandText;
        Write-Error $_;
    }
}

# Invokes Remove-SyncMailPublicFolder to remove a MEPF from AD
function RemoveMailEnabledPublicFolder()
{
    param ($remoteFolder)    

    $removeParams = @{};
    $removeParams.Add("Identity", $remoteFolder.DistinguishedName);
    $removeParams.Add("Confirm", $false);
    $removeParams.Add("WarningAction", [System.Management.Automation.ActionPreference]::SilentlyContinue);
    $removeParams.Add("ErrorAction", [System.Management.Automation.ActionPreference]::Stop);

    if ($WhatIf)
    {
        $removeParams.Add("WhatIf", $true);
    }
    
    [string]$commandText = (FormatCommand $script:RemoveSyncMailPublicFolderCommand $removeParams);

    if ($script:verbose)
    {
        WriteVerboseMessage $commandText;
    }
    
    try
    {
        &$script:RemoveSyncMailPublicFolderCommand @removeParams;
        WriteOperationSummary $remoteFolder $LocalizedStrings.RemoveOperationName $LocalizedStrings.CsvSuccessResult $commandText;

        if (-not $WhatIf)
        {
            $script:ObjectsDeleted++;
        }
    }
    catch
    {
        WriteErrorSummary $remoteFolder $LocalizedStrings.RemoveOperationName $_.Exception.Message $commandText;
        Write-Error $_;
    }
}

# Invokes Set-MailPublicFolder to update the properties of an existing MEPF
function UpdateMailEnabledPublicFolder()
{
    param ($localFolder, $remoteFolder)

    $localEmailAddresses = $localFolder.EmailAddresses.ToStringArray();
    $localEmailAddresses += ("x500:" + $localFolder.LegacyExchangeDN); # preserve the ability to reply via Outlook's nickname cache post-migration
    $emailAddresses = ConsolidateEmailAddresses $localEmailAddresses $remoteFolder.EmailAddresses $remoteFolder.LegacyExchangeDN;

    $setParams = @{};
    $setParams.Add("Identity", $remoteFolder.DistinguishedName);

    if ($script:mailEnabledSystemFolders.Contains($localFolder.Guid))
    {
        $setParams.Add("IgnoreMissingFolderLink", $true);
    }

    AddNewOrSetCommonParameters $localFolder $emailAddresses $setParams;

    [string]$commandText = (FormatCommand $script:SetMailPublicFolderCommand $setParams);

    if ($script:verbose)
    {
        WriteVerboseMessage $commandText;
    }

    try
    {
        &$script:SetMailPublicFolderCommand @setParams;
        WriteOperationSummary $remoteFolder $LocalizedStrings.UpdateOperationName $LocalizedStrings.CsvSuccessResult $commandText;

        if (-not $WhatIf)
        {
            $script:ObjectsUpdated++;
        }
    }
    catch
    {
        WriteErrorSummary $remoteFolder $LocalizedStrings.UpdateOperationName $_.Exception.Message $commandText;
        Write-Error $_;
    }
}

# Adds the common set of parameters between New and Set cmdlets to the given dictionary
function AddNewOrSetCommonParameters()
{
    param ($localFolder, $emailAddresses, [System.Collections.IDictionary]$parameters)

    $windowsEmailAddress = $localFolder.WindowsEmailAddress.ToString();
    if ($windowsEmailAddress -eq "")
    {
        $windowsEmailAddress = $localFolder.PrimarySmtpAddress.ToString();      
    }

    $parameters.Add("Alias", $localFolder.Alias.Trim());
    $parameters.Add("DisplayName", $localFolder.DisplayName.Trim());
    $parameters.Add("EmailAddresses", $emailAddresses);
    $parameters.Add("ExternalEmailAddress", $localFolder.PrimarySmtpAddress.ToString());
    $parameters.Add("HiddenFromAddressListsEnabled", $localFolder.HiddenFromAddressListsEnabled);
    $parameters.Add("Name", $localFolder.Name.Trim());
    $parameters.Add("OnPremisesObjectId", $localFolder.Guid);
    $parameters.Add("WindowsEmailAddress", $windowsEmailAddress);
    $parameters.Add("ErrorAction", [System.Management.Automation.ActionPreference]::Stop);

    if ($WhatIf)
    {
        $parameters.Add("WhatIf", $true);
    }
}

# Finds out the cloud-only email addresses and merges those with the values current persisted in the on-premises object
function ConsolidateEmailAddresses()
{
    param($localEmailAddresses, $remoteEmailAddresses, $remoteLegDN)

    # Check if the email address in the existing cloud object is present on-premises; if it is not, then the address was either:
    # 1. Deleted on-premises and must be removed from cloud
    # 2. or it is a cloud-authoritative address and should be kept
    $remoteAuthoritative = @();
    foreach ($remoteAddress in $remoteEmailAddresses)
    {
        if ($remoteAddress.StartsWith("SMTP:", [StringComparison]::InvariantCultureIgnoreCase))
        {
            $found = $false;
            $remoteAddressParts = $remoteAddress.Split($script:proxyAddressSeparators); # e.g. SMTP:alias@domain
            if ($remoteAddressParts.Length -ne 3)
            {
                continue; # Invalid SMTP proxy address (it will be removed)
            }

            foreach ($localAddress in $localEmailAddresses)
            {
                # note that the domain part of email addresses is case insensitive while the alias part is case sensitive
                $localAddressParts = $localAddress.Split($script:proxyAddressSeparators);
                if ($localAddressParts.Length -eq 3 -and
                    $remoteAddressParts[0].Equals($localAddressParts[0], [StringComparison]::InvariantCultureIgnoreCase) -and
                    $remoteAddressParts[1].Equals($localAddressParts[1], [StringComparison]::InvariantCulture) -and
                    $remoteAddressParts[2].Equals($localAddressParts[2], [StringComparison]::InvariantCultureIgnoreCase))
                {
                    $found = $true;
                    break;
                }
            }

            if (-not $found)
            {
                foreach ($domain in $script:authoritativeDomains)
                {
                    if ($remoteAddressParts[2] -eq $domain)
                    {
                        $found = $true;
                        break;
                    }
                }

                if (-not $found)
                {
                    # the address on the remote object is from a cloud authoritative domain and should not be removed
                    $remoteAuthoritative += $remoteAddress;
                }
            }
        }
        elseif ($remoteAddress.StartsWith("X500:", [StringComparison]::InvariantCultureIgnoreCase) -and
            $remoteAddress.Substring(5) -eq $remoteLegDN)
        {
            $remoteAuthoritative += $remoteAddress;
        }
    }

    return $localEmailAddresses + $remoteAuthoritative;
}

# Formats the command and its parameters to be printed on console or to file
function FormatCommand()
{
    param ([string]$command, [System.Collections.IDictionary]$parameters)

    $commandText = New-Object System.Text.StringBuilder;
    [void]$commandText.Append($command);
    foreach ($name in $parameters.Keys)
    {
        [void]$commandText.AppendFormat(" -{0}:",$name);

        $value = $parameters[$name];
        if ($value -isnot [Array])
        {
            [void]$commandText.AppendFormat("`"{0}`"", $value);
        }
        elseif ($value.Length -eq 0)
        {
            [void]$commandText.Append("@()");
        }
        else
        {
            [void]$commandText.Append("@(");
            foreach ($subValue in $value)
            {
                [void]$commandText.AppendFormat("`"{0}`",",$subValue);
            }
            
            [void]$commandText.Remove($commandText.Length - 1, 1);
            [void]$commandText.Append(")");
        }
    }

    return $commandText.ToString();
}

function ValidateMailEnabledPublicFolders()
{
    $validateMailEnabledPublicFoldersScriptFile = Join-Path $PWD "ValidateMailEnabledPublicFolders.ps1";
    if (!(Test-Path $validateMailEnabledPublicFoldersScriptFile))
    {
        try 
        {
            # Download validate-mepf script
            WriteInfoMessage $LocalizedStrings.DownloadingValidateMEPFScript;
            Invoke-WebRequest -Uri "https://aka.ms/validatemepf" -OutFile $validateMailEnabledPublicFoldersScriptFile;
        }
        catch 
        {
            WriteWarningMessage ($LocalizedStrings.DownloadValidateMEPFScriptFailed -f $PWD);
            return;
        }
    }

    .\ValidateMailEnabledPublicFolders.ps1;

    if ($FixInconsistencies)
    {
        FixInconsistenciesWithMEPF;
    }
}

function checkForInconsistenciesWithMEPF()
{
    $files = @(
        $script:foldersToMailDisableFile, 
        $script:mailPublicFolderOrphansFile, 
        $script:mailPublicFolderDuplicatesFile, 
        $script:emailAddressMergeScriptFile,
        $script:mailDisabledWithProxyGuidFile,
        $script:mailPublicFoldersDisconnectedFile
    );

    # If there are any inconsistencies with mail-enabled public folders, the validatemepf script outputs any of these files
    for ($i = 0; $i -lt $files.Length; $i++)
    {
        if (Test-Path $files[$i])
        {
            return $true;
        }
    }
    
    return $false;
}

function FixInconsistenciesWithMEPF()
{
    Get-MailPublicFolder -ResultSize Unlimited | Export-Clixml ("MailPublicFolders_{0}.xml" -f (Get-Date -f yyyy-MM-ddThh-mm-ss)) -Encoding UTF8;

    # FoldersToMailDisableFile contains Identities of those mail enabled PFs which have no AD objects
    if (Test-Path $script:foldersToMailDisableFile)
    {
        WriteInfoMessage ($LocalizedStrings.MailDisablePublicFoldersInFile -f $script:foldersToMailDisableFile);
        Get-Content $script:foldersToMailDisableFile | % { Set-PublicFolder $_ -MailEnabled $false };
        Move-Item -Path $script:foldersToMailDisableFile -Destination $script:foldersToMailDisableFile.Replace('.txt', '_Processed.txt') -Force;
    }

    # MailPublicFolderOrphansFile contains DistinguishedNames of orphaned MEPFs
    if (Test-Path $script:mailPublicFolderOrphansFile)
    {
        WriteInfoMessage ($LocalizedStrings.DeleteOrphanedMailPublicFoldersInFile -f $script:mailPublicFolderOrphansFile);
        Get-Content $script:mailPublicFolderOrphansFile | % { $folder = ([ADSI]("LDAP://$_")); $parent = ([ADSI]"$($folder.Parent)"); $parent.Children.Remove($folder) };
        Move-Item -Path $script:mailPublicFolderOrphansFile -Destination $script:mailPublicFolderOrphansFile.Replace('.txt', '_Processed.txt') -Force;
    }

    # MailPublicFolderDuplicatesFile contains DistinguishedNames of duplicate MEPFs (MEPFs wrongly associated with same PF)
    if (Test-Path $script:mailPublicFolderDuplicatesFile)
    {
        WriteInfoMessage ($LocalizedStrings.DeleteDuplicateMailPublicFoldersInFile -f $script:mailPublicFolderDuplicatesFile);
        Get-Content $script:mailPublicFolderDuplicatesFile | % { $folder = ([ADSI]("LDAP://$_")); $parent = ([ADSI]"$($folder.Parent)"); $parent.Children.Remove($folder) };
        Move-Item -Path $script:mailPublicFolderDuplicatesFile -Destination $script:mailPublicFolderDuplicatesFile.Replace('.txt', '_Processed.txt') -Force;
    }

    # EmailAddressMergeScriptFile contains the script to merge the email addresses (which might still be in use) of duplicate MEPFs which were deleted.
    if (Test-Path $script:emailAddressMergeScriptFile)
    {
        WriteInfoMessage $LocalizedStrings.AddAddressesFromDuplicates;
        .\AddAddressesFromDuplicates.ps1;
        Move-Item -Path $script:emailAddressMergeScriptFile -Destination $script:emailAddressMergeScriptFile.Replace('.ps1', '_Processed.ps1') -Force;
    }

    # MailDisabledWithProxyGuidFile contains Identities of those PFs which are mail-disabled but have a proxy GUID
    if (Test-Path $script:mailDisabledWithProxyGuidFile)
    {
        WriteInfoMessage ($LocalizedStrings.MailEnablePublicFoldersWithProxyGUIDinFile -f $script:mailDisabledWithProxyGuidFile);
        Get-Content $script:mailDisabledWithProxyGuidFile | % { Enable-MailPublicFolder $_ };
        Move-Item -Path $script:mailDisabledWithProxyGuidFile -Destination $script:mailDisabledWithProxyGuidFile.Replace('.txt', '_Processed.txt') -Force;
    }

    # MailPublicFoldersDisconnectedFile contains DistinguishedNames of those MEPFs which are not associated to any PF
    if (Test-Path $script:mailPublicFoldersDisconnectedFile) {
        WriteInfoMessage ($LocalizedStrings.MailEnablePFAssociatedToDisconnectedMEPFsInFile -f $script:mailPublicFoldersDisconnectedFile);
        $disconnectedMEPFDNs = Get-Content $script:MailPublicFoldersDisconnectedFile;
        foreach ($dN in $disconnectedMEPFDNs) {
            $mailPublicFolder = Get-MailPublicFolder $dN;
            $publicFolder = Get-PublicFolder $mailPublicFolder.EntryId;
            if (!$publicFolder.MailEnabled)
            {
                # Update the MailEnabled and MailRecipientGuid properties of the public folder
                Invoke-Command { Set-PublicFolder $publicFolder -MailEnabled:$true -MailRecipientGuid $mailPublicFolder.Guid } -ErrorVariable errorOutput;
          
                # If the above command fails, simply mail-enable the PF and add the emailAddresses to its MEPF
                if ($errorOutput -ne $null)
                {
                    Enable-MailPublicFolder $publicFolder;
                    RemoveMEPFAndAddEmailAddresses $mailPublicFolder $publicFolder;
                }
            }
            else
            {
                # This case arises when there are multiple disconnected mepfs pointing to same PF
                # Once the PF is MailEnabled and MailRecipientGuid of the first disconnected mepf found in the list is set, remaining are simply duplicate MEPFs
                # Remove these duplicates and add email addresses to the mepf connected to it's PF
                RemoveMEPFAndAddEmailAddresses $mailPublicFolder $publicFolder;
            }
        }
        Move-Item -Path $script:mailPublicFoldersDisconnectedFile -Destination $script:mailPublicFoldersDisconnectedFile.Replace('.txt', '_Processed.txt') -Force;
    }
}

function RemoveMEPFAndAddEmailAddresses()
{
    param ($duplicateMailPublicFolder, $publicFolder)

    # Remove duplicate mepf
    $folder = ([ADSI]("LDAP://$($duplicateMailPublicFolder.DistinguishedName)")); 
    $parent = ([ADSI]"$($folder.Parent)"); 
    $parent.Children.Remove($folder);

    # Add email addresses
    $emailAddressesToAdd = @();
    foreach ($emailAddress in $duplicateMailPublicFolder.EmailAddresses)
    {
        if ($emailAddress.ToString().StartsWith("SMTP"))
        {
            # Add the address as a secondary smtp address
            $emailAddressesToAdd += $emailAddress.ToString().Substring($emailAddress.ToString().IndexOf(':') + 1);
        }
        else
        {
            $emailAddressesToAdd += $emailAddress.ToString();
        }
    }
    Set-MailPublicFolder $publicFolder -EmailAddresses @{add=$emailAddressesToAdd};
}

################ DECLARING GLOBAL VARIABLES ################
$script:isConnectedToExchangeOnline = $false;
$script:verbose = $VerbosePreference -eq [System.Management.Automation.ActionPreference]::Continue;

$script:csvSpecialChars = @("`r", "`n");
$script:csvEscapeChar = '"';
$script:csvFieldDelimiter = ',';

$script:ObjectsCreated = $script:ObjectsUpdated = $script:ObjectsDeleted = 0;
$script:NewSyncMailPublicFolderCommand = "New-RemoteSyncMailPublicFolder";
$script:SetMailPublicFolderCommand = "Set-RemoteMailPublicFolder";
$script:RemoveSyncMailPublicFolderCommand = "Remove-RemoteSyncMailPublicFolder";
[char[]]$script:proxyAddressSeparators = ':','@';
$script:errorsEncountered = 0;
$script:authoritativeDomains = $null;
$script:mailEnabledSystemFolders = New-Object 'System.Collections.Generic.HashSet[Guid]'; 
$script:WellKnownSystemFolders = @(
    "\NON_IPM_SUBTREE\EFORMS REGISTRY",
    "\NON_IPM_SUBTREE\OFFLINE ADDRESS BOOK",
    "\NON_IPM_SUBTREE\SCHEDULE+ FREE BUSY",
    "\NON_IPM_SUBTREE\schema-root",
    "\NON_IPM_SUBTREE\Events Root");
$script:foldersToMailDisableFile = Join-Path $PWD "FoldersToMailDisable.txt";
$script:mailPublicFolderOrphansFile = Join-Path $PWD "MailPublicFolderOrphans.txt";
$script:mailPublicFolderDuplicatesFile = Join-Path $PWD "MailPublicFolderDuplicates.txt";
$script:emailAddressMergeScriptFile = Join-Path $PWD "AddAddressesFromDuplicates.ps1";
$script:mailDisabledWithProxyGuidFile = Join-Path $PWD "MailDisabledWithProxyGuid.txt";
$script:mailPublicFoldersDisconnectedFile = Join-Path $PWD "MailPublicFoldersDisconnected.txt";

#load hashtable of localized string
Import-LocalizedData -BindingVariable LocalizedStrings -FileName SyncModernMailPublicFolders.strings.psd1

#minimum supported exchange version to run this script
$minSupportedVersion = 8
################ END OF DECLARATION #################

try
{
    ValidateMailEnabledPublicFolders;
}
catch
{
    WriteWarningMessage $LocalizedStrings.ValidateMailEnabledPublicFoldersFailed;
    WriteWarningMessage $_;
}

if (Test-Path $CsvSummaryFile)
{
    Remove-Item $CsvSummaryFile -Confirm:$Confirm -Force;
}

# Write the output CSV headers
$csvFile = New-Item -Path $CsvSummaryFile -ItemType File -Force -ErrorAction:Stop -Value ("#{0},{1},{2},{3},{4}`r`n" -f $LocalizedStrings.TimestampCsvHeader,
    $LocalizedStrings.IdentityCsvHeader,
    $LocalizedStrings.OperationCsvHeader,
    $LocalizedStrings.ResultCsvHeader,
    $LocalizedStrings.CommandCsvHeader);

$localServerVersion = (Get-ExchangeServer $env:COMPUTERNAME -ErrorAction:Stop).AdminDisplayVersion;
# This script can run from Exchange 2007 Management shell and above
if ($localServerVersion.Major -lt $minSupportedVersion)
{
    Write-Error ($LocalizedStrings.LocalServerVersionNotSupported -f $localServerVersion) -ErrorAction:Continue;
    Exit;
}

try
{
    InitializeExchangeOnlineRemoteSession;

    WriteInfoMessage $LocalizedStrings.LocalMailPublicFolderEnumerationStart;

    # During finalization, Public Folders deployment is locked for migration, which means the script cannot invoke
    # Get-PublicFolder as that operation would fail. In that case, the script cannot determine which mail public folder
    # objects are linked to system folders under the NON_IPM_SUBTREE.
    $lockedForMigration = (Get-OrganizationConfig).PublicFolderMailboxesLockedForNewConnections;
    $allSystemFoldersInAD = @();
    if (-not $lockedForMigration)
    {
        # See https://technet.microsoft.com/en-us/library/bb397221(v=exchg.141).aspx#Trees
        # Certain WellKnownFolders in pre-E15 are created with prefix such as OWAScratchPad, StoreEvents.
        # For instance, StoreEvents folders have the following pattern: "\NON_IPM_SUBTREE\StoreEvents{46F83CF7-2A81-42AC-A0C6-68C7AA49FF18}\internal1"
        $storeEventAndOwaScratchPadFolders = @(Get-PublicFolder \NON_IPM_SUBTREE -GetChildren -ResultSize:Unlimited | ?{$_.Name -like "StoreEvents*" -or $_.Name -like "OWAScratchPad*"});
        $allSystemFolderParents = $storeEventAndOwaScratchPadFolders + @($script:WellKnownSystemFolders | Get-PublicFolder -ErrorAction:SilentlyContinue);
        $allSystemFoldersInAD = @($allSystemFolderParents | Get-PublicFolder -Recurse -ResultSize:Unlimited | Get-MailPublicFolder -ErrorAction:SilentlyContinue);

        foreach ($systemFolder in $allSystemFoldersInAD)
        {
            [void]$script:mailEnabledSystemFolders.Add($systemFolder.Guid);
        }
    }
    else
    {
        WriteWarningMessage $LocalizedStrings.UnableToDetectSystemMailPublicFolders;
    }

    if ($script:verbose)
    {
        WriteVerboseMessage ($LocalizedStrings.SystemFoldersSkipped -f $script:mailEnabledSystemFolders.Count);
        $allSystemFoldersInAD | Sort Alias | ft -a | Out-String | Write-Host -ForegroundColor Green -BackgroundColor Black;
    }

    $localFolders = @(Get-MailPublicFolder -ResultSize:Unlimited -IgnoreDefaultScope | Sort Guid);
    WriteInfoMessage ($LocalizedStrings.LocalMailPublicFolderEnumerationCompleted -f $localFolders.Length);

    if ($localFolders.Length -eq 0 -and $Force -eq $false)
    {
        WriteWarningMessage $LocalizedStrings.ForceParameterRequired;
        Exit;
    }

    WriteInfoMessage $LocalizedStrings.RemoteMailPublicFolderEnumerationStart;
    $remoteFolders = @(Get-RemoteMailPublicFolder -ResultSize:Unlimited | Sort OnPremisesObjectId);
    WriteInfoMessage ($LocalizedStrings.RemoteMailPublicFolderEnumerationCompleted -f $remoteFolders.Length);

    $missingOnPremisesGuid = @();
    $pendingRemoves = @();
    $pendingUpdates = @{};
    $pendingAdds = @{};

    $localIndex = 0;
    $remoteIndex = 0;
    while ($localIndex -lt $localFolders.Length -and $remoteIndex -lt $remoteFolders.Length)
    {
        $local = $localFolders[$localIndex];
        $remote = $remoteFolders[$remoteIndex];

        if ($remote.OnPremisesObjectId -eq "")
        {
            # This folder must be processed based on PrimarySmtpAddress
            $missingOnPremisesGuid += $remote;
            $remoteIndex++;
        }
        elseif ($local.Guid.ToString() -eq $remote.OnPremisesObjectId)
        {
            $pendingUpdates.Add($local.Guid, (New-Object PSObject -Property @{ Local=$local; Remote=$remote }));
            $localIndex++;
            $remoteIndex++;
        }
        elseif ($local.Guid.ToString() -lt $remote.OnPremisesObjectId)
        {
            if (-not $script:mailEnabledSystemFolders.Contains($local.Guid))
            {
                $pendingAdds.Add($local.Guid, $local);
            }

            $localIndex++;
        }
        else
        {
            $pendingRemoves += $remote;
            $remoteIndex++;
        }
    }

    # Remaining folders on $localFolders collection must be added to Exchange Online
    while ($localIndex -lt $localFolders.Length)
    {
        $local = $localFolders[$localIndex];

        if (-not $script:mailEnabledSystemFolders.Contains($local.Guid))
        {
            $pendingAdds.Add($local.Guid, $local);
        }

        $localIndex++;
    }

    # Remaining folders on $remoteFolders collection must be removed from Exchange Online
    while ($remoteIndex -lt $remoteFolders.Length)
    {
        $remote = $remoteFolders[$remoteIndex];
        if ($remote.OnPremisesObjectId  -eq "")
        {
            # This folder must be processed based on PrimarySmtpAddress
            $missingOnPremisesGuid += $remote;
        }
        else
        {
            $pendingRemoves += $remote;
        }
        
        $remoteIndex++;
    }

    if ($missingOnPremisesGuid.Length -gt 0)
    {
        # Process remote objects missing the OnPremisesObjectId using the PrimarySmtpAddress as a key instead.
        $missingOnPremisesGuid = @($missingOnPremisesGuid | Sort PrimarySmtpAddress);
        $localFolders = @($localFolders | Sort PrimarySmtpAddress);

        $localIndex = 0;
        $remoteIndex = 0;
        while ($localIndex -lt $localFolders.Length -and $remoteIndex -lt $missingOnPremisesGuid.Length)
        {
            $local = $localFolders[$localIndex];
            $remote = $missingOnPremisesGuid[$remoteIndex];

            if ($local.PrimarySmtpAddress.ToString() -eq $remote.PrimarySmtpAddress.ToString())
            {
                # Make sure the PrimarySmtpAddress has no duplicate on-premises; otherwise, skip updating all objects with duplicate address
                $j = $localIndex + 1;
                while ($j -lt $localFolders.Length)
                {
                    $next = $localFolders[$j];
                    if ($local.PrimarySmtpAddress.ToString() -ne $next.PrimarySmtpAddress.ToString())
                    {
                        break;
                    }

                    WriteErrorSummary $next $LocalizedStrings.UpdateOperationName ($LocalizedStrings.PrimarySmtpAddressUsedByAnotherFolder -f $local.PrimarySmtpAddress,$local.Guid) "";

                    # If there were a previous match based on OnPremisesObjectId, remove the folder operation from add and update collections
                    $pendingAdds.Remove($next.Guid);
                    $pendingUpdates.Remove($next.Guid);
                    $j++;
                }

                $duplicatesFound = $j - $localIndex - 1;
                if ($duplicatesFound -gt 0)
                {
                    # If there were a previous match based on OnPremisesObjectId, remove the folder operation from add and update collections
                    $pendingAdds.Remove($local.Guid);
                    $pendingUpdates.Remove($local.Guid);
                    $localIndex += $duplicatesFound + 1;

                    WriteErrorSummary $local $LocalizedStrings.UpdateOperationName ($LocalizedStrings.PrimarySmtpAddressUsedByOtherFolders -f $local.PrimarySmtpAddress,$duplicatesFound) "";
                    WriteWarningMessage ($LocalizedStrings.SkippingFoldersWithDuplicateAddress -f ($duplicatesFound + 1),$local.PrimarySmtpAddress);
                }
                elseif ($pendingUpdates.Contains($local.Guid))
                {
                    # If we get here, it means two different remote objects match the same local object (one by OnPremisesObjectId and another by PrimarySmtpAddress).
                    # Since that is an ambiguous resolution, let's skip updating the remote objects.
                    $ambiguousRemoteObj = $pendingUpdates[$local.Guid].Remote;
                    $pendingUpdates.Remove($local.Guid);

                    $errorMessage = ($LocalizedStrings.AmbiguousLocalMailPublicFolderResolution -f $local.Guid,$ambiguousRemoteObj.Guid,$remote.Guid);
                    WriteErrorSummary $local $LocalizedStrings.UpdateOperationName $errorMessage "";
                    WriteWarningMessage $errorMessage;
                }
                else
                {
                    # Since there was no match originally using OnPremisesObjectId, the local object was treated as an add to Exchange Online.
                    # In this way, since we now found a remote object (by PrimarySmtpAddress) to update, we must first remove the local object from the add list.
                    $pendingAdds.Remove($local.Guid);
                    $pendingUpdates.Add($local.Guid, (New-Object PSObject -Property @{ Local=$local; Remote=$remote }));
                }

                $localIndex++;
                $remoteIndex++;
            }
            elseif ($local.PrimarySmtpAddress.ToString() -gt $remote.PrimarySmtpAddress.ToString())
            {
                # There are no local objects using the remote object's PrimarySmtpAddress
                $pendingRemoves += $remote;
                $remoteIndex++;
            }
            else
            {
                $localIndex++;
            }
        }

        # All objects remaining on the $missingOnPremisesGuid list no longer exist on-premises
        while ($remoteIndex -lt $missingOnPremisesGuid.Length)
        {
            $pendingRemoves += $missingOnPremisesGuid[$remoteIndex];
            $remoteIndex++;
        }
    }

    $script:totalItems = $pendingRemoves.Length + $pendingUpdates.Count + $pendingAdds.Count;

    # At this point, we know all changes that need to be synced to Exchange Online. Let's prompt the admin for confirmation before proceeding.
    if ($Confirm -eq $true -and $script:totalItems -gt 0)
    {
        $title = $LocalizedStrings.ConfirmationTitle;
        $message = ($LocalizedStrings.ConfirmationQuestion -f $pendingAdds.Count,$pendingUpdates.Count,$pendingRemoves.Length);
        $yes = New-Object System.Management.Automation.Host.ChoiceDescription $LocalizedStrings.ConfirmationYesOption, `
            $LocalizedStrings.ConfirmationYesOptionHelp;

        $no = New-Object System.Management.Automation.Host.ChoiceDescription $LocalizedStrings.ConfirmationNoOption, `
            $LocalizedStrings.ConfirmationNoOptionHelp;

        [System.Management.Automation.Host.ChoiceDescription[]]$options = $no,$yes;
        $confirmation = $host.ui.PromptForChoice($title, $message, $options, 0);
        if ($confirmation -eq 0)
        {
            Exit;
        }
    }

    # Find out the authoritative AcceptedDomains on-premises so that we don't accidently remove cloud-only email addresses during updates
    $script:authoritativeDomains = @(Get-AcceptedDomain | ?{$_.DomainType -eq "Authoritative" } | foreach {$_.DomainName.ToString()});
    
    # Finally, let's perfom the actual operations against Exchange Online
    $script:itemsProcessed = 0;
    for ($i = 0; $i -lt $pendingRemoves.Length; $i++)
    {
        WriteProgress $LocalizedStrings.ProgressBarStatusRemoving $i $pendingRemoves.Length;
        RemoveMailEnabledPublicFolder $pendingRemoves[$i];
    }

    $script:itemsProcessed += $pendingRemoves.Length;
    $updatesProcessed = 0;
    foreach ($folderPair in $pendingUpdates.Values)
    {
        WriteProgress $LocalizedStrings.ProgressBarStatusUpdating $updatesProcessed $pendingUpdates.Count;
        UpdateMailEnabledPublicFolder $folderPair.Local $folderPair.Remote;
        $updatesProcessed++;
    }

    $script:itemsProcessed += $pendingUpdates.Count;
    $addsProcessed = 0;
    foreach ($localFolder in $pendingAdds.Values)
    {
        WriteProgress $LocalizedStrings.ProgressBarStatusCreating $addsProcessed $pendingAdds.Count;
        NewMailEnabledPublicFolder $localFolder;
        $addsProcessed++;
    }

    Write-Progress -Activity $LocalizedStrings.ProgressBarActivity -Status ($LocalizedStrings.ProgressBarStatusCreating -f $pendingAdds.Count,$pendingAdds.Count) -Completed;
    WriteInfoMessage ($LocalizedStrings.SyncMailPublicFolderObjectsComplete -f $script:ObjectsCreated,$script:ObjectsUpdated,$script:ObjectsDeleted);

    if ($script:errorsEncountered -gt 0)
    {
        WriteWarningMessage ($LocalizedStrings.ErrorsFoundDuringImport -f $script:errorsEncountered,(Get-Item $CsvSummaryFile).FullName);
    }
}
finally
{
    if (checkForInconsistenciesWithMEPF)
    {
        WriteWarningMessage $LocalizedStrings.FoundInconsistenciesWithMEPFs;
    }

    if ($script:isConnectedToExchangeOnline)
    {
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue;
    }
}
# SIG # Begin signature block
# MIIntwYJKoZIhvcNAQcCoIInqDCCJ6QCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAL+5vjsg2RW9lR
# mvQLuqnBiopkYhrNb1NbqEHvvQpwHqCCDYEwggX/MIID56ADAgECAhMzAAACUosz
# qviV8znbAAAAAAJSMA0GCSqGSIb3DQEBCwUAMH4xCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25p
# bmcgUENBIDIwMTEwHhcNMjEwOTAyMTgzMjU5WhcNMjIwOTAxMTgzMjU5WjB0MQsw
# CQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9u
# ZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMR4wHAYDVQQDExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24wggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIB
# AQDQ5M+Ps/X7BNuv5B/0I6uoDwj0NJOo1KrVQqO7ggRXccklyTrWL4xMShjIou2I
# sbYnF67wXzVAq5Om4oe+LfzSDOzjcb6ms00gBo0OQaqwQ1BijyJ7NvDf80I1fW9O
# L76Kt0Wpc2zrGhzcHdb7upPrvxvSNNUvxK3sgw7YTt31410vpEp8yfBEl/hd8ZzA
# v47DCgJ5j1zm295s1RVZHNp6MoiQFVOECm4AwK2l28i+YER1JO4IplTH44uvzX9o
# RnJHaMvWzZEpozPy4jNO2DDqbcNs4zh7AWMhE1PWFVA+CHI/En5nASvCvLmuR/t8
# q4bc8XR8QIZJQSp+2U6m2ldNAgMBAAGjggF+MIIBejAfBgNVHSUEGDAWBgorBgEE
# AYI3TAgBBggrBgEFBQcDAzAdBgNVHQ4EFgQUNZJaEUGL2Guwt7ZOAu4efEYXedEw
# UAYDVR0RBEkwR6RFMEMxKTAnBgNVBAsTIE1pY3Jvc29mdCBPcGVyYXRpb25zIFB1
# ZXJ0byBSaWNvMRYwFAYDVQQFEw0yMzAwMTIrNDY3NTk3MB8GA1UdIwQYMBaAFEhu
# ZOVQBdOCqhc3NyK1bajKdQKVMFQGA1UdHwRNMEswSaBHoEWGQ2h0dHA6Ly93d3cu
# bWljcm9zb2Z0LmNvbS9wa2lvcHMvY3JsL01pY0NvZFNpZ1BDQTIwMTFfMjAxMS0w
# Ny0wOC5jcmwwYQYIKwYBBQUHAQEEVTBTMFEGCCsGAQUFBzAChkVodHRwOi8vd3d3
# Lm1pY3Jvc29mdC5jb20vcGtpb3BzL2NlcnRzL01pY0NvZFNpZ1BDQTIwMTFfMjAx
# MS0wNy0wOC5jcnQwDAYDVR0TAQH/BAIwADANBgkqhkiG9w0BAQsFAAOCAgEAFkk3
# uSxkTEBh1NtAl7BivIEsAWdgX1qZ+EdZMYbQKasY6IhSLXRMxF1B3OKdR9K/kccp
# kvNcGl8D7YyYS4mhCUMBR+VLrg3f8PUj38A9V5aiY2/Jok7WZFOAmjPRNNGnyeg7
# l0lTiThFqE+2aOs6+heegqAdelGgNJKRHLWRuhGKuLIw5lkgx9Ky+QvZrn/Ddi8u
# TIgWKp+MGG8xY6PBvvjgt9jQShlnPrZ3UY8Bvwy6rynhXBaV0V0TTL0gEx7eh/K1
# o8Miaru6s/7FyqOLeUS4vTHh9TgBL5DtxCYurXbSBVtL1Fj44+Od/6cmC9mmvrti
# yG709Y3Rd3YdJj2f3GJq7Y7KdWq0QYhatKhBeg4fxjhg0yut2g6aM1mxjNPrE48z
# 6HWCNGu9gMK5ZudldRw4a45Z06Aoktof0CqOyTErvq0YjoE4Xpa0+87T/PVUXNqf
# 7Y+qSU7+9LtLQuMYR4w3cSPjuNusvLf9gBnch5RqM7kaDtYWDgLyB42EfsxeMqwK
# WwA+TVi0HrWRqfSx2olbE56hJcEkMjOSKz3sRuupFCX3UroyYf52L+2iVTrda8XW
# esPG62Mnn3T8AuLfzeJFuAbfOSERx7IFZO92UPoXE1uEjL5skl1yTZB3MubgOA4F
# 8KoRNhviFAEST+nG8c8uIsbZeb08SeYQMqjVEmkwggd6MIIFYqADAgECAgphDpDS
# AAAAAAADMA0GCSqGSIb3DQEBCwUAMIGIMQswCQYDVQQGEwJVUzETMBEGA1UECBMK
# V2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0
# IENvcnBvcmF0aW9uMTIwMAYDVQQDEylNaWNyb3NvZnQgUm9vdCBDZXJ0aWZpY2F0
# ZSBBdXRob3JpdHkgMjAxMTAeFw0xMTA3MDgyMDU5MDlaFw0yNjA3MDgyMTA5MDla
# MH4xCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdS
# ZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMT
# H01pY3Jvc29mdCBDb2RlIFNpZ25pbmcgUENBIDIwMTEwggIiMA0GCSqGSIb3DQEB
# AQUAA4ICDwAwggIKAoICAQCr8PpyEBwurdhuqoIQTTS68rZYIZ9CGypr6VpQqrgG
# OBoESbp/wwwe3TdrxhLYC/A4wpkGsMg51QEUMULTiQ15ZId+lGAkbK+eSZzpaF7S
# 35tTsgosw6/ZqSuuegmv15ZZymAaBelmdugyUiYSL+erCFDPs0S3XdjELgN1q2jz
# y23zOlyhFvRGuuA4ZKxuZDV4pqBjDy3TQJP4494HDdVceaVJKecNvqATd76UPe/7
# 4ytaEB9NViiienLgEjq3SV7Y7e1DkYPZe7J7hhvZPrGMXeiJT4Qa8qEvWeSQOy2u
# M1jFtz7+MtOzAz2xsq+SOH7SnYAs9U5WkSE1JcM5bmR/U7qcD60ZI4TL9LoDho33
# X/DQUr+MlIe8wCF0JV8YKLbMJyg4JZg5SjbPfLGSrhwjp6lm7GEfauEoSZ1fiOIl
# XdMhSz5SxLVXPyQD8NF6Wy/VI+NwXQ9RRnez+ADhvKwCgl/bwBWzvRvUVUvnOaEP
# 6SNJvBi4RHxF5MHDcnrgcuck379GmcXvwhxX24ON7E1JMKerjt/sW5+v/N2wZuLB
# l4F77dbtS+dJKacTKKanfWeA5opieF+yL4TXV5xcv3coKPHtbcMojyyPQDdPweGF
# RInECUzF1KVDL3SV9274eCBYLBNdYJWaPk8zhNqwiBfenk70lrC8RqBsmNLg1oiM
# CwIDAQABo4IB7TCCAekwEAYJKwYBBAGCNxUBBAMCAQAwHQYDVR0OBBYEFEhuZOVQ
# BdOCqhc3NyK1bajKdQKVMBkGCSsGAQQBgjcUAgQMHgoAUwB1AGIAQwBBMAsGA1Ud
# DwQEAwIBhjAPBgNVHRMBAf8EBTADAQH/MB8GA1UdIwQYMBaAFHItOgIxkEO5FAVO
# 4eqnxzHRI4k0MFoGA1UdHwRTMFEwT6BNoEuGSWh0dHA6Ly9jcmwubWljcm9zb2Z0
# LmNvbS9wa2kvY3JsL3Byb2R1Y3RzL01pY1Jvb0NlckF1dDIwMTFfMjAxMV8wM18y
# Mi5jcmwwXgYIKwYBBQUHAQEEUjBQME4GCCsGAQUFBzAChkJodHRwOi8vd3d3Lm1p
# Y3Jvc29mdC5jb20vcGtpL2NlcnRzL01pY1Jvb0NlckF1dDIwMTFfMjAxMV8wM18y
# Mi5jcnQwgZ8GA1UdIASBlzCBlDCBkQYJKwYBBAGCNy4DMIGDMD8GCCsGAQUFBwIB
# FjNodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpb3BzL2RvY3MvcHJpbWFyeWNw
# cy5odG0wQAYIKwYBBQUHAgIwNB4yIB0ATABlAGcAYQBsAF8AcABvAGwAaQBjAHkA
# XwBzAHQAYQB0AGUAbQBlAG4AdAAuIB0wDQYJKoZIhvcNAQELBQADggIBAGfyhqWY
# 4FR5Gi7T2HRnIpsLlhHhY5KZQpZ90nkMkMFlXy4sPvjDctFtg/6+P+gKyju/R6mj
# 82nbY78iNaWXXWWEkH2LRlBV2AySfNIaSxzzPEKLUtCw/WvjPgcuKZvmPRul1LUd
# d5Q54ulkyUQ9eHoj8xN9ppB0g430yyYCRirCihC7pKkFDJvtaPpoLpWgKj8qa1hJ
# Yx8JaW5amJbkg/TAj/NGK978O9C9Ne9uJa7lryft0N3zDq+ZKJeYTQ49C/IIidYf
# wzIY4vDFLc5bnrRJOQrGCsLGra7lstnbFYhRRVg4MnEnGn+x9Cf43iw6IGmYslmJ
# aG5vp7d0w0AFBqYBKig+gj8TTWYLwLNN9eGPfxxvFX1Fp3blQCplo8NdUmKGwx1j
# NpeG39rz+PIWoZon4c2ll9DuXWNB41sHnIc+BncG0QaxdR8UvmFhtfDcxhsEvt9B
# xw4o7t5lL+yX9qFcltgA1qFGvVnzl6UJS0gQmYAf0AApxbGbpT9Fdx41xtKiop96
# eiL6SJUfq/tHI4D1nvi/a7dLl+LrdXga7Oo3mXkYS//WsyNodeav+vyL6wuA6mk7
# r/ww7QRMjt/fdW1jkT3RnVZOT7+AVyKheBEyIXrvQQqxP/uozKRdwaGIm1dxVk5I
# RcBCyZt2WwqASGv9eZ/BvW1taslScxMNelDNMYIZjDCCGYgCAQEwgZUwfjELMAkG
# A1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQx
# HjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEoMCYGA1UEAxMfTWljcm9z
# b2Z0IENvZGUgU2lnbmluZyBQQ0EgMjAxMQITMwAAAlKLM6r4lfM52wAAAAACUjAN
# BglghkgBZQMEAgEFAKCBrjAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgor
# BgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAvBgkqhkiG9w0BCQQxIgQgk8zicYeu
# HJf9HpNc3N8m/G0670VN+JxVlXGCRi21OnUwQgYKKwYBBAGCNwIBDDE0MDKgFIAS
# AE0AaQBjAHIAbwBzAG8AZgB0oRqAGGh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbTAN
# BgkqhkiG9w0BAQEFAASCAQCgu37okvIYsUWE5nb2D2+qlphboUtGrvX0KYa99H+F
# qv3T4q3ndRKhkz6GYdtXLgWg/yTgGtfsy+5ysuUn9hNu0/VUe8s+vYCJttIDuFvK
# 0d5IfSJOdazsFflS/xSNB27/aflvO2r3qU5iExaeElr/N1gcvCoBcoBPSEw541Yd
# Skek90EuuCUsyiFTmIGz7KFwqfnYuX9/98be/lQpLJsVjQQ5pnA7UItjTnThbPsX
# vJHVUfwS73vVhyhVCNVvppDwadbY4JpM3hT2Ccp87QCysr2bl8+zFqSIA3JbqIii
# O7OoecnjV9MM5nW+MMiEhp5awQxpMzkSEpAErJmMrUS+oYIXFjCCFxIGCisGAQQB
# gjcDAwExghcCMIIW/gYJKoZIhvcNAQcCoIIW7zCCFusCAQMxDzANBglghkgBZQME
# AgEFADCCAVkGCyqGSIb3DQEJEAEEoIIBSASCAUQwggFAAgEBBgorBgEEAYRZCgMB
# MDEwDQYJYIZIAWUDBAIBBQAEIJm8Ktd/69eeLZcHP84BQNcLwRdsc55HA6XAjOJ2
# PoykAgZibEP3/dUYEzIwMjIwNTEyMDIyMDM1LjE5OFowBIACAfSggdikgdUwgdIx
# CzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRt
# b25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xLTArBgNVBAsTJE1p
# Y3Jvc29mdCBJcmVsYW5kIE9wZXJhdGlvbnMgTGltaXRlZDEmMCQGA1UECxMdVGhh
# bGVzIFRTUyBFU046MkFENC00QjkyLUZBMDExJTAjBgNVBAMTHE1pY3Jvc29mdCBU
# aW1lLVN0YW1wIFNlcnZpY2WgghFlMIIHFDCCBPygAwIBAgITMwAAAYZ45RmJ+CRL
# zAABAAABhjANBgkqhkiG9w0BAQsFADB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMK
# V2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0
# IENvcnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0Eg
# MjAxMDAeFw0yMTEwMjgxOTI3MzlaFw0yMzAxMjYxOTI3MzlaMIHSMQswCQYDVQQG
# EwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwG
# A1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMS0wKwYDVQQLEyRNaWNyb3NvZnQg
# SXJlbGFuZCBPcGVyYXRpb25zIExpbWl0ZWQxJjAkBgNVBAsTHVRoYWxlcyBUU1Mg
# RVNOOjJBRDQtNEI5Mi1GQTAxMSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFt
# cCBTZXJ2aWNlMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAwI3G2Wpv
# 6B4IjAfrgfJpndPOPYO1Yd8+vlfoIxMW3gdCDT+zIbafg14pOu0t0ekUQx60p7Pa
# dH4OjnqNIE1q6ldH9ntj1gIdl4Hq4rdEHTZ6JFdE24DSbVoqqR+R4Iw4w3GPbfc2
# Q3kfyyFyj+DOhmCWw/FZiTVTlT4bdejyAW6r/Jn4fr3xLjbvhITatr36VyyzgQ0Y
# 4Wr73H3gUcLjYu0qiHutDDb6+p+yDBGmKFznOW8wVt7D+u2VEJoE6JlK0EpVLZus
# dSzhecuUwJXxb2uygAZXlsa/fHlwW9YnlBqMHJ+im9HuK5X4x8/5B5dkuIoX5lWG
# jFMbD2A6Lu/PmUB4hK0CF5G1YaUtBrME73DAKkypk7SEm3BlJXwY/GrVoXWYUGEH
# yfrkLkws0RoEMpoIEgebZNKqjRynRJgR4fPCKrEhwEiTTAc4DXGci4HHOm64EQ1g
# /SDHMFqIKVSxoUbkGbdKNKHhmahuIrAy4we9s7rZJskveZYZiDmtAtBt/gQojxbZ
# 1vO9C11SthkrmkkTMLQf9cDzlVEBeu6KmHX2Sze6ggne3I4cy/5IULnHZ3rM4ZpJ
# c0s2KpGLHaVrEQy4x/mAn4yaYfgeH3MEAWkVjy/qTDh6cDCF/gyz3TaQDtvFnAK7
# 0LqtbEvBPdBpeCG/hk9l0laYzwiyyGY/HqMCAwEAAaOCATYwggEyMB0GA1UdDgQW
# BBQZtqNFA+9mdEu/h33UhHMN6whcLjAfBgNVHSMEGDAWgBSfpxVdAF5iXYP05dJl
# pxtTNRnpcjBfBgNVHR8EWDBWMFSgUqBQhk5odHRwOi8vd3d3Lm1pY3Jvc29mdC5j
# b20vcGtpb3BzL2NybC9NaWNyb3NvZnQlMjBUaW1lLVN0YW1wJTIwUENBJTIwMjAx
# MCgxKS5jcmwwbAYIKwYBBQUHAQEEYDBeMFwGCCsGAQUFBzAChlBodHRwOi8vd3d3
# Lm1pY3Jvc29mdC5jb20vcGtpb3BzL2NlcnRzL01pY3Jvc29mdCUyMFRpbWUtU3Rh
# bXAlMjBQQ0ElMjAyMDEwKDEpLmNydDAMBgNVHRMBAf8EAjAAMBMGA1UdJQQMMAoG
# CCsGAQUFBwMIMA0GCSqGSIb3DQEBCwUAA4ICAQDD7mehJY3fTHKC4hj+wBWB8544
# uaJiMMIHnhK9ONTM7VraTYzx0U/TcLJ6gxw1tRzM5uu8kswJNlHNp7RedsAiwviV
# QZV9AL8IbZRLJTwNehCwk+BVcY2gh3ZGZmx8uatPZrRueyhhTTD2PvFVLrfwh2li
# DG/dEPNIHTKj79DlEcPIWoOCUp7p0ORMwQ95kVaibpX89pvjhPl2Fm0CBO3pXXJg
# 0bydpQ5dDDTv/qb0+WYF/vNVEU/MoMEQqlUWWuXECTqx6TayJuLJ6uU7K5QyTkQ/
# l24IhGjDzf5AEZOrINYzkWVyNfUOpIxnKsWTBN2ijpZ/Tun5qrmo9vNIDT0lobgn
# ulae17NaEO9oiEJJH1tQ353dhuRi+A00PR781iYlzF5JU1DrEfEyNx8CWgERi90L
# KsYghZBCDjQ3DiJjfUZLqONeHrJfcmhz5/bfm8+aAaUPpZFeP0g0Iond6XNk4YiY
# bWPFoofc0LwcqSALtuIAyz6f3d+UaZZsp41U4hCIoGj6hoDIuU839bo/mZ/AgESw
# GxIXs0gZU6A+2qIUe60QdA969wWSzucKOisng9HCSZLF1dqc3QUawr0C0U41784K
# o9vckAG3akwYuVGcs6hM/SqEhoe9jHwe4Xp81CrTB1l9+EIdukCbP0kyzx0WZzte
# eiDN5rdiiQR9mBJuljCCB3EwggVZoAMCAQICEzMAAAAVxedrngKbSZkAAAAAABUw
# DQYJKoZIhvcNAQELBQAwgYgxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5n
# dG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9y
# YXRpb24xMjAwBgNVBAMTKU1pY3Jvc29mdCBSb290IENlcnRpZmljYXRlIEF1dGhv
# cml0eSAyMDEwMB4XDTIxMDkzMDE4MjIyNVoXDTMwMDkzMDE4MzIyNVowfDELMAkG
# A1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQx
# HjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWljcm9z
# b2Z0IFRpbWUtU3RhbXAgUENBIDIwMTAwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAw
# ggIKAoICAQDk4aZM57RyIQt5osvXJHm9DtWC0/3unAcH0qlsTnXIyjVX9gF/bErg
# 4r25PhdgM/9cT8dm95VTcVrifkpa/rg2Z4VGIwy1jRPPdzLAEBjoYH1qUoNEt6aO
# RmsHFPPFdvWGUNzBRMhxXFExN6AKOG6N7dcP2CZTfDlhAnrEqv1yaa8dq6z2Nr41
# JmTamDu6GnszrYBbfowQHJ1S/rboYiXcag/PXfT+jlPP1uyFVk3v3byNpOORj7I5
# LFGc6XBpDco2LXCOMcg1KL3jtIckw+DJj361VI/c+gVVmG1oO5pGve2krnopN6zL
# 64NF50ZuyjLVwIYwXE8s4mKyzbnijYjklqwBSru+cakXW2dg3viSkR4dPf0gz3N9
# QZpGdc3EXzTdEonW/aUgfX782Z5F37ZyL9t9X4C626p+Nuw2TPYrbqgSUei/BQOj
# 0XOmTTd0lBw0gg/wEPK3Rxjtp+iZfD9M269ewvPV2HM9Q07BMzlMjgK8QmguEOqE
# UUbi0b1qGFphAXPKZ6Je1yh2AuIzGHLXpyDwwvoSCtdjbwzJNmSLW6CmgyFdXzB0
# kZSU2LlQ+QuJYfM2BjUYhEfb3BvR/bLUHMVr9lxSUV0S2yW6r1AFemzFER1y7435
# UsSFF5PAPBXbGjfHCBUYP3irRbb1Hode2o+eFnJpxq57t7c+auIurQIDAQABo4IB
# 3TCCAdkwEgYJKwYBBAGCNxUBBAUCAwEAATAjBgkrBgEEAYI3FQIEFgQUKqdS/mTE
# mr6CkTxGNSnPEP8vBO4wHQYDVR0OBBYEFJ+nFV0AXmJdg/Tl0mWnG1M1GelyMFwG
# A1UdIARVMFMwUQYMKwYBBAGCN0yDfQEBMEEwPwYIKwYBBQUHAgEWM2h0dHA6Ly93
# d3cubWljcm9zb2Z0LmNvbS9wa2lvcHMvRG9jcy9SZXBvc2l0b3J5Lmh0bTATBgNV
# HSUEDDAKBggrBgEFBQcDCDAZBgkrBgEEAYI3FAIEDB4KAFMAdQBiAEMAQTALBgNV
# HQ8EBAMCAYYwDwYDVR0TAQH/BAUwAwEB/zAfBgNVHSMEGDAWgBTV9lbLj+iiXGJo
# 0T2UkFvXzpoYxDBWBgNVHR8ETzBNMEugSaBHhkVodHRwOi8vY3JsLm1pY3Jvc29m
# dC5jb20vcGtpL2NybC9wcm9kdWN0cy9NaWNSb29DZXJBdXRfMjAxMC0wNi0yMy5j
# cmwwWgYIKwYBBQUHAQEETjBMMEoGCCsGAQUFBzAChj5odHRwOi8vd3d3Lm1pY3Jv
# c29mdC5jb20vcGtpL2NlcnRzL01pY1Jvb0NlckF1dF8yMDEwLTA2LTIzLmNydDAN
# BgkqhkiG9w0BAQsFAAOCAgEAnVV9/Cqt4SwfZwExJFvhnnJL/Klv6lwUtj5OR2R4
# sQaTlz0xM7U518JxNj/aZGx80HU5bbsPMeTCj/ts0aGUGCLu6WZnOlNN3Zi6th54
# 2DYunKmCVgADsAW+iehp4LoJ7nvfam++Kctu2D9IdQHZGN5tggz1bSNU5HhTdSRX
# ud2f8449xvNo32X2pFaq95W2KFUn0CS9QKC/GbYSEhFdPSfgQJY4rPf5KYnDvBew
# VIVCs/wMnosZiefwC2qBwoEZQhlSdYo2wh3DYXMuLGt7bj8sCXgU6ZGyqVvfSaN0
# DLzskYDSPeZKPmY7T7uG+jIa2Zb0j/aRAfbOxnT99kxybxCrdTDFNLB62FD+Cljd
# QDzHVG2dY3RILLFORy3BFARxv2T5JL5zbcqOCb2zAVdJVGTZc9d/HltEAY5aGZFr
# DZ+kKNxnGSgkujhLmm77IVRrakURR6nxt67I6IleT53S0Ex2tVdUCbFpAUR+fKFh
# bHP+CrvsQWY9af3LwUFJfn6Tvsv4O+S3Fb+0zj6lMVGEvL8CwYKiexcdFYmNcP7n
# tdAoGokLjzbaukz5m/8K6TT4JDVnK+ANuOaMmdbhIurwJ0I9JZTmdHRbatGePu1+
# oDEzfbzL6Xu/OHBE0ZDxyKs6ijoIYn/ZcGNTTY3ugm2lBRDBcQZqELQdVTNYs6Fw
# ZvKhggLUMIICPQIBATCCAQChgdikgdUwgdIxCzAJBgNVBAYTAlVTMRMwEQYDVQQI
# EwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3Nv
# ZnQgQ29ycG9yYXRpb24xLTArBgNVBAsTJE1pY3Jvc29mdCBJcmVsYW5kIE9wZXJh
# dGlvbnMgTGltaXRlZDEmMCQGA1UECxMdVGhhbGVzIFRTUyBFU046MkFENC00Qjky
# LUZBMDExJTAjBgNVBAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2WiIwoB
# ATAHBgUrDgMCGgMVAAGu2DRzWkKljmXySX1korHL4fMnoIGDMIGApH4wfDELMAkG
# A1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQx
# HjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWljcm9z
# b2Z0IFRpbWUtU3RhbXAgUENBIDIwMTAwDQYJKoZIhvcNAQEFBQACBQDmJpPWMCIY
# DzIwMjIwNTEyMDM1ODE0WhgPMjAyMjA1MTMwMzU4MTRaMHQwOgYKKwYBBAGEWQoE
# ATEsMCowCgIFAOYmk9YCAQAwBwIBAAICDFcwBwIBAAICEWIwCgIFAOYn5VYCAQAw
# NgYKKwYBBAGEWQoEAjEoMCYwDAYKKwYBBAGEWQoDAqAKMAgCAQACAwehIKEKMAgC
# AQACAwGGoDANBgkqhkiG9w0BAQUFAAOBgQCmRYAyIBtx8p/+mS6VyXH79ZFyqWwF
# Ov63UbYgNvbct4rUh0YjBTZd/vps2/TjvDOQ514nMGTKpvw0388axSuFhw1RJwXi
# 0JOBNaI7KLPEeV8TlUmtUDIXpOGf0jPZxxPu1dWX/ihLYpYExWHXn+61UDzKUWm2
# MjkW1b+8GZOU8jGCBA0wggQJAgEBMIGTMHwxCzAJBgNVBAYTAlVTMRMwEQYDVQQI
# EwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3Nv
# ZnQgQ29ycG9yYXRpb24xJjAkBgNVBAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1wIFBD
# QSAyMDEwAhMzAAABhnjlGYn4JEvMAAEAAAGGMA0GCWCGSAFlAwQCAQUAoIIBSjAa
# BgkqhkiG9w0BCQMxDQYLKoZIhvcNAQkQAQQwLwYJKoZIhvcNAQkEMSIEIIAytBOc
# HXvTH7e8BFsbUxSW4lSLr2V8YKk0aFQchaJBMIH6BgsqhkiG9w0BCRACLzGB6jCB
# 5zCB5DCBvQQgGpmI4LIsCFTGiYyfRAR7m7Fa2guxVNIw17mcAiq8Qn4wgZgwgYCk
# fjB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMH
# UmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSYwJAYDVQQD
# Ex1NaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAxMAITMwAAAYZ45RmJ+CRLzAAB
# AAABhjAiBCB5CI13uVv9lTv5LFf3EKIM7qTACgQLfipcOGpzq4q0nDANBgkqhkiG
# 9w0BAQsFAASCAgBnNy5Clv1Yq4LtLRY7plLi8glHtdg9hepoY6nc/2ucl2EivHGH
# zaGnNSKgvRhdjcB4GZne6RGmGTi+HY6KffukbVo444Fh0245d3STWBH0bfFahLtr
# LCua9gdmhXo4ulBm/jUj3xrLigJdH+rY5v/nTOzD2Shv7kLSM5IW7+aC0ZKA5EoD
# nF0lf4WDUG54WAYxdrVpMYIzaSQkev9LYO/nlCdxwtwYfDwxo38lwlWT8/wXSo/Q
# wZPNPWZegDICyPt4+dWiO3a6iM5VeobZGWc20W3O+iP0r0W97ORJTZB+sPoU4NSF
# y104vqeHcPH+vixD2J7vAfEf3MMgd81SBN/RQhUyqQK+ZI4SczgT61GDVA3PyBJx
# DEh4G4v6kGrBMwRQu0bUuI1MMWZY3n+nxUdzrrPiG1pA3h23d6iZC5hIhjsvEcCx
# TojDOv2irFjtsAUbJc1gpqz2Mpk1n0BAhbrjcyJWDFn+eFPnRazb0l67g0fWd4B9
# HBaPJ/6l9YBjAyVb0xG39sTk6PKtvY124CkUrdbtlkoLUKKxdxoQw5D7I0Z+FgDV
# TlT12DjE2gtN4HQmJJrOl+bXog0Lt0DBIuqrd8xyjV5p+hbGD+8Gl5c5Dzkp7bwn
# I/ZXHTiuCAW1lMasK0TbSX3BtfNv8pAg8H+QD1iEJjK2ZLFytBKIkZq6sg==
# SIG # End signature block
