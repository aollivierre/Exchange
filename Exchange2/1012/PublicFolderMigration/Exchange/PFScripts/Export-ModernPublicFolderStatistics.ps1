# .SYNOPSIS
# Export-ModernPublicFolderStatistics.ps1
#    Generates a CSV file that contains the list of public folders and their individual sizes
#
# .DESCRIPTION
#
# Copyright (c) 2016 Microsoft Corporation. All rights reserved.
#
# THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK
# OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

Param(
    # File to export to
    [Parameter(
        Mandatory=$true,
        HelpMessage = "Full path of the output file to be generated. If only filename is specified, then the output file will be generated in the current directory.")]
    [ValidateNotNull()]
    [string] $ExportFile
    )

#load hashtable of localized string
Import-LocalizedData -BindingVariable ModernPublicFolderStatistics_LocalizedStrings -FileName Export-ModernPublicFolderStatistics.strings.psd1

################ START OF DEFAULTS ################

$ErrorActionPreference = 'Stop'
$WarningPreference = 'SilentlyContinue';

$script:Exchange15MajorVersion = 15;
$script:Exchange15MinorVersion = 0;
$script:Exchange15CUBuild = 1263;
$script:Exchange16MajorVersion = 15;
$script:Exchange16MinorVersion = 1;
$script:Exchange16CUBuild = 669;

################ END OF DEFAULTS #################

# Function that determines if the given NON IPM folder should be included.
function ShouldIncludeNonIpmFolder()
{
    param($PublicFolderIdentity);

    # Append a "\" to the path. Since paths in the whitelist end with a "\",
    # this is required to ensure the root gets matched correctly.
    $folderPath = $PublicFolderIdentity.TrimEnd('\') + '\';

    foreach ($includedFolder in $script:IncludedNonIpmSubtree)
    {
        if ($folderPath.StartsWith($includedFolder))
        {
            return $true;
        }
    }

    return $false;
}

# Recurse through IPM_SUBTREE to get the folder path foreach Public Folder
function ReadIpmSubtree()
{
    $badFolders = @();
    $ipmSubtreeFolderList = Get-PublicFolder "\" -Recurse -ResultSize:Unlimited;

    foreach ($folder in $ipmSubtreeFolderList)
    {
        if (IsValidFolderName $folder.Name)
        {
            $nameAndDumpsterEntryId = New-Object PSObject -Property @{FolderIdentity = $folder.Identity.ToString(); DumpsterEntryId = $folder.DumpsterEntryId}
            $script:IdToNameAndDumpsterMap.Add($folder.EntryId, $nameAndDumpsterEntryId);
        }
        else
        {
            # Mark the folder as invalid but continue so that we can find all the bad ones.
            $badFolders += $folder;
        }
    }

    # Ensure there are no folders with invalid names.
    AssertAllFolderNamesValid $badFolders;

    # Root path will have a "\" at the end while other folders doesn't. Normalize by removing it.
    $ipmSubtreeRoot = $ipmSubtreeFolderList | Where-Object { $null -eq $_.ParentPath };
    $nameAndDumpsterEntryId = New-Object PSObject -Property @{FolderIdentity = ""; DumpsterEntryId = $ipmSubtreeRoot.DumpsterEntryId}
    $script:IdToNameAndDumpsterMap[$ipmSubtreeRoot.EntryId] = $nameAndDumpsterEntryId;

    return $ipmSubtreeFolderList.Count;
}

# Recurse through NON_IPM_SUBTREE to get the folder path foreach Public Folder
function ReadNonIpmSubtree()
{
    $badFolders = @();
    $nonIpmSubtreeFolderList = Get-PublicFolder "\NON_IPM_SUBTREE" -Recurse -ResultSize:Unlimited;

    foreach ($folder in $nonIpmSubtreeFolderList)
    {
        $folderIdentity = $folder.Identity.ToString();

        if (ShouldIncludeNonIpmFolder($folderIdentity))
        {
            if (IsValidFolderName $folder.Name)
            {
                $nameAndDumpsterEntryId = New-Object PSObject -Property @{FolderIdentity = $folder.Identity.ToString(); DumpsterEntryId = ""+$folder.DumpsterEntryId};
                $script:IdToNameAndDumpsterMap.Add($folder.EntryId, $nameAndDumpsterEntryId);

                $script:NonIpmSubtreeFolders.Add($folder.EntryId, $folderIdentity);
            }
            else
            {
                # Mark the folder as invalid but continue so that we can find all the bad ones.
                $badFolders += $folder;
            }
        }
    }

    # Ensure there are no folders with invalid names.
    AssertAllFolderNamesValid $badFolders;

    # Add the root folder to the list since this wouldn't otherwise be included.
    $nonIpmSubtreeRoot = $nonIpmSubtreeFolderList | Where-Object { $null -eq $_.ParentPath };
    $nameAndDumpsterEntryId = New-Object PSObject -Property @{FolderIdentity = $nonIpmSubtreeRoot.Identity.ToString(); DumpsterEntryId = $nonIpmSubtreeRoot.DumpsterEntryId};
    $script:IdToNameAndDumpsterMap.Add($nonIpmSubtreeRoot.EntryId, $nameAndDumpsterEntryId);
    $script:NonIpmSubtreeFolders.Add($nonIpmSubtreeRoot.EntryId, $nonIpmSubtreeRoot.Identity.ToString());

    return $nonIpmSubtreeFolderList.Count;
}

# Function that executes statistics
function GatherStatistics()
{
    $index = 0;
    $PFIDENTITY_INDEX = 2;

    #Prepare and call the Get-PublicFolderStatistics cmdlet
    $publicFolderStatistics = @(Get-PublicFolderStatistics -ResultSize:Unlimited);

    #Explcitly get statistics for NON_IPM_SUBTREE since this is not included by default.
    $publicFolderStatistics += @($script:NonIpmSubtreeFolders.Values | Get-PublicFolderStatistics -ResultSize:Unlimited);

    #Fill Folder Statistics
    while ($index -lt $publicFolderStatistics.Count)
    {
        $publicFolderEntryId = $($publicFolderStatistics[$index].EntryId);
        $dumpsterEntryId = $script:IdToNameAndDumpsterMap[$publicFolderEntryId].DumpsterEntryId;
        $publicFolderIdentity = $script:IdToNameAndDumpsterMap[$publicFolderEntryId].FolderIdentity;
        
        # We have a public folder in NON_IPM_SUBTREE\DUMPSTER_ROOT
        # Check if its a normal folder or dumpster folder
        if($publicFolderIdentity.StartsWith("\NON_IPM_SUBTREE\DUMPSTER_ROOT\"))
        {
            # Continue if dumpster is not set or
            # Folder is not present in NON IPM Subtree or
            # Folder's dumpster is not present in "\NON_IPM_SUBTREE\DUMPSTER_ROOT"
            if([String]::IsNullOrEmpty($dumpsterEntryId) -or (!$script:NonIpmSubtreeFolders.ContainsKey($dumpsterEntryId)) -or
               (!$script:NonIpmSubtreeFolders[$dumpsterEntryId].StartsWith("\NON_IPM_SUBTREE\DUMPSTER_ROOT\")))
            {
                $index++;
                continue;
            }
            
            if($script:FolderStatistics.ContainsKey($dumpsterEntryId))
            {
                # We already have processed its dumpster
                # Check which is the deepest folder, deepest one is the folder and shorter one is its dumpster
                # When processing dumpster we dont want to have a deleted folder and its dumpster size taken
                # into account twice. So we always take deleted folder into account instead of its dumpster.
                $dumpsterFolderIdentity = $script:FolderStatistics[$dumpsterEntryId][$PFIDENTITY_INDEX];
                $dumpsterDepth = ([Regex]::Matches($dumpsterFolderIdentity, "\\")).Count;
                $folderDepth = ([Regex]::Matches($publicFolderIdentity, "\\")).Count;
                if($folderDepth -gt $dumpsterDepth)
                {
                    $script:FolderStatistics.Remove($dumpsterEntryId);
                }
                else
                {
                    $index++;
                    continue;
                }
            }
        }
        $newFolder = @();
        $newFolder += $($publicFolderStatistics[$index].TotalItemSize.ToBytes());
        $newFolder += $($publicFolderStatistics[$index].TotalDeletedItemSize.ToBytes())
        $newFolder += $publicFolderIdentity;
        $newFolder += $dumpsterEntryId;
        $script:FolderStatistics[$publicFolderEntryId] = $newFolder;
        $index++;
    }
}

# Writes the current progress
function WriteProgress()
{
    param($statusFormat, $statusProcessed, $statusTotal)
    Write-Progress -Activity $ModernPublicFolderStatistics_LocalizedStrings.ProgressBarActivity `
        -Status ($statusFormat -f $statusProcessed,$statusTotal) `
        -PercentComplete (100*($statusProcessed/$statusTotal));
}

# Function that creates folder objects in right way for exporting
function CreateFolderObjects()
{
    $index = 1;
    $PFSIZE_INDEX = 0;
    $PFDELETEDSIZE_INDEX = 1;
    $PFIDENTITY_INDEX = 2;

    foreach ($publicFolderEntryId in $script:FolderStatistics.Keys)
    {
        $IsNonIpmSubtreeFolder = $script:NonIpmSubtreeFolders.ContainsKey($publicFolderEntryId);
        $publicFolderIdentity = "";

        if ($IsNonIpmSubtreeFolder)
        {
            $publicFolderIdentity = $script:FolderStatistics[$publicFolderEntryId][$PFIDENTITY_INDEX];
            $dumpsterSize = $script:FolderStatistics[$publicFolderEntryId][$PFDELETEDSIZE_INDEX];
            $folderSize = $script:FolderStatistics[$publicFolderEntryId][$PFSIZE_INDEX];
        }
        else
        {
            $publicFolderIdentity = "\IPM_SUBTREE" + $script:FolderStatistics[$publicFolderEntryId][$PFIDENTITY_INDEX];
            $dumpsterSize = $script:FolderStatistics[$publicFolderEntryId][$PFDELETEDSIZE_INDEX];
            $folderSize = $script:FolderStatistics[$publicFolderEntryId][$PFSIZE_INDEX];
        }

        if ($publicFolderIdentity -ne "")
        {
            WriteProgress $ModernPublicFolderStatistics_LocalizedStrings.ProcessedFolders $index $script:FolderStatistics.Keys.Count

            # Create a folder object to be exported to a CSV
            $newFolderObject = New-Object PSObject -Property @{FolderName = $publicFolderIdentity; FolderSize = $folderSize; DeletedItemSize = $dumpsterSize}
            [void]$script:ExportFolders.Add($newFolderObject);
            $index++;
        }
    }

    WriteProgress $ModernPublicFolderStatistics_LocalizedStrings.ProcessedFolders $script:FolderStatistics.Keys.Count $script:FolderStatistics.Keys.Count
}

# Check if Exchange version of all public folder servers are greater than required CU
function AssertMinVersion()
{
    $servers = Get-ExchangeServer;
    $serversWithPf = (Get-Mailbox -PublicFolder | select ServerName | Sort-Object -Unique ServerName).ServerName.ToLower()
    $failedServers = @();

    foreach ($server in $servers)
    {
         # Check only those Exchange servers which have public folders mailboxes
        if(!$serversWithPf.Contains($server.Name.ToLower()))
        {
            continue;
        }

        $version = $server.AdminDisplayVersion;
        $hasMinE15Version = (($version.Major -eq $script:Exchange15MajorVersion) -and
            ($version.Minor -eq $script:Exchange15MinorVersion) -and
            ($version.Build -ge $script:Exchange15CUBuild));
        $hasMinE16Version = (($version.Major -eq $script:Exchange16MajorVersion) -and
            ($version.Minor -eq $script:Exchange16MinorVersion) -and
            ($version.Build -ge $script:Exchange16CUBuild));

        # If version is less than minimum version of Exchange15 or Exchange16, or if the version belongs to Exchange19 onwards, then add the server to failed list.
        if (!$hasMinE15Version -and !$hasMinE16Version -and ($version.Minor -le $script:Exchange16MinorVersion))
        {
            $failedServers += $server.Fqdn;
        }
    }

    if ($failedServers.Count -gt 0)
    {
        Write-Error ($ModernPublicFolderStatistics_LocalizedStrings.VersionErrorMessage -f ($failedServers -join "`n`t"))
        exit;
    }
}

# Validate public folders are present.
function AssertPublicFoldersPresent()
{
    [void](Get-PublicFolder -ErrorAction Stop)
}

# Validate path to the ExportFile exists.
function AssertExportFileValid()
{
    param($ExportFile);

    # Check if the path leading upto the item is valid.
    $parent = Split-Path $ExportFile -Parent;
    $parentValid = ($parent -eq "") -or (Test-Path $parent -PathType Container);

    if ($parentValid)
    {
        # In case the item already exists, it should be a file.
        $isDirectory = Test-Path $ExportFile -PathType Container;

        if (!$isDirectory)
        {
            return;
        }
    }

    Write-Error ($ModernPublicFolderStatistics_LocalizedStrings.InvalidExportFile -f $ExportFile);
    exit;
}

# Validate public folder names does not have invalid characters in it.
function IsValidFolderName()
{
    param($Name);

    return !($Name.Contains('\') -or $Name.Contains('/'));
}

# Ensure there are no folders with invalid characters, or fail otherwise.
function AssertAllFolderNamesValid()
{
    param($BadFolders);

    if ($BadFolders.Count -gt 0)
    {
        $folderList = ($BadFolders | ForEach-Object { $_.ParentPath + ' -> ' + $_.Name }) -join "`n`t";
        Write-Error ($ModernPublicFolderStatistics_LocalizedStrings.InvalidFolderNames -f $folderList);
        exit;
    }
}


####################################################################################################
# Script starts here
####################################################################################################

# Assert pre-requisites.
AssertMinVersion
AssertPublicFoldersPresent
AssertExportFileValid $ExportFile

# Array of folder objects for exporting
$script:ExportFolders = $null;

# Hash table that contains the folder list (IPM_SUBTREE via Get-PublicFolderStatistics)
$script:FolderStatistics = @{};

# Hash table that contains the folder list (NON_IPM_SUBTREE via Get-PublicFolder)
$script:NonIpmSubtreeFolders = @{};

# Hash table EntryId to Name to map FolderPath
$script:IdToNameAndDumpsterMap = @{};

# Folders from NON_IPM_SUBTREE that are to be included while computing statistics
$script:IncludedNonIpmSubtree = @("\NON_IPM_SUBTREE\EFORMS REGISTRY", "\NON_IPM_SUBTREE\DUMPSTER_ROOT");


# Just making sure that all the paths in the whitelist have a trailing '\'.
# This will be of significance later on when the filtering happens.
$script:IncludedNonIpmSubtree = @($script:IncludedNonIpmSubtree | ForEach-Object { $_.TrimEnd('\') + '\' })

# Recurse through IPM_SUBTREE to get the folder path for each Public Folder
# Remarks:
# This is done so we can overcome a limitation of Get-PublicFolderStatistics
# where it fails to display Unicode chars in the FolderPath value, but
# Get-PublicFolder properly renders these characters
Write-Host "[$($(Get-Date).ToString())]" $ModernPublicFolderStatistics_LocalizedStrings.ProcessingIpmSubtree;
$folderCount = ReadIpmSubtree;
Write-Host "[$($(Get-Date).ToString())]" ($ModernPublicFolderStatistics_LocalizedStrings.ProcessingIpmSubtreeComplete -f $folderCount);

# Recurse through NON_IPM_SUBTREE to get the folder path for each Public Folder
Write-Host "[$($(Get-Date).ToString())]" $ModernPublicFolderStatistics_LocalizedStrings.ProcessingNonIpmSubtree;
$folderCount = ReadNonIpmSubtree;
Write-Host "[$($(Get-Date).ToString())]" ($ModernPublicFolderStatistics_LocalizedStrings.ProcessingNonIpmSubtreeComplete -f $folderCount);

# Gathering statistics
Write-Host "[$($(Get-Date).ToString())]" ($ModernPublicFolderStatistics_LocalizedStrings.RetrievingStatistics);
GatherStatistics;
Write-Host "[$($(Get-Date).ToString())]" ($ModernPublicFolderStatistics_LocalizedStrings.RetrievingStatisticsComplete -f $script:FolderStatistics.Count);

# Creating folder objects for exporting to a CSV
Write-Host "[$($(Get-Date).ToString())]" $ModernPublicFolderStatistics_LocalizedStrings.ExportToCSV;
$script:ExportFolders = New-Object System.Collections.ArrayList -ArgumentList ($script:FolderStatistics.Count);
CreateFolderObjects;

# Export the folders to CSV file
$script:ExportFolders | Sort-Object -Property FolderName | Select FolderSize, DeletedItemSize, FolderName | Export-CSV -Path $ExportFile -Force -NoTypeInformation -Encoding "Unicode";

# SIG # Begin signature block
# MIIdvwYJKoZIhvcNAQcCoIIdsDCCHawCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUamUMCmI6V+IFPsBWhe/1szv8
# ucqgghhUMIIEwjCCA6qgAwIBAgITMwAAAMDeLD0HlORJeQAAAAAAwDANBgkqhkiG
# 9w0BAQUFADB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
# A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSEw
# HwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EwHhcNMTYwOTA3MTc1ODUw
# WhcNMTgwOTA3MTc1ODUwWjCBsjELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjEMMAoGA1UECxMDQU9DMScwJQYDVQQLEx5uQ2lwaGVyIERTRSBFU046
# N0FCNS0yREYyLURBM0YxJTAjBgNVBAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNl
# cnZpY2UwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQDoiKVSfklVAB4E
# Oc9+r95kl32muXNITYcTbaRtuJl+MQzEnD0eU2JUXx2mI06ONnTfFW39ZQPF1pvU
# WkHBrS6m8oKy7Em4Ol91RJ5Knwy1VvY2Tawqh+VxwdARRgOeFtFm0S+Pa+BrXtVU
# hTtGl0BGMsKGEQKdDNGJD259Iq47qPLw3CmllE3/YFw1GGoJ9C3ry+I7ntxIjJYB
# LXA122vw93OOD/zWFh1SVq2AejPxcjKtHH2hjoeTKwkFeMNtIekrUSvhbuCGxW5r
# 54KW0Yus4o8392l9Vz8lSEn2j/TgPTqD6EZlzkpw54VSwede/vyqgZIrRbat0bAh
# b8doY8vDAgMBAAGjggEJMIIBBTAdBgNVHQ4EFgQUFf5K2jOJ0xmF1WRZxNxTQRBP
# tzUwHwYDVR0jBBgwFoAUIzT42VJGcArtQPt2+7MrsMM1sw8wVAYDVR0fBE0wSzBJ
# oEegRYZDaHR0cDovL2NybC5taWNyb3NvZnQuY29tL3BraS9jcmwvcHJvZHVjdHMv
# TWljcm9zb2Z0VGltZVN0YW1wUENBLmNybDBYBggrBgEFBQcBAQRMMEowSAYIKwYB
# BQUHMAKGPGh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2kvY2VydHMvTWljcm9z
# b2Z0VGltZVN0YW1wUENBLmNydDATBgNVHSUEDDAKBggrBgEFBQcDCDANBgkqhkiG
# 9w0BAQUFAAOCAQEAGeJAuzJMR+kovMi8RK/LtfrKazWlR5Lx02hM9GFmMk1zWCSc
# pfVY6xqfzWFllCFHBtOaJZqLiV97jfNCLpG0PULz24CWSkG7jJ+mZaCSicZ7ZC3b
# WDh1zpc54llYVyyTkRVYx/mtc9GujqbS8CBZgjaT/JsECnvGAPUcLYuSGt53CU1b
# UuiNwuzAhai4glcYyq3/7qMmmAtbnbCZhR5ySoMy7BwdzN70drLtafCJQncfAHXV
# O5r6SX4U/2J2zvWhA8lqhZu9SRulFGRvf81VTf+k5rJ2TjL6dYtSchooJ5YVvUk6
# i7bfV0VBN8xpaUhk8jbBnxhDPKIvDvnZlhPuJjCCBgEwggPpoAMCAQICEzMAAADE
# 6Yn4eoFQ6f8AAAAAAMQwDQYJKoZIhvcNAQELBQAwfjELMAkGA1UEBhMCVVMxEzAR
# BgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1p
# Y3Jvc29mdCBDb3Jwb3JhdGlvbjEoMCYGA1UEAxMfTWljcm9zb2Z0IENvZGUgU2ln
# bmluZyBQQ0EgMjAxMTAeFw0xNzA4MTEyMDIwMjRaFw0xODA4MTEyMDIwMjRaMHQx
# CzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRt
# b25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xHjAcBgNVBAMTFU1p
# Y3Jvc29mdCBDb3Jwb3JhdGlvbjCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoC
# ggEBAIiKuCTDB4+agHkV/CZg/HKILPr0o5eIlka3o8tfiS86My4ekXj6fKkfggG1
# essavAPKRuvFmff7BB3yhQr/Im6h8mc9xScY5Sgf9QSUQWPs47oVjO0TmjXeOHBU
# bzvsrUUJMEnBvo8wmQzLdsn3c5UWd9GLu5THCIUg7R6oNfFxwuB0AEuK0tyR69Z4
# /o36rWCIPb25H65il7/FhLGQrtavK9NU+zXazXGS5h7/7HFry38IdnTgEFFI1PEA
# yEhMowc15VkN/XycyOZa44X11poPH46m5IQXwdbKnx0Bx/1IpxOSM5chSDL4wiSi
# ALK+U8qDbilbge84boDzu+wTC+sCAwEAAaOCAYAwggF8MB8GA1UdJQQYMBYGCisG
# AQQBgjdMCAEGCCsGAQUFBwMDMB0GA1UdDgQWBBTL1mKEz2A56v9nwlzSyLurt8MT
# mDBSBgNVHREESzBJpEcwRTENMAsGA1UECxMETU9QUjE0MDIGA1UEBRMrMjMwMDEy
# K2M4MDRiNWVhLTQ5YjQtNDIzOC04MzYyLWQ4NTFmYTIyNTRmYzAfBgNVHSMEGDAW
# gBRIbmTlUAXTgqoXNzcitW2oynUClTBUBgNVHR8ETTBLMEmgR6BFhkNodHRwOi8v
# d3d3Lm1pY3Jvc29mdC5jb20vcGtpb3BzL2NybC9NaWNDb2RTaWdQQ0EyMDExXzIw
# MTEtMDctMDguY3JsMGEGCCsGAQUFBwEBBFUwUzBRBggrBgEFBQcwAoZFaHR0cDov
# L3d3dy5taWNyb3NvZnQuY29tL3BraW9wcy9jZXJ0cy9NaWNDb2RTaWdQQ0EyMDEx
# XzIwMTEtMDctMDguY3J0MAwGA1UdEwEB/wQCMAAwDQYJKoZIhvcNAQELBQADggIB
# AAYWH9tXwlDII0+iUXjX7fj9zb3VwPH5G1btU8hpRwXVxMvs4vyZW5VfETgowAVF
# E+CaeYi8Zqvbu+sCVSO3PSN4QW2u+PEAWpSZihzMCZXQmhxEMKmlFse6R1v1KzSL
# n49YN8NOHK8iyhDN2IIQqTXwriLIjySmgYvfJxzkZh2JPi7/VwNNwW6DoDLrtLMv
# UFZdBrEVjMgdY7dzDOPWeiYPKpZFpzKDPpY+V0l3I4n+sRDHiuUIFVHFK1oxWzlq
# lqikiGuWKG/xxK7qvUUXzGJOgbVUGkeOmKVtwG4nxvgnH8jtIKkLsfHOC5qU4mqd
# aYOhNtdtIP6F1f/DuJc2Cf49FMGYFKnAhszvgsGrVSRDGLVIhXiG0PnSnT8Z2RSJ
# 542faCSIaDupx4BOJucIIUxj/ZyTFU0ztVZgT9dKuTiO/y7dsV+kQ2vJeM+xu2uP
# g2yHcqrqpfuf3RrWOfxkyW0+COV8g7GtvKO6e8+WVqR6WMsSR2LSIe/8PMQxC/cv
# PmSlN29gUD+3RJBPoAuLvn5Y9sdnh2HbnpjEyIzLb0fhwC6U7bH2sDBt7GpJqOmW
# dsi9CMT+O/WuczcGslbPGdS79ZTKhxzygGoBT7YbgXOz01siPzpYGN+I7mfESacv
# 3CWLPV7Q7DREkR28kQx2gj7vxNgtoQQCjkj5790CzwOiMIIGBzCCA++gAwIBAgIK
# YRZoNAAAAAAAHDANBgkqhkiG9w0BAQUFADBfMRMwEQYKCZImiZPyLGQBGRYDY29t
# MRkwFwYKCZImiZPyLGQBGRYJbWljcm9zb2Z0MS0wKwYDVQQDEyRNaWNyb3NvZnQg
# Um9vdCBDZXJ0aWZpY2F0ZSBBdXRob3JpdHkwHhcNMDcwNDAzMTI1MzA5WhcNMjEw
# NDAzMTMwMzA5WjB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQ
# MA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9u
# MSEwHwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EwggEiMA0GCSqGSIb3
# DQEBAQUAA4IBDwAwggEKAoIBAQCfoWyx39tIkip8ay4Z4b3i48WZUSNQrc7dGE4k
# D+7Rp9FMrXQwIBHrB9VUlRVJlBtCkq6YXDAm2gBr6Hu97IkHD/cOBJjwicwfyzMk
# h53y9GccLPx754gd6udOo6HBI1PKjfpFzwnQXq/QsEIEovmmbJNn1yjcRlOwhtDl
# KEYuJ6yGT1VSDOQDLPtqkJAwbofzWTCd+n7Wl7PoIZd++NIT8wi3U21StEWQn0gA
# SkdmEScpZqiX5NMGgUqi+YSnEUcUCYKfhO1VeP4Bmh1QCIUAEDBG7bfeI0a7xC1U
# n68eeEExd8yb3zuDk6FhArUdDbH895uyAc4iS1T/+QXDwiALAgMBAAGjggGrMIIB
# pzAPBgNVHRMBAf8EBTADAQH/MB0GA1UdDgQWBBQjNPjZUkZwCu1A+3b7syuwwzWz
# DzALBgNVHQ8EBAMCAYYwEAYJKwYBBAGCNxUBBAMCAQAwgZgGA1UdIwSBkDCBjYAU
# DqyCYEBWJ5flJRP8KuEKU5VZ5KShY6RhMF8xEzARBgoJkiaJk/IsZAEZFgNjb20x
# GTAXBgoJkiaJk/IsZAEZFgltaWNyb3NvZnQxLTArBgNVBAMTJE1pY3Jvc29mdCBS
# b290IENlcnRpZmljYXRlIEF1dGhvcml0eYIQea0WoUqgpa1Mc1j0BxMuZTBQBgNV
# HR8ESTBHMEWgQ6BBhj9odHRwOi8vY3JsLm1pY3Jvc29mdC5jb20vcGtpL2NybC9w
# cm9kdWN0cy9taWNyb3NvZnRyb290Y2VydC5jcmwwVAYIKwYBBQUHAQEESDBGMEQG
# CCsGAQUFBzAChjhodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpL2NlcnRzL01p
# Y3Jvc29mdFJvb3RDZXJ0LmNydDATBgNVHSUEDDAKBggrBgEFBQcDCDANBgkqhkiG
# 9w0BAQUFAAOCAgEAEJeKw1wDRDbd6bStd9vOeVFNAbEudHFbbQwTq86+e4+4LtQS
# ooxtYrhXAstOIBNQmd16QOJXu69YmhzhHQGGrLt48ovQ7DsB7uK+jwoFyI1I4vBT
# Fd1Pq5Lk541q1YDB5pTyBi+FA+mRKiQicPv2/OR4mS4N9wficLwYTp2Oawpylbih
# OZxnLcVRDupiXD8WmIsgP+IHGjL5zDFKdjE9K3ILyOpwPf+FChPfwgphjvDXuBfr
# Tot/xTUrXqO/67x9C0J71FNyIe4wyrt4ZVxbARcKFA7S2hSY9Ty5ZlizLS/n+YWG
# zFFW6J1wlGysOUzU9nm/qhh6YinvopspNAZ3GmLJPR5tH4LwC8csu89Ds+X57H21
# 46SodDW4TsVxIxImdgs8UoxxWkZDFLyzs7BNZ8ifQv+AeSGAnhUwZuhCEl4ayJ4i
# IdBD6Svpu/RIzCzU2DKATCYqSCRfWupW76bemZ3KOm+9gSd0BhHudiG/m4LBJ1S2
# sWo9iaF2YbRuoROmv6pH8BJv/YoybLL+31HIjCPJZr2dHYcSZAI9La9Zj7jkIeW1
# sMpjtHhUBdRBLlCslLCleKuzoJZ1GtmShxN1Ii8yqAhuoFuMJb+g74TKIdbrHk/J
# mu5J4PcBZW+JC33Iacjmbuqnl84xKf8OxVtc2E0bodj6L54/LlUWa8kTo/0wggd6
# MIIFYqADAgECAgphDpDSAAAAAAADMA0GCSqGSIb3DQEBCwUAMIGIMQswCQYDVQQG
# EwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwG
# A1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMTIwMAYDVQQDEylNaWNyb3NvZnQg
# Um9vdCBDZXJ0aWZpY2F0ZSBBdXRob3JpdHkgMjAxMTAeFw0xMTA3MDgyMDU5MDla
# Fw0yNjA3MDgyMTA5MDlaMH4xCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5n
# dG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9y
# YXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25pbmcgUENBIDIwMTEw
# ggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQCr8PpyEBwurdhuqoIQTTS6
# 8rZYIZ9CGypr6VpQqrgGOBoESbp/wwwe3TdrxhLYC/A4wpkGsMg51QEUMULTiQ15
# ZId+lGAkbK+eSZzpaF7S35tTsgosw6/ZqSuuegmv15ZZymAaBelmdugyUiYSL+er
# CFDPs0S3XdjELgN1q2jzy23zOlyhFvRGuuA4ZKxuZDV4pqBjDy3TQJP4494HDdVc
# eaVJKecNvqATd76UPe/74ytaEB9NViiienLgEjq3SV7Y7e1DkYPZe7J7hhvZPrGM
# XeiJT4Qa8qEvWeSQOy2uM1jFtz7+MtOzAz2xsq+SOH7SnYAs9U5WkSE1JcM5bmR/
# U7qcD60ZI4TL9LoDho33X/DQUr+MlIe8wCF0JV8YKLbMJyg4JZg5SjbPfLGSrhwj
# p6lm7GEfauEoSZ1fiOIlXdMhSz5SxLVXPyQD8NF6Wy/VI+NwXQ9RRnez+ADhvKwC
# gl/bwBWzvRvUVUvnOaEP6SNJvBi4RHxF5MHDcnrgcuck379GmcXvwhxX24ON7E1J
# MKerjt/sW5+v/N2wZuLBl4F77dbtS+dJKacTKKanfWeA5opieF+yL4TXV5xcv3co
# KPHtbcMojyyPQDdPweGFRInECUzF1KVDL3SV9274eCBYLBNdYJWaPk8zhNqwiBfe
# nk70lrC8RqBsmNLg1oiMCwIDAQABo4IB7TCCAekwEAYJKwYBBAGCNxUBBAMCAQAw
# HQYDVR0OBBYEFEhuZOVQBdOCqhc3NyK1bajKdQKVMBkGCSsGAQQBgjcUAgQMHgoA
# UwB1AGIAQwBBMAsGA1UdDwQEAwIBhjAPBgNVHRMBAf8EBTADAQH/MB8GA1UdIwQY
# MBaAFHItOgIxkEO5FAVO4eqnxzHRI4k0MFoGA1UdHwRTMFEwT6BNoEuGSWh0dHA6
# Ly9jcmwubWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1Y3RzL01pY1Jvb0NlckF1
# dDIwMTFfMjAxMV8wM18yMi5jcmwwXgYIKwYBBQUHAQEEUjBQME4GCCsGAQUFBzAC
# hkJodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpL2NlcnRzL01pY1Jvb0NlckF1
# dDIwMTFfMjAxMV8wM18yMi5jcnQwgZ8GA1UdIASBlzCBlDCBkQYJKwYBBAGCNy4D
# MIGDMD8GCCsGAQUFBwIBFjNodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpb3Bz
# L2RvY3MvcHJpbWFyeWNwcy5odG0wQAYIKwYBBQUHAgIwNB4yIB0ATABlAGcAYQBs
# AF8AcABvAGwAaQBjAHkAXwBzAHQAYQB0AGUAbQBlAG4AdAAuIB0wDQYJKoZIhvcN
# AQELBQADggIBAGfyhqWY4FR5Gi7T2HRnIpsLlhHhY5KZQpZ90nkMkMFlXy4sPvjD
# ctFtg/6+P+gKyju/R6mj82nbY78iNaWXXWWEkH2LRlBV2AySfNIaSxzzPEKLUtCw
# /WvjPgcuKZvmPRul1LUdd5Q54ulkyUQ9eHoj8xN9ppB0g430yyYCRirCihC7pKkF
# DJvtaPpoLpWgKj8qa1hJYx8JaW5amJbkg/TAj/NGK978O9C9Ne9uJa7lryft0N3z
# Dq+ZKJeYTQ49C/IIidYfwzIY4vDFLc5bnrRJOQrGCsLGra7lstnbFYhRRVg4MnEn
# Gn+x9Cf43iw6IGmYslmJaG5vp7d0w0AFBqYBKig+gj8TTWYLwLNN9eGPfxxvFX1F
# p3blQCplo8NdUmKGwx1jNpeG39rz+PIWoZon4c2ll9DuXWNB41sHnIc+BncG0Qax
# dR8UvmFhtfDcxhsEvt9Bxw4o7t5lL+yX9qFcltgA1qFGvVnzl6UJS0gQmYAf0AAp
# xbGbpT9Fdx41xtKiop96eiL6SJUfq/tHI4D1nvi/a7dLl+LrdXga7Oo3mXkYS//W
# syNodeav+vyL6wuA6mk7r/ww7QRMjt/fdW1jkT3RnVZOT7+AVyKheBEyIXrvQQqx
# P/uozKRdwaGIm1dxVk5IRcBCyZt2WwqASGv9eZ/BvW1taslScxMNelDNMYIE1TCC
# BNECAQEwgZUwfjELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAO
# BgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEo
# MCYGA1UEAxMfTWljcm9zb2Z0IENvZGUgU2lnbmluZyBQQ0EgMjAxMQITMwAAAMTp
# ifh6gVDp/wAAAAAAxDAJBgUrDgMCGgUAoIHpMBkGCSqGSIb3DQEJAzEMBgorBgEE
# AYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMCMGCSqGSIb3DQEJ
# BDEWBBS4QF1cNZ2Pn1WQ3/vlXDK7Z3UCxjCBiAYKKwYBBAGCNwIBDDF6MHigUIBO
# AEUAeABwAG8AcgB0AC0ATQBvAGQAZQByAG4AUAB1AGIAbABpAGMARgBvAGwAZABl
# AHIAUwB0AGEAdABpAHMAdABpAGMAcwAuAHAAcwAxoSSAImh0dHA6Ly93d3cubWlj
# cm9zb2Z0LmNvbS9leGNoYW5nZSAwDQYJKoZIhvcNAQEBBQAEggEAJ/GJs5mXNhaC
# c+wBRKTBZ/rn3jZVgQyA/rofxWOifd9tK00a4gWKav4kiWP0RXaKH6UOUq7jl3d/
# XdXIR6AgZ/Iq7fUFMmKejchshHAvHjnhpP1qsDr9loDnuTgFNCaRYAmr9+iJAYhS
# sY1tM68flj09B3oamFANV4CfKt9zY3o//KnscEHSjGbm/fRFh+vciYmk24rDynh9
# ROwK1/elQoKzSH+FYbS9BzCnEwNLLhUxC2E/02WRxxDu2SgqvJZNmUAMSETHvHNd
# W1fQumsWi9sdnxDfq8Q0ezWGWLd9V9bm/r6a0fK1Fwl9zIhD8Sp+J9kM3oJ1L0PY
# reky8VjpXKGCAigwggIkBgkqhkiG9w0BCQYxggIVMIICEQIBATCBjjB3MQswCQYD
# VQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEe
# MBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSEwHwYDVQQDExhNaWNyb3Nv
# ZnQgVGltZS1TdGFtcCBQQ0ECEzMAAADA3iw9B5TkSXkAAAAAAMAwCQYFKw4DAhoF
# AKBdMBgGCSqGSIb3DQEJAzELBgkqhkiG9w0BBwEwHAYJKoZIhvcNAQkFMQ8XDTE4
# MDMwODAxMjQxN1owIwYJKoZIhvcNAQkEMRYEFPcwQkC0SkYMY68sh1yj/uDXluTQ
# MA0GCSqGSIb3DQEBBQUABIIBAOMTsQq8YgVhl8mn6tZsztu3ujSRJK2fAvwB3Vfh
# XWotwC/pV5W9ZXtjkseOEAvwBLc1evti+nWmvylBeMnM7CJItG5M4pISb6ksToCJ
# 2nCw8jK1D2q6saVR6LXRB14VLXhF+3xoRuUA3Tcfatg2xpzHapuY2PqKJKXgyhvU
# 4rx+1VW3VS5r0BYXoTBEy5NXcBnQ6EyLzcD25bQQDDi22N+wcgKZzjHaiEYS5i2i
# DG9T8gSfShzhiT9vmeNoS7SExGH+wmyCOVn5C2QlpSrhBtFQ3/g7lzSRovH/uw0Y
# RmktbFpmjMCXJZsuju9dS5KU5UGYQVMYDW979wHnxwEk5+Q=
# SIG # End signature block
