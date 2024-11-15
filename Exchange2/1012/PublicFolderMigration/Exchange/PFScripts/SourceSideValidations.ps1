<#
    MIT License

    Copyright (c) Microsoft Corporation.

    Permission is hereby granted, free of charge, to any person obtaining a copy
    of this software and associated documentation files (the "Software"), to deal
    in the Software without restriction, including without limitation the rights
    to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    copies of the Software, and to permit persons to whom the Software is
    furnished to do so, subject to the following conditions:

    The above copyright notice and this permission notice shall be included in all
    copies or substantial portions of the Software.

    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
    SOFTWARE
#>

# Version 23.04.21.1447

[CmdletBinding(DefaultParameterSetName = "Default", SupportsShouldProcess)]
param (
    [Parameter(Mandatory = $false, ParameterSetName = "Default")]
    [bool]
    $StartFresh = $true,

    [Parameter(Mandatory = $false, ParameterSetName = "Default")]
    [switch]
    $SlowTraversal,

    [Parameter(Mandatory = $true, ParameterSetName = "RemoveInvalidPermissions")]
    [switch]
    $RemoveInvalidPermissions,

    [Parameter(Mandatory = $true, ParameterSetName = "SummarizePreviousResults")]
    [Switch]
    $SummarizePreviousResults,

    [Parameter(ParameterSetName = "Default")]
    [Parameter(ParameterSetName = "RemoveInvalidPermissions")]
    [Parameter(ParameterSetName = "SummarizePreviousResults")]
    [string]
    $ResultsFile = (Join-Path $PSScriptRoot "ValidationResults.csv"),

    [Parameter()]
    [switch]
    $SkipVersionCheck,

    [Parameter(Mandatory = $false, ParameterSetName = "Default")]
    [ValidateSet("Dumpsters", "Limits", "Names", "MailEnabled", "Permissions")]
    [string[]]
    $Tests = @("Dumpsters", "Limits", "Names", "MailEnabled", "Permissions")
)




function New-TestResult {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'No state change.')]
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [string]
        $TestName,

        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [string]
        $ResultType,

        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [ValidateSet("Information", "Warning", "Error")]
        [string]
        $Severity,

        [Parameter(ValueFromPipelineByPropertyName = $true)]
        [string]
        $FolderIdentity,

        [Parameter(ValueFromPipelineByPropertyName = $true)]
        [string]
        $FolderEntryId,

        [Parameter(ValueFromPipelineByPropertyName = $true)]
        [string]
        $ResultData
    )

    process {
        [PSCustomObject]@{
            TestName       = $TestName
            ResultType     = $ResultType
            Severity       = $Severity
            FolderIdentity = $FolderIdentity
            FolderEntryId  = $FolderEntryId
            ResultData     = $ResultData
        }
    }
}

function Test-DumpsterMapping {
    [CmdletBinding()]
    [OutputType([System.Object[]])]
    param (
        [Parameter()]
        [PSCustomObject]
        $FolderData
    )

    begin {
        function Test-DumpsterValid {
            [CmdletBinding()]
            [OutputType([bool])]
            param (
                [Parameter()]
                [PSCustomObject]
                $Folder,

                [Parameter()]
                [PSCustomObject]
                $FolderData
            )

            begin {
                $valid = $true
            }

            process {
                $dumpster = $FolderData.NonIpmEntryIdDictionary[$Folder.DumpsterEntryId]

                if ($null -eq $dumpster -or
                    (-not $dumpster.Identity.StartsWith("\NON_IPM_SUBTREE\DUMPSTER_ROOT", "OrdinalIgnoreCase")) -or
                    $dumpster.DumpsterEntryId -ne $Folder.EntryId) {

                    $valid = $false
                }
            }

            end {
                return $valid
            }
        }

        function NewTestDumpsterMappingResult {
            [CmdletBinding()]
            param (
                [Parameter(Position = 0)]
                [object]
                $Folder
            )

            process {
                $params = @{
                    TestName       = "DumpsterMapping"
                    ResultType     = "BadDumpsterMapping"
                    Severity       = "Error"
                    FolderIdentity = $Folder.Identity
                    FolderEntryId  = $Folder.EntryId
                }

                New-TestResult @params
            }
        }

        $startTime = Get-Date
        $progressCount = 0
        $sw = New-Object System.Diagnostics.Stopwatch
        $sw.Start()
        $progressParams = @{
            Activity = "Checking dumpster mappings"
            Id       = 2
            ParentId = 1
        }
    }

    process {
        $FolderData.IpmSubtree | ForEach-Object {
            $progressCount++
            if ($sw.ElapsedMilliseconds -gt 1000) {
                $sw.Restart()
                Write-Progress @progressParams -Status $progressCount -PercentComplete ($progressCount * 100 / $FolderData.IpmSubtree.Count)
            }

            if (-not (Test-DumpsterValid $_ $FolderData)) {
                NewTestDumpsterMappingResult $_
            }
        }

        Write-Progress @progressParams -Status "Checking EFORMS dumpster mappings"

        $FolderData.NonIpmSubtree | Where-Object { $_.Identity -like "\NON_IPM_SUBTREE\EFORMS REGISTRY\*" } | ForEach-Object {
            if (-not (Test-DumpsterValid $_ $FolderData)) {
                NewTestDumpsterMappingResult $_
            }
        }
    }

    end {
        Write-Progress @progressParams -Completed

        $params = @{
            TestName   = "DumpsterMapping"
            ResultType = "Duration"
            Severity   = "Information"
            ResultData = ((Get-Date) - $startTime)
        }

        New-TestResult @params
    }
}


function Get-ResultSummary {
    [CmdletBinding()]
    param (
        [string]
        $ResultType = $(throw "ResultType is mandatory"),

        [ValidateSet("Information", "Warning", "Error")]
        [string]
        $Severity = $(throw "Severity is mandatory"),

        [int]
        $Count = $(throw "Count is mandatory"),

        [string]
        $Action = $(throw "Action is mandatory")
    )

    process {
        [PSCustomObject]@{
            ResultType = $ResultType
            Severity   = $Severity
            Count      = $Count
            Action     = $Action
        }
    }
}

function Write-TestDumpsterMappingResult {
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipeline = $true)]
        [object]
        $TestResult
    )

    begin {
        $badDumpsters = New-Object System.Collections.ArrayList
    }

    process {
        if ($TestResult.TestName -eq "DumpsterMapping" -and $TestResult.ResultType -eq "BadDumpsterMapping") {
            $badDumpsters += $TestResult
        }
    }

    end {
        if ($badDumpsters.Count -gt 0) {
            Get-ResultSummary -ResultType $badDumpsters[0].ResultType -Severity $badDumpsters[0].Severity -Count $badDumpsters.Count -Action `
                "Use the -ExcludeDumpsters switch to skip these folders during migration, or delete the folders."
        }
    }
}



function Test-FolderLimit {
    <#
    .SYNOPSIS
        Flags folders that exceed the child count limit, depth limit,
        or item limit.
    #>
    [CmdletBinding()]
    param (
        [Parameter()]
        [PSObject]
        $FolderData
    )

    begin {
        $startTime = Get-Date
        $progressCount = 0
        $sw = New-Object System.Diagnostics.Stopwatch
        $sw.Start()
        $progressParams = @{
            Activity = "Checking limits in IPM_SUBTREE"
            Id       = 2
            ParentId = 1
        }
        $testResultParams = @{
            TestName = "Limit"
            Severity = "Error"
        }
        $folderCountMigrationLimit = 250000
        $aggregateChildItemCounts = @{}
    }

    process {
        # We start from the deepest folders and work upwards so we can calculate the aggregate child
        # counts in one pass
        foreach ($folder in ($FolderData.IpmSubtree | Sort-Object FolderPathDepth -Descending)) {
            $progressCount++
            if ($sw.ElapsedMilliseconds -gt 1000) {
                $sw.Restart()
                Write-Progress @progressParams -Status $progressCount -PercentComplete ($progressCount * 100 / $FolderData.IpmSubtree.Count)
            }

            # If we failed to get statistics for some reason, assume we have content
            [int]$itemCount = 1
            [Int64]$totalItemSize = 0
            $aggregateChildItemCount = $aggregateChildItemCounts[$folder.EntryId]

            $stats = $FolderData.StatisticsDictionary[$folder.EntryId]
            if ($null -ne $stats) {
                [int]$itemCount = $stats.ItemCount
                [Int64]$totalItemSize = $stats.TotalItemSize
            } else {
                $noStatisticsResult = @{
                    TestName       = "Limit"
                    Severity       = "Warning"
                    ResultType     = "NoStatistics"
                    FolderIdentity = $folder.Identity.ToString()
                    FolderEntryId  = $folder.EntryId.ToString()
                }
                New-TestResult @noStatisticsResult
            }

            $parent = $FolderData.EntryIdDictionary[$folder.ParentEntryId]
            if ($null -ne $parent) {
                $aggregateChildItemCounts[$parent.EntryId] += $itemCount
                if ($null -ne $aggregateChildItemCount) {
                    $aggregateChildItemCounts[$parent.EntryId] += $aggregateChildItemCount
                }
            }

            if ($itemCount -lt 1 -and $aggregateChildItemCounts[$folder.EntryId] -lt 1 -and $folder.FolderPathDepth -gt 0) {
                $emptyFolderInformation = @{
                    TestName       = "Limit"
                    Severity       = "Information"
                    ResultType     = "EmptyFolder"
                    FolderIdentity = $folder.Identity.ToString()
                    FolderEntryId  = $folder.EntryId.ToString()
                }
                New-TestResult @emptyFolderInformation
            }

            if ($FolderData.ParentEntryIdCounts[$folder.EntryId] -gt 10000) {
                $testResultParams.ResultType = "ChildCount"
                $testResultParams.FolderIdentity = $folder.Identity.ToString()
                $testResultParams.FolderEntryId = $folder.EntryId.ToString()
                New-TestResult @testResultParams
            }

            if ($folder.FolderPathDepth -gt 299) {
                $testResultParams.ResultType = "FolderPathDepth"
                $testResultParams.FolderIdentity = $folder.Identity.ToString()
                $testResultParams.FolderEntryId = $folder.EntryId.ToString()
                New-TestResult @testResultParams
            }

            if ($itemCount -gt 1000000) {
                $testResultParams.ResultType = "ItemCount"
                $testResultParams.FolderIdentity = $folder.Identity.ToString()
                $testResultParams.FolderEntryId = $folder.EntryId.ToString()
                New-TestResult @testResultParams
            }

            if ($totalItemSize -gt 25000000000) {
                $testResultParams.ResultType = "TotalItemSize"
                $testResultParams.FolderIdentity = $folder.Identity.ToString()
                $testResultParams.FolderEntryId = $folder.EntryId.ToString()
                New-TestResult @testResultParams
            }
        }

        if ($folderData.IpmSubtree.Count -gt $folderCountMigrationLimit) {
            $testResultParams.ResultType = "HierarchyCount"
            $testResultParams.FolderIdentity = ""
            $testResultParams.FolderEntryId = ""
            $testResultParams.ResultData = $folderData.IpmSubtree.Count
            New-TestResult @testResultParams
        } elseif ($folderData.IpmSubtree.Count * 2 -gt $folderCountMigrationLimit) {
            $testResultParams.ResultType = "HierarchyAndDumpsterCount"
            $testResultParams.FolderIdentity = ""
            $testResultParams.FolderEntryId = ""
            $testResultParams.ResultData = $folderData.IpmSubtree.Count
            New-TestResult @testResultParams
        }

        $progressParams.Activity = "Checking limits in NON_IPM_SUBTREE"
        $progressCount = 0

        foreach ($folder in ($FolderData.NonIpmSubtree | Sort-Object FolderPathDepth -Descending)) {
            $progressCount++
            if ($sw.ElapsedMilliseconds -gt 1000) {
                $sw.Restart()
                Write-Progress @progressParams -Status $progressCount -PercentComplete ($progressCount * 100 / $FolderData.NonIpmSubtree.Count)
            }

            if ($FolderData.ParentEntryIdCounts[$folder.EntryId] -gt 10000) {
                $testResultParams.ResultType = "ChildCount"
                $testResultParams.FolderIdentity = $folder.Identity.ToString()
                $testResultParams.FolderEntryId = $folder.EntryId.ToString()
                New-TestResult @testResultParams
            }
        }
    }

    end {
        Write-Progress @progressParams -Completed

        $params = @{
            TestName       = $testResultParams.TestName
            ResultType     = "Duration"
            Severity       = "Information"
            FolderIdentity = ""
            FolderEntryId  = ""
            ResultData     = ((Get-Date) - $startTime)
        }

        New-TestResult @params
    }
}


function Write-TestFolderLimitResult {
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipeline = $true)]
        [object]
        $TestResult
    )

    begin {
        $childCountResults = New-Object System.Collections.ArrayList
        $folderPathDepthResults = New-Object System.Collections.ArrayList
        $itemCountResults = New-Object System.Collections.ArrayList
        $totalItemSizeResults = New-Object System.Collections.ArrayList
        $emptyFolderResults = New-Object System.Collections.ArrayList
        $noStatisticsResults = New-Object System.Collections.ArrayList
        $hierarchyCountResult = $null
        $hierarchyAndDumpsterCountResult = $null
        $folderCountMigrationLimit = 250000
    }

    process {
        if ($TestResult.TestName -eq "Limit") {
            switch ($TestResult.ResultType) {
                "EmptyFolder" { [void]$emptyFolderResults.Add($TestResult) }
                "ChildCount" { [void]$childCountResults.Add($TestResult) }
                "FolderPathDepth" { [void]$folderPathDepthResults.Add($TestResult) }
                "ItemCount" { [void]$itemCountResults.Add($TestResult) }
                "TotalItemSize" { [void]$totalItemSizeResults.Add($TestResult) }
                "HierarchyCount" { $hierarchyCountResult = $TestResult }
                "HierarchyAndDumpsterCount" { $hierarchyAndDumpsterCountResult = $TestResult }
                "NoStatistics" { [void]$noStatisticsResults.Add($TestResult) }
            }
        }
    }

    end {
        if ($childCountResults.Count -gt 0) {
            Get-ResultSummary -ResultType $childCountResults[0].ResultType -Severity $childCountResults[0].Severity -Count $childCountResults.Count -Action (
                "Under each of the listed folders, child folders should be relocated or deleted to reduce " +
                "the number of child folders to 10,000 or less.")
        }

        if ($folderPathDepthResults.Count -gt 0) {
            Get-ResultSummary -ResultType $folderPathDepthResults[0].ResultType -Severity $folderPathDepthResults[0].Severity -Count $folderPathDepthResults.Count -Action (
                "These folders should be relocated to reduce the path depth to 299 or less.")
        }

        if ($itemCountResults.Count -gt 0) {
            Get-ResultSummary -ResultType $itemCountResults[0].ResultType -Severity $itemCountResults[0].Severity -Count $itemCountResults.Count -Action (
                "Items should be deleted from these folders to reduce the item count in each folder to 1 million items or less.")
        }

        if ($totalItemSizeResults.Count -gt 0) {
            Get-ResultSummary -ResultType $totalItemSizeResults[0].ResultType -Severity $totalItemSizeResults[0].Severity -Count $totalItemSizeResults.Count -Action (
                "Items should be deleted from these folders until the folder size is less than 25 GB.")
        }

        if ($null -ne $hierarchyCountResult) {
            Get-ResultSummary -ResultType $hierarchyCountResult.ResultType -Severity $hierarchyCountResult.Severity -Count 1 -Action (
                "There are $($hierarchyCountResult.ResultData) public folders in the hierarchy. This exceeds " +
                "the supported migration limit of $folderCountMigrationLimit for Exchange Online. The number " +
                "of public folders must be reduced prior to migrating to Exchange Online.")
        }

        if ($null -ne $hierarchyAndDumpsterCountResult) {
            Get-ResultSummary -ResultType $hierarchyAndDumpsterCountResult.ResultType -Severity $hierarchyAndDumpsterCountResult.Severity -Count 1 -Action (
                "There are $($hierarchyAndDumpsterCountResult.ResultData) public folders in the hierarchy. Because each of these " +
                "has a dumpster folder, the total number of folders to migrate will be twice as many. " +
                "This exceeds the supported migration limit of $folderCountMigrationLimit for Exchange Online. " +
                "New-MigrationBatch can be run with the -ExcludeDumpsters switch to skip the dumpster " +
                "folders, or public folders may be deleted to reduce the number of folders.")
        }

        if ($emptyFolderResults.Count -gt 0) {
            Get-ResultSummary -ResultType $emptyFolderResults[0].ResultType -Severity $emptyFolderResults[0].Severity -Count $emptyFolderResults.Count -Action (
                "Folders contain no items and have only empty SubFolders. " +
                "These will not cause a migration issue, but they may be pruned if desired.")
        }

        if ($noStatisticsResults.Count -gt 0) {
            Get-ResultSummary -ResultType $noStatisticsResults[0].ResultType -Severity $noStatisticsResults[0].Severity -Count $noStatisticsResults.Count -Action (
                "Public folder statistics could not be retrieved for these folders. " +
                "ItemCount, TotalItemSize, and EmptyFolder tests were skipped for these folders.")
        }
    }
}



function Test-FolderName {
    [CmdletBinding()]
    param (
        [Parameter()]
        [PSObject]
        $FolderData
    )

    begin {
        $startTime = Get-Date
        $progressCount = 0
        $sw = New-Object System.Diagnostics.Stopwatch
        $sw.Start()
        $progressParams = @{
            Activity = "Checking names"
            Id       = 2
            ParentId = 1
        }
        $testResultParams = @{
            TestName   = "FolderName"
            Severity   = "Error"
            ResultType = "SpecialCharacters"
        }
    }

    process {
        $FolderData.IpmSubtree | ForEach-Object {
            $progressCount++
            if ($sw.ElapsedMilliseconds -gt 1000) {
                $sw.Restart()
                Write-Progress @progressParams -Status $progressCount -PercentComplete ($progressCount * 100 / $FolderData.IpmSubtree.Count)
            }

            if ($_.Name -match "@|/|\\") {
                $testResultParams.FolderIdentity = $_.Identity.ToString()
                $testResultParams.FolderEntryId = $_.EntryId.ToString()
                $testResultParams.ResultData = $_.Name
                New-TestResult @testResultParams
            }
        }
    }

    end {
        Write-Progress @progressParams -Completed

        $params = @{
            TestName       = $testResultParams.TestName
            ResultType     = "Duration"
            Severity       = "Information"
            FolderIdentity = ""
            FolderEntryId  = ""
            ResultData     = ((Get-Date) - $startTime)
        }

        New-TestResult @params
    }
}


function Write-TestFolderNameResult {
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipeline = $true)]
        [object]
        $TestResult
    )

    begin {
        $badNames = New-Object System.Collections.ArrayList
    }

    process {
        if ($TestResult.TestName -eq "FolderName" -and $TestResult.ResultType -eq "SpecialCharacters") {
            [void]$badNames.Add($TestResult)
        }
    }

    end {
        if ($badNames.Count -gt 0) {
            Get-ResultSummary -ResultType $badNames[0].ResultType -Severity $badNames[0].Severity -Count $badNames.Count -Action (
                "Folders have characters @, /, or \ in the folder name. " +
                "These folders should be renamed prior to migrating. The following command " +
                "can be used:`n`n" +
                "Import-Csv .\ValidationResults.csv |`n" +
                " ? ResultType -eq SpecialCharacters |`n" +
                " % {`n" +
                "  `$newName = (`$_.ResultData -replace `"@|/|\\`", `" `").Trim()`n" +
                "  Set-PublicFolder `$_.FolderEntryId -Name `$newName`n" +
                " }")
        }
    }
}



function Test-MailEnabledFolder {
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param (
        [Parameter()]
        [PSCustomObject]
        $FolderData
    )

    begin {
        function GetCommandToMergeEmailAddresses($publicFolder, $orphanedMailPublicFolder) {
            $linkedMailPublicFolder = Get-PublicFolder $publicFolder.Identity | Get-MailPublicFolder
            $emailAddressesOnGoodObject = @($linkedMailPublicFolder.EmailAddresses | Where-Object { $_.ToString().StartsWith("smtp:", "OrdinalIgnoreCase") } | ForEach-Object { $_.ToString().Substring($_.ToString().IndexOf(':') + 1) })
            $emailAddressesOnBadObject = @($orphanedMailPublicFolder.EmailAddresses | Where-Object { $_.ToString().StartsWith("smtp:", "OrdinalIgnoreCase") } | ForEach-Object { $_.ToString().Substring($_.ToString().IndexOf(':') + 1) })
            $emailAddressesToAdd = $emailAddressesOnBadObject | Where-Object { -not $emailAddressesOnGoodObject.Contains($_) }
            $emailAddressesToAdd = $emailAddressesToAdd | ForEach-Object { "`"" + $_ + "`"" }
            if ($emailAddressesToAdd.Count -gt 0) {
                $emailAddressesToAddString = [string]::Join(",", $emailAddressesToAdd)
                $command = "Get-PublicFolder `"$($publicFolder.Identity)`" | Get-MailPublicFolder | Set-MailPublicFolder -EmailAddresses @{add=$emailAddressesToAddString}"
                return $command
            } else {
                return $null
            }
        }

        function NewTestMailEnabledFolderResult {
            [CmdletBinding()]
            param (
                [Parameter(Position = 0)]
                [string]
                $Identity,

                [Parameter(Position = 1)]
                [string]
                $EntryId,

                [Parameter(Position = 2)]
                [ValidateSet("Duration", "MailEnabledSystemFolder", "MailEnabledWithNoADObject", "MailDisabledWithProxyGuid", "OrphanedMPF", "OrphanedMPFDuplicate", "OrphanedMPFDisconnected")]
                [string]
                $ResultType,

                [Parameter(Position = 3)]
                [string]
                $ResultData
            )

            $params = @{
                TestName       = "MailEnabledFolder"
                ResultType     = $ResultType
                Severity       = "Error"
                FolderIdentity = $Identity
                FolderEntryId  = $EntryId
            }

            if ($null -ne $ResultData) {
                $params.ResultData = $ResultData
            }

            New-TestResult @params
        }

        $startTime = Get-Date
        $progressCount = 0
        $sw = New-Object System.Diagnostics.Stopwatch
        $sw.Start()
        $progressParams = @{
            Activity = "Validating mail-enabled public folders"
            Id       = 2
            ParentId = 1
        }
    }

    process {
        $FolderData.NonIpmSubtree | Where-Object { $_.MailEnabled -eq $true } | ForEach-Object { NewTestMailEnabledFolderResult -Identity $_.Identity -EntryId $_.EntryId -ResultType "MailEnabledSystemFolder" }
        $ipmSubtreeMailEnabled = @($FolderData.IpmSubtree | Where-Object { $_.MailEnabled -eq $true })
        $mailDisabledWithProxyGuid = @($FolderData.IpmSubtree | Where-Object { $_.MailEnabled -ne $true -and -not [string]::IsNullOrEmpty($_.MailRecipientGuid) -and [Guid]::Empty -ne $_.MailRecipientGuid } | ForEach-Object { $_.Identity.ToString() })
        $mailDisabledWithProxyGuid | ForEach-Object {
            $params = @{
                Identity   = $_.Identity
                EntryId    = $_.EntryId
                ResultType = "MailDisabledWithProxyGuid"
            }

            NewTestMailEnabledFolderResult @params
        }

        $mailPublicFoldersLinked = New-Object 'System.Collections.Generic.Dictionary[string, object]'
        $progressParams.CurrentOperation = "Checking for missing AD objects"
        $startTimeForThisCheck = Get-Date
        for ($i = 0; $i -lt $ipmSubtreeMailEnabled.Count; $i++) {
            $progressCount++
            if ($sw.ElapsedMilliseconds -gt 1000) {
                $sw.Restart()
                $elapsed = ((Get-Date) - $startTimeForThisCheck)
                $estimatedRemaining = [TimeSpan]::FromTicks($ipmSubtreeMailEnabled.Count / $progressCount * $elapsed.Ticks - $elapsed.Ticks).ToString("hh\:mm\:ss")
                Write-Progress @progressParams -PercentComplete ($i * 100 / $ipmSubtreeMailEnabled.Count) -Status ("$i of $($ipmSubtreeMailEnabled.Count) Estimated time remaining: $estimatedRemaining")
            }
            $result = Get-MailPublicFolder $ipmSubtreeMailEnabled[$i].Identity -ErrorAction SilentlyContinue
            if ($null -eq $result) {
                $params = @{
                    Identity   = $ipmSubtreeMailEnabled[$i].Identity
                    EntryId    = $ipmSubtreeMailEnabled[$i].EntryId
                    ResultType = "MailEnabledWithNoADObject"
                }

                NewTestMailEnabledFolderResult @params
            } else {
                $guidString = $result.Guid.ToString()
                if (-not $mailPublicFoldersLinked.ContainsKey($guidString)) {
                    $mailPublicFoldersLinked.Add($guidString, $result) | Out-Null
                }
            }
        }

        $progressCount = 0
        $progressParams.CurrentOperation = "Getting all MailPublicFolder objects"
        $allMailPublicFolders = @(Get-MailPublicFolder -ResultSize Unlimited | ForEach-Object {
                $progressCount++
                if ($sw.ElapsedMilliseconds -gt 1000) {
                    $sw.Restart()
                    Write-Progress @progressParams -Status "$progressCount"
                }

                $_
            })

        $progressCount = 0
        $progressParams.CurrentOperation = "Checking for orphaned MailPublicFolders"
        $orphanedMailPublicFolders = @($allMailPublicFolders | ForEach-Object {
                $progressCount++
                if ($sw.ElapsedMilliseconds -gt 1000) {
                    $sw.Restart()
                    Write-Progress @progressParams -PercentComplete ($progressCount * 100 / $allMailPublicFolders.Count) -Status ("$progressCount of $($allMailPublicFolders.Count)")
                }

                if (!($mailPublicFoldersLinked.ContainsKey($_.Guid.ToString()))) {
                    $_
                }
            })

        $progressParams.CurrentOperation = "Building EntryId HashSets"
        Write-Progress @progressParams
        $byEntryId = New-Object 'System.Collections.Generic.Dictionary[string, object]'
        $FolderData.IpmSubtree | ForEach-Object { $byEntryId.Add($_.EntryId.ToString(), $_) }
        $byPartialEntryId = New-Object 'System.Collections.Generic.Dictionary[string, object]'
        $FolderData.IpmSubtree | ForEach-Object { $byPartialEntryId.Add($_.EntryId.ToString().Substring(44), $_) }

        $progressParams.CurrentOperation = "Checking for orphans that point to a valid folder"
        for ($i = 0; $i -lt $orphanedMailPublicFolders.Count; $i++) {
            if ($sw.ElapsedMilliseconds -gt 1000) {
                $sw.Restart()
                Write-Progress @progressParams -PercentComplete ($i * 100 / $orphanedMailPublicFolders.Count) -Status ("$i of $($orphanedMailPublicFolders.Count)")
            }

            $thisMPF = $orphanedMailPublicFolders[$i]
            $pf = $null
            if ($null -ne $thisMPF.ExternalEmailAddress -and $thisMPF.ExternalEmailAddress.ToString().StartsWith("exPf".ToLower())) {
                $partialEntryId = $thisMPF.ExternalEmailAddress.ToString().Substring(5).Replace("-", "")
                $partialEntryId += "0000"
                if ($byPartialEntryId.TryGetValue($partialEntryId, [ref]$pf)) {
                    if ($pf.MailEnabled -eq $true) {

                        $command = GetCommandToMergeEmailAddresses $pf $thisMPF

                        $params = @{
                            Identity   = $thisMPF.DistinguishedName.Replace("/", "\/")
                            EntryId    = $pf.EntryId
                            ResultType = "OrphanedMPFDuplicate"
                            ResultData = $command
                        }

                        NewTestMailEnabledFolderResult @params
                    } else {
                        $params = @{
                            Identity   = $thisMPF.DistinguishedName.Replace("/", "\/")
                            EntryId    = $pf.EntryId
                            ResultType = "OrphanedMPFDisconnected"
                        }

                        NewTestMailEnabledFolderResult @params
                    }

                    continue
                }
            }

            if ($null -ne $thisMPF.EntryId -and $byEntryId.TryGetValue($thisMPF.EntryId.ToString(), [ref]$pf)) {
                if ($pf.MailEnabled -eq $true) {

                    $command = GetCommandToMergeEmailAddresses $pf $thisMPF

                    $params = @{
                        Identity   = $thisMPF.DistinguishedName.Replace("/", "\/")
                        EntryId    = $pf.EntryId
                        ResultType = "OrphanedMPFDuplicate"
                    }

                    if ($null -ne $command) {
                        $params.ResultData = $command
                    }

                    NewTestMailEnabledFolderResult @params
                } else {
                    $params = @{
                        Identity   = $thisMPF.DistinguishedName.Replace("/", "\/")
                        EntryId    = $pf.EntryId
                        ResultType = "OrphanedMPFDisconnected"
                    }

                    NewTestMailEnabledFolderResult @params
                }
            } else {
                $params = @{
                    Identity   = $thisMPF.DistinguishedName.Replace("/", "\/")
                    EntryId    = ""
                    ResultType = "OrphanedMPF"
                }

                NewTestMailEnabledFolderResult @params
            }
        }
    }

    end {
        Write-Progress @progressParams -Completed

        $params = @{
            TestName       = "MailEnabledFolder"
            ResultType     = "Duration"
            Severity       = "Information"
            FolderIdentity = ""
            FolderEntryId  = ""
            ResultData     = ((Get-Date) - $startTime)
        }

        New-TestResult @params
    }
}


function Write-TestMailEnabledFolderResult {
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipeline = $true)]
        [object]
        $TestResult
    )

    begin {
        $mailEnabledSystemFolderResults = New-Object System.Collections.ArrayList
        $mailEnabledWithNoADObjectResults = New-Object System.Collections.ArrayList
        $mailDisabledWithProxyGuidResults = New-Object System.Collections.ArrayList
        $orphanedMPFResults = New-Object System.Collections.ArrayList
        $orphanedMPFDuplicateResults = New-Object System.Collections.ArrayList
        $orphanedMPFDisconnectedResults = New-Object System.Collections.ArrayList
    }

    process {
        if ($TestResult.TestName -eq "MailEnabledFolder") {
            switch ($TestResult.ResultType) {
                "MailEnabledSystemFolder" { [void]$mailEnabledSystemFolderResults.Add($TestResult) }
                "MailEnabledWithNoADObject" { [void]$mailEnabledWithNoADObjectResults.Add($TestResult) }
                "MailDisabledWithProxyGuid" { [void]$mailDisabledWithProxyGuidResults.Add($TestResult) }
                "OrphanedMPF" { [void]$orphanedMPFResults.Add($TestResult) }
                "OrphanedMPFDuplicate" { [void]$orphanedMPFDuplicateResults.Add($TestResult) }
                "OrphanedMPFDisconnected" { [void]$orphanedMPFDisconnectedResults.Add($TestResult) }
            }
        }
    }

    end {
        if ($mailEnabledSystemFolderResults.Count -gt 0) {
            Get-ResultSummary -ResultType $mailEnabledSystemFolderResults[0].ResultType -Severity $mailEnabledSystemFolderResults[0].Severity -Count $mailEnabledSystemFolderResults.Count -Action (
                "System folders are mail-enabled. These folders should be mail-disabled. " +
                "After confirming the accuracy of the results, you can mail-disable them with the following command:`n`n" +
                "Import-Csv .\ValidationResults.csv |`n" +
                " ? ResultType -eq MailEnabledSystemFolder |`n" +
                " % { Disable-MailPublicFolder `$_.FolderIdentity }")
        }

        if ($mailEnabledWithNoADObjectResults.Count -gt 0) {
            Get-ResultSummary -ResultType $mailEnabledWithNoADObjectResults[0].ResultType -Severity $mailEnabledWithNoADObjectResults[0].Severity -Count $mailEnabledWithNoADObjectResults.Count -Action (
                "Folders are mail-enabled, but have no AD object. These folders should be mail-disabled. " +
                "After confirming the accuracy of the results, you can mail-disable them with the following command:`n`n" +
                "Import-Csv .\ValidationResults.csv | `n" +
                " ? ResultType -eq MailEnabledWithNoADObject |`n" +
                " % { Disable-MailPublicFolder `$_.FolderIdentity }")
        }

        if ($mailDisabledWithProxyGuidResults.Count -gt 0) {
            Get-ResultSummary -ResultType $mailDisabledWithProxyGuidResults[0].ResultType -Severity $mailDisabledWithProxyGuidResults[0].Severity -Count $mailDisabledWithProxyGuidResults.Count -Action (
                "Folders are mail-disabled, but have proxy GUID values. These folders should be mail-enabled. " +
                "After confirming the accuracy of the results, you can mail-enable them with the following command:`n`n" +
                "Import-Csv .\ValidationResults.csv |`n" +
                " ? ResultType -eq MailDisabledWithProxyGuid |`n" +
                " % { Enable-MailPublicFolder `$_.FolderIdentity }")
        }

        if ($orphanedMPFResults.Count -gt 0) {
            Get-ResultSummary -ResultType $orphanedMPFResults[0].ResultType -Severity $orphanedMPFResults[0].Severity -Count $orphanedMPFResults.Count -Action (
                "Mail public folders are orphaned. They exist in Active Directory " +
                "but are not linked to any public folder. Therefore, they should be deleted. " +
                "After confirming the accuracy of the results, you can delete them manually, " +
                "or use a command like this to delete them all:`n`n" +
                "Import-Csv .\ValidationResults.csv |`n" +
                " ? ResultType -eq OrphanedMPF |`n" +
                " % {`n" +
                "  `$folder = ([ADSI](`"LDAP://`$(`$_.FolderIdentity)`"))`n" +
                "  `$parent = ([ADSI]`"`$(`$folder.Parent)`")`n" +
                "  `$parent.Children.Remove(`$folder)`n" +
                " }")
        }

        if ($orphanedMPFDuplicateResults.Count -gt 0) {
            Get-ResultSummary -ResultType $orphanedMPFDuplicateResults[0].ResultType -Severity $orphanedMPFDuplicateResults[0].Severity -Count $orphanedMPFDuplicateResults.Count -Action (
                "Mail public folders point to public folders that point to a different directory object. " +
                "These should be deleted. Their email addresses may be merged onto the linked object. " +
                "After confirming the accuracy of the results, you can delete them manually, " +
                "or use a command like this:`n`n" +
                "Import-Csv .\ValidationResults.csv |`n" +
                " ? ResultType -eq OrphanedMPFDuplicate |`n" +
                " % {`n" +
                "  `$folder = ([ADSI](`"LDAP://`$(`$_.FolderIdentity)`"))`n" +
                "  `$parent = ([ADSI]`"`$(`$folder.Parent)`")`n" +
                "  `$parent.Children.Remove(`$folder)`n" +
                " }`n`n" +
                "After these objects are deleted, the email addresses can be merged onto the linked objects:`n`n" +
                "Import-Csv .\ValidationResults.csv |`n" +
                " ? ResultType -eq OrphanedMPFDuplicate |`n" +
                " % { Invoke-Expression `$_.ResultData }")
        }

        if ($orphanedMPFDisconnectedResults.Count -gt 0) {
            Get-ResultSummary -ResultType $orphanedMPFDisconnectedResults[0].ResultType -Severity $orphanedMPFDisconnectedResults[0].Severity -Count $orphanedMPFDisconnectedResults.Count -Action (
                "Mail public folders point to public folders that are mail-disabled. " +
                "These require manual intervention. Either the directory object should be deleted, or the folder should be mail-enabled, or both. " +
                "Open the ValidationResults.csv and filter for ResultType of OrphanedMPFDisconnected to identify these folders. " +
                "The FolderIdentity provides the DN of the mail object. The FolderEntryId provides the EntryId of the folder.")
        }
    }
}



function Test-BadPermissionJob {
    [CmdletBinding()]
    param (
        [Parameter(Position = 0)]
        [string]
        $Server,

        [Parameter(Position = 1)]
        [string]
        $Mailbox,

        [Parameter(Position = 2)]
        [PSCustomObject[]]
        $Folders
    )

    begin {
        $WarningPreference = "SilentlyContinue"
        Import-PSSession (New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://$Server/powershell" -Authentication Kerberos) | Out-Null
        $startTime = Get-Date
        $progressCount = 0
        $sw = New-Object System.Diagnostics.Stopwatch
        $sw.Start()
        $progressParams = @{
            Activity = "Checking permissions in mailbox $Mailbox"
        }
    }

    process {
        $Folders | ForEach-Object {
            $progressCount++
            if ($sw.ElapsedMilliseconds -gt 1000) {
                $sw.Restart()
                $elapsed = ((Get-Date) - $startTime)
                $estimatedRemaining = [TimeSpan]::FromTicks($Folders.Count / $progressCount * $elapsed.Ticks - $elapsed.Ticks).ToString("hh\:mm\:ss")
                Write-Progress @progressParams -Status "$progressCount / $($Folders.Count) Estimated time remaining: $estimatedRemaining" -PercentComplete ($progressCount * 100 / $Folders.Count)
            }

            $identity = $_.Identity.ToString()
            $entryId = $_.EntryId.ToString()
            Get-PublicFolderClientPermission $entryId | ForEach-Object {
                if (
                    ($_.User.DisplayName -ne "Default") -and
                    ($_.User.DisplayName -ne "Anonymous") -and
                    ($null -eq $_.User.ADRecipient) -and
                    ($_.User.UserType.ToString() -eq "Unknown")
                ) {
                    # We can't use New-TestResult here since we are inside a job
                    [PSCustomObject]@{
                        TestName       = "Permission"
                        ResultType     = "BadPermission"
                        Severity       = "Error"
                        FolderIdentity = $identity
                        FolderEntryId  = $entryId
                        ResultData     = $_.User.DisplayName
                    }
                }
            }
        }
    }

    end {
        Write-Progress @progressParams -Completed
        [PSCustomObject]@{
            TestName       = "Permission"
            ResultType     = "$Mailbox Duration"
            Severity       = "Information"
            FolderIdentity = ""
            FolderEntryId  = ""
            ResultData     = ((Get-Date) - $startTime)
        }
    }
}

function Test-Permission {
    [CmdletBinding()]
    param (
        [Parameter()]
        [PSCustomObject]
        $FolderData
    )

    begin {
        $startTime = Get-Date
    }

    process {
        $folderData.IpmSubtreeByMailbox | ForEach-Object {
            $argumentList = $FolderData.MailboxToServerMap[$_.Name], $_.Name, $_.Group
            $name = $_.Name
            $scriptBlock = ${Function:Test-BadPermissionJob}
            Add-JobQueueJob @{
                ArgumentList = $argumentList
                Name         = "$name Permissions Check"
                ScriptBlock  = $scriptBlock
            }
        }

        Wait-QueuedJob
    }

    end {
        $params = @{
            TestName       = "Permission"
            ResultType     = "Duration"
            Severity       = "Information"
            FolderIdentity = ""
            FolderEntryId  = ""
            ResultData     = ((Get-Date) - $startTime)
        }

        New-TestResult @params
    }
}


function Write-TestPermissionResult {
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipeline = $true)]
        [object]
        $TestResult
    )

    begin {
        $badPermissionResults = New-Object System.Collections.ArrayList
    }

    process {
        if ($TestResult.TestName -eq "Permission" -and $TestResult.ResultType -eq "BadPermission") {
            [void]$badPermissionResults.Add($TestResult)
        }
    }

    end {
        if ($badPermissionResults.Count -gt 0) {
            Get-ResultSummary -ResultType $badPermissionResults[0].ResultType -Severity $badPermissionResults[0].Severity -Count $badPermissionResults.Count -Action (
                "Invalid permissions were found. These can be removed using the RemoveInvalidPermissions switch as follows:`n`n" +
                ".\SourceSideValidations.ps1 -RemoveInvalidPermissions")
        }
    }
}

function Remove-InvalidPermission {
    [CmdletBinding(SupportsShouldProcess)]
    param (
        [Parameter(ValueFromPipeline = $true)]
        [object]
        $TestResult
    )

    begin {
        $progressParams = @{
            Activity = "Repairing folder permissions"
        }

        $progressCount = 0
        $entryIdsProcessed = New-Object 'System.Collections.Generic.HashSet[string]'
        $badPermissions = New-Object System.Collections.ArrayList

        $sw = New-Object System.Diagnostics.Stopwatch
        $sw.Start()
    }

    process {
        if ($TestResult.TestName -eq "Permission" -and $TestResult.ResultType -eq "BadPermission") {
            [void]$badPermissions.Add($TestResult)
        }
    }

    end {
        foreach ($result in $badPermissions) {
            $progressCount++

            if ($sw.ElapsedMilliseconds -gt 1000) {
                $sw.Restart()
                Write-Progress @progressParams -Status "$progressCount / $($badPermissions.Count)" -PercentComplete ($progressCount * 100 / $badPermissions.Count) -CurrentOperation $permission.Identity
            }

            if ($entryIdsProcessed.Add($result.FolderEntryId)) {
                $permsOnFolder = Get-PublicFolderClientPermission -Identity $result.FolderEntryId
                foreach ($perm in $permsOnFolder) {
                    if (
                        ($perm.User.DisplayName -ne "Default") -and
                        ($perm.User.DisplayName -ne "Anonymous") -and
                        ($null -eq $perm.User.ADRecipient) -and
                        ($perm.User.UserType -eq "Unknown")
                    ) {
                        if ($PSCmdlet.ShouldProcess("$($result.FolderIdentity)", "Remove $($perm.User.DisplayName)")) {
                            Write-Host "Removing $($perm.User.DisplayName) from folder $($result.FolderIdentity)"
                            $perm | Remove-PublicFolderClientPermission -Confirm:$false
                        }
                    }
                }
            }
        }
    }
}


function Get-IpmSubtree {
    [CmdletBinding()]
    param (
        [Parameter(Position = 0)]
        [string]
        $Server,

        [Parameter(Position = 1)]
        [bool]
        $SlowTraversal = $false
    )

    begin {
        $WarningPreference = "SilentlyContinue"
        $progressCount = 0
        $maxRetries = 10
        $retryDelay = [TimeSpan]::FromMinutes(5)
        $ipmSubtree = New-Object System.Collections.ArrayList
        $sw = New-Object System.Diagnostics.Stopwatch
        $sw.Start()
        $progressParams = @{
            Activity = "Retrieving IPM_SUBTREE folders"
        }

        # Only used for slow traversal to save progress in case of failure
        $foldersProcessed = New-Object 'System.Collections.Generic.HashSet[string]'

        # This must be defined in the function scope because this function is runs as a job
        function Get-FoldersRecursive {
            [CmdletBinding()]
            param (
                [Parameter(Position = 0)]
                [object]
                $Folder,

                [Parameter(Position = 1)]
                [object]
                $FoldersProcessed
            )

            $children = Get-PublicFolder $Folder.EntryId -GetChildren -ResultSize Unlimited
            foreach ($child in $children) {
                if (-not $FoldersProcessed.Contains($child.EntryId.ToString())) {
                    if ($child.HasSubFolders) {
                        Get-FoldersRecursive $child $FoldersProcessed
                    }

                    $child
                }
            }
        }
    }

    process {
        $getCommand = { Get-PublicFolder -Recurse -ResultSize Unlimited }

        if ($SlowTraversal) {
            $getCommand = { $top = Get-PublicFolder "\"; Get-FoldersRecursive $top $foldersProcessed; $top }
        }

        $outputResultsScriptBlock = {
            [CmdletBinding()]
            param (
                [Parameter(ValueFromPipeline = $true)]
                [object]
                $Folder
            )

            process {
                $progressCount++

                if ($sw.ElapsedMilliseconds -gt 1000) {
                    $sw.Restart()
                    Write-Progress @progressParams -Status $progressCount
                }

                $result = [PSCustomObject]@{
                    Name              = $Folder.Name
                    Identity          = $Folder.Identity.ToString()
                    EntryId           = $Folder.EntryId.ToString()
                    ParentEntryId     = $Folder.ParentFolder.ToString()
                    DumpsterEntryId   = if ($Folder.DumpsterEntryId) { $Folder.DumpsterEntryId.ToString() } else { $null }
                    FolderSize        = $Folder.FolderSize
                    HasSubFolders     = $Folder.HasSubFolders
                    ContentMailbox    = $Folder.ContentMailboxName
                    MailEnabled       = $Folder.MailEnabled
                    MailRecipientGuid = $Folder.MailRecipientGuid
                }

                [void]$ipmSubtree.Add($result)

                [void]$foldersProcessed.Add($Folder.EntryId.ToString())
            }
        }

        for ($retryCount = 1; $retryCount -le $maxRetries; $retryCount++) {
            try {
                Get-PSSession | Remove-PSSession
                Import-PSSession (New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://$Server/powershell" -Authentication Kerberos) -AllowClobber | Out-Null
                Invoke-Command $getCommand | &$outputResultsScriptBlock
                break
            } catch {
                if (-not $SlowTraversal) {
                    throw
                }

                $sw.Restart()
                while ($sw.ElapsedMilliseconds -lt $retryDelay.TotalMilliseconds) {
                    Write-Progress @progressParams -Status "Retry $retryCount of $maxRetries. Error: $($_.Message)"
                    Start-Sleep -Seconds 5
                    $remainingMilliseconds = $retryDelay.TotalMilliseconds - $sw.ElapsedMilliseconds
                    if ($remainingMilliseconds -lt 0) { $remainingMilliseconds = 0 }
                    Write-Progress @progressParams -Status "Retry $retryCount of $maxRetries. Will retry in $([TimeSpan]::FromMilliseconds($remainingMilliseconds))"
                    Start-Sleep -Seconds 5
                }
            }
        }
    }

    end {
        Write-Progress @progressParams -Completed

        return [PSCustomObject]@{
            IpmSubtree = $ipmSubtree
        }
    }
}

function Get-NonIpmSubtree {
    [CmdletBinding()]
    param (
        [Parameter(Position = 0)]
        [string]
        $Server,

        [Parameter(Position = 1)]
        [bool]
        $SlowTraversal = $false
    )

    begin {
        $WarningPreference = "SilentlyContinue"
        $progressCount = 0
        $maxRetries = 10
        $retryDelay = [TimeSpan]::FromMinutes(5)
        $nonIpmSubtree = New-Object System.Collections.ArrayList
        $sw = New-Object System.Diagnostics.Stopwatch
        $sw.Start()
        $progressParams = @{
            Activity = "Retrieving NON_IPM_SUBTREE folders"
        }

        # Only used for slow traversal to save progress in case of failure
        $foldersProcessed = New-Object 'System.Collections.Generic.HashSet[string]'

        # This must be defined in the function scope because this function is runs as a job
        function Get-FoldersRecursive {
            [CmdletBinding()]
            param (
                [Parameter(Position = 0)]
                [object]
                $Folder,

                [Parameter(Position = 1)]
                [object]
                $FoldersProcessed
            )

            $children = Get-PublicFolder $Folder.EntryId -GetChildren -ResultSize Unlimited
            foreach ($child in $children) {
                if (-not $FoldersProcessed.Contains($child.EntryId.ToString())) {
                    if ($child.HasSubFolders) {
                        Get-FoldersRecursive $child $FoldersProcessed
                    }

                    $child
                }
            }
        }
    }

    process {
        $getCommand = { Get-PublicFolder "\non_ipm_subtree" -Recurse -ResultSize Unlimited }

        if ($SlowTraversal) {
            $getCommand = { $top = Get-PublicFolder "\non_ipm_subtree"; Get-FoldersRecursive $top $foldersProcessed; $top }
        }

        $outputResultsScriptBlock = {
            [CmdletBinding()]
            param (
                [Parameter(ValueFromPipeline = $true)]
                [object]
                $Folder
            )

            process {
                $progressCount++
                if ($sw.ElapsedMilliseconds -gt 1000) {
                    $sw.Restart()
                    Write-Progress @progressParams -Status $progressCount
                }

                $result = [PSCustomObject]@{
                    Identity        = $Folder.Identity.ToString()
                    EntryId         = $Folder.EntryId.ToString()
                    ParentEntryId   = $Folder.ParentFolder.ToString()
                    DumpsterEntryId = if ($Folder.DumpsterEntryId) { $Folder.DumpsterEntryId.ToString() } else { $null }
                    MailEnabled     = $Folder.MailEnabled
                }

                $null = $nonIpmSubtree.Add($result)

                $null = $foldersProcessed.Add($Folder.EntryId.ToString())
            }
        }

        for ($retryCount = 1; $retryCount -le $maxRetries; $retryCount++) {
            try {
                Get-PSSession | Remove-PSSession
                Import-PSSession (New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://$Server/powershell" -Authentication Kerberos) -AllowClobber | Out-Null
                Invoke-Command $getCommand | &$outputResultsScriptBlock
                break
            } catch {
                if (-not $SlowTraversal) {
                    throw
                }

                $sw.Restart()
                while ($sw.ElapsedMilliseconds -lt $retryDelay.TotalMilliseconds) {
                    Write-Progress @progressParams -Status "Retry $retryCount of $maxRetries. Error: $($_.Message)"
                    Start-Sleep -Seconds 5
                    $remainingMilliseconds = $retryDelay.TotalMilliseconds - $sw.ElapsedMilliseconds
                    if ($remainingMilliseconds -lt 0) { $remainingMilliseconds = 0 }
                    Write-Progress @progressParams -Status "Retry $retryCount of $maxRetries. Will retry in $([TimeSpan]::FromMilliseconds($remainingMilliseconds))"
                    Start-Sleep -Seconds 5
                }
            }
        }
    }

    end {
        Write-Progress @progressParams -Completed

        return [PSCustomObject]@{
            NonIpmSubtree = $nonIpmSubtree
        }
    }
}


function Get-StatisticsJob {
    [CmdletBinding()]
    param (
        [Parameter(Position = 0)]
        [string]
        $Server,

        [Parameter(Position = 1)]
        [string]
        $Mailbox,

        [Parameter(Position = 2)]
        [PSCustomObject[]]
        $Folders
    )

    begin {
        $statistics = New-Object System.Collections.ArrayList
        $errors = New-Object System.Collections.ArrayList
        $permanentFailureOccurred = $false
        $permanentFailures = @(
            "Kerberos",
            "Cannot process argument transformation on parameter 'Identity'",
            "Starting a command on the remote server failed"
        )
        $WarningPreference = "SilentlyContinue"
        $Error.Clear()
        Import-PSSession (New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://$Server/powershell" -Authentication Kerberos) -AllowClobber | Out-Null
        if ($Error.Count -gt 0) {
            $permanentFailureOccurred = $true
            foreach ($err in $Error) {
                $errorReport = @{
                    TestName       = "Get-Statistics"
                    ResultType     = "ImportSessionFailure"
                    Severity       = "Error"
                    FolderIdentity = ""
                    FolderEntryId  = ""
                    ResultData     = $err.ToString()
                }

                [void]$errors.Add($errorReport)
            }
        }

        if (-not $permanentFailureOccurred -and $null -eq (Get-Command Get-PublicFolderStatistics -ErrorAction SilentlyContinue)) {
            $permanentFailureOccurred = $true
            $errorReport = @{
                TestName       = "Get-Statistics"
                ResultType     = "CommandNotFound"
                Severity       = "Error"
                FolderIdentity = ""
                FolderEntryId  = ""
                ResultData     = ""
            }

            [void]$errors.Add($errorReport)
        }

        $startTime = Get-Date
        $progressCount = 0
        $sw = New-Object System.Diagnostics.Stopwatch
        $sw.Start()
        $progressParams = @{
            Activity = "Getting public folder statistics"
        }
    }

    process {
        if ($permanentFailureOccurred) {
            return
        }

        $ErrorActionPreference = "Stop" # So our try/catch works
        $statistics = New-Object System.Collections.ArrayList
        foreach ($folder in $Folders) {
            $progressCount++
            if ($sw.ElapsedMilliseconds -gt 1000) {
                $sw.Restart()
                Write-Progress @progressParams -Status $progressCount
            }

            try {
                if ([string]::IsNullOrEmpty($folder.EntryId)) {
                    $folderObject = $folder | Format-List | Out-String
                    $foldersCollection = $Folders | Format-List | Out-String
                    $errorDetails = "$folderObject`n`n$foldersCollection"
                    $errorReport = @{
                        TestName       = "Get-Statistics"
                        ResultType     = "NullEntryId"
                        Severity       = "Error"
                        FolderIdentity = $folder.Identity
                        FolderEntryId  = $folder.EntryId
                        ResultData     = $errorDetails
                    }

                    [void]$errors.Add($errorReport)
                }
                $stats = Get-PublicFolderStatistics $folder.EntryId | Select-Object EntryId, ItemCount, TotalItemSize

                [Int64]$totalItemSize = -1
                if ($stats.TotalItemSize.ToString() -match "\(([\d|,|.]+) bytes\)") {
                    $totalItemSize = [Int64]::Parse($Matches[1], "AllowThousands")
                }

                [void]$statistics.Add([PSCustomObject]@{
                        EntryId       = $stats.EntryId
                        ItemCount     = $stats.ItemCount
                        TotalItemSize = $totalItemSize
                    })
            } catch {
                $errorText = $_.ToString()
                $isPermanentFailure = $null -ne ($permanentFailures | Where-Object { $errorText.Contains($_) })
                if ($isPermanentFailure) {
                    $errorReport = @{
                        TestName       = "Get-Statistics"
                        ResultType     = "JobFailure"
                        Severity       = "Error"
                        FolderIdentity = $folder.Identity
                        FolderEntryId  = $folder.EntryId
                        ResultData     = $errorText
                    }

                    [void]$errors.Add($errorReport)
                    $permanentFailureOccurred = $true
                    break
                } else {
                    $errorReport = @{
                        TestName       = "Get-Statistics"
                        ResultType     = "CouldNotGetStatistics"
                        Severity       = "Error"
                        FolderIdentity = $folder.Identity
                        FolderEntryId  = $folder.EntryId
                        ResultData     = $errorText
                    }

                    [void]$errors.Add($errorReport)
                }
            }
        }
    }

    end {
        Write-Progress @progressParams -Completed
        $duration = ((Get-Date) - $startTime)
        return [PSCustomObject]@{
            Statistics       = $statistics
            Errors           = $errors
            PermanentFailure = $permanentFailureOccurred
            Server           = $Server
            Mailbox          = $Mailbox
            Folders          = $Folders
            Duration         = $duration
        }
    }
}

function Get-Statistics {
    <#
    .SYNOPSIS
        Gets the item count for each folder.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Position = 0)]
        [string]
        $Server,

        [Parameter(Position = 1)]
        [PSCustomObject]
        $FolderData = $null
    )

    begin {
        Write-Verbose "$($MyInvocation.MyCommand) called."

        $progressCount = 0
        $sw = New-Object System.Diagnostics.Stopwatch
        $sw.Start()
        $progressParams = @{
            Activity = "Getting public folder statistics"
        }

        $statistics = New-Object System.Collections.ArrayList
        $errors = New-Object System.Collections.ArrayList
    }

    process {
        if ($null -eq $FolderData) {
            $WarningPreference = "SilentlyContinue"
            Import-PSSession (New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://$Server/powershell" -Authentication Kerberos) | Out-Null
            $statistics = Get-PublicFolderStatistics -ResultSize Unlimited | ForEach-Object {
                $progressCount++
                if ($sw.ElapsedMilliseconds -gt 1000) {
                    $sw.Restart()
                    Write-Progress @progressParams -Status $progressCount
                }

                [Int64]$totalItemSize = -1
                if ($_.TotalItemSize.ToString() -match "\(([\d|,|.]+) bytes\)") {
                    $numberString = $Matches[1] -replace "\D", ""
                    $totalItemSize = [Int64]::Parse($numberString)
                }

                [PSCustomObject]@{
                    EntryId       = $_.EntryId
                    ItemCount     = $_.ItemCount
                    TotalItemSize = $totalItemSize
                }
            }
        } else {
            $batchSize = 10000
            $jobsToCreate = New-Object 'System.Collections.Generic.Dictionary[string, System.Collections.ArrayList]'
            foreach ($group in $folderData.IpmSubtreeByMailbox) {
                # MailboxToServerMap is not populated yet, so we can't use it here
                $server = (Get-MailboxDatabase (Get-Mailbox -PublicFolder $group.Name).Database).Server.Name
                [int]$mailboxBatchCount = ($group.Group.Count / $batchSize) + 1
                Write-Verbose "Creating $mailboxBatchCount statistics jobs for $($group.Group.Count) folders in mailbox $($group.Name) on server $server."
                $jobsForThisMailbox = New-Object System.Collections.ArrayList
                for ($i = 0; $i -lt $mailboxBatchCount; $i++) {
                    $batch = $group.Group | Select-Object -First $batchSize -Skip ($batchSize * $i)
                    if ($batch.Count -gt 0) {
                        $argumentList = $server, $group.Name, $batch
                        [void]$jobsForThisMailbox.Add(@{
                                ArgumentList = $argumentList
                                Name         = "Statistics $($group.Name) Job $($i + 1)"
                                ScriptBlock  = ${Function:Get-StatisticsJob}
                            })
                    }
                }

                [void]$jobsToCreate.Add($group.Name, $jobsForThisMailbox)
            }

            # Add the jobs by round-robin among the mailboxes so we don't execute all jobs
            # for one mailbox in parallel unless we have to
            $jobsAddedThisRound = 0
            $index = 0
            do {
                $jobsAddedThisRound = 0
                foreach ($mailboxName in $jobsToCreate.Keys) {
                    $batchesForThisMailbox = $jobsToCreate[$mailboxName]
                    if ($batchesForThisMailbox.Count -gt $index) {
                        $jobParams = $batchesForThisMailbox[$index]
                        Add-JobQueueJob $jobParams
                        $jobsAddedThisRound++
                    }
                }

                $index++
            } while ($jobsAddedThisRound -gt 0)

            $hierarchyMailbox = Get-Mailbox -PublicFolder (Get-OrganizationConfig).RootPublicFolderMailbox.ToString()
            $serverWithHierarchy = (Get-MailboxDatabase $hierarchyMailbox.Database).Server.Name
            $retryJobNumber = 1

            Wait-QueuedJob | ForEach-Object {
                $finishedJob = $_
                $statistics.AddRange($finishedJob.Statistics)
                $errors.AddRange($finishedJob.Errors)
                Write-Verbose "Retrieved item counts for $($statistics.Count) folders so far. $($errors.Count) errors encountered."
                if ($finishedJob.PermanentFailure) {
                    # If a permanent failure occurred, re-queue remaining items on the server that has the writable
                    # hierarchy, and hope it works there.
                    Write-Host "Job experienced a permanent failure."
                    if ($finishedJob.Server -eq $serverWithHierarchy) {
                        Write-Host "Permanent failure on root mailbox server is not retryable."
                    } else {
                        $entryIdsProcessed = New-Object 'System.Collections.Generic.HashSet[string]'
                        $finishedJob.Statistics | ForEach-Object { [void]$entryIdsProcessed.Add($_.EntryId) }
                        $foldersRemaining = @($finishedJob.Folders | Where-Object { -not $entryIdsProcessed.Contains($_.EntryId) })
                        if ($foldersRemaining.Count -gt 0) {
                            Write-Host "$($foldersRemaining.Count) folders remaining in the failed job. Re-queueing for $serverWithHierarchy."
                            $retryJob = @{
                                ArgumentList = $serverWithHierarchy, $hierarchyMailbox.Name, $foldersRemaining
                                Name         = "Statistics Retry Job $($retryJobNumber++)"
                                ScriptBlock  = ${Function:Get-StatisticsJob}
                            }

                            Add-JobQueueJob $retryJob
                        }
                    }
                }
            }
        }
    }

    end {
        Write-Progress @progressParams -Completed

        return [PSCustomObject]@{
            Statistics = $statistics
            Errors     = $errors
        }
    }
}

function Get-FolderData {
    [CmdletBinding()]
    param (
        [Parameter()]
        [bool]
        $StartFresh = $true,

        [Parameter()]
        [bool]
        $SlowTraversal = $false
    )

    begin {
        Write-Verbose "$($MyInvocation.MyCommand) called."
        $startTime = Get-Date
        $serverName = (Get-MailboxDatabase (Get-Mailbox -PublicFolder (Get-OrganizationConfig).RootPublicFolderMailbox.HierarchyMailboxGuid.ToString()).Database).Server.Name
        $folderData = [PSCustomObject]@{
            IpmSubtree              = $null
            IpmSubtreeByMailbox     = $null
            ParentEntryIdCounts     = @{}
            EntryIdDictionary       = @{}
            NonIpmSubtree           = $null
            NonIpmEntryIdDictionary = @{}
            MailboxToServerMap      = @{}
            Statistics              = @()
            StatisticsDictionary    = @{}
            Errors                  = New-Object System.Collections.ArrayList
        }
    }

    process {
        if (-not $StartFresh -and (Test-Path $PSScriptRoot\IpmSubtree.csv)) {
            $folderData.IpmSubtree = Import-Csv $PSScriptRoot\IpmSubtree.csv
        } else {
            Add-JobQueueJob @{
                ArgumentList = $serverName, $SlowTraversal
                Name         = "Get-IpmSubtree"
                ScriptBlock  = ${Function:Get-IpmSubtree}
            }
        }

        if (-not $StartFresh -and (Test-Path $PSScriptRoot\NonIpmSubtree.csv)) {
            $folderData.NonIpmSubtree = Import-Csv $PSScriptRoot\NonIpmSubtree.csv
        } else {
            Add-JobQueueJob @{
                ArgumentList = $serverName, $SlowTraversal
                Name         = "Get-NonIpmSubtree"
                ScriptBlock  = ${Function:Get-NonIpmSubtree}
            }
        }

        # If we're not doing slow traversal, we can get the stats concurrently with the other jobs
        if (-not $SlowTraversal) {
            if (-not $StartFresh -and (Test-Path $PSScriptRoot\Statistics.csv)) {
                $folderData.Statistics = Import-Csv $PSScriptRoot\Statistics.csv
            } else {
                Add-JobQueueJob @{
                    ArgumentList = $serverName
                    Name         = "Get-Statistics"
                    ScriptBlock  = ${Function:Get-Statistics}
                }
            }
        }

        $completedJobs = Wait-QueuedJob

        foreach ($job in $completedJobs) {
            if ($null -ne $job.IpmSubtree) {
                $folderData.IpmSubtree = $job.IpmSubtree
                $folderData.IpmSubtree | Export-Csv $PSScriptRoot\IpmSubtree.csv
            }

            if ($null -ne $job.NonIpmSubtree) {
                $folderData.NonIpmSubtree = $job.NonIpmSubtree
                $folderData.NonIpmSubtree | Export-Csv $PSScriptRoot\NonIpmSubtree.csv
            }

            if ($null -ne $job.Statistics) {
                $folderData.Statistics = $job.Statistics
                $folderData.Statistics | Export-Csv $PSScriptRoot\Statistics.csv
            }
        }

        $folderData.IpmSubtreeByMailbox = $folderData.IpmSubtree | Group-Object ContentMailbox
        $folderData.IpmSubtree | ForEach-Object { $folderData.ParentEntryIdCounts[$_.ParentEntryId] += 1 }
        $folderData.NonIpmSubtree | ForEach-Object { $folderData.ParentEntryIdCounts[$_.ParentEntryId] += 1 }
        $folderData.IpmSubtree | ForEach-Object { $folderData.EntryIdDictionary[$_.EntryId] = $_ }
        # We can't count on $folder.Path.Depth being available in remote powershell,
        # so we calculate the depth by walking the parent entry IDs.
        $folderData.IpmSubtree | ForEach-Object {
            $pathDepth = 0
            $parent = $folderData.EntryIdDictionary[$_.ParentEntryId]
            while ($null -ne $parent) {
                $pathDepth++
                $parent = $folderData.EntryIdDictionary[$parent.ParentEntryId]
            }

            Add-Member -InputObject $_ -MemberType NoteProperty -Name FolderPathDepth -Value $pathDepth
        }
        $folderData.NonIpmSubtree | ForEach-Object { $folderData.NonIpmEntryIdDictionary[$_.EntryId] = $_ }

        # If we're doing slow traversal, we have to get the stats after we have the hierarchy
        # grouped by mailbox.
        if ($SlowTraversal) {
            if (-not $StartFresh -and (Test-Path $PSScriptRoot\Statistics.csv)) {
                $folderData.Statistics = Import-Csv $PSScriptRoot\Statistics.csv
            } else {
                Write-Verbose "Starting slow traversal item count."
                $statisticsResult = Get-Statistics $serverName $folderData
                $folderData.Statistics = $statisticsResult.Statistics
                $folderData.Statistics | Export-Csv $PSScriptRoot\Statistics.csv
                foreach ($errorParam in $statisticsResult.Errors) {
                    $errorResult = New-TestResult @errorParam
                    $folderData.Errors.Add($errorResult)
                }
            }
        }

        $folderData.Statistics | ForEach-Object { $folderData.StatisticsDictionary[$_.EntryId] = $_ }
    }

    end {
        Write-Host "Get-FolderData duration $((Get-Date) - $startTime)"
        Write-Host "    IPM_SUBTREE folder count: $($folderData.IpmSubtree.Count)"
        Write-Host "    NON_IPM_SUBTREE folder count: $($folderData.NonIpmSubtree.Count)"

        return $folderData
    }
}

$jobsQueued = New-Object 'System.Collections.Generic.Queue[object]'

function Add-JobQueueJob {
    [CmdletBinding()]
    param (
        [Parameter()]
        [PSCustomObject]
        $JobParameters
    )

    begin {
    }

    process {
        $jobsQueued.Enqueue($JobParameters)
        Write-Host "Added job $($JobParameters.Name) to queue."
    }

    end {
    }
}

function Wait-QueuedJob {
    [CmdletBinding()]
    [OutputType([System.Object[]])]
    param (

    )

    begin {
        $jobsRunning = @()
        $jobQueueMaxConcurrency = 5
    }

    process {
        while ($jobsQueued.Count -gt 0 -or $jobsRunning.Count -gt 0) {
            if ($jobsRunning.Count -lt $jobQueueMaxConcurrency -and $jobsQueued.Count -gt 0) {
                $jobArgs = $jobsQueued.Dequeue()
                $newJob = Start-Job @jobArgs
                $jobsRunning += $newJob
                Write-Host "Started executing job $($jobArgs.Name)."
            }

            $justFinished = @($jobsRunning | Where-Object { $_.State -ne "Running" })
            if ($justFinished.Count -gt 0) {
                foreach ($job in $justFinished) {
                    $result = Receive-Job $job
                    $lastProgress = $job.ChildJobs.Progress | Select-Object -Last 1
                    if ($lastProgress) {
                        Write-Progress -Activity $lastProgress.Activity -ParentId 1 -Id ($job.Id + 1) -Completed
                    }
                    Write-Host $job.Name "job finished."
                    Remove-Job $job -Force
                    $result
                }

                $jobsRunning = @($jobsRunning | Where-Object { -not $justFinished.Contains($_) })
            }

            for ($i = 0; $i -lt $jobQueueMaxConcurrency; $i++) {
                if ($jobsRunning.Count -gt $i) {
                    $lastProgress = $jobsRunning[$i].ChildJobs.Progress | Select-Object -Last 1
                    if ($lastProgress) {
                        Write-Progress -Activity $lastProgress.Activity -Status $lastProgress.StatusDescription -PercentComplete $lastProgress.PercentComplete -ParentId 1 -Id ($jobsRunning[$i].Id + 1)
                    }
                }
            }

            Start-Sleep 1
        }
    }
}




function Confirm-ProxyServer {
    [CmdletBinding()]
    [OutputType([bool])]
    param (
        [Parameter(Mandatory = $true)]
        [string]
        $TargetUri
    )

    Write-Verbose "Calling $($MyInvocation.MyCommand)"
    try {
        $proxyObject = ([System.Net.WebRequest]::GetSystemWebProxy()).GetProxy($TargetUri)
        if ($TargetUri -ne $proxyObject.OriginalString) {
            Write-Verbose "Proxy server configuration detected"
            Write-Verbose $proxyObject.OriginalString
            return $true
        } else {
            Write-Verbose "No proxy server configuration detected"
            return $false
        }
    } catch {
        Write-Verbose "Unable to check for proxy server configuration"
        return $false
    }
}

function WriteErrorInformationBase {
    [CmdletBinding()]
    param(
        [object]$CurrentError = $Error[0],
        [ValidateSet("Write-Host", "Write-Verbose")]
        [string]$Cmdlet
    )

    if ($null -ne $CurrentError.OriginInfo) {
        & $Cmdlet "Error Origin Info: $($CurrentError.OriginInfo.ToString())"
    }

    & $Cmdlet "$($CurrentError.CategoryInfo.Activity) : $($CurrentError.ToString())"

    if ($null -ne $CurrentError.Exception -and
        $null -ne $CurrentError.Exception.StackTrace) {
        & $Cmdlet "Inner Exception: $($CurrentError.Exception.StackTrace)"
    } elseif ($null -ne $CurrentError.Exception) {
        & $Cmdlet "Inner Exception: $($CurrentError.Exception)"
    }

    if ($null -ne $CurrentError.InvocationInfo.PositionMessage) {
        & $Cmdlet "Position Message: $($CurrentError.InvocationInfo.PositionMessage)"
    }

    if ($null -ne $CurrentError.Exception.SerializedRemoteInvocationInfo.PositionMessage) {
        & $Cmdlet "Remote Position Message: $($CurrentError.Exception.SerializedRemoteInvocationInfo.PositionMessage)"
    }

    if ($null -ne $CurrentError.ScriptStackTrace) {
        & $Cmdlet "Script Stack: $($CurrentError.ScriptStackTrace)"
    }
}

function Write-VerboseErrorInformation {
    [CmdletBinding()]
    param(
        [object]$CurrentError = $Error[0]
    )
    WriteErrorInformationBase $CurrentError "Write-Verbose"
}

function Write-HostErrorInformation {
    [CmdletBinding()]
    param(
        [object]$CurrentError = $Error[0]
    )
    WriteErrorInformationBase $CurrentError "Write-Host"
}

function Invoke-WebRequestWithProxyDetection {
    [CmdletBinding(DefaultParameterSetName = "Default")]
    param (
        [Parameter(Mandatory = $true, ParameterSetName = "Default")]
        [string]
        $Uri,

        [Parameter(Mandatory = $false, ParameterSetName = "Default")]
        [switch]
        $UseBasicParsing,

        [Parameter(Mandatory = $true, ParameterSetName = "ParametersObject")]
        [hashtable]
        $ParametersObject,

        [Parameter(Mandatory = $false, ParameterSetName = "Default")]
        [string]
        $OutFile
    )

    Write-Verbose "Calling $($MyInvocation.MyCommand)"
    if ([System.String]::IsNullOrEmpty($Uri)) {
        $Uri = $ParametersObject.Uri
    }

    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    if (Confirm-ProxyServer -TargetUri $Uri) {
        $webClient = New-Object System.Net.WebClient
        $webClient.Headers.Add("User-Agent", "PowerShell")
        $webClient.Proxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
    }

    if ($null -eq $ParametersObject) {
        $params = @{
            Uri     = $Uri
            OutFile = $OutFile
        }

        if ($UseBasicParsing) {
            $params.UseBasicParsing = $true
        }
    } else {
        $params = $ParametersObject
    }

    try {
        Invoke-WebRequest @params
    } catch {
        Write-VerboseErrorInformation
    }
}

<#
    Determines if the script has an update available.
#>
function Get-ScriptUpdateAvailable {
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param (
        [Parameter(Mandatory = $false)]
        [string]
        $VersionsUrl = "https://github.com/microsoft/CSS-Exchange/releases/latest/download/ScriptVersions.csv"
    )

    $BuildVersion = "23.04.21.1447"

    $scriptName = $script:MyInvocation.MyCommand.Name
    $scriptPath = [IO.Path]::GetDirectoryName($script:MyInvocation.MyCommand.Path)
    $scriptFullName = (Join-Path $scriptPath $scriptName)

    $result = [PSCustomObject]@{
        ScriptName     = $scriptName
        CurrentVersion = $BuildVersion
        LatestVersion  = ""
        UpdateFound    = $false
        Error          = $null
    }

    if ((Get-AuthenticodeSignature -FilePath $scriptFullName).Status -eq "NotSigned") {
        Write-Warning "This script appears to be an unsigned test build. Skipping version check."
    } else {
        try {
            $versionData = [Text.Encoding]::UTF8.GetString((Invoke-WebRequestWithProxyDetection -Uri $VersionsUrl -UseBasicParsing).Content) | ConvertFrom-Csv
            $latestVersion = ($versionData | Where-Object { $_.File -eq $scriptName }).Version
            $result.LatestVersion = $latestVersion
            if ($null -ne $latestVersion) {
                $result.UpdateFound = ($latestVersion -ne $BuildVersion)
            } else {
                Write-Warning ("Unable to check for a script update as no script with the same name was found." +
                    "`r`nThis can happen if the script has been renamed. Please check manually if there is a newer version of the script.")
            }

            Write-Verbose "Current version: $($result.CurrentVersion) Latest version: $($result.LatestVersion) Update found: $($result.UpdateFound)"
        } catch {
            Write-Verbose "Unable to check for updates: $($_.Exception)"
            $result.Error = $_
        }
    }

    return $result
}


function Confirm-Signature {
    [CmdletBinding()]
    [OutputType([bool])]
    param (
        [Parameter(Mandatory = $true)]
        [string]
        $File
    )

    $IsValid = $false
    $MicrosoftSigningRoot2010 = 'CN=Microsoft Root Certificate Authority 2010, O=Microsoft Corporation, L=Redmond, S=Washington, C=US'
    $MicrosoftSigningRoot2011 = 'CN=Microsoft Root Certificate Authority 2011, O=Microsoft Corporation, L=Redmond, S=Washington, C=US'

    try {
        $sig = Get-AuthenticodeSignature -FilePath $File

        if ($sig.Status -ne 'Valid') {
            Write-Warning "Signature is not trusted by machine as Valid, status: $($sig.Status)."
            throw
        }

        $chain = New-Object -TypeName System.Security.Cryptography.X509Certificates.X509Chain
        $chain.ChainPolicy.VerificationFlags = "IgnoreNotTimeValid"

        if (-not $chain.Build($sig.SignerCertificate)) {
            Write-Warning "Signer certificate doesn't chain correctly."
            throw
        }

        if ($chain.ChainElements.Count -le 1) {
            Write-Warning "Certificate Chain shorter than expected."
            throw
        }

        $rootCert = $chain.ChainElements[$chain.ChainElements.Count - 1]

        if ($rootCert.Certificate.Subject -ne $rootCert.Certificate.Issuer) {
            Write-Warning "Top-level certificate in chain is not a root certificate."
            throw
        }

        if ($rootCert.Certificate.Subject -ne $MicrosoftSigningRoot2010 -and $rootCert.Certificate.Subject -ne $MicrosoftSigningRoot2011) {
            Write-Warning "Unexpected root cert. Expected $MicrosoftSigningRoot2010 or $MicrosoftSigningRoot2011, but found $($rootCert.Certificate.Subject)."
            throw
        }

        Write-Host "File signed by $($sig.SignerCertificate.Subject)"

        $IsValid = $true
    } catch {
        $IsValid = $false
    }

    $IsValid
}

<#
.SYNOPSIS
    Overwrites the current running script file with the latest version from the repository.
.NOTES
    This function always overwrites the current file with the latest file, which might be
    the same. Get-ScriptUpdateAvailable should be called first to determine if an update is
    needed.

    In many situations, updates are expected to fail, because the server running the script
    does not have internet access. This function writes out failures as warnings, because we
    expect that Get-ScriptUpdateAvailable was already called and it successfully reached out
    to the internet.
#>
function Invoke-ScriptUpdate {
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    [OutputType([boolean])]
    param ()

    $scriptName = $script:MyInvocation.MyCommand.Name
    $scriptPath = [IO.Path]::GetDirectoryName($script:MyInvocation.MyCommand.Path)
    $scriptFullName = (Join-Path $scriptPath $scriptName)

    $oldName = [IO.Path]::GetFileNameWithoutExtension($scriptName) + ".old"
    $oldFullName = (Join-Path $scriptPath $oldName)
    $tempFullName = (Join-Path $env:TEMP $scriptName)

    if ($PSCmdlet.ShouldProcess("$scriptName", "Update script to latest version")) {
        try {
            Invoke-WebRequestWithProxyDetection -Uri "https://github.com/microsoft/CSS-Exchange/releases/latest/download/$scriptName" -OutFile $tempFullName
        } catch {
            Write-Warning "AutoUpdate: Failed to download update: $($_.Exception.Message)"
            return $false
        }

        try {
            if (Confirm-Signature -File $tempFullName) {
                Write-Host "AutoUpdate: Signature validated."
                if (Test-Path $oldFullName) {
                    Remove-Item $oldFullName -Force -Confirm:$false -ErrorAction Stop
                }
                Move-Item $scriptFullName $oldFullName
                Move-Item $tempFullName $scriptFullName
                Remove-Item $oldFullName -Force -Confirm:$false -ErrorAction Stop
                Write-Host "AutoUpdate: Succeeded."
                return $true
            } else {
                Write-Warning "AutoUpdate: Signature could not be verified: $tempFullName."
                Write-Warning "AutoUpdate: Update was not applied."
            }
        } catch {
            Write-Warning "AutoUpdate: Failed to apply update: $($_.Exception.Message)"
        }
    }

    return $false
}

<#
    Determines if the script has an update available. Use the optional
    -AutoUpdate switch to make it update itself. Pass -Confirm:$false
    to update without prompting the user. Pass -Verbose for additional
    diagnostic output.

    Returns $true if an update was downloaded, $false otherwise. The
    result will always be $false if the -AutoUpdate switch is not used.
#>
function Test-ScriptVersion {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSShouldProcess', '', Justification = 'Need to pass through ShouldProcess settings to Invoke-ScriptUpdate')]
    [CmdletBinding(SupportsShouldProcess)]
    [OutputType([bool])]
    param (
        [Parameter(Mandatory = $false)]
        [switch]
        $AutoUpdate,
        [Parameter(Mandatory = $false)]
        [string]
        $VersionsUrl = "https://github.com/microsoft/CSS-Exchange/releases/latest/download/ScriptVersions.csv"
    )

    $updateInfo = Get-ScriptUpdateAvailable $VersionsUrl
    if ($updateInfo.UpdateFound) {
        if ($AutoUpdate) {
            return Invoke-ScriptUpdate
        } else {
            Write-Warning "$($updateInfo.ScriptName) $BuildVersion is outdated. Please download the latest, version $($updateInfo.LatestVersion)."
        }
    }

    return $false
}

<#
.SYNOPSIS
    Outputs a table of objects with certain values colorized.
.EXAMPLE
    PS C:\> <example usage>
    Explanation of what the example does
.INPUTS
    Inputs (if any)
.OUTPUTS
    Output (if any)
.NOTES
    General notes
#>
function Out-Columns {
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipeline = $true)]
        [object[]]
        $InputObject,

        [Parameter(Mandatory = $false, Position = 0)]
        [string[]]
        $Properties,

        [Parameter(Mandatory = $false, Position = 1)]
        [ScriptBlock[]]
        $ColorizerFunctions = @(),

        [Parameter(Mandatory = $false)]
        [int]
        $IndentSpaces = 0,

        [Parameter(Mandatory = $false)]
        [int]
        $LinesBetweenObjects = 0,

        [Parameter(Mandatory = $false)]
        [ref]
        $StringOutput
    )

    begin {
        function WrapLine {
            param([string]$line, [int]$width)
            if ($line.Length -le $width -and $line.IndexOf("`n") -lt 0) {
                return $line
            }

            $lines = New-Object System.Collections.ArrayList

            $noLF = $line.Replace("`r", "")
            $lineSplit = $noLF.Split("`n")
            foreach ($l in $lineSplit) {
                if ($l.Length -le $width) {
                    [void]$lines.Add($l)
                } else {
                    $split = $l.Split(" ")
                    $sb = New-Object System.Text.StringBuilder
                    for ($i = 0; $i -lt $split.Length; $i++) {
                        if ($sb.Length -eq 0 -and $sb.Length + $split[$i].Length -lt $width) {
                            [void]$sb.Append($split[$i])
                        } elseif ($sb.Length -gt 0 -and $sb.Length + $split[$i].Length + 1 -lt $width) {
                            [void]$sb.Append(" " + $split[$i])
                        } elseif ($sb.Length -gt 0) {
                            [void]$lines.Add($sb.ToString())
                            [void]$sb.Clear()
                            $i--
                        } else {
                            if ($split[$i].Length -le $width) {
                                [void]$lines.Add($split[$i])
                            } else {
                                [void]$lines.Add($split[$i].Substring(0, $width))
                                $split[$i] = $split[$i].Substring($width)
                                $i--
                            }
                        }
                    }

                    if ($sb.Length -gt 0) {
                        [void]$lines.Add($sb.ToString())
                    }
                }
            }

            return $lines
        }

        function GetLineObjects {
            param($obj, $props, $colWidths)
            $linesNeededForThisObject = 1
            $multiLineProps = @{}
            for ($i = 0; $i -lt $props.Length; $i++) {
                $p = $props[$i]
                $val = $obj."$p"

                if ($val -isnot [array]) {
                    $val = WrapLine -line $val -width $colWidths[$i]
                } elseif ($val -is [array]) {
                    $val = $val | Where-Object { $null -ne $_ }
                    $val = $val | ForEach-Object { WrapLine -line $_ -width $colWidths[$i] }
                }

                if ($val -is [array]) {
                    $multiLineProps[$p] = $val
                    if ($val.Length -gt $linesNeededForThisObject) {
                        $linesNeededForThisObject = $val.Length
                    }
                }
            }

            if ($linesNeededForThisObject -eq 1) {
                $obj
            } else {
                for ($i = 0; $i -lt $linesNeededForThisObject; $i++) {
                    $lineProps = @{}
                    foreach ($p in $props) {
                        if ($null -ne $multiLineProps[$p] -and $multiLineProps[$p].Length -gt $i) {
                            $lineProps[$p] = $multiLineProps[$p][$i]
                        } elseif ($i -eq 0) {
                            $lineProps[$p] = $obj."$p"
                        } else {
                            $lineProps[$p] = $null
                        }
                    }

                    [PSCustomObject]$lineProps
                }
            }
        }

        function GetColumnColors {
            param($obj, $props, $functions)

            $consoleHost = (Get-Host).Name -eq "ConsoleHost"
            $colColors = New-Object string[] $props.Count
            for ($i = 0; $i -lt $props.Count; $i++) {
                if ($consoleHost) {
                    $fgColor = (Get-Host).ui.RawUi.ForegroundColor
                } else {
                    $fgColor = "White"
                }
                foreach ($func in $functions) {
                    $result = $func.Invoke($obj, $props[$i])
                    if (-not [string]::IsNullOrEmpty($result)) {
                        $fgColor = $result
                        break # The first colorizer that takes action wins
                    }
                }

                $colColors[$i] = $fgColor
            }

            $colColors
        }

        function GetColumnWidths {
            param($objects, $props)

            $colWidths = New-Object int[] $props.Count

            # Start with the widths of the property names
            for ($i = 0; $i -lt $props.Count; $i++) {
                $colWidths[$i] = $props[$i].Length
            }

            # Now check the widths of the widest values
            foreach ($thing in $objects) {
                for ($i = 0; $i -lt $props.Count; $i++) {
                    $val = $thing."$($props[$i])"
                    if ($null -ne $val) {
                        $width = 0
                        if ($val -isnot [array]) {
                            $val = $val.ToString().Split("`n")
                        }

                        $width = ($val | ForEach-Object {
                                if ($null -ne $_) { $_.ToString() } else { "" }
                            } | Sort-Object Length -Descending | Select-Object -First 1).Length

                        if ($width -gt $colWidths[$i]) {
                            $colWidths[$i] = $width
                        }
                    }
                }
            }

            # If we're within the window width, we're done
            $totalColumnWidth = $colWidths.Length * $padding + ($colWidths | Measure-Object -Sum).Sum + $IndentSpaces
            $windowWidth = (Get-Host).UI.RawUI.WindowSize.Width
            if ($windowWidth -lt 1 -or $totalColumnWidth -lt $windowWidth) {
                return $colWidths
            }

            # Take size away from one or more columns to make them fit
            while ($totalColumnWidth -ge $windowWidth) {
                $startingTotalWidth = $totalColumnWidth
                $widest = $colWidths | Sort-Object -Descending | Select-Object -First 1
                $newWidest = [Math]::Floor($widest * 0.95)
                for ($i = 0; $i -lt $colWidths.Length; $i++) {
                    if ($colWidths[$i] -eq $widest) {
                        $colWidths[$i] = $newWidest
                        break
                    }
                }

                $totalColumnWidth = $colWidths.Length * $padding + ($colWidths | Measure-Object -Sum).Sum + $IndentSpaces
                if ($totalColumnWidth -ge $startingTotalWidth) {
                    # Somehow we didn't reduce the size at all, so give up
                    break
                }
            }

            return $colWidths
        }

        $objects = New-Object System.Collections.ArrayList
        $padding = 2
        $stb = New-Object System.Text.StringBuilder
    }

    process {
        foreach ($thing in $InputObject) {
            [void]$objects.Add($thing)
        }
    }

    end {
        if ($objects.Count -gt 0) {
            $props = $null

            if ($null -ne $Properties) {
                $props = $Properties
            } else {
                $props = $objects[0].PSObject.Properties.Name
            }

            $colWidths = GetColumnWidths $objects $props

            Write-Host
            [void]$stb.Append([System.Environment]::NewLine)

            Write-Host (" " * $IndentSpaces) -NoNewline
            [void]$stb.Append(" " * $IndentSpaces)

            for ($i = 0; $i -lt $props.Count; $i++) {
                Write-Host ("{0,$(-1 * ($colWidths[$i] + $padding))}" -f $props[$i]) -NoNewline
                [void]$stb.Append("{0,$(-1 * ($colWidths[$i] + $padding))}" -f $props[$i])
            }

            Write-Host
            [void]$stb.Append([System.Environment]::NewLine)

            Write-Host (" " * $IndentSpaces) -NoNewline
            [void]$stb.Append(" " * $IndentSpaces)

            for ($i = 0; $i -lt $props.Count; $i++) {
                Write-Host ("{0,$(-1 * ($colWidths[$i] + $padding))}" -f ("-" * $props[$i].Length)) -NoNewline
                [void]$stb.Append("{0,$(-1 * ($colWidths[$i] + $padding))}" -f ("-" * $props[$i].Length))
            }

            Write-Host
            [void]$stb.Append([System.Environment]::NewLine)

            foreach ($o in $objects) {
                $colColors = GetColumnColors -obj $o -props $props -functions $ColorizerFunctions
                $lineObjects = @(GetLineObjects -obj $o -props $props -colWidths $colWidths)
                foreach ($lineObj in $lineObjects) {
                    Write-Host (" " * $IndentSpaces) -NoNewline
                    [void]$stb.Append(" " * $IndentSpaces)
                    for ($i = 0; $i -lt $props.Count; $i++) {
                        $val = $o."$($props[$i])"
                        Write-Host ("{0,$(-1 * ($colWidths[$i] + $padding))}" -f $lineObj."$($props[$i])") -NoNewline -ForegroundColor $colColors[$i]
                        [void]$stb.Append("{0,$(-1 * ($colWidths[$i] + $padding))}" -f $lineObj."$($props[$i])")
                    }

                    Write-Host
                    [void]$stb.Append([System.Environment]::NewLine)
                }

                for ($i = 0; $i -lt $LinesBetweenObjects; $i++) {
                    Write-Host
                    [void]$stb.Append([System.Environment]::NewLine)
                }
            }

            Write-Host
            [void]$stb.Append([System.Environment]::NewLine)

            if ($null -ne $StringOutput) {
                $StringOutput.Value = $stb.ToString()
            }
        }
    }
}

# For HashSet support
Add-Type -AssemblyName System.Core -ErrorAction Stop

try {
    if (-not $SkipVersionCheck) {
        if (Test-ScriptVersion -AutoUpdate) {
            # Update was downloaded, so stop here.
            Write-Host "Script was updated. Please rerun the command."
            return
        }
    }

    $errorColor = "Red"
    $configuredErrorColor = (Get-Host).PrivateData.ErrorForegroundColor
    if ($configuredErrorColor -is [ConsoleColor]) {
        $errorColor = $configuredErrorColor
    }

    $warningColor = "Yellow"
    $configuredWarningColor = (Get-Host).PrivateData.WarningForegroundColor
    if ($configuredWarningColor -is [ConsoleColor]) {
        $warningColor = $configuredWarningColor
    }

    $severityColorizer = {
        param($o, $propName)
        if ($propName -eq "Severity") {
            switch ($o.$propName) {
                "Error" { $errorColor }
                "Warning" { $warningColor }
            }
        }
    }

    if ($SummarizePreviousResults) {
        $results = Import-Csv $ResultsFile
        $summary = New-Object System.Collections.ArrayList
        $summary.AddRange(@($results | Write-TestDumpsterMappingResult))
        $summary.AddRange(@($results | Write-TestFolderLimitResult))
        $summary.AddRange(@($results | Write-TestFolderNameResult))
        $summary.AddRange(@($results | Write-TestMailEnabledFolderResult))
        $summary.AddRange(@($results | Write-TestPermissionResult))
        $summary | Out-Columns -LinesBetweenObjects 1 -ColorizerFunctions $severityColorizer
        return
    }

    if ($RemoveInvalidPermissions) {
        if (-not (Test-Path $ResultsFile)) {
            Write-Error "File not found: $ResultsFile. Please specify -ResultsFile or run without -RemoveInvalidPermissions to generate a results file."
        } else {
            Import-Csv $ResultsFile | Remove-InvalidPermission
        }

        return
    }

    $startTime = Get-Date

    if ($null -eq (Get-Command Set-ADServerSettings -ErrorAction:SilentlyContinue)) {
        Write-Warning "Exchange Server cmdlets are not present in this shell."
        return
    }

    Set-ADServerSettings -ViewEntireForest $true

    $progressParams = @{
        Activity = "Validating public folders"
        Id       = 1
    }

    Write-Progress @progressParams -Status "Step 1 of 6"

    $folderData = Get-FolderData -StartFresh $StartFresh -SlowTraversal $SlowTraversal

    if ($folderData.IpmSubtree.Count -lt 1) {
        return
    }

    $script:anyDatabaseDown = $false
    Get-Mailbox -PublicFolder | ForEach-Object {
        try {
            $db = Get-MailboxDatabase $_.Database -Status
            if ($db.Mounted) {
                $folderData.MailboxToServerMap[$_.DisplayName] = $db.Server
            } else {
                Write-Error "Database $db is not mounted. This database holds PF mailbox $_ and must be mounted."
                $script:anyDatabaseDown = $true
            }
        } catch {
            Write-Error $_
            $script:anyDatabaseDown = $true
        }
    }

    if ($script:anyDatabaseDown) {
        Write-Host "One or more PF mailboxes cannot be reached. Unable to proceed."
        return
    }

    # Now we're ready to do the checks

    if (Test-Path $ResultsFile) {
        $directory = [System.IO.Path]::GetDirectoryName($ResultsFile)
        $fileName = [System.IO.Path]::GetFileNameWithoutExtension($ResultsFile)
        $timeString = (Get-Item $ResultsFile).LastWriteTime.ToString("yyMMdd-HHmm")
        Move-Item -Path $ResultsFile -Destination (Join-Path $directory "$($fileName)-$timeString.csv")
    }

    if ($folderData.Errors.Count -gt 0) {
        $folderData.Errors | Export-Csv $ResultsFile -NoTypeInformation
    }

    if ("Dumpsters" -in $Tests) {
        Write-Progress @progressParams -Status "Step 2 of 6"

        $badDumpsters = Test-DumpsterMapping -FolderData $folderData
        $badDumpsters | Export-Csv $ResultsFile -NoTypeInformation -Append
    }

    if ("Limits" -in $Tests) {
        Write-Progress @progressParams -Status "Step 3 of 6"

        # This test emits results in a weird order, so sort them.
        $limitsExceeded = Test-FolderLimit -FolderData $folderData | Sort-Object FolderIdentity
        $limitsExceeded | Export-Csv $ResultsFile -NoTypeInformation -Append
    }

    if ("Names" -in $Tests) {
        Write-Progress @progressParams -Status "Step 4 of 6"

        $badNames = Test-FolderName -FolderData $folderData
        $badNames | Export-Csv $ResultsFile -NoTypeInformation -Append
    }

    if ("MailEnabled" -in $Tests) {
        Write-Progress @progressParams -Status "Step 5 of 6"

        $badMailEnabled = Test-MailEnabledFolder -FolderData $folderData
        $badMailEnabled | Export-Csv $ResultsFile -NoTypeInformation -Append
    }

    if ("Permissions" -in $Tests) {
        Write-Progress @progressParams -Status "Step 6 of 6"

        $badPermissions = Test-Permission -FolderData $folderData
        $badPermissions | Export-Csv $ResultsFile -NoTypeInformation -Append
    }

    # Output the results

    $results = New-Object System.Collections.ArrayList
    $results.AddRange(@($badDumpsters | Write-TestDumpsterMappingResult))
    $results.AddRange(@($limitsExceeded | Write-TestFolderLimitResult))
    $results.AddRange(@($badNames | Write-TestFolderNameResult))
    $results.AddRange(@($badMailEnabled | Write-TestMailEnabledFolderResult))
    $results.AddRange(@($badPermissions | Write-TestPermissionResult))
    $results | Out-Columns -LinesBetweenObjects 1

    Write-Host
    Write-Host "Validation results were written to file:"
    Write-Host $ResultsFile -ForegroundColor Green

    $private:endTime = Get-Date

    Write-Host
    Write-Host "SourceSideValidations complete. Total duration" ($endTime - $startTime)
} finally {
    Write-Host
    Write-Host "Liked the script or had a problem? Let us know at ExToolsFeedback@microsoft.com"
}

# SIG # Begin signature block
# MIIoQgYJKoZIhvcNAQcCoIIoMzCCKC8CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDATnTPvhiAdwMC
# nBB3m6k8c9CNNQ190211F4Fxgu2OZKCCDXYwggX0MIID3KADAgECAhMzAAADTrU8
# esGEb+srAAAAAANOMA0GCSqGSIb3DQEBCwUAMH4xCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25p
# bmcgUENBIDIwMTEwHhcNMjMwMzE2MTg0MzI5WhcNMjQwMzE0MTg0MzI5WjB0MQsw
# CQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9u
# ZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMR4wHAYDVQQDExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24wggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIB
# AQDdCKiNI6IBFWuvJUmf6WdOJqZmIwYs5G7AJD5UbcL6tsC+EBPDbr36pFGo1bsU
# p53nRyFYnncoMg8FK0d8jLlw0lgexDDr7gicf2zOBFWqfv/nSLwzJFNP5W03DF/1
# 1oZ12rSFqGlm+O46cRjTDFBpMRCZZGddZlRBjivby0eI1VgTD1TvAdfBYQe82fhm
# WQkYR/lWmAK+vW/1+bO7jHaxXTNCxLIBW07F8PBjUcwFxxyfbe2mHB4h1L4U0Ofa
# +HX/aREQ7SqYZz59sXM2ySOfvYyIjnqSO80NGBaz5DvzIG88J0+BNhOu2jl6Dfcq
# jYQs1H/PMSQIK6E7lXDXSpXzAgMBAAGjggFzMIIBbzAfBgNVHSUEGDAWBgorBgEE
# AYI3TAgBBggrBgEFBQcDAzAdBgNVHQ4EFgQUnMc7Zn/ukKBsBiWkwdNfsN5pdwAw
# RQYDVR0RBD4wPKQ6MDgxHjAcBgNVBAsTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEW
# MBQGA1UEBRMNMjMwMDEyKzUwMDUxNjAfBgNVHSMEGDAWgBRIbmTlUAXTgqoXNzci
# tW2oynUClTBUBgNVHR8ETTBLMEmgR6BFhkNodHRwOi8vd3d3Lm1pY3Jvc29mdC5j
# b20vcGtpb3BzL2NybC9NaWNDb2RTaWdQQ0EyMDExXzIwMTEtMDctMDguY3JsMGEG
# CCsGAQUFBwEBBFUwUzBRBggrBgEFBQcwAoZFaHR0cDovL3d3dy5taWNyb3NvZnQu
# Y29tL3BraW9wcy9jZXJ0cy9NaWNDb2RTaWdQQ0EyMDExXzIwMTEtMDctMDguY3J0
# MAwGA1UdEwEB/wQCMAAwDQYJKoZIhvcNAQELBQADggIBAD21v9pHoLdBSNlFAjmk
# mx4XxOZAPsVxxXbDyQv1+kGDe9XpgBnT1lXnx7JDpFMKBwAyIwdInmvhK9pGBa31
# TyeL3p7R2s0L8SABPPRJHAEk4NHpBXxHjm4TKjezAbSqqbgsy10Y7KApy+9UrKa2
# kGmsuASsk95PVm5vem7OmTs42vm0BJUU+JPQLg8Y/sdj3TtSfLYYZAaJwTAIgi7d
# hzn5hatLo7Dhz+4T+MrFd+6LUa2U3zr97QwzDthx+RP9/RZnur4inzSQsG5DCVIM
# pA1l2NWEA3KAca0tI2l6hQNYsaKL1kefdfHCrPxEry8onJjyGGv9YKoLv6AOO7Oh
# JEmbQlz/xksYG2N/JSOJ+QqYpGTEuYFYVWain7He6jgb41JbpOGKDdE/b+V2q/gX
# UgFe2gdwTpCDsvh8SMRoq1/BNXcr7iTAU38Vgr83iVtPYmFhZOVM0ULp/kKTVoir
# IpP2KCxT4OekOctt8grYnhJ16QMjmMv5o53hjNFXOxigkQWYzUO+6w50g0FAeFa8
# 5ugCCB6lXEk21FFB1FdIHpjSQf+LP/W2OV/HfhC3uTPgKbRtXo83TZYEudooyZ/A
# Vu08sibZ3MkGOJORLERNwKm2G7oqdOv4Qj8Z0JrGgMzj46NFKAxkLSpE5oHQYP1H
# tPx1lPfD7iNSbJsP6LiUHXH1MIIHejCCBWKgAwIBAgIKYQ6Q0gAAAAAAAzANBgkq
# hkiG9w0BAQsFADCBiDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24x
# EDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlv
# bjEyMDAGA1UEAxMpTWljcm9zb2Z0IFJvb3QgQ2VydGlmaWNhdGUgQXV0aG9yaXR5
# IDIwMTEwHhcNMTEwNzA4MjA1OTA5WhcNMjYwNzA4MjEwOTA5WjB+MQswCQYDVQQG
# EwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwG
# A1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSgwJgYDVQQDEx9NaWNyb3NvZnQg
# Q29kZSBTaWduaW5nIFBDQSAyMDExMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIIC
# CgKCAgEAq/D6chAcLq3YbqqCEE00uvK2WCGfQhsqa+laUKq4BjgaBEm6f8MMHt03
# a8YS2AvwOMKZBrDIOdUBFDFC04kNeWSHfpRgJGyvnkmc6Whe0t+bU7IKLMOv2akr
# rnoJr9eWWcpgGgXpZnboMlImEi/nqwhQz7NEt13YxC4Ddato88tt8zpcoRb0Rrrg
# OGSsbmQ1eKagYw8t00CT+OPeBw3VXHmlSSnnDb6gE3e+lD3v++MrWhAfTVYoonpy
# 4BI6t0le2O3tQ5GD2Xuye4Yb2T6xjF3oiU+EGvKhL1nkkDstrjNYxbc+/jLTswM9
# sbKvkjh+0p2ALPVOVpEhNSXDOW5kf1O6nA+tGSOEy/S6A4aN91/w0FK/jJSHvMAh
# dCVfGCi2zCcoOCWYOUo2z3yxkq4cI6epZuxhH2rhKEmdX4jiJV3TIUs+UsS1Vz8k
# A/DRelsv1SPjcF0PUUZ3s/gA4bysAoJf28AVs70b1FVL5zmhD+kjSbwYuER8ReTB
# w3J64HLnJN+/RpnF78IcV9uDjexNSTCnq47f7Fufr/zdsGbiwZeBe+3W7UvnSSmn
# Eyimp31ngOaKYnhfsi+E11ecXL93KCjx7W3DKI8sj0A3T8HhhUSJxAlMxdSlQy90
# lfdu+HggWCwTXWCVmj5PM4TasIgX3p5O9JawvEagbJjS4NaIjAsCAwEAAaOCAe0w
# ggHpMBAGCSsGAQQBgjcVAQQDAgEAMB0GA1UdDgQWBBRIbmTlUAXTgqoXNzcitW2o
# ynUClTAZBgkrBgEEAYI3FAIEDB4KAFMAdQBiAEMAQTALBgNVHQ8EBAMCAYYwDwYD
# VR0TAQH/BAUwAwEB/zAfBgNVHSMEGDAWgBRyLToCMZBDuRQFTuHqp8cx0SOJNDBa
# BgNVHR8EUzBRME+gTaBLhklodHRwOi8vY3JsLm1pY3Jvc29mdC5jb20vcGtpL2Ny
# bC9wcm9kdWN0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFfMDNfMjIuY3JsMF4GCCsG
# AQUFBwEBBFIwUDBOBggrBgEFBQcwAoZCaHR0cDovL3d3dy5taWNyb3NvZnQuY29t
# L3BraS9jZXJ0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFfMDNfMjIuY3J0MIGfBgNV
# HSAEgZcwgZQwgZEGCSsGAQQBgjcuAzCBgzA/BggrBgEFBQcCARYzaHR0cDovL3d3
# dy5taWNyb3NvZnQuY29tL3BraW9wcy9kb2NzL3ByaW1hcnljcHMuaHRtMEAGCCsG
# AQUFBwICMDQeMiAdAEwAZQBnAGEAbABfAHAAbwBsAGkAYwB5AF8AcwB0AGEAdABl
# AG0AZQBuAHQALiAdMA0GCSqGSIb3DQEBCwUAA4ICAQBn8oalmOBUeRou09h0ZyKb
# C5YR4WOSmUKWfdJ5DJDBZV8uLD74w3LRbYP+vj/oCso7v0epo/Np22O/IjWll11l
# hJB9i0ZQVdgMknzSGksc8zxCi1LQsP1r4z4HLimb5j0bpdS1HXeUOeLpZMlEPXh6
# I/MTfaaQdION9MsmAkYqwooQu6SpBQyb7Wj6aC6VoCo/KmtYSWMfCWluWpiW5IP0
# wI/zRive/DvQvTXvbiWu5a8n7dDd8w6vmSiXmE0OPQvyCInWH8MyGOLwxS3OW560
# STkKxgrCxq2u5bLZ2xWIUUVYODJxJxp/sfQn+N4sOiBpmLJZiWhub6e3dMNABQam
# ASooPoI/E01mC8CzTfXhj38cbxV9Rad25UAqZaPDXVJihsMdYzaXht/a8/jyFqGa
# J+HNpZfQ7l1jQeNbB5yHPgZ3BtEGsXUfFL5hYbXw3MYbBL7fQccOKO7eZS/sl/ah
# XJbYANahRr1Z85elCUtIEJmAH9AAKcWxm6U/RXceNcbSoqKfenoi+kiVH6v7RyOA
# 9Z74v2u3S5fi63V4GuzqN5l5GEv/1rMjaHXmr/r8i+sLgOppO6/8MO0ETI7f33Vt
# Y5E90Z1WTk+/gFcioXgRMiF670EKsT/7qMykXcGhiJtXcVZOSEXAQsmbdlsKgEhr
# /Xmfwb1tbWrJUnMTDXpQzTGCGiIwghoeAgEBMIGVMH4xCzAJBgNVBAYTAlVTMRMw
# EQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVN
# aWNyb3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNp
# Z25pbmcgUENBIDIwMTECEzMAAANOtTx6wYRv6ysAAAAAA04wDQYJYIZIAWUDBAIB
# BQCggcYwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEO
# MAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEINukBqVVE6jlJdQvjdPZEVw/
# naR784u2qSbFnu+71nRDMFoGCisGAQQBgjcCAQwxTDBKoBqAGABDAFMAUwAgAEUA
# eABjAGgAYQBuAGcAZaEsgCpodHRwczovL2dpdGh1Yi5jb20vbWljcm9zb2Z0L0NT
# Uy1FeGNoYW5nZSAwDQYJKoZIhvcNAQEBBQAEggEAIxrTllVuHud34GTS/+tBplcv
# n1QO66/u9MzPQioi1SLElu+7wvz+RUzDrSPZ2rSB1QedxXp4OUAjsFNf8170MZZp
# xjD5taJybVnu7WPpILo4N83U72Zrhe/qySg1gqcCGgfHOxBxi05yrzZSk3STrFpa
# /r61c1M03NYgRMe46cYAlrUeVB6K3FN1WNP2hgyqaWNq2VPmoqtm0FuD+kygiQzC
# z8J/unygUSAnFTBvg2Ax6YB77BUQ3DWbSUTS60TEGo9UkxpztYqM1jh7PclUyOm2
# s4o0AtQyDPYayXNZoEzArrRbAS5C2z88GIicF+L6Rxjb6DSWIPqPDRcdIntZj6GC
# F5QwgheQBgorBgEEAYI3AwMBMYIXgDCCF3wGCSqGSIb3DQEHAqCCF20wghdpAgED
# MQ8wDQYJYIZIAWUDBAIBBQAwggFSBgsqhkiG9w0BCRABBKCCAUEEggE9MIIBOQIB
# AQYKKwYBBAGEWQoDATAxMA0GCWCGSAFlAwQCAQUABCCC74ssu0VvKbwqqReulQd9
# /HhmijHotQ980TexcS5a3wIGZNTIoA8cGBMyMDIzMDgxNjAwMDg0MC4xNDlaMASA
# AgH0oIHRpIHOMIHLMQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQ
# MA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9u
# MSUwIwYDVQQLExxNaWNyb3NvZnQgQW1lcmljYSBPcGVyYXRpb25zMScwJQYDVQQL
# Ex5uU2hpZWxkIFRTUyBFU046OEQwMC0wNUUwLUQ5NDcxJTAjBgNVBAMTHE1pY3Jv
# c29mdCBUaW1lLVN0YW1wIFNlcnZpY2WgghHqMIIHIDCCBQigAwIBAgITMwAAAc1V
# ByrnysGZHQABAAABzTANBgkqhkiG9w0BAQsFADB8MQswCQYDVQQGEwJVUzETMBEG
# A1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWlj
# cm9zb2Z0IENvcnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQgVGltZS1TdGFt
# cCBQQ0EgMjAxMDAeFw0yMzA1MjUxOTEyMDVaFw0yNDAyMDExOTEyMDVaMIHLMQsw
# CQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9u
# ZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSUwIwYDVQQLExxNaWNy
# b3NvZnQgQW1lcmljYSBPcGVyYXRpb25zMScwJQYDVQQLEx5uU2hpZWxkIFRTUyBF
# U046OEQwMC0wNUUwLUQ5NDcxJTAjBgNVBAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1w
# IFNlcnZpY2UwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQDTOCLVS2jm
# EWOqxzygW7s6YLmm29pjvA+Ch6VL7HlTL8yUt3Z0KIzTa2O/Hvr/aJza1qEVklq7
# NPiOrpBAIz657LVxwEc4BxJiv6B68a8DQiF6WAFFNaK3WHi7TfxRnqLohgNz7vZP
# ylZQX795r8MQvX56uwjj/R4hXnR7Na4Llu4mWsml/wp6VJqCuxZnu9jX4qaUxngc
# rfFT7+zvlXClwLah2n0eGKna1dOjOgyK00jYq5vtzr5NZ+qVxqaw9DmEsj9vfqYk
# fQZry2JO5wmgXX79Ox7PLMUfqT4+8w5JkdSMoX32b1D6cDKWRUv5qjiYh4o/a9eh
# E/KAkUWlSPbbDR/aGnPJLAGPy2qA97YCBeeIJjRKURgdPlhE5O46kOju8nYJnIvx
# buC2Qp2jxwc6rD9M6Pvc8sZIcQ10YKZVYKs94YPSlkhwXwttbRY+jZnQiDm2ZFjH
# 8SPe1I6ERcfeYX1zCYjEzdwWcm+fFZmlJA9HQW7ZJAmOECONtfK28EREEE5yzq+T
# 3QMVPhiEfEhgcYsh0DeoWiYGsDiKEuS+FElMMyT456+U2ZRa2hbRQ97QcbvaAd6O
# VQLp3TQqNEu0es5Zq0wg2CADf+QKQR/Y6+fGgk9qJNJW3Mu771KthuPlNfKss0B1
# zh0xa1yN4qC3zoE9Uq6T8r7G3/OtSFms4wIDAQABo4IBSTCCAUUwHQYDVR0OBBYE
# FKGT+aY2aZrBAJVIZh5kicokfNWaMB8GA1UdIwQYMBaAFJ+nFV0AXmJdg/Tl0mWn
# G1M1GelyMF8GA1UdHwRYMFYwVKBSoFCGTmh0dHA6Ly93d3cubWljcm9zb2Z0LmNv
# bS9wa2lvcHMvY3JsL01pY3Jvc29mdCUyMFRpbWUtU3RhbXAlMjBQQ0ElMjAyMDEw
# KDEpLmNybDBsBggrBgEFBQcBAQRgMF4wXAYIKwYBBQUHMAKGUGh0dHA6Ly93d3cu
# bWljcm9zb2Z0LmNvbS9wa2lvcHMvY2VydHMvTWljcm9zb2Z0JTIwVGltZS1TdGFt
# cCUyMFBDQSUyMDIwMTAoMSkuY3J0MAwGA1UdEwEB/wQCMAAwFgYDVR0lAQH/BAww
# CgYIKwYBBQUHAwgwDgYDVR0PAQH/BAQDAgeAMA0GCSqGSIb3DQEBCwUAA4ICAQBS
# qG3ppKIU+i/EMwwtotoxnKfw0SX/3T16EPbjwsAImWOZ5nLAbatopl8zFY841gb5
# eiL1j81h4DiEiXt+BJgHIA2LIhKhSscd79oMbr631DiEqf9X5LZR3V3KIYstU3K7
# f5Dk7tbobuHu+6fYM/gOx44sgRU7YQ+YTYHvv8k4mMnuiahJRlU/F2vavcHU5uhX
# i078K4nSRAPnWyX7gVi6iVMBBUF4823oPFznEcHup7VNGRtGe1xvnlMd1CuyxctM
# 8d/oqyTsxwlJAM5F/lDxnEWoSzAkad1nWvkaAeMV7+39IpXhuf9G3xbffKiyBnj3
# cQeiA4SxSwCdnx00RBlXS6r9tGDa/o9RS01FOABzKkP5CBDpm4wpKdIU74KtBH2s
# E5QYYn7liYWZr2f/U+ghTmdOEOPkXEcX81H4dRJU28Tj/gUZdwL81xah8Kn+cB7v
# M/Hs3/J8tF13ZPP+8NtX3vu4NrchHDJYgjOi+1JuSf+4jpF/pEEPXp9AusizmSmk
# BK4iVT7NwVtRnS1ts8qAGHGPg2HPa4b2u9meueUoqNVtMhbumI1y+d9ZkThNXBXz
# 2aItT2C99DM3T3qYqAUmvKUryVSpMLVpse4je5WN6VVlCDFKWFRH202YxEVWsZ5b
# aN9CaqCbCS0Ea7s9OFLaEM5fNn9m5s69lD/ekcW2qTCCB3EwggVZoAMCAQICEzMA
# AAAVxedrngKbSZkAAAAAABUwDQYJKoZIhvcNAQELBQAwgYgxCzAJBgNVBAYTAlVT
# MRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQK
# ExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xMjAwBgNVBAMTKU1pY3Jvc29mdCBSb290
# IENlcnRpZmljYXRlIEF1dGhvcml0eSAyMDEwMB4XDTIxMDkzMDE4MjIyNVoXDTMw
# MDkzMDE4MzIyNVowfDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24x
# EDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlv
# bjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBIDIwMTAwggIiMA0G
# CSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQDk4aZM57RyIQt5osvXJHm9DtWC0/3u
# nAcH0qlsTnXIyjVX9gF/bErg4r25PhdgM/9cT8dm95VTcVrifkpa/rg2Z4VGIwy1
# jRPPdzLAEBjoYH1qUoNEt6aORmsHFPPFdvWGUNzBRMhxXFExN6AKOG6N7dcP2CZT
# fDlhAnrEqv1yaa8dq6z2Nr41JmTamDu6GnszrYBbfowQHJ1S/rboYiXcag/PXfT+
# jlPP1uyFVk3v3byNpOORj7I5LFGc6XBpDco2LXCOMcg1KL3jtIckw+DJj361VI/c
# +gVVmG1oO5pGve2krnopN6zL64NF50ZuyjLVwIYwXE8s4mKyzbnijYjklqwBSru+
# cakXW2dg3viSkR4dPf0gz3N9QZpGdc3EXzTdEonW/aUgfX782Z5F37ZyL9t9X4C6
# 26p+Nuw2TPYrbqgSUei/BQOj0XOmTTd0lBw0gg/wEPK3Rxjtp+iZfD9M269ewvPV
# 2HM9Q07BMzlMjgK8QmguEOqEUUbi0b1qGFphAXPKZ6Je1yh2AuIzGHLXpyDwwvoS
# CtdjbwzJNmSLW6CmgyFdXzB0kZSU2LlQ+QuJYfM2BjUYhEfb3BvR/bLUHMVr9lxS
# UV0S2yW6r1AFemzFER1y7435UsSFF5PAPBXbGjfHCBUYP3irRbb1Hode2o+eFnJp
# xq57t7c+auIurQIDAQABo4IB3TCCAdkwEgYJKwYBBAGCNxUBBAUCAwEAATAjBgkr
# BgEEAYI3FQIEFgQUKqdS/mTEmr6CkTxGNSnPEP8vBO4wHQYDVR0OBBYEFJ+nFV0A
# XmJdg/Tl0mWnG1M1GelyMFwGA1UdIARVMFMwUQYMKwYBBAGCN0yDfQEBMEEwPwYI
# KwYBBQUHAgEWM2h0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2lvcHMvRG9jcy9S
# ZXBvc2l0b3J5Lmh0bTATBgNVHSUEDDAKBggrBgEFBQcDCDAZBgkrBgEEAYI3FAIE
# DB4KAFMAdQBiAEMAQTALBgNVHQ8EBAMCAYYwDwYDVR0TAQH/BAUwAwEB/zAfBgNV
# HSMEGDAWgBTV9lbLj+iiXGJo0T2UkFvXzpoYxDBWBgNVHR8ETzBNMEugSaBHhkVo
# dHRwOi8vY3JsLm1pY3Jvc29mdC5jb20vcGtpL2NybC9wcm9kdWN0cy9NaWNSb29D
# ZXJBdXRfMjAxMC0wNi0yMy5jcmwwWgYIKwYBBQUHAQEETjBMMEoGCCsGAQUFBzAC
# hj5odHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpL2NlcnRzL01pY1Jvb0NlckF1
# dF8yMDEwLTA2LTIzLmNydDANBgkqhkiG9w0BAQsFAAOCAgEAnVV9/Cqt4SwfZwEx
# JFvhnnJL/Klv6lwUtj5OR2R4sQaTlz0xM7U518JxNj/aZGx80HU5bbsPMeTCj/ts
# 0aGUGCLu6WZnOlNN3Zi6th542DYunKmCVgADsAW+iehp4LoJ7nvfam++Kctu2D9I
# dQHZGN5tggz1bSNU5HhTdSRXud2f8449xvNo32X2pFaq95W2KFUn0CS9QKC/GbYS
# EhFdPSfgQJY4rPf5KYnDvBewVIVCs/wMnosZiefwC2qBwoEZQhlSdYo2wh3DYXMu
# LGt7bj8sCXgU6ZGyqVvfSaN0DLzskYDSPeZKPmY7T7uG+jIa2Zb0j/aRAfbOxnT9
# 9kxybxCrdTDFNLB62FD+CljdQDzHVG2dY3RILLFORy3BFARxv2T5JL5zbcqOCb2z
# AVdJVGTZc9d/HltEAY5aGZFrDZ+kKNxnGSgkujhLmm77IVRrakURR6nxt67I6Ile
# T53S0Ex2tVdUCbFpAUR+fKFhbHP+CrvsQWY9af3LwUFJfn6Tvsv4O+S3Fb+0zj6l
# MVGEvL8CwYKiexcdFYmNcP7ntdAoGokLjzbaukz5m/8K6TT4JDVnK+ANuOaMmdbh
# IurwJ0I9JZTmdHRbatGePu1+oDEzfbzL6Xu/OHBE0ZDxyKs6ijoIYn/ZcGNTTY3u
# gm2lBRDBcQZqELQdVTNYs6FwZvKhggNNMIICNQIBATCB+aGB0aSBzjCByzELMAkG
# A1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQx
# HjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjElMCMGA1UECxMcTWljcm9z
# b2Z0IEFtZXJpY2EgT3BlcmF0aW9uczEnMCUGA1UECxMeblNoaWVsZCBUU1MgRVNO
# OjhEMDAtMDVFMC1EOTQ3MSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBT
# ZXJ2aWNloiMKAQEwBwYFKw4DAhoDFQBoqfem2KKzuRZjISYifGolVOdyBKCBgzCB
# gKR+MHwxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQH
# EwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJjAkBgNV
# BAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQSAyMDEwMA0GCSqGSIb3DQEBCwUA
# AgUA6IaHDDAiGA8yMDIzMDgxNTIzMjE0OFoYDzIwMjMwODE2MjMyMTQ4WjB0MDoG
# CisGAQQBhFkKBAExLDAqMAoCBQDohocMAgEAMAcCAQACAiDnMAcCAQACAhMyMAoC
# BQDoh9iMAgEAMDYGCisGAQQBhFkKBAIxKDAmMAwGCisGAQQBhFkKAwKgCjAIAgEA
# AgMHoSChCjAIAgEAAgMBhqAwDQYJKoZIhvcNAQELBQADggEBABUvGq9AH8z/5GLW
# 4YbnaeBtgUyTcx3+pipj9HCx3TkRFF/nYmVhSSZnptizjRFHRTcIiIqO3x+rC+eE
# cKnApk8KIGMqakoytZpAIhovKklzfMKe25J9mPE8cuJqMNf0cwNttvN+aI/8YhUS
# h+9Mfdb0NdSK9oW/nbuvTgnvGKMPOyTwRdaYs6O47rjPWzIZHhX9vQf8bNiHuheg
# yUMQbEIdqdhlVEXJdhenaXPfqhXHiRC559LX9qFF0+hY16HSxGLcIYUo0UjTKsWx
# q4dzzROuE/NuULGTJxnxM9XEBisnrjEmjtQTdqIimZPYRWbM9oGeX1mosWM1//OY
# oLs3OAQxggQNMIIECQIBATCBkzB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2Fz
# aGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENv
# cnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAx
# MAITMwAAAc1VByrnysGZHQABAAABzTANBglghkgBZQMEAgEFAKCCAUowGgYJKoZI
# hvcNAQkDMQ0GCyqGSIb3DQEJEAEEMC8GCSqGSIb3DQEJBDEiBCBzNk/YK1FsRKK6
# HgFF5YswE7WFvJzU303PlQvf37BXJTCB+gYLKoZIhvcNAQkQAi8xgeowgecwgeQw
# gb0EIOJmpfitVr1PZGgvTEdTpStUc6GNh7LNroQBKwpURpkKMIGYMIGApH4wfDEL
# MAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1v
# bmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWlj
# cm9zb2Z0IFRpbWUtU3RhbXAgUENBIDIwMTACEzMAAAHNVQcq58rBmR0AAQAAAc0w
# IgQgOY6jS98hyAF8SB0B9WnrF7F9HIf+r86pM47N3Ko5okowDQYJKoZIhvcNAQEL
# BQAEggIAw6khhH9c+DBehOHHtW5XywU7//uB8+gtS1+vB8moqlTo3x8d3RJeUogE
# KQnkUzffwH2vzfW03Ih9JU5rN2iN+pxLPsPlFOiYbAfL38LKQi81nT9NusxUuflg
# QvZUTGkEworEpsmuW97HHswN0dyn8STfwSFj+dLEUxy2QI6X7PivwucvQ4eeFZ/y
# cEmAnJ5XQ+s0oxff71y/PbFcMzgueljB0OzZXONKDqQXDTK33JvwaPbpv5KYBK55
# 1PaLpHkSslpgL3g8PrO6FRYOhaHPzEIthKzhTVHAmvE6d0eAOvogROP48ACAWYoZ
# h3TrHc7/k7opqo+tpZwc88XfOsCUvkUnQHfXbCrrXbMhagRxBrMxD5XWvpCWAe8b
# KVqWhDZ6gHE3EYqV/9m6vzV8TrCLZbkNhzw5j4UN+5VA0NKp0A5vBwH45x/X5c/Y
# uUO0wOIsKwN2OcA7LarPqzvCxOihvW//h/VI7sDXL2ZfodRmr3my8e4c/vPGAx2O
# W+JoF2rwAk5xOw6yeTRFjzZLL3VEp4xG3ZlEV/0naXxvQ2m6k6MC6bvLZe5izM/Q
# IV7SrHEJj1SB9PkEjqu7t6Xa+H1QtcnqY/2iUxhpTDRPQKqhUyFzsLzr9uZygK6U
# CjW1e2AJnKRmbTncOFqiIBp7DFD9hhHiFDaFTHeM43/rvHeh5gQ=
# SIG # End signature block
