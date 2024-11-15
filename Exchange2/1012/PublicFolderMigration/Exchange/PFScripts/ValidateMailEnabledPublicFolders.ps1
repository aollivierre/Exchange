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

# Version 23.01.13.1832

# ValidateMailEnabledPublicFolders.ps1
#
# Note: If running on Exchange 2010, the ExFolders tool must be in the V14\bin folder
# in order to run one of the checks. ExFolders can be downloaded here:
# https://techcommunity.microsoft.com/gxcuf89792/attachments/gxcuf89792/Exchange/12412/2/ExFolders-SP1+.zip

Set-ADServerSettings -ViewEntireForest $true

Write-Host "Checking for mail-enabled System folders..."

$nonIpmSubtreeMailEnabled = @(Get-PublicFolder \non_ipm_subtree -Recurse -ResultSize Unlimited | Where-Object { $_.MailEnabled })

Write-Host "Found $($nonIpmSubtreeMailEnabled.Count) mail-enabled System folders."

Write-Host "Getting all public folders. This might take a while..."

$allIpmSubtree = @(Get-PublicFolder -Recurse -ResultSize Unlimited | Select-Object Identity, MailEnabled, EntryId)

Write-Host "Found $($allIpmSubtree.Count) public folders."

if ($allIpmSubtree.Count -lt 1) {
    return
}

$ipmSubtreeMailEnabled = @($allIpmSubtree | Where-Object { $_.MailEnabled })

Write-Host "$($ipmSubtreeMailEnabled.Count) of those are mail-enabled."

$mailDisabledWithProxyGuid = $null

if ($null -ne (Get-PublicFolder).DumpsterEntryId) {
    $mailDisabledWithProxyGuid = @($allIpmSubtree | Where-Object { -not $_.MailEnabled -and $null -ne $_.MailRecipientGuid -and [Guid]::Empty -ne $_.MailRecipientGuid } | ForEach-Object { $_.Identity.ToString() })
} else {
    $registryPath = "HKCU:\Software\Microsoft\Exchange\ExFolders"
    $valueName = "PublicFolderPropertiesSelected"
    $value = @("PR_PF_PROXY: 0x671D0102", "PR_PF_PROXY_REQUIRED: 0x671F000B", "DS:legacyExchangeDN")

    if (-not (Test-Path $registryPath)) {
        New-Item -Path $registryPath -Force | Out-Null
    }

    if (-not (Test-Path $registryPath)) {
        Write-Error "Could not create ExFolders registry key."
        return
    }

    New-ItemProperty -Path $registryPath -Name $valueName -Value $value -PropertyType MultiString -Force | Out-Null

    $result = (Get-ItemProperty -Path $registryPath -Name $valueName).PublicFolderPropertiesSelected

    if ($result[0] -ne $value[0] -or $result[1] -ne $value[1] -or $result[2] -ne $value[2]) {
        Write-Error "Could not set PublicFolderPropertiesSelected value for ExFolders in the registry."
        return
    }

    $msiInstallPath = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\ExchangeServer\v14\Setup" -Name "MsiInstallPath").MsiInstallPath
    $exFoldersExe = "$msiInstallPath\bin\ExFolders.exe"

    $pfDbToUse = Get-PublicFolderDatabase | Select-Object -First 1

    Write-Host "Generating ExFolders export."
    Write-Warning "NOTE: ExFolders will appear to be Not Responding during the export. That is normal."
    Write-Host "Waiting for export to finish..."

    $exFoldersExportFile = (Join-Path $PWD "ExFoldersMailEnabledPropertyExport.txt")

    & $exFoldersExe -ConnectTo $pfDbToUse.Name -export -properties -f (Join-Path $PWD "ExFoldersMailEnabledPropertyExport.txt") | Out-Null

    if (-not (Test-Path $exFoldersExportFile)) {
        Write-Error "Failed to generate ExFolders export. Continuing with other tests. If"
        Write-Error "any mail-disabled folders have invalid proxy GUIDs, those will be missed."
    } else {
        $exportResults = Import-Csv .\ExFoldersMailEnabledPropertyExport.txt -Delimiter `t
        $mailDisabledWithProxyGuid = @($exportResults | Where-Object { $_."PR_PF_PROXY_REQUIRED: 0x671F000B" -ne "True" -and $_."PR_PF_PROXY: 0x671D0102" -ne "PropertyError: NotFound" -and $_."DS:legacyExchangeDN".length -lt 1 } | ForEach-Object { $_."Folder Path" })
    }
}

$mailEnabledFoldersWithNoADObject = @()

$mailPublicFoldersLinked = New-Object 'System.Collections.Generic.Dictionary[string, object]'

for ($i = 0; $i -lt $ipmSubtreeMailEnabled.Count; $i++) {
    Write-Progress -Activity "Checking for missing AD objects" -PercentComplete ($i * 100 / $ipmSubtreeMailEnabled.Count) -Status ("$i of $($ipmSubtreeMailEnabled.Count)")
    $result = $ipmSubtreeMailEnabled[$i] | Get-MailPublicFolder -ErrorAction SilentlyContinue
    if ($null -eq $result) {
        $mailEnabledFoldersWithNoADObject += $ipmSubtreeMailEnabled[$i]
    } else {
        $guidString = $result.Guid.ToString()
        if (-not $mailPublicFoldersLinked.ContainsKey($guidString)) {
            $mailPublicFoldersLinked.Add($guidString, $result) | Out-Null
        }
    }
}

Write-Host "$($mailEnabledFoldersWithNoADObject.Count) folders are mail-enabled with no AD object."

Write-Host "$($mailPublicFoldersLinked.Keys.Count) folders are mail-enabled and are properly linked to an existing AD object."

Write-Host "Getting all MailPublicFolder objects..."

$allMailPublicFolders = @(Get-MailPublicFolder -ResultSize Unlimited)

$orphanedMailPublicFolders = @()

for ($i = 0; $i -lt $allMailPublicFolders.Count; $i++) {
    Write-Progress -Activity "Checking for orphaned MailPublicFolders" -PercentComplete ($i * 100 / $allMailPublicFolders.Count) -Status ("$i of $($allMailPublicFolders.Count)")
    if (!($mailPublicFoldersLinked.ContainsKey($allMailPublicFolders[$i].Guid.ToString()))) {
        $orphanedMailPublicFolders += $allMailPublicFolders[$i]
    }
}

Write-Host "$($orphanedMailPublicFolders.Count) MailPublicFolders are orphaned."

Write-Host "Building EntryId HashSets..."

$byEntryId = New-Object 'System.Collections.Generic.Dictionary[string, object]'
$allIpmSubtree | ForEach-Object { $byEntryId.Add($_.EntryId.ToString(), $_) }

$byPartialEntryId = New-Object 'System.Collections.Generic.Dictionary[string, object]'
$allIpmSubtree | ForEach-Object { $byPartialEntryId.Add($_.EntryId.ToString().Substring(44), $_) }

$orphanedMPFsThatPointToAMailDisabledFolder = @()
$orphanedMPFsThatPointToAMailEnabledFolder = @()
$orphanedMPFsThatPointToNothing = @()
$emailAddressMergeCommands = @()

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

for ($i = 0; $i -lt $orphanedMailPublicFolders.Count; $i++) {
    Write-Progress -Activity "Checking for orphans that point to a valid folder" -PercentComplete ($i * 100 / $orphanedMailPublicFolders.Count) -Status ("$i of $($orphanedMailPublicFolders.Count)")
    $thisMPF = $orphanedMailPublicFolders[$i]
    $pf = $null
    if ($null -ne $thisMPF.ExternalEmailAddress -and $thisMPF.ExternalEmailAddress.ToString().StartsWith("exPf".ToLower())) {
        $partialEntryId = $thisMPF.ExternalEmailAddress.ToString().Substring(5).Replace("-", "")
        $partialEntryId += "0000"
        if ($byPartialEntryId.TryGetValue($partialEntryId, [ref]$pf)) {
            if ($pf.MailEnabled) {

                $command = GetCommandToMergeEmailAddresses $pf $thisMPF
                if ($null -ne $command) {
                    $emailAddressMergeCommands += $command
                }

                $orphanedMPFsThatPointToAMailEnabledFolder += $thisMPF
            } else {
                $orphanedMPFsThatPointToAMailDisabledFolder += $thisMPF
            }

            continue
        }
    }

    if ($null -ne $thisMPF.EntryId -and $byEntryId.TryGetValue($thisMPF.EntryId.ToString(), [ref]$pf)) {
        if ($pf.MailEnabled) {

            $command = GetCommandToMergeEmailAddresses $pf $thisMPF
            if ($null -ne $command) {
                $emailAddressMergeCommands += $command
            }

            $orphanedMPFsThatPointToAMailEnabledFolder += $thisMPF
        } else {
            $orphanedMPFsThatPointToAMailDisabledFolder += $thisMPF
        }
    } else {
        $orphanedMPFsThatPointToNothing += $thisMPF
    }
}

Write-Host $orphanedMailPublicFolders.Count "orphaned MailPublicFolder objects."
Write-Host $orphanedMPFsThatPointToAMailEnabledFolder.Count "of those orphans point to mail-enabled folders that point to some other object."
Write-Host $orphanedMPFsThatPointToAMailDisabledFolder.Count "of those orphans point to mail-disabled folders."

$foldersToMailDisableFile = Join-Path $PWD "FoldersToMailDisable.txt"
$foldersToMailDisable = @()
$nonIpmSubtreeMailEnabled | ForEach-Object { $foldersToMailDisable += $_.Identity.ToString() }
$mailEnabledFoldersWithNoADObject | ForEach-Object { $foldersToMailDisable += $_.Identity }

if ($foldersToMailDisable.Count -gt 0) {
    Set-Content -Path $foldersToMailDisableFile -Value $foldersToMailDisable

    Write-Host
    Write-Host "Results:"
    Write-Host
    Write-Host $foldersToMailDisable.Count "folders should be mail-disabled, either because the MailRecipientGuid"
    Write-Host "does not exist, or because they are system folders. These are listed in the file called:"
    Write-Host $foldersToMailDisableFile -ForegroundColor Green
    if ($null -ne $allIpmSubtree[0].DumpsterEntryId) {
        # This is modern public folders, which means we can just toggle the attribute
        Write-Host "After confirming the accuracy of the results, you can mail-disable them with the following command:"
        Write-Host "Get-Content `"$foldersToMailDisableFile`" | % { Set-PublicFolder `$_ -MailEnabled `$false }" -ForegroundColor Green
    } else {
        # This is 2010. We can just mail-disable.
        Write-Host "After confirming the accuracy of the results, you can mail-disable them with the following command:"
        Write-Host "Get-Content `"$foldersToMailDisableFile`" | % { Disable-MailPublicFolder `$_ }" -ForegroundColor Green
    }
} else {
    Write-Host
    Write-Host "No folders need to be mail-disabled."
}

$mailPublicFoldersToDeleteFile = Join-Path $PWD "MailPublicFolderOrphans.txt"
$mailPublicFoldersToDelete = @()
$orphanedMPFsThatPointToNothing | ForEach-Object { $mailPublicFoldersToDelete += $_.DistinguishedName.Replace("/", "\/") }

if ($orphanedMPFsThatPointToNothing.Count -gt 0) {
    Set-Content -Path $mailPublicFoldersToDeleteFile -Value $mailPublicFoldersToDelete

    Write-Host
    Write-Host $mailPublicFoldersToDelete.Count "MailPublicFolders are orphans and should be deleted. They exist in Active Directory"
    Write-Host "but are not linked to any public folder. These are listed in a file called:"
    Write-Host $mailPublicFoldersToDeleteFile -ForegroundColor Green
    Write-Host "After confirming the accuracy of the results, you can delete them with the following command:"
    Write-Host "Get-Content `"$mailPublicFoldersToDeleteFile`" | % { `$folder = ([ADSI](`"LDAP://`$_`")); `$parent = ([ADSI]`"`$(`$folder.Parent)`"); `$parent.Children.Remove(`$folder) }" -ForegroundColor Green
} else {
    Write-Host
    Write-Host "No orphaned MailPublicFolders were found."
}

$mailPublicFolderDuplicatesFile = Join-Path $PWD "MailPublicFolderDuplicates.txt"
$mailPublicFolderDuplicates = @()
$orphanedMPFsThatPointToAMailEnabledFolder | ForEach-Object { $mailPublicFolderDuplicates += $_.DistinguishedName }

if ($orphanedMPFsThatPointToAMailEnabledFolder.Count -gt 0) {
    Set-Content -Path $mailPublicFolderDuplicatesFile -Value $mailPublicFolderDuplicates

    Write-Host
    Write-Host $mailPublicFolderDuplicates.Count "MailPublicFolders are duplicates and should be deleted. They exist in Active Directory"
    Write-Host "and point to a valid folder, but that folder points to some other directory object."
    Write-Host "These are listed in a file called:"
    Write-Host $mailPublicFolderDuplicatesFile -ForegroundColor Green
    Write-Host "After confirming the accuracy of the results, you can delete them with the following command:"
    Write-Host "Get-Content `"$mailPublicFolderDuplicatesFile`" | % { `$folder = ([ADSI](`"LDAP://`$_`")); `$parent = ([ADSI]`"`$(`$folder.Parent)`"); `$parent.Children.Remove(`$folder) }" -ForegroundColor Green

    if ($emailAddressMergeCommands.Count -gt 0) {
        $emailAddressMergeScriptFile = Join-Path $PWD "AddAddressesFromDuplicates.ps1"
        Set-Content -Path $emailAddressMergeScriptFile -Value $emailAddressMergeCommands
        Write-Host "The duplicates we are deleting contain email addresses that might still be in use."
        Write-Host "To preserve these, we generated a script that will add these to the linked objects for those folders."
        Write-Host "After deleting the duplicate objects using the command above, run the script as follows to"
        Write-Host "populate these addresses:"
        Write-Host ".\$emailAddressMergeScriptFile" -ForegroundColor Green
    }
} else {
    Write-Host
    Write-Host "No duplicate MailPublicFolders were found."
}

$mailDisabledWithProxyGuidFile = Join-Path $PWD "MailDisabledWithProxyGuid.txt"

if ($mailDisabledWithProxyGuid.Count -gt 0) {
    Set-Content -Path $mailDisabledWithProxyGuidFile -Value $mailDisabledWithProxyGuid

    Write-Host
    Write-Host $mailDisabledWithProxyGuid.Count "public folders have proxy GUIDs even though the folders are mail-disabled."
    Write-Host "These folders should be mail-enabled. They can be mail-disabled again afterwards if desired."
    Write-Host "To mail-enable these folders, run:"
    Write-Host "Get-Content `"$mailDisabledWithProxyGuidFile`" | % { Enable-MailPublicFolder `$_ }" -ForegroundColor Green
} else {
    Write-Host
    Write-Host "No mail-disabled public folders with proxy GUIDs were found."
}

$mailPublicFoldersDisconnectedFile = Join-Path $PWD "MailPublicFoldersDisconnected.txt"
$mailPublicFoldersDisconnected = @()
$orphanedMPFsThatPointToAMailDisabledFolder | ForEach-Object { $mailPublicFoldersDisconnected += $_.DistinguishedName }

if ($orphanedMPFsThatPointToAMailDisabledFolder.Count -gt 0) {
    Set-Content -Path $mailPublicFoldersDisconnectedFile -Value $mailPublicFoldersDisconnected

    Write-Host
    Write-Host $mailPublicFoldersDisconnected.Count "MailPublicFolders are disconnected from their folders. This means they exist in"
    Write-Host "Active Directory and the folders are probably functioning as mail-enabled folders,"
    Write-Host "even while the properties of the public folders themselves say they are not mail-enabled."
    Write-Host "This can be complex to fix. Either the directory object should be deleted, or the public folder"
    Write-Host "should be mail-enabled, or both. These directory objects are listed in a file called:"
    Write-Host $mailPublicFoldersDisconnectedFile -ForegroundColor Green
} else {
    Write-Host
    Write-Host "No disconnected MailPublicFolders were found."
}

Write-Host
Write-Host "Done!"

# SIG # Begin signature block
# MIIoQgYJKoZIhvcNAQcCoIIoMzCCKC8CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCsJ6fSaxWIGUjM
# o/owrUI0wYUJBTi9weXvQJGMe2mcBqCCDXYwggX0MIID3KADAgECAhMzAAADTrU8
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
# MAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEIH/v/YMzGOoLSaSFpswAy7wc
# MDU0oGBQbloeeMHPH65NMFoGCisGAQQBgjcCAQwxTDBKoBqAGABDAFMAUwAgAEUA
# eABjAGgAYQBuAGcAZaEsgCpodHRwczovL2dpdGh1Yi5jb20vbWljcm9zb2Z0L0NT
# Uy1FeGNoYW5nZSAwDQYJKoZIhvcNAQEBBQAEggEAawILPCRBjbSijSdSCf4yDHRH
# DLwBf62y0zaDHG/mVIgdRsQWJKvrp8haBhKCHVVtJAIADOoKboPNvQJ2W5uqe6F6
# 1xlX4R4wKxORLTU3YTW/XvgsVawkhgzqiREUn9xfREpb1q0Jn/DocI8bv00f9EyS
# z8rrbBVYdv4Pz8z4dSuIpcrJM+xy80RIlpBlqj2PwV7Q+P1yUjF1wdN1U1V0v+22
# aezAv610SffflC8fPm5cpQNg4M5Ly2iOBdqb6cxdzkGIkeSvY9l17TAotEzVUUX2
# zEVz1szB8OowUJC2pDRZpIvTpEHqK2dE5O+hQgVd/TL9tIGWcy0tvgfkjuY1uqGC
# F5QwgheQBgorBgEEAYI3AwMBMYIXgDCCF3wGCSqGSIb3DQEHAqCCF20wghdpAgED
# MQ8wDQYJYIZIAWUDBAIBBQAwggFSBgsqhkiG9w0BCRABBKCCAUEEggE9MIIBOQIB
# AQYKKwYBBAGEWQoDATAxMA0GCWCGSAFlAwQCAQUABCB21/3hYCibreSzNEtpD8TF
# h2sITY+W04tFA/MQyr0rDQIGZNTI8wCNGBMyMDIzMDgxNjAwMDg0MC4wODRaMASA
# AgH0oIHRpIHOMIHLMQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQ
# MA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9u
# MSUwIwYDVQQLExxNaWNyb3NvZnQgQW1lcmljYSBPcGVyYXRpb25zMScwJQYDVQQL
# Ex5uU2hpZWxkIFRTUyBFU046QTkzNS0wM0UwLUQ5NDcxJTAjBgNVBAMTHE1pY3Jv
# c29mdCBUaW1lLVN0YW1wIFNlcnZpY2WgghHqMIIHIDCCBQigAwIBAgITMwAAAdGy
# W0AobC7SRQABAAAB0TANBgkqhkiG9w0BAQsFADB8MQswCQYDVQQGEwJVUzETMBEG
# A1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWlj
# cm9zb2Z0IENvcnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQgVGltZS1TdGFt
# cCBQQ0EgMjAxMDAeFw0yMzA1MjUxOTEyMThaFw0yNDAyMDExOTEyMThaMIHLMQsw
# CQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9u
# ZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSUwIwYDVQQLExxNaWNy
# b3NvZnQgQW1lcmljYSBPcGVyYXRpb25zMScwJQYDVQQLEx5uU2hpZWxkIFRTUyBF
# U046QTkzNS0wM0UwLUQ5NDcxJTAjBgNVBAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1w
# IFNlcnZpY2UwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQCZTNo0OeGz
# 2XFd2gLg5nTlBm8XOpuwJIiXsMU61rwq1ZKDpa443RrSG/pH8Gz6XNnFQKGnCqNC
# tmvoKULApwrT/s7/e1X0lNFKmj7U7X4p00S0uQbW6LwSn/zWHaG2c54ZXsGY+BYf
# hWDgbFpCTxRzTnRCG62bkWPp6ZHbZPg4Ht1CRCAMhhOGTR8wI4G7wwWZwdMc6UvU
# Ulq0ql9AxAfzkYRpi2tRvDHMdmZ3vyXpqhFwvRG8cgCH/TTCjW5q6aNbdqKL3BFD
# PzUtuCNsPXL3/E0dR2bDMqa0aNH+iIfhGC4/vcwuteOMCPUIDVSqDCNfIaPDEwYc
# i1fd9gu1zVw+HEhDZM7Ea3nxIUrzt+Rfp5ToMMj4QAmJ6Uadm+TPbDbo8kFIK70S
# hmW8wn8fJk9ReQQEpTtIN43eRv9QmXy3Ued80osOBE+WkdMvSCFh+qgCsKdzQxQJ
# G62cTeoU2eqNhH3oppXmyfVUwbsefQzMPtbinCZd0FUlmlM/dH+4OniqQyaHvrtY
# y3wqIafY3zeFITlVAoP9q9vF4W7KHR/uF0mvTpAL5NaTDN1plQS0MdjMkgzZK5gt
# wqOe/3rTlqBzxwa7YYp3urP5yWkTzISGnhNWIZOxOyQIOxZfbiIbAHbm3M8hj73K
# QWcCR5JavgkwUmncFHESaQf4Drqs+/1L1QIDAQABo4IBSTCCAUUwHQYDVR0OBBYE
# FAuO8UzF7DcH0mmsF4XQxxHQvS2jMB8GA1UdIwQYMBaAFJ+nFV0AXmJdg/Tl0mWn
# G1M1GelyMF8GA1UdHwRYMFYwVKBSoFCGTmh0dHA6Ly93d3cubWljcm9zb2Z0LmNv
# bS9wa2lvcHMvY3JsL01pY3Jvc29mdCUyMFRpbWUtU3RhbXAlMjBQQ0ElMjAyMDEw
# KDEpLmNybDBsBggrBgEFBQcBAQRgMF4wXAYIKwYBBQUHMAKGUGh0dHA6Ly93d3cu
# bWljcm9zb2Z0LmNvbS9wa2lvcHMvY2VydHMvTWljcm9zb2Z0JTIwVGltZS1TdGFt
# cCUyMFBDQSUyMDIwMTAoMSkuY3J0MAwGA1UdEwEB/wQCMAAwFgYDVR0lAQH/BAww
# CgYIKwYBBQUHAwgwDgYDVR0PAQH/BAQDAgeAMA0GCSqGSIb3DQEBCwUAA4ICAQCb
# u9rTAHV24mY0qoG5eEnImz5akGXTviBwKp2Y51s26w8oDrWor+m00R4/3BcDmYlU
# K8Nrx/auYFYidZddcUjw42QxSStmv/qWnCQi/2OnH32KVHQ+kMOZPABQTG1XkcnY
# PUOOEEor6f/3Js1uj4wjHzE4V4aumYXBAsr4L5KR8vKes5tFxhMkWND/O7W/RaHY
# wJMjMkxVosBok7V21sJAlxScEXxfJa+/qkqUr7CZgw3R4jCHRkPqQhMWibXPMYar
# /iF0ZuLB9O89DMJNhjK9BSf6iqgZoMuzIVt+EBoTzpv/9p4wQ6xoBCs29mkj/EIW
# Fdc+5a30kuCQOSEOj07+WI29A4k6QIRB5w+eMmZ0Jec0sSyeQB5KjxE51iYMhtlM
# rUKcr06nBqCsSKPYsSAITAzgssJD+Z/cTS7Cu35fJrWhM9NYX24uAxYLAW0ipNtW
# ptIeV6akuZEeEV6BNtM3VTk+mAlV5/eC/0Y17aVSjK5/gyDoLNmrgVwv5TAaBmq/
# wgRRFHmW9UJ3zv8Lmk6mIoAyTpqBbuUjMLyrtajuSsA/m2DnKMO0Qiz1v+FSVbqM
# 38J/PTlhCTUbFOx0kLT7Y/7+ZyrilVCzyAYfFIinDIjWlM85tDeU8ZfJCjFKwq3D
# sRxV4JY18xww8TTmod3lkr9NqGQ54LmyPVc+5ibNrjCCB3EwggVZoAMCAQICEzMA
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
# OkE5MzUtMDNFMC1EOTQ3MSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBT
# ZXJ2aWNloiMKAQEwBwYFKw4DAhoDFQBHJY2Fv+GhLQtRDR2vIzBaSv/7LKCBgzCB
# gKR+MHwxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQH
# EwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJjAkBgNV
# BAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQSAyMDEwMA0GCSqGSIb3DQEBCwUA
# AgUA6IaHizAiGA8yMDIzMDgxNTIzMjM1NVoYDzIwMjMwODE2MjMyMzU1WjB0MDoG
# CisGAQQBhFkKBAExLDAqMAoCBQDohoeLAgEAMAcCAQACAi33MAcCAQACAhMcMAoC
# BQDoh9kLAgEAMDYGCisGAQQBhFkKBAIxKDAmMAwGCisGAQQBhFkKAwKgCjAIAgEA
# AgMHoSChCjAIAgEAAgMBhqAwDQYJKoZIhvcNAQELBQADggEBAHn4ZwL6UxdORSnT
# nUe8JYuq7t8sZInUFcs72TFalUADfXIWWrBMCPYyTlDOb1SES6kuBLxY5m2BXaU9
# CXBprc2Hz2FwmhGoWxjOC7nGm2VCU6baCW20aU0imh/6MLHQORsOLJ6tpyJTV2KT
# gjLPK+8KU9CEF1RkBpvdZ2wAjV/d2sJj+5EpcePHFtljFl0P2P5k3z+skX0Pp935
# kTydeJrjCINykT7sK0xthUIYxvd+azEquvIlDnpds2BqTEENeS3lgU9E+C3I1ccn
# LAwPsPVe/BjwAAhc4qqVsU2YWL0geVmFlhVZhdgjgcN4OxceBQlpR+Qdo0wOEjxj
# JzkGVkExggQNMIIECQIBATCBkzB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2Fz
# aGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENv
# cnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAx
# MAITMwAAAdGyW0AobC7SRQABAAAB0TANBglghkgBZQMEAgEFAKCCAUowGgYJKoZI
# hvcNAQkDMQ0GCyqGSIb3DQEJEAEEMC8GCSqGSIb3DQEJBDEiBCCvagM6tugn9oP0
# AxgOtXfXCsLE+Uwc7fzPcmBWmn7VRDCB+gYLKoZIhvcNAQkQAi8xgeowgecwgeQw
# gb0EIMy8YXkCALv57c5sRhrPTub1q4TwJ6oVA36k8IiI/AcMMIGYMIGApH4wfDEL
# MAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1v
# bmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWlj
# cm9zb2Z0IFRpbWUtU3RhbXAgUENBIDIwMTACEzMAAAHRsltAKGwu0kUAAQAAAdEw
# IgQgQsKY27z27JDTUmRSPGw9o0lPvZQFWNGAm1F2k/wMuvIwDQYJKoZIhvcNAQEL
# BQAEggIAYN7xA/c9y1Sy+oehaodhr7/+EHnGt0v/YOytB4CvZJH7DXHJm7qDpDm3
# e4OC9Nh39vD1a7Eb6gdIvXG1Xi+Yc676hZGp6tDP8IAm3Sb2+iSGq3k9bSxPIbcf
# Xv9uz/CMWLXZzsY4xPxjM+YS+znWkxdhh+xkMRAaPl4r1xdUqzwglPRqrZMQAENv
# ocVGRpR7ZaHzqZwnL0OsoQeLb1gxDXK/UFscDeQFXYL+QcXt+j6ECgdOVmj69k4d
# is+w8+TrOieo8d1VVxHSGVoDZ2n1z37Nzqxi7LGFALy1q0rzDxXCJ/lgYSRx5aJ2
# 5XzUmo6xnhnbAHynEYkSFTKvIv7fKjhUPcLuUfq38ERiZScxnWp0EJBmW4jLAVqZ
# 9c1PIn7gMwnxX5lWG+q5K5SNPnr4irdG1CHm1ZbBxfow/DX4rdYfaEIbkd1D/PSd
# 9y3lMM3LLsZFyGS4tkkrINUWYPyO7LP0JoWVnkLbYFNhlsQUj1uEjjsH1HebNMUt
# BpXdOLopI71FV9AfHfRowSYiXtDapJTQ0K3fBZalCNqdJh2VToGobYGk5NyKgItW
# mBmVUjN6svWdyjrU0tNqgyeQXSWfneTqFIX0PLnNnF12nLhbna4EX27TbONvE6QK
# CQiak+229ubPg6jIw5TkPjw6z1CBS65QevNRHk3J3uzxrRhrXJI=
# SIG # End signature block
