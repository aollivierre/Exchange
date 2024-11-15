$groups = Get-DistributionGroup

$groupInfo = $groups | ForEach-Object {
    $members = Get-DistributionGroupMember -Identity $_.DistinguishedName

    # Fetch all proxy addresses (both primary and secondary)
    $proxyAddresses = $_.EmailAddresses | ForEach-Object { $_.SmtpAddress }

    # Convert ADMultiValuedProperty to comma-separated strings
    $managedBy = if ($_.ManagedBy -ne $null) { ($_.ManagedBy | ForEach-Object { $_.Name }) -join ', ' } else { $null }
    $acceptOnlyFrom = if ($_.AcceptMessagesOnlyFrom -ne $null) { ($_.AcceptMessagesOnlyFrom | ForEach-Object { $_.Name }) -join ', ' } else { $null }
    $rejectFrom = if ($_.RejectMessagesFrom -ne $null) { ($_.RejectMessagesFrom | ForEach-Object { $_.Name }) -join ', ' } else { $null }

    [PSCustomObject]@{
        'Name of DL'                     = $_.Name
        'Proxy Addresses'                = ($proxyAddresses -join ', ')
        'List of Members'                = ($members.Name -join ', ')
        'GroupType'                      = $_.GroupType
        'RecipientTypeDetails'           = $_.RecipientTypeDetails
        'ManagedBy'                      = $managedBy
        'HiddenFromAddressListsEnabled'  = $_.HiddenFromAddressListsEnabled
        'AcceptMessagesOnlyFrom'         = $acceptOnlyFrom
        'RejectMessagesFrom'             = $rejectFrom
        # ... Add more settings as needed
    }
}


$groupInfo | out-gridview

$groupInfo | Export-Csv -Path "C:\Code\Exchange\Exports\CPDMH_October_3rd_2023_12_00_PM_DL_PowerShell_Export_V9.csv" -NoTypeInformation
