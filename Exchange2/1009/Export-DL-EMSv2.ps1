$groups = Get-DistributionGroup

$groupInfo = $groups | ForEach-Object {
    $members = Get-DistributionGroupMember -Identity $_.DistinguishedName

    # Fetch all proxy addresses (both primary and secondary)
    $proxyAddresses = $_.EmailAddresses | ForEach-Object { $_.SmtpAddress }

    [PSCustomObject]@{
        'Name of DL'          = $_.Name
        'Proxy Addresses'     = ($proxyAddresses -join ', ')
        'List of Members'     = ($members.Name -join ', ')
        'GroupType'           = $_.GroupType
        'RecipientTypeDetails'= $_.RecipientTypeDetails
        'ManagedBy'           = $_.ManagedBy
        'HiddenFromAddressListsEnabled' = $_.HiddenFromAddressListsEnabled
        'AcceptMessagesOnlyFrom' = $_.AcceptMessagesOnlyFrom
        'RejectMessagesFrom' = $_.RejectMessagesFrom
        # ... Add more settings as needed
    }
}



$groupInfo | Out-GridView

$groupInfo | Export-Csv -Path "C:\Code\Exchange\Exports\CPDMH_October_3rd_2023_12_00_PM_DL_PowerShell_Export_V10.csv" -NoTypeInformation
