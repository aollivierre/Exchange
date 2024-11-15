# Fetch distribution groups
$groups = Get-DistributionGroup

$groupInfo = $groups | ForEach-Object {
    $members = Get-DistributionGroupMember -Identity $_.DistinguishedName
    
    # Getting the list of proxy addresses
    $proxyAddresses = $_.EmailAddresses | Where-Object { $_.PrefixString -eq "smtp" } | ForEach-Object { $_.SmtpAddress }

    [PSCustomObject]@{
        'Name of DL'          = $_.Name
        'Proxy Addresses'     = ($proxyAddresses -join ', ')
        'List of Members'     = ($members.Name -join ', ')
        'DistinguishedName'   = $_.DistinguishedName
        'GroupScope'          = $_.GroupScope
        'ObjectClass'         = $_.ObjectClass
        'SamAccountName'      = $_.SamAccountName
        'SID'                 = $_.SID
    }
}


$groupInfo | Out-GridView

$groupInfo | Export-Csv -Path "C:\Code\Exchange\Exports\CPDMH_October_3rd_2023_12_00_PM_DL_PowerShell_Export_V5.csv" -NoTypeInformation
