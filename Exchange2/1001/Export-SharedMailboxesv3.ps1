$sharedMailboxes = Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails SharedMailbox

$sharedMailboxInfo = $sharedMailboxes | ForEach-Object {
    $sendAs = Get-ADPermission $_.Identity | Where-Object { $_.ExtendedRights -like "*Send-As*" -and $_.User -notlike "NT AUTHORITY\SELF" } | ForEach-Object { $_.User }
    $sendOnBehalf = $_.GrantSendOnBehalfTo -join ', '
    $fullAccess = Get-MailboxPermission $_.Identity | Where-Object { $_.IsInherited -eq $false -and $_.User -ne "NT AUTHORITY\SELF" } | ForEach-Object { $_.User.Value }
    $proxyAddresses = ($_.EmailAddresses | Where-Object { $_.PrefixString -eq "SMTP" -or $_.PrefixString -eq "smtp" } | ForEach-Object { $_.SmtpAddress }) -join ', '

    [PSCustomObject]@{
        'Name of SM' = $_.Name
        'Proxy Addresses' = $proxyAddresses
        'Send As' = ($sendAs -join ', ')
        'Send On Behalf' = $sendOnBehalf
        'Full Access' = ($fullAccess -join ', ')
        'Issue Warning Quota' = $_.IssueWarningQuota
        'Prohibit Send Quota' = $_.ProhibitSendQuota
        'Prohibit Send Receive Quota' = $_.ProhibitSendReceiveQuota
        'Hidden from Address Lists' = $_.HiddenFromAddressListsEnabled
        'Organizational Unit' = $_.OrganizationalUnit
        'Primary SMTP Address' = $_.PrimarySmtpAddress
        'Database' = $_.Database
        # You can add additional properties if required.
    }
}

$sharedMailboxInfo | Export-Csv -Path "C:\Code\Exchange\Exports\AGH_October_3rd_2023_12_34_PM_SharedMailbox_PowerShell_Export_V7.csv" -NoTypeInformation
