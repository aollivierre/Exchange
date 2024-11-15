# Import necessary modules for Active Directory
Import-Module ActiveDirectory

$users = Get-ADUser -Filter {Enabled -eq $true} -Properties DisplayName, GivenName, Surname, Title, Department, Company, Manager, UserPrincipalName, SamAccountName 

$userInfo = $users | ForEach-Object {
    $manager = if($_.Manager) { (Get-ADUser $_.Manager).Name } else { $null }

    # Fetch mailbox details from Exchange for the user
    $mailUser = Get-Mailbox -Identity $_.SamAccountName -ErrorAction SilentlyContinue
    $smtpAddresses = if ($mailUser) {
        $mailUser.EmailAddresses | Where-Object { $_.PrefixString -eq "SMTP" -or $_.PrefixString -eq "smtp" } | ForEach-Object { $_.SmtpAddress }
    } else {
        $null
    }

    # Get primary SMTP address directly
    $primarySMTP = if ($mailUser) { $mailUser.PrimarySmtpAddress } else { $null }

    [PSCustomObject]@{
        'Full Name/Display Name'      = $_.DisplayName
        'First Name'                  = $_.GivenName
        'Last Name'                   = $_.Surname
        'Job Title'                   = $_.Title
        'Department'                  = $_.Department
        'Company Name'                = $_.Company
        'Manager Name/Reporting To'   = $manager
        'UserPrincipalName'           = $_.UserPrincipalName
        'Samaccountname'              = $_.SamAccountName
        'Primary SMTP Address'        = $primarySMTP
        'All Proxy Addresses'         = ($smtpAddresses -join ', ')
    }
}

$userInfo | Export-Csv -Path "C:\Code\Exchange\Exports\CPDMH_October_3rd_2023_16_10_PM_AD_Users_PowerShell_Export_V12.csv" -NoTypeInformation
