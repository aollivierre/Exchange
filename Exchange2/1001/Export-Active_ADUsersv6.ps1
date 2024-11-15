# Import necessary modules for Active Directory
Import-Module ActiveDirectory

$users = Get-ADUser -Filter {Enabled -eq $true} -Properties DisplayName, GivenName, Surname, Title, Department, Company, Manager, UserPrincipalName, SamAccountName 

$userInfo = $users | ForEach-Object {
    $manager = if($_.Manager) { (Get-ADUser $_.Manager).Name } else { $null }

    # Fetch proxy addresses from Exchange for the user
    $mailUser = Get-Mailbox -Identity $_.SamAccountName -ErrorAction SilentlyContinue
    $smtpAddresses = if ($mailUser) {
        $mailUser.EmailAddresses | Where-Object { $_.PrefixString -eq "SMTP" -or $_.PrefixString -eq "smtp" } | ForEach-Object { $_.SmtpAddress }
    } else {
        $null
    }

    # Extract primary SMTP (uppercase "SMTP")
    $primarySMTP = ($mailUser.EmailAddresses | Where-Object { $_.PrefixString -eq "SMTP" }).SmtpAddress

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

$userInfo | Export-Csv -Path "C:\Code\Exchange\Exports\AGH_October_3rd_2023_16_45_PM_AD_Users_PowerShell_Export_V11.csv" -NoTypeInformation
