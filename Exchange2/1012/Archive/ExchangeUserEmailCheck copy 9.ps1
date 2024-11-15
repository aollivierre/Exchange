# Connect to Exchange Server
# $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://<FQDN of your Exchange Server>/PowerShell/ -Authentication Kerberos
# Import-PSSession $Session

$allUsers = Get-User -ResultSize Unlimited
$usersWithMailboxes = @()

foreach($user in $allUsers){
    try{
        $mailbox = Get-Mailbox -Identity $user.UserPrincipalName -ErrorAction SilentlyContinue
        if ($null -ne $mailbox) {
            $usersWithMailboxes += $user
        }
    }catch{
        try{
            $remoteMailbox = Get-RemoteMailbox -Identity $user.UserPrincipalName -ErrorAction SilentlyContinue
            if ($null -ne $remoteMailbox) {
                $usersWithMailboxes += $user
            }
        }catch{
            Write-Host "$(Get-Date) - [WARNING] Could not find mailbox for user $($user.UserPrincipalName)" -ForegroundColor Yellow
        }
    }
}

$results = foreach($user in $usersWithMailboxes){
    $mailboxType = $null
    $emailAddressCheck = $user.WindowsEmailAddress.Local -like "*@glebecentre.mail.onmicrosoft.com"
    $isRemoteUser = $null -ne (Get-Recipient $user.UserPrincipalName | Where-Object {$_.RecipientTypeDetails -eq "RemoteUserMailbox"})

    if($isRemoteUser){
        $mailboxType = 'Remote'
    }else{
        $mailboxType = 'Local'
    }

    $userProperties = @{
        UserPrincipalName = $user.UserPrincipalName
        EmailAddresses = $user.WindowsEmailAddress.Local
        EmailAddressCheck = $emailAddressCheck
        MailboxType = $mailboxType
    }

    New-Object PsObject -Property $userProperties
}

$results | Out-GridView

$totalUsers = $usersWithMailboxes.count
$remoteUsers = $results.Where({ $_.MailboxType -eq "Remote" }).count
$localUsers = $results.Where({ $_.MailboxType -eq "Local" }).count
$passedCheck = $results.Where({ $_.EmailAddressCheck -eq $true }).count
$failedCheck = $results.Where({ $_.EmailAddressCheck -eq $false }).count

Write-Host "$(Get-Date) - [INFO] Total Users in Exchange: $totalUsers" -ForegroundColor Green
Write-Host "$(Get-Date) - [INFO] Total Remote Users: $remoteUsers" -ForegroundColor Green
Write-Host "$(Get-Date) - [INFO] Total Local Users: $localUsers" -ForegroundColor Green
Write-Host "$(Get-Date) - [INFO] Total Users Passed Email Check: $passedCheck" -ForegroundColor Green
Write-Host "$(Get-Date) - [INFO] Total Users Failed Email Check: $failedCheck" -ForegroundColor Red

# $results | Export-Csv -Path 'MailboxDetails.csv' -NoTypeInformation

# Remove-PSSession $Session
