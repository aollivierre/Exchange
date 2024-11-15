# # Connect to Exchange Online
# $UserCredential = Get-Credential
# $ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
# Import-PSSession $ExchangeSession


# Connect-ExchangeOnline

# Define the user's email or UPN
# $UserEmail = "bnickoloff@bcclsp.org"
$UserEmail = "bev.nickoloff@sympatico.ca"
# $UserEmail = "sandrew@bcclsp.org"

# Get all mailboxes and check permissions for the defined user
$Mailboxes = Get-Mailbox -ResultSize Unlimited

foreach ($Mailbox in $Mailboxes) {
    $Permissions = Get-MailboxPermission -Identity $Mailbox.Identity | Where-Object { $_.User -like "*$UserEmail*" -and $_.IsInherited -eq $false }
    if ($Permissions) {
        $Permissions | Select-Object @{n='Mailbox';e={$Mailbox.PrimarySmtpAddress}}, User, AccessRights, Deny | Format-Table -AutoSize
    }
}




# Define the user's email or UPN
# $UserEmail = "user@example.com"

# Get all mailboxes
$Mailboxes = Get-Mailbox -ResultSize Unlimited

foreach ($Mailbox in $Mailboxes) {
    # Check for permissions on the Calendar folder
    $Permissions = Get-MailboxFolderPermission -Identity "$($Mailbox.PrimarySmtpAddress):\Calendar" | Where-Object { $_.User -like "*$UserEmail*" }

    if ($Permissions) {
        $Permissions | Select-Object @{n='Mailbox';e={$Mailbox.PrimarySmtpAddress}}, User, AccessRights | Format-Table -AutoSize
    }
}


# Close the session
# Remove-PSSession $ExchangeSession
