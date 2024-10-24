# Step 2: Check On-Premises Exchange Management Server
$mailbox = Get-Mailbox -Identity "JHeuser@arnpriorhealth.ca" -ErrorAction SilentlyContinue
$mailUser = Get-MailUser -Identity "JHeuser@arnpriorhealth.ca" -ErrorAction SilentlyContinue

if ($mailbox) {
    Write-Output "Mailbox found: $($mailbox.PrimarySmtpAddress)"
} elseif ($mailUser) {
    Write-Output "Mail user found: $($mailUser.PrimarySmtpAddress)"
} else {
    Write-Output "No mailbox or mail user found for JHeuser@arnpriorhealth.ca."
}

# Step 3: Identify Disabled Mailboxes (correct the recipient type)
$disabledMailbox = Get-User -RecipientTypeDetails DisabledUser -Identity "JHeuser@arnpriorhealth.ca" -ErrorAction SilentlyContinue

if ($disabledMailbox) {
    Write-Output "Disabled user found: $($disabledMailbox.PrimarySmtpAddress)"
    Enable-Mailbox -Identity "JHeuser@arnpriorhealth.ca"
    Write-Output "Mailbox enabled."
} else {
    Write-Output "No disabled mailbox found."
}

# Step 4: Enable or Create a Remote Mailbox
if (-not $mailbox -and $mailUser) {
    Enable-RemoteMailbox -Identity "JHeuser@arnpriorhealth.ca" -RemoteRoutingAddress "JHeuser@arnpriorhealth.mail.onmicrosoft.com"
    Write-Output "Remote mailbox created."
} elseif (-not $mailbox) {
    New-RemoteMailbox -Name "JHeuser" -Alias "JHeuser" -UserPrincipalName "JHeuser@arnpriorhealth.ca" -PrimarySmtpAddress "JHeuser@arnpriorhealth.ca" -RemoteRoutingAddress "JHeuser@arnpriorhealth.mail.onmicrosoft.com"
    Write-Output "New remote mailbox created."
}

# Step 5: Verify the Mailbox in Exchange Online
# $mailboxOnline = Get-Mailbox -Identity "JHeuser@arnpriorhealth.ca" -ErrorAction SilentlyContinue

# if ($mailboxOnline) {
#     Write-Output "Mailbox found in Exchange Online: $($mailboxOnline.PrimarySmtpAddress)"
# } else {
#     Write-Output "No mailbox found in Exchange Online."
# }