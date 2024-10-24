# Step 2: Check On-Premises Exchange Management Server
$mailbox = Get-Mailbox -Identity "enter email address" -ErrorAction SilentlyContinue
$mailUser = Get-MailUser -Identity "enter email address" -ErrorAction SilentlyContinue

if ($mailbox) {
    Write-Output "Mailbox found: $($mailbox.PrimarySmtpAddress)"
} elseif ($mailUser) {
    Write-Output "Mail user found: $($mailUser.PrimarySmtpAddress)"
} else {
    Write-Output "No mailbox or mail user found for enter email address."
}

# Step 3: Identify Disabled Mailboxes (correct the recipient type)
$disabledMailbox = Get-User -RecipientTypeDetails DisabledUser -Identity "enter email address" -ErrorAction SilentlyContinue

if ($disabledMailbox) {
    Write-Output "Disabled user found: $($disabledMailbox.PrimarySmtpAddress)"
    Enable-Mailbox -Identity "enter email address"
    Write-Output "Mailbox enabled."
} else {
    Write-Output "No disabled mailbox found."
}

# Step 4: Enable or Create a Remote Mailbox
if (-not $mailbox -and $mailUser) {
    Enable-RemoteMailbox -Identity "enter email address" -RemoteRoutingAddress "JHeuser@contoso.com.mail.onmicrosoft.com"
    Write-Output "Remote mailbox created."
} elseif (-not $mailbox) {
    New-RemoteMailbox -Name "JHeuser" -Alias "JHeuser" -UserPrincipalName "enter email address" -PrimarySmtpAddress "enter email address" -RemoteRoutingAddress "JHeuser@contoso.com.mail.onmicrosoft.com"
    Write-Output "New remote mailbox created."
}

# Step 5: Verify the Mailbox in Exchange Online
# $mailboxOnline = Get-Mailbox -Identity "enter email address" -ErrorAction SilentlyContinue

# if ($mailboxOnline) {
#     Write-Output "Mailbox found in Exchange Online: $($mailboxOnline.PrimarySmtpAddress)"
# } else {
#     Write-Output "No mailbox found in Exchange Online."
# }