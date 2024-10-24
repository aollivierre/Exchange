$mailbox = Get-Mailbox -Identity "enter email address" -ErrorAction SilentlyContinue
$mailUser = Get-MailUser -Identity "enter email address"

if ($mailbox) {
    Write-Output "Mailbox found: $($mailbox.PrimarySmtpAddress)"
} elseif ($mailUser) {
    Write-Output "Mail user found: $($mailUser.PrimarySmtpAddress)"
} else {
    Write-Output "No mailbox or mail user found for enter email address."
}



if (-not $mailbox) {
    Enable-RemoteMailbox -Identity "enter email address" -RemoteRoutingAddress "JHeuser@arnpriorhealth.mail.onmicrosoft.com"
    Write-Output "Remote mailbox created."
}
