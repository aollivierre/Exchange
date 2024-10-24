$mailbox = Get-Mailbox -Identity "JHeuser@arnpriorhealth.ca" -ErrorAction SilentlyContinue
$mailUser = Get-MailUser -Identity "JHeuser@arnpriorhealth.ca"

if ($mailbox) {
    Write-Output "Mailbox found: $($mailbox.PrimarySmtpAddress)"
} elseif ($mailUser) {
    Write-Output "Mail user found: $($mailUser.PrimarySmtpAddress)"
} else {
    Write-Output "No mailbox or mail user found for JHeuser@arnpriorhealth.ca."
}



if (-not $mailbox) {
    Enable-RemoteMailbox -Identity "JHeuser@arnpriorhealth.ca" -RemoteRoutingAddress "JHeuser@arnpriorhealth.mail.onmicrosoft.com"
    Write-Output "Remote mailbox created."
}
