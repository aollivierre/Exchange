$mailbox = Get-Mailbox -Identity "enter email address" -ErrorAction SilentlyContinue
$mailUser = Get-MailUser -Identity "enter email address"

if ($mailbox) {
    Write-Output "Mailbox found: $($mailbox.PrimarySmtpAddress)"
} elseif ($mailUser) {
    Write-Output "Mail user found: $($mailUser.PrimarySmtpAddress)"
} else {
    Write-Output "No mailbox or mail user found for enter email address."
}
