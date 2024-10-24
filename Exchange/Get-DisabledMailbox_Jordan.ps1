$disabledMailbox = Get-Mailbox -RecipientTypeDetails DisabledMailbox -Identity "enter email address"

if ($disabledMailbox) {
    Write-Output "Disabled mailbox found: $($disabledMailbox.PrimarySmtpAddress)"
    Enable-Mailbox -Identity "enter email address"
    Write-Output "Mailbox enabled."
} else {
    Write-Output "No disabled mailbox found."
}
