$disabledMailbox = Get-Mailbox -RecipientTypeDetails DisabledMailbox -Identity "JHeuser@arnpriorhealth.ca"

if ($disabledMailbox) {
    Write-Output "Disabled mailbox found: $($disabledMailbox.PrimarySmtpAddress)"
    Enable-Mailbox -Identity "JHeuser@arnpriorhealth.ca"
    Write-Output "Mailbox enabled."
} else {
    Write-Output "No disabled mailbox found."
}
