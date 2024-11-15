$DateCutoff = (Get-Date).AddYears(-4)
Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited | Where-Object {$_.WhenMailboxCreated -lt $DateCutoff} | Select-Object DisplayName, PrimarySmtpAddress, WhenMailboxCreated | Export-Csv -Path "$env:USERPROFILE\Desktop\OldSharedMailboxes.csv" -NoTypeInformation
