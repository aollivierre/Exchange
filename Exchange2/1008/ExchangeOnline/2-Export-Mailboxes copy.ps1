$DateCutoff = (Get-Date).AddYears(-4)


Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited | Where-Object {$_.WhenMailboxCreated -lt $DateCutoff -and $_.WhenMailboxCreated -lt (Get-Date).AddYears(-4).AddDays(-1)} | Select-Object DisplayName, UserPrincipalName, PrimarySmtpAddress, WhenMailboxCreated | Export-Csv -Path "$env:USERPROFILE\Desktop\OldSharedMailboxes3.csv" -NoTypeInformation
