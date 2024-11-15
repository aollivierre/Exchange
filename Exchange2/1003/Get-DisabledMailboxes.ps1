# # Get all mailboxes
# $mailboxes = Get-Mailbox -ResultSize Unlimited

# # Loop through each mailbox to check if the associated user account is disabled
# foreach ($mailbox in $mailboxes) {
#     $user = Get-User $mailbox.DistinguishedName
#     if ($user.UserAccountControl -like "*AccountDisabled*") {
#         Write-Host "Disabled mailbox: $($mailbox.DisplayName)"
#     }
# }




$dbs = Get-MailboxDatabase
$dbs | ForEach-Object {Get-MailboxStatistics -Database $_.DistinguishedName} | Where-Object {$_.DisconnectReason -eq "Disabled"} | Format-Table DisplayName,Database,DisconnectDate