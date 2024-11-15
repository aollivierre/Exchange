$sharedMailboxes = Get-Mailbox -RecipientTypeDetails SharedMailbox
$sharedMailboxesCount = $sharedMailboxes.Count

if ($sharedMailboxesCount -gt 0) {
    if ($sharedMailboxesCount -lt 50) {
        Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - There are $($sharedMailboxesCount) shared mailboxes." -ForegroundColor Green
    } elseif ($sharedMailboxesCount -lt 100) {
        Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - There are $($sharedMailboxesCount) shared mailboxes." -ForegroundColor Yellow
    } else {
        Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - There are $($sharedMailboxesCount) shared mailboxes." -ForegroundColor Red
    }
} else {
    Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - No shared mailboxes found." -ForegroundColor Gray
}
