$mailboxes = Get-Mailbox -ResultSize Unlimited
$mailboxesCount = $mailboxes.Count

if ($mailboxesCount -gt 0) {
    if ($mailboxesCount -lt 50) {
        Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - There are $($mailboxesCount) mailboxes." -ForegroundColor Green
    } elseif ($mailboxesCount -lt 100) {
        Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - There are $($mailboxesCount) mailboxes." -ForegroundColor Yellow
    } else {
        Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - There are $($mailboxesCount) mailboxes." -ForegroundColor Red
    }
} else {
    Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - No mailboxes found." -ForegroundColor Gray
}
