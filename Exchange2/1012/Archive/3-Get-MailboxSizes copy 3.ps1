$mailboxes = Get-Mailbox -ResultSize Unlimited

$results = foreach ($mailbox in $mailboxes) {
    $stats = Get-MailboxStatistics -Identity $mailbox.UserPrincipalName

    $sizeRaw, $unit = $stats.TotalItemSize.Value -split " ", 2
    $sizeRaw = [double]::Parse($sizeRaw)

    $size = $null
    switch ($unit) {
        "B"  { $size = $sizeRaw / 1MB }     # Convert bytes to megabytes
        "KB" { $size = $sizeRaw / 1KB }     # Convert kilobytes to megabytes
        "MB" { $size = $sizeRaw }           # Value is already in megabytes
        "GB" { $size = $sizeRaw * 1KB }     # Convert gigabytes to megabytes
        default { $size = "Unknown Unit: $unit" }
    }

    $userProperties = @{
        UserPrincipalName = $mailbox.UserPrincipalName
        MailboxSize_MB = $size
    }

    New-Object PsObject -Property $userProperties
}

$results = $results | Sort-Object -Property MailboxSize_MB -Descending
$results | Out-GridView
