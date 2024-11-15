$mailboxes = Get-Mailbox -ResultSize Unlimited

$results = foreach ($mailbox in $mailboxes) {
    $stats = Get-MailboxStatistics -Identity $mailbox.UserPrincipalName

    $sizeRaw = $stats.TotalItemSize.Value -replace "(.*\()|,.*"
    $unit = $stats.TotalItemSize.Value -replace ".*\("
    $size = $null

    switch ($unit) {
        "B)"   { $size = $sizeRaw / 1MB }     # Convert bytes to megabytes
        # "KB)"  { $size = $sizeRaw / 1K }      # Convert kilobytes to megabytes
        "MB)"  { $size = $sizeRaw }           # Value is already in megabytes
        # "GB)"  { $size = $sizeRaw * 1K }      # Convert gigabytes to megabytes
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