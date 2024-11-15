#Get Local Mailboxes
$mailboxes = Get-Mailbox -ResultSize Unlimited

$results = foreach ($mailbox in $mailboxes) {
    $stats = Get-MailboxStatistics -Identity $mailbox.UserPrincipalName
    
    if ($stats.TotalItemSize.Value -match "(?<Size>\d+(?:\.\d+)?)\s?(?<Unit>\w+)") {
        $sizeRaw = [double]::Parse($Matches.Size)
        $unit = $Matches.Unit

        $size = $null
        switch ($unit) {
            "B"   { $size = $sizeRaw / 1MB }      # Convert bytes to megabytes
            "KB"  { $size = $sizeRaw / 1KB }      # Convert kilobytes to megabytes
            "MB"  { $size = $sizeRaw }            # Value is already in megabytes
            "GB"  { $size = $sizeRaw * 1KB }      # Convert gigabytes to megabytes
            "TB"  { $size = $sizeRaw * 1MB }      # Convert terabytes to megabytes
            default { $size = "Unknown Unit: $unit" }
        }
    }

    $userProperties = @{
        UserPrincipalName = $mailbox.UserPrincipalName
        MailboxSize_MB = $size
        MailboxType = $mailbox.RecipientTypeDetails
    }

    New-Object PsObject -Property $userProperties
}

$results = $results | Sort-Object -Property MailboxSize_MB -Descending

# Output total statistics
$totalUsers = $mailboxes.Count
$totalMailboxSize = ($results | Measure-Object -Property MailboxSize_MB -Sum).Sum
Write-Host "$(Get-Date) - [INFO] Total Users: $totalUsers" -ForegroundColor Green
Write-Host "$(Get-Date) - [INFO] Total Mailbox Size (MB): $totalMailboxSize" -ForegroundColor Green

$results | Out-GridView
