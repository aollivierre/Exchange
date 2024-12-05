# Configuration
$exportPath = "\\$env:COMPUTERNAME\C$\ExchangeArchives"
$userEmail = "JMcKitrick@tunngavik.com"
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"

# Create export directory and set permissions
try {
    if (-not (Test-Path $exportPath)) {
        $null = New-Item -ItemType Directory -Force -Path $exportPath
    }

    $acl = Get-Acl $exportPath
    $identities = @(
        "Exchange Trusted Subsystem",
        "SYSTEM",
        "NETWORK SERVICE"
    )

    foreach ($identity in $identities) {
        Write-Host "Adding permissions for $identity"
        $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule(
            $identity,
            "FullControl",
            "ContainerInherit,ObjectInherit",
            "None",
            "Allow"
        )
        $acl.AddAccessRule($accessRule)
    }

    Set-Acl -Path $exportPath -AclObject $acl
    Write-Host "Permissions set successfully on $exportPath"
    
    Write-Host "`nCurrent permissions on export folder:"
    Get-Acl $exportPath | Select-Object -ExpandProperty Access | Format-Table IdentityReference, FileSystemRights -AutoSize
} catch {
    Write-Error "Failed to set permissions: $_"
    exit
}

# Get mailbox info
$mailbox = Get-Mailbox $userEmail
$archiveStats = Get-MailboxStatistics -Identity $mailbox.Identity -Archive

Write-Host "`nStarting archive export for $($mailbox.DisplayName)"
Write-Host "Archive size: $($archiveStats.TotalItemSize)"
Write-Host "Item count: $($archiveStats.ItemCount) items"
Write-Host "Export path: $exportPath"

# Clean up any existing requests
Get-MailboxExportRequest -Name "Archive_Export_$($mailbox.Alias)*" | 
    Remove-MailboxExportRequest -Confirm:$false -ErrorAction SilentlyContinue

$pstPath = "$exportPath\$($mailbox.Alias)_archive_$timestamp.pst"
$exportParams = @{
    Mailbox = $mailbox.Identity
    IsArchive = $true
    FilePath = $pstPath
    Name = "Archive_Export_$($mailbox.Alias)_$timestamp"
}

try {
    Write-Host "`nCreating export request to: $pstPath"
    $startTime = Get-Date
    $request = New-MailboxExportRequest @exportParams
    Write-Host "Export request created successfully. Request Status: $($request.Status)"
    Write-Host "Monitoring progress..."
    
    do {
        Start-Sleep -Seconds 10
        $status = Get-MailboxExportRequest -Name $exportParams.Name
        if ($status.Status -like "*Failed*") {
            throw "Export failed: $($status.LastError)"
        }
        $percentComplete = if ($status.PercentComplete) { "$($status.PercentComplete)%" } else { "calculating..." }
        Write-Host "Status: $($status.Status) - Percent Complete: $percentComplete"
    } while ($status.Status -notlike "*Completed*")
    
    $endTime = Get-Date
    $duration = $endTime - $startTime
    
    Write-Host "`nExport completed successfully!"
    Write-Host "PST file location: $pstPath"
    Write-Host "`nExport Summary:"
    Write-Host "Start Time: $($startTime.ToString('yyyy-MM-dd HH:mm:ss'))"
    Write-Host "End Time: $($endTime.ToString('yyyy-MM-dd HH:mm:ss'))"
    Write-Host "Duration: $([math]::Round($duration.TotalMinutes, 2)) minutes"
    Write-Host "Original Archive Size: $($archiveStats.TotalItemSize)"
    Write-Host "Original Item Count: $($archiveStats.ItemCount) items"
    
    if (Test-Path $pstPath) {
        $pstSize = (Get-Item $pstPath).Length
        Write-Host "Final PST Size: $([math]::Round($pstSize / 1MB, 2)) MB"
    }
    
} catch {
    Write-Error "Export failed: $_"
    Get-MailboxExportRequest -Name $exportParams.Name | 
        Remove-MailboxExportRequest -Confirm:$false -ErrorAction SilentlyContinue
}