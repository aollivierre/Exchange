# # Import Microsoft Graph PowerShell module
# Import-Module Microsoft.Graph.Intune

# # Set variables
# $tenantId = "yourtenantid"
# $appId = "yourappid"
# $appSecret = "yourappsecret"
# $accessToken = Get-MgAccessToken -TenantId $tenantId -ClientId $appId -ClientSecret $appSecret
$devices = Get-MgDevice -All

# Set last login threshold to 90 days
$lastLoginThreshold = (Get-Date).AddDays(-90)

# Filter devices that have not logged in for more than 90 days
$inactiveDevices = $devices | Where-Object { $_.ApproximateLastSignInDateTime -lt $lastLoginThreshold }

# Get total device count
$totalDeviceCount = $devices.Count

# Get total enabled device count
$enabledDeviceCount = ($devices | Where-Object { $_.accountEnabled }).Count

# Get total inactive device count
$inactiveDeviceCount = $inactiveDevices.Count

# Display summary information
Write-Host $("$(Get-Date) Total Device Count: $($totalDeviceCount.ToString('N0'))") -ForegroundColor Green
Write-Host $("$(Get-Date) Enabled Device Count: $($enabledDeviceCount.ToString('N0'))") -ForegroundColor Green
Write-Host $("$(Get-Date) Inactive Device Count: $($inactiveDeviceCount.ToString('N0'))") -ForegroundColor Yellow

# Export inactive devices to CSV
if ($null -ne $inactiveDevices -and $inactiveDevices.Count -gt 0) {
    $exportPath = "C:\Code\Entra\Exports"
    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $inactiveExportFileName = "InactiveDevices_$timestamp.csv"
    $inactiveExportFilePath = Join-Path $exportPath $inactiveExportFileName

    if (-not (Test-Path $exportPath)) {
        New-Item -ItemType Directory -Path $exportPath -Force | Out-Null
    }

    # Export inactive device details to CSV
    $inactiveDevices | Sort-Object ApproximateLastSignInDateTime | Select-Object DisplayName, DeviceId, ApproximateLastSignInDateTime, accountEnabled, Id | Export-Csv -Path $inactiveExportFilePath -NoTypeInformation

    Write-Host $("$(Get-Date) $inactiveDeviceCount inactive devices exported to $inactiveExportFilePath") -ForegroundColor Green
}
else {
    Write-Warning "No inactive devices found for export."
}

# Export all device details for inactive devices to CSV
if ($null -ne $inactiveDevices -and $inactiveDevices.Count -gt 0) {
    $allExportFileName = "InactiveDevices_AllDetails_$timestamp.csv"
    $allExportFilePath = Join-Path $exportPath $allExportFileName

    # Export all device details to CSV
    $inactiveDevices | Sort-Object ApproximateLastSignInDateTime | Export-Csv -Path $allExportFilePath -NoTypeInformation

    Write-Host $("$(Get-Date) $inactiveDeviceCount inactive devices and their details exported to $allExportFilePath") -ForegroundColor Green
}
else {
    Write-Warning "No inactive devices found for all device details export."
}
