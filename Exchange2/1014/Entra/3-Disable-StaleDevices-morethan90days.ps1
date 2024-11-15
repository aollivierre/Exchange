# Import the Microsoft Graph PowerShell module
# Import-Module Microsoft.Graph.Intune

# Set the last login threshold to 90 days
$lastLoginThreshold = (Get-Date).AddDays(-90)

# Import the CSV file containing the stale devices to disable
# $staleDevices = Import-Csv -Path "C:\Code\CB\Entra\Exports\InactiveDevices_20230402_162120.csv"
$staleDevices = Import-Csv -Path "C:\Code\Entra\Exports\InactiveDevices_20240108_114916.csv"

# Set the log path
$logPath = "C:\Code\Entra\Logs\Disable-StaleDevices.log"

# Loop through each device in the CSV file
foreach ($device in $staleDevices) {
    Write-Host "Processing $($device.DisplayName)..."
    
    # Get the device details from Microsoft Graph
    $graphDevice = Get-MgDevice -Filter "deviceId eq '$($device.DeviceId)'"

    if ($graphDevice) {
        # Check if the device has not logged in for more than 90 days
        if ($graphDevice.ApproximateLastSignInDateTime -lt $lastLoginThreshold) {
            # Disable the device
            $params = @{
                AccountEnabled = $false
            }
            Update-MgDevice -DeviceId $graphDevice.Id -BodyParameter $params
            
            Write-Host "Device $($device.DisplayName) has been disabled." -ForegroundColor Green
            Add-Content -Path $logPath -Value "$($device.DisplayName) - Disabled"
        }
        else {
            Write-Host "Device $($device.DisplayName) has logged in within the last 90 days." -ForegroundColor Yellow
            Add-Content -Path $logPath -Value "$($device.DisplayName) - Skipped"
        }
    }
    else {
        Write-Host "Device $($device.DisplayName) not found in Microsoft Graph." -ForegroundColor Red
        Add-Content -Path $logPath -Value "$($device.DisplayName) - Not found"
    }
}
Write-Host "Script execution completed." -ForegroundColor Green
Add-Content -Path $logPath -Value "Script execution completed."
