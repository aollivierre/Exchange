# Initialize time stamp
$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

# Initialize counters
$keysBeforeCount = 0
$keysAfterCount = 0

# Function to write color-coded and timestamped messages
function Write-ColorHost {
    param (
        [string]$Message,
        [ConsoleColor]$Color = "White"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Host "$timestamp - $Message" -ForegroundColor $Color
}

try {
    # Check if the key exists
    $keyPath = "HKCU:\Software\Microsoft\Office\16.0\Outlook\AutoDiscover"

    if (Test-Path $keyPath) {
        # Count existing keys
        $keysBeforeCount = (Get-Item $keyPath).Property.Count
        Write-ColorHost "Existing keys before operation: $keysBeforeCount" -Color Green

        # List existing values
        Write-ColorHost "Current Registry Settings:" -Color Yellow
        Get-ItemProperty -Path $keyPath | Select-Object PreferLocalXML, ExcludeHttpRedirect, ExcludeHttpsAutoDiscoverDomain, ExcludeHttpsRootDomain, ExcludeScpLookup, ExcludeSrvRecord, ExcludeLastKnownGoodURL, ExcludeExplicitO365Endpoint | Format-Table -AutoSize
    }
    else {
        Write-ColorHost "Registry path doesn't exist. It will be created." -Color Yellow
    }

    # Setting registry keys
    Write-ColorHost "Setting Registry keys..." -Color Yellow

    # Create key if it doesn't exist
    if (-Not (Test-Path $keyPath)) {
        New-Item -Path $keyPath -Force
    }

    # Set registry values
    Set-ItemProperty -Path $keyPath -Name "PreferLocalXML" -Value 1
    Set-ItemProperty -Path $keyPath -Name "ExcludeHttpRedirect" -Value 0
    Set-ItemProperty -Path $keyPath -Name "ExcludeHttpsAutoDiscoverDomain" -Value 0
    Set-ItemProperty -Path $keyPath -Name "ExcludeHttpsRootDomain" -Value 0
    Set-ItemProperty -Path $keyPath -Name "ExcludeScpLookup" -Value 0
    Set-ItemProperty -Path $keyPath -Name "ExcludeSrvRecord" -Value 0
    Set-ItemProperty -Path $keyPath -Name "ExcludeLastKnownGoodURL" -Value 0
    Set-ItemProperty -Path $keyPath -Name "ExcludeExplicitO365Endpoint" -Value 1

    # Count existing keys after setting
    $keysAfterCount = (Get-Item $keyPath).Property.Count
    Write-ColorHost "Existing keys after operation: $keysAfterCount" -Color Green

    # List updated values
    Write-ColorHost "Updated Registry Settings:" -Color Yellow
    Get-ItemProperty -Path $keyPath | Select-Object PreferLocalXML, ExcludeHttpRedirect, ExcludeHttpsAutoDiscoverDomain, ExcludeHttpsRootDomain, ExcludeScpLookup, ExcludeSrvRecord, ExcludeLastKnownGoodURL, ExcludeExplicitO365Endpoint | Format-Table -AutoSize

    Write-ColorHost "Registry keys set successfully." -Color Green

} catch {
    Write-ColorHost "An error occurred: $_" -Color Red
}