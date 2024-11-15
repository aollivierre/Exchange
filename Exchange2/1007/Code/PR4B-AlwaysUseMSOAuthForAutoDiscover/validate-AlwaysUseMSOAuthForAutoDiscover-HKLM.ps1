$keyPath = "HKLM:\Software\Microsoft\Exchange"
$valueName = "AlwaysUseMSOAuthForAutoDiscover"

if (Test-Path $keyPath) {
    $value = Get-ItemProperty -Path $keyPath -Name $valueName -ErrorAction SilentlyContinue
    if ($null -ne $value) {
        Write-Host "Registry key exists and has value: $($value.$valueName)"
    }
    else {
        Write-Host "Registry key exists but value is not set."
    }
}
else {
    Write-Host "Registry key does not exist."
}
