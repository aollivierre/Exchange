# Define the registry key path and value
$regPath = "HKLM:\Software\Microsoft\Exchange"
$regName = "AlwaysUseMSOAuthForAutoDiscover"

# Remove the registry value if it exists
if (Test-Path $regPath) {
    Remove-ItemProperty -Path $regPath -Name $regName -ErrorAction SilentlyContinue

    # If no other properties are present, remove the key
    if (-Not (Get-ItemProperty -Path $regPath -ErrorAction SilentlyContinue)) {
        Remove-Item -Path $regPath -Recurse -Force
    }
}
