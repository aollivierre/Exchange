# Import the necessary module
# Import-Module Exchange

# Get the current date
$currentDate = Get-Date

# Fetch all Exchange certificates and filter only the expired ones
$expiredCerts = Get-ExchangeCertificate | Where-Object { $_.NotAfter -lt $currentDate }

# Counter for successfully removed certificates
$removedCount = 0

# Remove the expired certificates
foreach ($cert in $expiredCerts) {
    try {
        Write-Host "Attempting to remove expired certificate with thumbprint: $($cert.Thumbprint)"
        Remove-ExchangeCertificate -Thumbprint $cert.Thumbprint -Confirm:$false | Out-Null
        # Increment the counter only if removal is successful
        $removedCount++
    } catch {
        Write-Host "Error removing certificate with thumbprint: $($cert.Thumbprint). Error: $_"
    }
}

# Output a message
if ($removedCount -eq 0) {
    Write-Host "No expired certificates removed."
} else {
    Write-Host "$removedCount expired certificate(s) removed."
}
