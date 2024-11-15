# Define the backup directory
$backupDirectory = "C:\code\backup-certs\"

# Create the backup directory if it doesn't exist
if (-not (Test-Path $backupDirectory)) {
    New-Item -Path $backupDirectory -ItemType Directory
}

# Get all expired Exchange certificates
$expiredCerts = Get-ExchangeCertificate | Where-Object { $_.NotAfter -lt (Get-Date) }

# Backup and remove each expired certificate
foreach ($cert in $expiredCerts) {
    # Backup the certificate to the defined directory in .cer format with a timestamp to ensure uniqueness
    $timestamp = Get-Date -Format "yyyyMMddHHmmss"
    $certFile = Join-Path -Path $backupDirectory -ChildPath ("cert_" + $cert.Thumbprint + "_" + $timestamp + ".pfx")
    
    try {
        Export-ExchangeCertificate -Thumbprint $cert.Thumbprint -BinaryEncoded:$true -Password (ConvertTo-SecureString -String "Cj9WsYv,x5k04shNu5ZASAr>=Q>N[H;F" -AsPlainText -Force) -FileName $certFile
    } catch {
        Write-Error "Failed to export certificate with thumbprint: $cert.Thumbprint. Error: $_"
        continue
    }

    # Remove the certificate from Exchange
    try {
        # $cert | Remove-ExchangeCertificate -Confirm:$false
    } catch {
        Write-Error "Failed to remove certificate with thumbprint: $cert.Thumbprint. Error: $_"
        continue
    }
}

Write-Output "Expired certificates have been backed up"
