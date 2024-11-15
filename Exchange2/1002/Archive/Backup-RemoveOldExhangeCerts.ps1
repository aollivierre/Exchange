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
    # Backup the certificate to the defined directory
    $certFile = Join-Path -Path $backupDirectory -ChildPath ("cert_" + $cert.Thumbprint + ".pfx")
    Export-PfxCertificate -Cert $cert.PSPath -FilePath $certFile -Password (ConvertTo-SecureString -String "8,6DE!)1M1ZeRI2+U)J{4<o)F(=7[{+[" -Force -AsPlainText)
    
    # Remove the certificate from Exchange
    $cert | Remove-ExchangeCertificate -Confirm:$false
}

Write-Output "Expired certificates have been backed up and removed."
