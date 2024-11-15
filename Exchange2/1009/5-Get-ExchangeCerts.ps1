# let's break down the requirements:

# List all Exchange certificates.
# Search for these certificates in the Local Computer certificate store.
# Search for these certificates in the Current User certificate store.
# Indicate which certificates are found.
# Count the total number of certificates, expired certificates, certificates found in Exchange, and certificates found in both certificate stores.
# Here's a PowerShell script that accomplishes these tasks:


# This script will give you an overview of the Exchange certificates and their presence in both the Local Computer and Current User certificate stores.

# Before you run the script, ensure you have the necessary permissions to access Exchange and the certificate stores. The script assumes that you have the Exchange Management Shell installed and loaded. If you run this script on an Exchange Server directly, it should work without any modifications. If you run it remotely, you'll need to set up a remote session to the Exchange server.




# Initialize counters
$totalCerts = 0
$expiredCerts = 0
$foundInExchange = 0
$foundInComputerStore = 0
$foundInUserStore = 0

# Get all Exchange certificates
$exchangeCerts = Get-ExchangeCertificate
$totalCerts += $exchangeCerts.Count
$foundInExchange += $exchangeCerts.Count

# Loop through each Exchange certificate
foreach ($cert in $exchangeCerts) {
    # Check if the certificate is expired
    if ($cert.NotAfter -lt (Get-Date)) {
        $expiredCerts++
    }
    
    # Check if the certificate exists in Local Computer certificate store
    $localCert = Get-ChildItem -Path Cert:\LocalMachine\My | Where-Object { $_.Thumbprint -eq $cert.Thumbprint }
    if ($localCert) {
        $foundInComputerStore++
    }

    # Check if the certificate exists in Current User certificate store
    $userCert = Get-ChildItem -Path Cert:\CurrentUser\My | Where-Object { $_.Thumbprint -eq $cert.Thumbprint }
    if ($userCert) {
        $foundInUserStore++
    }
}

# Output the results
Write-Host "Total Certificates: $totalCerts"
Write-Host "Expired Certificates: $expiredCerts"
Write-Host "Certificates found in Exchange: $foundInExchange"
Write-Host "Certificates found in Local Computer store: $foundInComputerStore"
Write-Host "Certificates found in Current User store: $foundInUserStore"

