function Get-ExchangeCertificateStatus {
    param (
        [string]$ExchangeServer
    )

    # Load Exchange Management Shell if not already loaded
    if (-not (Get-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction SilentlyContinue)) {
        Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction Stop
    }

    # Fetch Exchange certificates and display relevant information
    try {
        $certificates = Get-ExchangeCertificate -Server $ExchangeServer
        $certificates | Select-Object Thumbprint, Services, Subject, @{Name="NotBefore";Expression={$_.NotBefore}}, @{Name="NotAfter";Expression={$_.NotAfter}}, Status
    } catch {
        Write-Error "Failed to retrieve Exchange certificates. Error: $_"
    }
}

# Usage: Get-ExchangeCertificateStatus -ExchangeServer "YourExchangeServerName"
