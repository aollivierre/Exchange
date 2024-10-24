# [17:02:23] Certificate details:
# [17:02:23] 

# Subject      : CN=AGH.com, O=IT, O=Almonte General Hospital., L=Almonte, S=Ontario, C=CA
# Issuer       : CN=AGH.com, O=IT, O=Almonte General Hospital., L=Almonte, S=Ontario, C=CA
# Thumbprint   : 38BE999AF50350CF591CFB365EF0C3EBA34616C8
# FriendlyName : LDAPS-SSO-FortiMail-AGH
# NotBefore    : 10/14/2023 4:51:20 PM
# NotAfter     : 10/11/2033 4:51:20 PM
# Extensions   : {System.Security.Cryptography.Oid, System.Security.Cryptography.Oid, System.Security.Cryptography.Oid, 
#                System.Security.Cryptography.Oid...}


function Get-CertDetails {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Thumbprint
    )

    # Helper function to write color-coded, timestamped messages
    function Write-TimestampedMessage {
        param(
            [Parameter(Mandatory = $true)]
            [string]$Message,

            [Parameter(Mandatory = $false)]
            [ConsoleColor]$Color = 'White'
        )

        $timestamp = Get-Date -Format "HH:mm:ss"
        Write-Host "[$timestamp] $Message" -ForegroundColor $Color
    }

    # Load the certificate store
    $certStore = New-Object System.Security.Cryptography.X509Certificates.X509Store("Root", "LocalMachine")
    $certStore.Open("ReadOnly")

    try {
        # Find the certificate based on the thumbprint
        $cert = $certStore.Certificates | Where-Object { $_.Thumbprint -eq $Thumbprint }

        if ($null -ne $cert) {
            Write-TimestampedMessage "Certificate details:" -Color Cyan

            # Print certificate details
            $cert | Format-List | Out-String | ForEach-Object { Write-TimestampedMessage $_ }

        } else {
            Write-TimestampedMessage "Certificate not found!" -Color Red
        }
    } finally {
        # Close the certificate store
        $certStore.Close()
    }
}

# To use the function
Get-CertDetails -Thumbprint "38BE999AF50350CF591CFB365EF0C3EBA34616C8"
