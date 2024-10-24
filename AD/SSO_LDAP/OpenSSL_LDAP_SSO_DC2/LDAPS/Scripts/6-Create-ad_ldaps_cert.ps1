<#
.SYNOPSIS
    A short one-line action-based description, e.g. 'Tests if a function is valid'
.DESCRIPTION
    A longer description of the function, its purpose, common use cases, etc.
.NOTES
    Information or caveats about the function e.g. 'This function is not supported in Linux'
.LINK
    Specify a URI to a help page, this will show when Get-Help -Online is used.
.EXAMPLE

[17:17:53] Creating ad_ldaps_cert by signing the csr...
Certificate request self-signature ok
subject=CN = AGH.com
Enter pass phrase for ca.key:
[17:18:12] Successfully created the certificate at ad_ldaps_cert.crt

#>




function CreateADLDAPSCert {
    param (
        [Parameter(Mandatory=$true)]
        [string]$CsrPath,
        [Parameter(Mandatory=$true)]
        [string]$CaCrtPath,
        [Parameter(Mandatory=$true)]
        [string]$CaKeyPath,
        [Parameter(Mandatory=$true)]
        [string]$ExtfilePath,
        [Parameter(Mandatory=$true)]
        [string]$OutputCertPath
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

    # Set the path for openssl
    $opensslPath = "C:\Program Files\OpenSSL-Win64\bin\openssl.exe"

    Write-TimestampedMessage "Creating ad_ldaps_cert by signing the csr..." -Color Cyan

    # Run the OpenSSL command
    & $opensslPath x509 -req -days 825 -in $CsrPath -CA $CaCrtPath -CAkey $CaKeyPath -extfile $ExtfilePath -set_serial 01 -out $OutputCertPath

    if ($?) { # Check last command status
        Write-TimestampedMessage "Successfully created the certificate at $OutputCertPath" -Color Green
    } else {
        Write-TimestampedMessage "Failed to create the certificate!" -Color Red
    }
}

# To use the function
CreateADLDAPSCert -CsrPath "ad.csr" -CaCrtPath "ca.crt" -CaKeyPath "ca.key" -ExtfilePath "v3ext.txt" -OutputCertPath "ad_ldaps_cert.crt"