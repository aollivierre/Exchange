function GenerateLDAPSCertificates {
    # Set the path for openssl
    $opensslPath = "C:\Program Files\OpenSSL-Win64\bin\openssl.exe"

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

    # Create LDAPS directory and change directory
    Write-TimestampedMessage "Creating LDAPS directory..." -Color Cyan
    # New-Item -ItemType Directory -Name "LDAPS" -Force | Set-Location

    # Generate the ca key, create a password and keep it for use throughout this guide.
    Write-TimestampedMessage "Generating ca key..." -Color Cyan
    & $opensslPath genrsa -des3 -out ca.key 4096

    # Create ca cert with valid of 10 years with info based off the provided ca_san.conf file
    Write-TimestampedMessage "Creating ca cert based on provided ca_san.conf file..." -Color Cyan
    & $opensslPath req -new -x509 -extensions v3_ca -days 3650 -key ca.key -out ca.crt -config ca_san.conf

    # List the contents
    Write-TimestampedMessage "Listing the generated files..." -Color Cyan
    Get-ChildItem -Name
}

# To run the function
GenerateLDAPSCertificates
