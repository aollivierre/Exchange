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
    Test-MyTestFunction -Verbose
    Explanation of the function or its result. You can include multiple examples with additional .EXAMPLE lines


    

#>



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

# Domain Controller and Port
$domainController = "AGH-DC01.AGH.com"
# $domainController = [System.Net.Dns]::GetHostEntry([System.Net.Dns]::GetHostName()).HostName
$port = 636

# Get the script's directory
$scriptDir = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent

# CA Certificate Path
$caCertFileName = "ca.crt"
$caCertPath = Join-Path -Path $scriptDir -ChildPath $caCertFileName

Write-TimestampedMessage "Attempting to establish a secure connection to $domainController on port $port..." -Color Cyan

# Run the OpenSSL command
$OpenSSLarguments = @('s_client', '-connect', "${domainController}:${port}", '-CAfile', $caCertPath)
& $opensslPath $OpenSSLarguments


# if ($?) { # Check last command status
#     Write-TimestampedMessage "Connection attempt complete. Check above for detailed connection output." -Color Green
# } else {
#     Write-TimestampedMessage "Failed to establish a connection!" -Color Red
# }