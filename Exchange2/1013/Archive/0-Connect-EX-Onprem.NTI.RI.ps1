# Define the domain name of the Exchange server
$domainName = 'NTI-RI-EX02.RI.nti.local'
# $domainName = '192.168.0.22'

# uncomment the following connection tests when debugging
ping $domainName
nslookup $domainName
Test-NetConnection $domainName -Port 5985

# Path to the credentials file
$SecretsFile = Join-Path -Path $PSScriptRoot -ChildPath "secrets.exchange.$domainName.psd1"

function Get-ExchangeCredential {
    param (
        [string]$SecretsFilePath
    )

    if (Test-Path $SecretsFilePath) {
        try {
            $Cred = Import-Clixml -Path $SecretsFilePath
            # Test the credentials
            $TestSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://$domainName/PowerShell/" -Authentication Kerberos -Credential $Cred -ErrorAction Stop
            Remove-PSSession $TestSession
            return $Cred

            # logic to handle failed authentication and prompt the user to update their credentials. Automate Credential Refresh:
        } catch {
            Write-Warning "Stored credentials are invalid or expired. Please enter new credentials."
            Remove-Item -Path $SecretsFilePath -Force
            return Get-ExchangeCredential -SecretsFilePath $SecretsFilePath
        }
    } else {
        $Cred = Get-Credential
        $Cred | Export-Clixml -Path $SecretsFilePath
        return $Cred
    }
}

# Retrieve the credentials
$UserCredential = Get-ExchangeCredential -SecretsFilePath $SecretsFile

# Establish a new PowerShell session with the Exchange server
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://$domainName/PowerShell/" -Authentication Kerberos -Credential $UserCredential

# Import the session
Import-PSSession $Session -DisableNameChecking