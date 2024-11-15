#run the following on EMS on-prem to get the Exchange Server name

# $exchangeServer = Get-ExchangeServer | Where-Object { $_.IsE14OrLater -eq $true } | Select-Object -First 1
# $domainName = $exchangeServer.Fqdn


# Add-Type -AssemblyName "System.DirectoryServices.AccountManagement"
# $contextType = [System.DirectoryServices.AccountManagement.ContextType]::Domain
# $principalContext = New-Object -TypeName System.DirectoryServices.AccountManagement.PrincipalContext -ArgumentList $contextType
# $domainName = $principalContext.ConnectedServer



$domainName = "GLB-EX01.GLEBE.LOCAL"

$UserCredential = Get-Credential
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://$domainName/PowerShell/" -Authentication Kerberos -Credential $UserCredential
Import-PSSession $Session -DisableNameChecking