$exchangeServer = Get-ExchangeServer | Where-Object { $_.IsE14OrLater -eq $true } | Select-Object -First 1
$domainName = $exchangeServer.Fqdn

$UserCredential = Get-Credential
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://$domainName/PowerShell/" -Authentication Kerberos -Credential $UserCredential
Import-PSSession $Session -DisableNameChecking


# $UserCredential = Get-Credential
# $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://ambico-mail.doors-ambico.com/PowerShell/ -Authentication Kerberos -Credential $UserCredential
# Import-PSSession $Session -DisableNameChecking