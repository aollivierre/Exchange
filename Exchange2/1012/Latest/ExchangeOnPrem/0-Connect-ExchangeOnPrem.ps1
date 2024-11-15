# [System.Net.Dns]::GetHostEntry([System.Net.Dns]::GetHostName()).HostName

#Exchange Server FQDN
$ExFQDN = "GLB-EX01.GLEBE.LOCAL"

$UserCredential = Get-Credential
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://$ExFQDN/PowerShell/" -Authentication Kerberos -Credential $UserCredential
Import-PSSession $Session -DisableNameChecking