<#
.SYNOPSIS
.DESCRIPTION
	A longer description of the function, its purpose, common use cases, etc.
.NOTES
	Information or caveats about the function e.g. 'This function is not supported in Linux'
.LINK
	Specify a URI to a help page, this will show when Get-Help -Online is used.
.EXAMPLE
	Test-MyTestFunction -Verbose
	Explanation of the function or its result. You can include multiple examples with additional .EXAMPLE lines

Active Directory server not configured for SSL, test connection to LDAP://gbrhcdcldc05.AGH.com:636 did not work.
Active Directory server correctly configured for SSL, test connection to LDAP://ALNDC.AGH.com:636 completed.
Active Directory server not configured for SSL, test connection to LDAP://gbrhcdcldc06.AGH.com:636 did not work.
Active Directory server correctly configured for SSL, test connection to LDAP://AGH-DC01.AGH.com:636 completed.
Active Directory server correctly configured for SSL, test connection to LDAP://AGH-DC02.AGH.com:636 completed.

#>



##################
#### TEST ALL AD DCs for LDAPS
##################
# $AllDCs = Get-ADDomainController -Filter * -Server AGH-DC01.AGH.com | Select-Object Hostname
$AllDCs = Get-ADDomainController -Filter * -Server "AGH.com" | Select-Object Hostname
 foreach ($dc in $AllDCs) {
	$LDAPS = [ADSI]"LDAP://$($dc.hostname):636"
	#write-host $LDAPS
	try {
   	$Connection = [adsi]($LDAPS)
	} Catch {
	}
	If ($Connection.Path) {
   	Write-Host "Active Directory server correctly configured for SSL, test connection to $($LDAPS.Path) completed."
	} Else {
   	Write-Host "Active Directory server not configured for SSL, test connection to LDAP://$($dc.hostname):636 did not work."
	}
 }