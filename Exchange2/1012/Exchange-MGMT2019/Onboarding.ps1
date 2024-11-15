 $Credentials = Get-Credential
New-RemoteMailBox -Name "Tara Gvozenovic" -Password $Credentials.Password -UserPrincipalName TGvozdenovic@glebecentre.ca