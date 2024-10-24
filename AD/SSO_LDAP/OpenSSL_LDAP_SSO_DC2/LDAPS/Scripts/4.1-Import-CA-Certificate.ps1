# VERBOSE: Performing the operation "Import certificate" on target "Item:
# C:\Code\AD\SSO_LDAP\OpenSSL_LDAP_SSO_DC2\LDAPS\ca.crt Destination: Root".


#    PSParentPath: Microsoft.PowerShell.Security\Certificate::LocalMachine\Root

# Thumbprint                                Subject
# ----------                                -------
# 38BE999AF50350CF591CFB365EF0C3EBA34616C8  CN=AGH.com, O=IT, O=Almonte General Hospital., L=Almonte, S=Ontario, C=CA


#import the cert as a trusted CA on the domain controller
Import-Certificate -FilePath ca.crt  -CertStoreLocation 'Cert:\LocalMachine\Root' -Verbose