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
    

   PSParentPath: Microsoft.PowerShell.Security\Certificate::LocalMachine\My

Thumbprint                                Subject
----------                                -------
FB3B4DCBD1325C4C660FEC41CC490D40B05C69C3  CN=agh.com Root Cert
D7789B2F3FC69F77D561FE5A1492D66C91CBCFF7  CN=agh-dc01.agh.com
9B1715B6E26464080A795D43A35E59E6D3CA4277  CN=AGH.com << this is the new cert imported into Cert:\LocalMachine\My

#>

# We can check that the cert has been imported by running the following powershell. We should see CN=example.com

Get-ChildItem "Cert:\LocalMachine\My"