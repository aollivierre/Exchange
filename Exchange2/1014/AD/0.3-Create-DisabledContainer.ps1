# Import the ActiveDirectory module
Import-Module ActiveDirectory

# Get the root domain distinguished name
$rootDomain = (Get-ADDomain).DistinguishedName

# Create the new container called "Disabled Computers" in the root domain
New-ADObject -Type container -Name "Disabled Users - Container" -Path $rootDomain
