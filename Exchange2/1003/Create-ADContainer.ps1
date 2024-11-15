# Import the Active Directory module if not already loaded
Import-Module ActiveDirectory

# Specify the parent DN where the new container will be created
$parentDN = "DC=example,DC=com" # Change this to your domain distinguished name

# Specify the name of the new container
$containerName = "Nova_Admin_Do_Not_AADSync"

# Create the new container
New-ADObject -Type container -Name $containerName -Path $parentDN
