# Define the path for the CSV file
$exportPath = "C:\code\AD\exports\AD_OUs.csv" # Change this to your desired path

# Export the list of OUs to a CSV file
Get-ADOrganizationalUnit -Filter * | Select-Object Name, DistinguishedName | Export-Csv -Path $exportPath -NoTypeInformation

Get-ADOrganizationalUnit -Filter * | Select-Object Name, DistinguishedName | Out-GridView