# Get current domain name
$domainName = (Get-ADDomain).DNSRoot

# # Define AAD Connect sync configuration
# $containerInclusions = @(
#     "DC=ott,DC=nti,DC=local",
#     "OU=Staff,OU=NTIOTT,DC=ott,DC=nti,DC=local"
# )
# $containerExclusions = @(
#     "CN=LostAndFound,DC=ott,DC=nti,DC=local",
#     "CN=Configuration,DC=ott,DC=nti,DC=local",
#     "OU=NTIOTT,DC=ott,DC=nti,DC=local",
#     "CN=Program Data,DC=ott,DC=nti,DC=local",
#     "CN=System,DC=ott,DC=nti,DC=local",
#     "CN=Microsoft Exchange System Objects,DC=ott,DC=nti,DC=local",
#     "OU=Microsoft Exchange Security Groups,DC=ott,DC=nti,DC=local",
#     "CN=Managed Service Accounts,DC=ott,DC=nti,DC=local",
#     "CN=Infrastructure,DC=ott,DC=nti,DC=local",
#     "CN=ForeignSecurityPrincipals,DC=ott,DC=nti,DC=local",
#     "CN=Builtin,DC=ott,DC=nti,DC=local",
#     "OU=NotSyncedtoEID,DC=ott,DC=nti,DC=local"
# )





$containerInclusions = @(
    "DC=iq,DC=nti,DC=local"
)

  $containerExclusions= @(
   "CN=LostAndFound,DC=iq,DC=nti,DC=local",
    "CN=Managed Service Accounts,DC=iq,DC=nti,DC=local",
    "OU=Microsoft Exchange Security Groups,DC=iq,DC=nti,DC=local",
    "CN=Microsoft Exchange System Objects,DC=iq,DC=nti,DC=local",
    "OU=NotSyncedtoEID,DC=iq,DC=nti,DC=local",
    "CN=System,DC=iq,DC=nti,DC=local",
    "OU=NotSyncedContactsEID,DC=iq,DC=nti,DC=local"
  )




# Get all OUs and Containers in the domain
$allOUsAndContainers = @()
$allOUsAndContainers += Get-ADOrganizationalUnit -Filter * -Properties CanonicalName, DistinguishedName, Created
$allOUsAndContainers += Get-ADObject -Filter {ObjectClass -eq "container"} -Properties CanonicalName, DistinguishedName, Created

# Define complete list of default AD containers and OUs
$defaultContainers = @(
    "CN=Builtin",
    "CN=Computers",
    "CN=Domain Controllers",
    "CN=ForeignSecurityPrincipals",
    "CN=Keys",
    "CN=LostAndFound",
    "CN=Managed Service Accounts",
    "CN=Microsoft Exchange Security Groups",
    "CN=Microsoft Exchange System Objects",
    "CN=NTDS Quotas",
    "CN=Program Data",
    "CN=System",
    "CN=TPM Devices",
    "CN=Infrastructure",
    "CN=Users"
)

# Function to determine sync status
function Get-SyncStatus {
    param (
        [string]$dn
    )
    
    if ($containerInclusions -contains $dn) {
        return "Explicitly Included"
    }
    elseif ($containerExclusions -contains $dn) {
        return "Explicitly Excluded"
    }
    
    foreach ($inclusion in $containerInclusions) {
        if ($dn.EndsWith($inclusion)) {
            $isExcluded = $false
            foreach ($exclusion in $containerExclusions) {
                if ($dn.EndsWith($exclusion)) {
                    $isExcluded = $true
                    break
                }
            }
            if (-not $isExcluded) {
                return "Included (Inherited)"
            }
        }
    }
    
    foreach ($exclusion in $containerExclusions) {
        if ($dn.EndsWith($exclusion)) {
            return "Excluded (Inherited)"
        }
    }
    
    return "Not Configured"
}

# Initialize arrays to store results
$OUsWithComputers = @()
$OUsWithoutComputers = @()
$DefaultContainersAndOUs = @()

# Check each OU/Container for objects
foreach ($object in $allOUsAndContainers) {
    # Get counts for different object types
    $userCount = @(Get-ADUser -Filter * -SearchBase $object.DistinguishedName -SearchScope OneLevel -ErrorAction SilentlyContinue).Count
    $contactCount = @(Get-ADObject -Filter {ObjectClass -eq "contact"} -SearchBase $object.DistinguishedName -SearchScope OneLevel -ErrorAction SilentlyContinue).Count
    $groupCount = @(Get-ADGroup -Filter * -SearchBase $object.DistinguishedName -SearchScope OneLevel -ErrorAction SilentlyContinue).Count
    $computerCount = @(Get-ADComputer -Filter * -SearchBase $object.DistinguishedName -SearchScope OneLevel -ErrorAction SilentlyContinue).Count
    
    # Create custom object with details
    $objectInfo = [PSCustomObject]@{
        Name = $object.Name
        Type = if ($object.ObjectClass -eq "container") { "Container" } else { "OU" }
        Created = $object.Created
        CanonicalName = $object.CanonicalName
        DistinguishedName = $object.DistinguishedName
        UserCount = $userCount
        ContactCount = $contactCount
        GroupCount = $groupCount
        ComputerCount = $computerCount
        TotalObjects = $userCount + $contactCount + $groupCount + $computerCount
        SyncStatus = Get-SyncStatus -dn $object.DistinguishedName
    }
    
    # Check if this is a default container/OU
    $isDefault = $false
    foreach ($defaultContainer in $defaultContainers) {
        if ($object.DistinguishedName -match [regex]::Escape($defaultContainer)) {
            $isDefault = $true
            break
        }
    }
    
    # Categorize based on computer presence and default status
    if ($isDefault) {
        $DefaultContainersAndOUs += $objectInfo
    }
    elseif ($computerCount -gt 0) {
        $OUsWithComputers += $objectInfo
    }
    else {
        $OUsWithoutComputers += $objectInfo
    }
}

# Generate timestamp for files
$timestamp = Get-Date -Format "yyyyMMdd-HHmm"

# Export to CSV
$OUsWithComputers | Export-Csv -Path "OUs_With_Computers_$domainName`_$timestamp.csv" -NoTypeInformation
$OUsWithoutComputers | Export-Csv -Path "OUs_Without_Computers_$domainName`_$timestamp.csv" -NoTypeInformation
$DefaultContainersAndOUs | Export-Csv -Path "Default_Containers_and_OUs_$domainName`_$timestamp.csv" -NoTypeInformation

# Create and display HTML reports using Out-HTMLView
$OUsWithComputers | Sort-Object Name | 
    Select-Object Name, Type, Created, CanonicalName, DistinguishedName, UserCount, ContactCount, GroupCount, ComputerCount, TotalObjects, SyncStatus | 
    Out-HtmlView -Title "$domainName - Custom OUs/Containers With Computer Objects - Traditional Entra Connect Sync ($timestamp)" -FilePath "OUs_With_Computers_$domainName`_$timestamp.html"

$OUsWithoutComputers | Sort-Object Name | 
    Select-Object Name, Type, Created, CanonicalName, DistinguishedName, UserCount, ContactCount, GroupCount, ComputerCount, TotalObjects, SyncStatus | 
    Out-HtmlView -Title "$domainName - Custom OUs/Containers Without Computer Objects - Modern Cloud Sync ($timestamp)" -FilePath "OUs_Without_Computers_$domainName`_$timestamp.html"

$DefaultContainersAndOUs | Sort-Object Name | 
    Select-Object Name, Type, Created, CanonicalName, DistinguishedName, UserCount, ContactCount, GroupCount, ComputerCount, TotalObjects, SyncStatus | 
    Out-HtmlView -Title "$domainName - Default AD/Exchange Containers and OUs ($timestamp)" -FilePath "Default_Containers_and_OUs_$domainName`_$timestamp.html"

# Display summary in console with fixed counting
Write-Host "`nDomain: $domainName" -ForegroundColor Yellow
Write-Host "`nSummary:" -ForegroundColor Yellow
Write-Host "Total OUs and Containers: $($allOUsAndContainers.Count)"
Write-Host "Custom OUs/Containers with computers: $($OUsWithComputers.Count)"
Write-Host "Custom OUs/Containers without computers: $($OUsWithoutComputers.Count)"
Write-Host "Default AD/Exchange Containers and OUs: $($DefaultContainersAndOUs.Count)"
Write-Host "`nObject Counts:"

# Calculate totals properly by combining all three collections
$allObjects = @() + $OUsWithComputers + $OUsWithoutComputers + $DefaultContainersAndOUs

$totalUsers = ($allObjects | ForEach-Object { $_.UserCount } | Measure-Object -Sum).Sum
$totalContacts = ($allObjects | ForEach-Object { $_.ContactCount } | Measure-Object -Sum).Sum
$totalGroups = ($allObjects | ForEach-Object { $_.GroupCount } | Measure-Object -Sum).Sum
$totalComputers = ($allObjects | ForEach-Object { $_.ComputerCount } | Measure-Object -Sum).Sum

Write-Host "Total Users: $totalUsers"
Write-Host "Total Contacts: $totalContacts"
Write-Host "Total Groups: $totalGroups"
Write-Host "Total Computers: $totalComputers"

Write-Host "`nSync Status Summary:"
Write-Host "Explicitly Included: $(($allObjects | Where-Object {$_.SyncStatus -eq 'Explicitly Included'}).Count)"
Write-Host "Explicitly Excluded: $(($allObjects | Where-Object {$_.SyncStatus -eq 'Explicitly Excluded'}).Count)"
Write-Host "Included (Inherited): $(($allObjects | Where-Object {$_.SyncStatus -eq 'Included (Inherited)'}).Count)"
Write-Host "Excluded (Inherited): $(($allObjects | Where-Object {$_.SyncStatus -eq 'Excluded (Inherited)'}).Count)"
Write-Host "Not Configured: $(($allObjects | Where-Object {$_.SyncStatus -eq 'Not Configured'}).Count)"
Write-Host "`nReports have been exported to CSV and HTML files in the current directory."