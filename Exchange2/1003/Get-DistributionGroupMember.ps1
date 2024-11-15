# Initialize an empty array to hold the results
$results = @()

# Get all Distribution Groups
$groups = Get-DistributionGroup -ResultSize Unlimited

# Loop through each Distribution Group
foreach ($group in $groups) {
    $groupName = $group.DisplayName
    $groupEmail = $group.PrimarySmtpAddress
    $groupAlias = $group.Alias
    $groupType = $group.GroupType
    $managedBy = $group.ManagedBy -join ', '

    # Get members of the Distribution Group
    $members = Get-DistributionGroupMember -Identity $groupName -ResultSize Unlimited

    # Initialize an empty array to hold member names
    $memberNames = @()

    # Loop through each member and add to the memberNames array
    foreach ($member in $members) {
        $memberNames += $member.DisplayName
    }

    # Create a PSObject to hold the group attributes and member names
    $result = [PSCustomObject]@{
        'GroupName'         = $groupName
        'PrimarySmtpAddress' = $groupEmail
        'Alias'             = $groupAlias
        'GroupType'         = $groupType
        'ManagedBy'         = $managedBy
        'Members'           = -join ($memberNames -join ', ')
    }

    # Add the result to the results array
    $results += $result
}

# Export to a CSV file
$results | Format-Table -AutoSize
$results | Export-Csv -Path "C:\code\exports\Glebe_DL_Sept_05_2023.csv" -NoTypeInformation
