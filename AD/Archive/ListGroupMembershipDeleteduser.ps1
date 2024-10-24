# Define the username of the deleted user
$deletedUserName = "Jordan Heuser"

# Import the Active Directory module
Import-Module ActiveDirectory

# Retrieve the Security logs where user deletion events are recorded
# Event ID 4726 indicates a user was deleted
$events = Get-WinEvent -LogName Security -FilterHashtable @{Id=4726} -ErrorAction SilentlyContinue

# Iterate over the events to find the one related to the deleted user
$deletedUserEvents = $events | Where-Object {
    $_.Properties[0].Value -eq $deletedUserName
}

# Initialize an array to store the groups
$userGroups = @()

# For each deletion event, check the user SID and list group memberships
foreach ($event in $deletedUserEvents) {
    $userSID = $event.Properties[0].Value

    # Query group memberships from the logs or backup (hypothetical example)
    # Note: This part depends on how you have auditing or backups set up.
    # You might need to adjust this part to fit your specific environment.
    $groupEvents = Get-WinEvent -LogName Security -FilterHashtable @{Id=4732; Message="*'$userSID'*"} -ErrorAction SilentlyContinue
    $userGroups += $groupEvents | ForEach-Object { $_.Properties[1].Value }
}

# Remove duplicates and format the output
$userGroups = $userGroups | Select-Object -Unique

# Display the groups
if ($userGroups) {
    Write-Host "The user '$deletedUserName' was a member of the following groups:"
    $userGroups | ForEach-Object { Write-Host $_ }
} else {
    Write-Host "No group memberships found for the deleted user '$deletedUserName'."
}
