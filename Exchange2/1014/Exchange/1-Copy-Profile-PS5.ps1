# PowerShell Script to copy profile.ps1 to WindowsPowerShell folders
# Note: Exchange Remote Management PowerShell snap-in is only supported in PowerShell 5.1


# PowerShell Script to copy profile.ps1 to WindowsPowerShell folders
<#
IMPORTANT NOTE:
This script copies the profile to PowerShell 5.1 locations to enable automatic loading of the Exchange 
PowerShell snap-in. This approach is critical for managing hybrid Exchange environments after 
decommissioning the last on-premises Exchange server, as the Exchange Management Shell (EMS) 
will no longer be available. The snap-in provides continued access to Exchange PowerShell cmdlets 
for managing hybrid configuration and recipients.

The snap-in is only compatible with PowerShell 5.1, which is why this script targets WindowsPowerShell 
locations only.
#>


# Define the source file path
$sourceFile = "C:\Code\Exchange\profile.ps1"

# Define the destination folders for Windows PowerShell
$destFoldersWindowsPowerShell = @(
    "C:\Windows\System32\WindowsPowerShell\v1.0\", 
    "C:\Windows\SysWOW64\WindowsPowerShell\v1.0\"
)

# Loop through each destination folder and copy the file
foreach ($folder in $destFoldersWindowsPowerShell) {
    # Check if folder exists before attempting to copy
    if (Test-Path -Path $folder) {
        # Construct the destination file path
        $destFile = Join-Path -Path $folder -ChildPath "profile.ps1"

        # Copy the file
        Copy-Item -Path $sourceFile -Destination $destFile -Force

        # Output the status
        if (Test-Path -Path $destFile) {
            Write-Output "Copied $sourceFile to $destFile successfully."
        } else {
            Write-Output "Failed to copy $sourceFile to $destFile."
        }
    } else {
        Write-Output "Destination folder $folder does not exist."
    }
}

Write-Output "Note: Profile copied only to PowerShell 5.1 locations as Exchange Remote Management PowerShell snap-in is not supported in PowerShell 7+


This script copies the profile to PowerShell 5.1 locations to enable automatic loading of the Exchange 
PowerShell snap-in. This approach is critical for managing hybrid Exchange environments after 
decommissioning the last on-premises Exchange server, as the Exchange Management Shell (EMS) 
will no longer be available. The snap-in provides continued access to Exchange PowerShell cmdlets 
for managing hybrid configuration and recipients.

The snap-in is only compatible with PowerShell 5.1, which is why this script targets WindowsPowerShell 
locations only.


"