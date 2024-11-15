# Define the destination folders for PowerShell Core (7+)
$destFoldersPowerShellCore = @(
    [System.IO.Path]::Combine($env:ProgramFiles, "PowerShell\7\"),
    [System.IO.Path]::Combine($env:USERPROFILE, "Documents\PowerShell\"),
    [System.IO.Path]::Combine($env:ProgramFiles, "PowerShell\7-preview\") # If you're also using preview versions
)

# Loop through each destination folder and remove profile.ps1 if it exists
foreach ($folder in $destFoldersPowerShellCore) {
    # Construct the file path for profile.ps1 in the current folder
    $profilePath = Join-Path -Path $folder -ChildPath "profile.ps1"
    
    # Check if profile.ps1 exists at the path
    if (Test-Path -Path $profilePath) {
        # Remove profile.ps1
        Remove-Item -Path $profilePath -Force
        
        # Output status
        Write-Output "Removed $profilePath successfully."
    } else {
        # Output status if file does not exist
        Write-Output "File $profilePath does not exist."
    }
}
