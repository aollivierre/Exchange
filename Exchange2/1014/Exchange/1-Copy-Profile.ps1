# PowerShell Script to copy profile.ps1 to WindowsPowerShell and PowerShell 7+ folders

# Define the source file path
$sourceFile = "C:\Code\Exchange\profile.ps1"

# Define the destination folders for Windows PowerShell
$destFoldersWindowsPowerShell = @(
    "C:\Windows\System32\WindowsPowerShell\v1.0\", 
    "C:\Windows\SysWOW64\WindowsPowerShell\v1.0\"
)

# Define the destination folders for PowerShell Core (7+)
# Note: These paths are for Windows. Adjust accordingly for other OSes.
$destFoldersPowerShellCore = @(
    [System.IO.Path]::Combine($env:ProgramFiles, "PowerShell\7\"),
    [System.IO.Path]::Combine($env:USERPROFILE, "Documents\PowerShell\"),
    [System.IO.Path]::Combine($env:ProgramFiles, "PowerShell\7-preview\") # If you're also using preview versions
)

# Combine all folders
$destFolders = $destFoldersWindowsPowerShell + $destFoldersPowerShellCore

# Loop through each destination folder and copy the file
foreach ($folder in $destFolders) {
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
