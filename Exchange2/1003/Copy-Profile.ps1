# PowerShell Script to copy profile.ps1 to WindowsPowerShell folder in System32 and SysWOW64

# Define the source file path
$sourceFile = "C:\path\to\profile.ps1"

# Define the destination folders
$destFolders = @(
    "C:\Windows\System32\WindowsPowerShell\v1.0\", 
    "C:\Windows\SysWOW64\WindowsPowerShell\v1.0\"
)

# Loop through each destination folder and copy the file
foreach ($folder in $destFolders) {
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
}
