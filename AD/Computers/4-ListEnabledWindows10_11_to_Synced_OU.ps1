# Import the ActiveDirectory module
Import-Module ActiveDirectory

# Set the export path for the CSV file
$exportPath = "C:\code\AD\Exports\Computers_$(Get-Date -Format 'yyyy-MM-dd_HH-mm-ss').csv" # Adjust this path as needed

# Retrieve all enabled computers running Windows 10 or Windows 11 from the domain
$targetComputers = Get-ADComputer -Filter {(Enabled -eq $True) -and (OperatingSystem -like "Windows 10*" -or OperatingSystem -like "Windows 11*")}

# Export the data to a CSV file
$targetComputers | Export-Csv $exportPath -NoTypeInformation

# Display the data in a grid view
$targetComputers | Out-GridView

# Output the total count of computers with time stamp
$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
Write-Host "Timestamp: $timestamp - Total number of enabled computers running Windows 10 or Windows 11: $($targetComputers.Count)" -ForegroundColor Green
