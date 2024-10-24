# AD Recycle Bin Search and Export Tool

# Function to show search examples
function Show-SearchExamples {
    $examples = @"
Available search options:
1. Search by last name
2. Search by email/UPN
3. Search by display name
4. Search by SAM account name
5. Search all fields (recommended)

The search is not case sensitive and will find partial matches.
Example: Searching for 'gau' will find 'Gauthier'
"@
    Write-Host $examples -ForegroundColor Cyan
}

# Function to export selected objects
function Export-SelectedObjects {
    param (
        [Parameter(Mandatory=$true)]
        [Object[]]$Objects,
        [string]$SearchCriteria
    )
    
    # Generate export filename with timestamp
    $timestamp = Get-Date -Format "yyyyMMddHHmmss"
    $defaultPath = "$env:USERPROFILE\Documents"
    $filename = "DeletedUsers_${SearchCriteria}_$timestamp.xml"
    
    # Prompt for export location
    $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveDialog.InitialDirectory = $defaultPath
    $saveDialog.Filter = "XML Files (*.xml)|*.xml|All Files (*.*)|*.*"
    $saveDialog.FileName = $filename
    
    if ($saveDialog.ShowDialog() -eq 'OK') {
        $outputPath = $saveDialog.FileName
        $Objects | Export-Clixml -Path $outputPath
        Write-Host "`nSelected objects exported to: $outputPath" -ForegroundColor Green
    } else {
        Write-Host "`nExport cancelled by user." -ForegroundColor Yellow
    }
}

# Import required modules
Import-Module ActiveDirectory
Add-Type -AssemblyName System.Windows.Forms

# Check if the Active Directory Recycle Bin is enabled
$recycleBinEnabled = Get-ADOptionalFeature -Filter { Name -like "Recycle Bin Feature" }

if (-not $recycleBinEnabled) {
    Write-Host "The Active Directory Recycle Bin is not enabled. Please enable it to retrieve deleted objects." -ForegroundColor Red
    exit
}

# Show examples to the user
Show-SearchExamples

# Prompt for search type
Write-Host "`nSelect search type (enter number 1-5):" -ForegroundColor Green
$searchType = Read-Host "Choice"
$searchValue = Read-Host "Enter search value"

Write-Host "`nSearching for deleted objects..." -ForegroundColor Yellow

# Retrieve all deleted user objects from the AD Recycle Bin
$deletedUsers = Get-ADObject -Filter {
    IsDeleted -eq $true -and
    ObjectClass -eq "user"
} -IncludeDeletedObjects -Properties *

# Filter based on search type
$filteredUsers = switch ($searchType) {
    "1" { # Last name
        $deletedUsers | Where-Object { 
            $_.DistinguishedName -like "*$searchValue*" -or 
            $_.Name -like "*$searchValue*" -or 
            $_.Surname -like "*$searchValue*"
        }
    }
    "2" { # Email/UPN
        $deletedUsers | Where-Object { 
            $_.UserPrincipalName -like "*$searchValue*" -or 
            $_.mail -like "*$searchValue*" -or
            $_.proxyAddresses -like "*$searchValue*"
        }
    }
    "3" { # Display name
        $deletedUsers | Where-Object { $_.DisplayName -like "*$searchValue*" }
    }
    "4" { # SAM account
        $deletedUsers | Where-Object { $_.SamAccountName -like "*$searchValue*" }
    }
    "5" { # All fields
        $deletedUsers | Where-Object { 
            $_.DistinguishedName -like "*$searchValue*" -or 
            $_.Name -like "*$searchValue*" -or 
            $_.Surname -like "*$searchValue*" -or
            $_.UserPrincipalName -like "*$searchValue*" -or 
            $_.mail -like "*$searchValue*" -or
            $_.proxyAddresses -like "*$searchValue*" -or
            $_.DisplayName -like "*$searchValue*" -or
            $_.SamAccountName -like "*$searchValue*" -or
            $_.GivenName -like "*$searchValue*"
        }
    }
    default {
        Write-Host "Invalid search type selected. Using all fields search." -ForegroundColor Yellow
        $deletedUsers | Where-Object { 
            $_.DistinguishedName -like "*$searchValue*" -or 
            $_.Name -like "*$searchValue*" -or 
            $_.Surname -like "*$searchValue*" -or
            $_.UserPrincipalName -like "*$searchValue*" -or 
            $_.mail -like "*$searchValue*" -or
            $_.proxyAddresses -like "*$searchValue*" -or
            $_.DisplayName -like "*$searchValue*" -or
            $_.SamAccountName -like "*$searchValue*" -or
            $_.GivenName -like "*$searchValue*"
        }
    }
}

# Check if any filtered users were found
if ($null -eq $filteredUsers -or $filteredUsers.Count -eq 0) {
    Write-Host "`nNo deleted users found matching: $searchValue" -ForegroundColor Yellow
    
    # Show what was searched
    Write-Host "`nDebug Information:" -ForegroundColor Yellow
    Write-Host "Total deleted users found in recycle bin: $($deletedUsers.Count)" -ForegroundColor Yellow
    Write-Host "Search value used: $searchValue" -ForegroundColor Yellow
    Write-Host "Search type used: $searchType" -ForegroundColor Yellow
} else {
    # Create a custom object with relevant properties for GridView
    $gridViewObjects = $filteredUsers | Select-Object @(
        @{Name='Name'; Expression={$_.Name}},
        @{Name='Display Name'; Expression={$_.DisplayName}},
        @{Name='User Principal Name'; Expression={$_.UserPrincipalName}},
        @{Name='SAM Account'; Expression={$_.SamAccountName}},
        @{Name='Last Known OU'; Expression={$_.LastKnownParent}},
        @{Name='When Deleted'; Expression={$_.whenChanged}},
        @{Name='When Created'; Expression={$_.whenCreated}},
        @{Name='Object GUID'; Expression={$_.ObjectGUID}},
        @{Name='Distinguished Name'; Expression={$_.DistinguishedName}}
    )
    
    Write-Host "`nFound $($filteredUsers.Count) deleted users matching your criteria." -ForegroundColor Green
    
    # Show results in GridView with selection enabled
    $selected = $gridViewObjects | Out-GridView -Title "Deleted Users - Select objects to export (Ctrl+Click for multiple)" -PassThru
    
    # Export selected objects if any were chosen
    if ($selected) {
        $selectedFull = $filteredUsers | Where-Object { 
            $guid = $_.ObjectGUID
            $selected.'Object GUID' -contains $guid
        }
        Export-SelectedObjects -Objects $selectedFull -SearchCriteria $searchValue
    } else {
        Write-Host "`nNo objects selected for export." -ForegroundColor Yellow
    }
}

Write-Host "`nScript execution completed." -ForegroundColor Green