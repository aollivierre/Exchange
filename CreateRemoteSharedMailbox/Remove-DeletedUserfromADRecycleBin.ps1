<#
.SYNOPSIS
Removes a deleted object from the Active Directory Recycle Bin.

.DESCRIPTION
This function removes a deleted object from the Active Directory Recycle Bin. It prompts the user to enter the name of the object to be removed, and then checks if the object is still in the recycle bin. If the object is found, it is permanently removed from the recycle bin. If the object is not found, a message is displayed indicating that no object was found with the specified name.

.PARAMETER Name
The name of the object to be removed from the recycle bin.

.EXAMPLE
RemoveFromADRecycleBin -Name "John Doe"
Removes the deleted object with the name "John Doe" from the Active Directory Recycle Bin.

.NOTES
Author: Unknown
Date: Unknown
#>

function RemoveFromADRecycleBin {
    param (
        [string]$Name = Read-Host -Prompt "Enter the name for the mailbox"
    )

    try {
        # Check if the deleted object is still in the recycle bin
        $filter = "isDeleted -eq `$True -and Name -like '*$Name*'"
        $deletedObject = Get-ADObject -Filter $filter -IncludeDeletedObjects -ErrorAction Ignore
        if ($null -ne $deletedObject) {
            # Permanently remove the object from the recycle bin
            $deletedObject | Remove-ADObject -Confirm:$false
            Write-Host ("[" + (Get-Date) + "] Removed object $Name from AD Recycle Bin successfully.") -ForegroundColor Green
        }
        else {
            Write-Host ("[" + (Get-Date) + "] No object found in AD Recycle Bin with name $Name.") -ForegroundColor Yellow
        }
    }
    catch {
        Write-Error "An error occurred: $_"
    }
}

# Call the function
RemoveFromADRecycleBin
