function RemoveFromADRecycleBin {
    param (
        [string]$Name = "DMARC3"
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
