# List all deleted objects in the AD Recycle Bin
try {
    $deletedObjects = Get-ADObject -Filter {isDeleted -eq $True} -IncludeDeletedObjects -ErrorAction Stop
    if ($deletedObjects) {
        # $deletedObjects | Format-Table Name, ObjectClass, LastKnownParent -AutoSize
        $deletedObjects | Format-Table Name
    }
    else {
        Write-Host "No deleted objects found in the AD Recycle Bin." -ForegroundColor Yellow
    }
}
catch {
    Write-Error "An error occurred: $_"
}
