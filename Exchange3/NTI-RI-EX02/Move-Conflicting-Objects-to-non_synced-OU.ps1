# Define target OU
$targetOU = Get-ADOrganizationalUnit -Filter "Name -eq 'NotSyncedtoEID'" | Select-Object -ExpandProperty DistinguishedName

if (-not $targetOU) {
    Write-Host "Target OU 'NotSyncedtoEID' not found!" -ForegroundColor Red
    exit
}

# Array of conflicting object details from the XML
$conflicts = @(
    @{
        ProxyAddress = "SMTP:MarineMonitoring@tunngavik.com"
        ConflictGuid = "68826cc2-0e66-480b-a2d3-d0b06d05b7e1"
    },
    @{
        ProxyAddress = "SMTP:PKotierk@tunngavik.com"
        ConflictGuid = "0b179282-72a9-4863-807f-54856372de9a"
    },
    @{
        ProxyAddress = "SMTP:sedmunds@tunngavik.com"
        ConflictGuid = "23dfafcf-30f2-4e25-b0f3-1443e589b3ed"
    },
    @{
        ProxyAddress = "SMTP:TAoudla-Henrie3@tunngavik.com"
        ConflictGuid = "6fb8f768-2ece-41c6-9ba0-ec868d3551ec"
    }
)

foreach ($conflict in $conflicts) {
    Write-Host "`nProcessing conflict for: $($conflict.ProxyAddress)" -ForegroundColor Yellow
    
    # Find both the conflicting objects
    $objects = Get-ADObject -LDAPFilter "(proxyAddresses=$($conflict.ProxyAddress))" -Properties ObjectClass, ProxyAddresses, DisplayName, DistinguishedName
    
    if ($objects) {
        foreach ($object in $objects) {
            Write-Host "`nFound object:" -ForegroundColor Green
            Write-Host "Display Name: $($object.DisplayName)" -ForegroundColor Green
            Write-Host "Object Class: $($object.ObjectClass)" -ForegroundColor Green
            Write-Host "Current Location: $($object.DistinguishedName)" -ForegroundColor Green
            
            # Only move if not already in NotSyncedtoEID
            if ($object.DistinguishedName -notlike "*$targetOU*") {
                try {
                    Move-ADObject -Identity $object.DistinguishedName -TargetPath $targetOU
                    Write-Host "Successfully moved object to NotSyncedtoEID" -ForegroundColor Green
                    
                    # Verify the move
                    $movedObject = Get-ADObject -Identity $object.ObjectGUID -Properties DistinguishedName
                    Write-Host "New location: $($movedObject.DistinguishedName)" -ForegroundColor Green
                }
                catch {
                    Write-Host "Error moving object: $_" -ForegroundColor Red
                }
            }
            else {
                Write-Host "Object already in NotSyncedtoEID OU" -ForegroundColor Yellow
            }
        }
    }
    else {
        Write-Host "No objects found with proxy address: $($conflict.ProxyAddress)" -ForegroundColor Red
    }
}

Write-Host "`nAll conflicting objects have been processed." -ForegroundColor Cyan