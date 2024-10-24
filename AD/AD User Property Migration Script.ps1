# AD User Property Migration Script

# Function to get detailed user properties
function Get-DetailedUserProperties {
    param (
        [Parameter(Mandatory=$true)]
        [Object]$User
    )
    
    return [ordered]@{
        "Profile Settings" = @{
            "Home Directory" = $User.HomeDirectory
            "Home Drive" = $User.HomeDrive
            "Profile Path" = $User.ProfilePath
            "Script Path" = $User.ScriptPath
            "Terminal Services Home Directory" = $User.TerminalServicesHomeDirectory
            "Terminal Services Home Drive" = $User.TerminalServicesHomeDrive
            "Terminal Services Profile Path" = $User.TerminalServicesProfilePath
        }
        "User Information" = @{
            "Display Name" = $User.DisplayName
            "First Name" = $User.givenName
            "Last Name" = $User.sn
            "Sam Account Name" = $User.SamAccountName
            "User Principal Name" = $User.UserPrincipalName
            "Distinguished Name" = $User.DistinguishedName
            "Description" = $User.Description
            "Email" = $User.mail
            "Mail Nickname" = $User.mailNickname
            "Target Address" = $User.targetAddress
            "Proxy Addresses" = ($User.proxyAddresses -join "; ")
        }
        "Employment Details" = @{
            "Department" = $User.Department
            "Title" = $User.Title
            "Company" = $User.Company
            "Manager" = $User.Manager
            "Office" = $User.physicalDeliveryOfficeName
        }
        "Contact Information" = @{
            "Street Address" = $User.StreetAddress
            "City" = $User.l
            "State/Province" = $User.st
            "Postal Code" = $User.PostalCode
            "Country" = $User.c
            "Phone Number" = $User.telephoneNumber
            "Mobile" = $User.Mobile
            "Fax" = $User.Fax
        }
        "Account Details" = @{
            "Account Created" = $User.whenCreated
            "Account Modified" = $User.whenChanged
            "Last Logon" = if($User.lastLogon){[datetime]::FromFileTime($User.lastLogon)}else{$null}
            "Password Last Set" = if($User.pwdLastSet){[datetime]::FromFileTime($User.pwdLastSet)}else{$null}
            "Account Expires" = if($User.accountExpires -and $User.accountExpires -ne 9223372036854775807){[datetime]::FromFileTime($User.accountExpires)}else{"Never"}
            "Primary Group ID" = $User.primaryGroupID
        }
        "Groups" = @{
            "Member Of" = $User.memberOf
        }
    }
}

function Get-DeletedUserDetails {
    param (
        [string]$LastName
    )
    
    try {
        # Retrieve all deleted user objects from the AD Recycle Bin
        $deletedUsers = Get-ADObject -Filter {
            IsDeleted -eq $true -and
            ObjectClass -eq "user"
        } -IncludeDeletedObjects -Properties *

        # Filter for users with matching last name
        $filteredUsers = $deletedUsers | Where-Object { 
            $_.DistinguishedName -like "*$LastName*" -or 
            $_.Name -like "*$LastName*" -or 
            $_.Surname -like "*$LastName*"
        }

        if ($null -eq $filteredUsers) {
            Write-Host "No deleted users found with last name: $LastName" -ForegroundColor Yellow
            return $null
        }

        # Create GridView-friendly objects
        $gridViewObjects = $filteredUsers | Select-Object @(
            @{Name='Display Name'; Expression={$_.DisplayName}},
            @{Name='Sam Account Name'; Expression={$_.SamAccountName}},
            @{Name='User Principal Name'; Expression={$_.UserPrincipalName}},
            @{Name='Department'; Expression={$_.Department}},
            @{Name='Title'; Expression={$_.Title}},
            @{Name='State'; Expression={$_.st}},
            @{Name='Email'; Expression={$_.mail}},
            @{Name='When Deleted'; Expression={$_.whenChanged}},
            @{Name='Distinguished Name'; Expression={$_.DistinguishedName}}
        )

        # Let user select the correct user if multiple found
        $selectedUser = $gridViewObjects | Out-GridView -Title "Select the correct deleted user" -PassThru

        if ($null -eq $selectedUser) {
            Write-Host "No user selected." -ForegroundColor Yellow
            return $null
        }

        # Get full user object for selected user
        $fullUserDetails = $filteredUsers | Where-Object { 
            $_.DistinguishedName -eq $selectedUser.'Distinguished Name' 
        } | Select-Object * -First 1

        return $fullUserDetails
    }
    catch {
        Write-Host "Error retrieving deleted user: $($_.Exception.Message)" -ForegroundColor Red
        return $null
    }
}

function Show-UserProperties {
    param (
        [object]$UserObject,
        [string]$Title = "User Properties"
    )

    $properties = Get-DetailedUserProperties -User $UserObject

    # Display properties in GridView
    $gridViewData = foreach ($section in $properties.Keys) {
        foreach ($prop in $properties[$section].Keys) {
            [PSCustomObject]@{
                Section = $section
                Property = $prop
                Value = $properties[$section][$prop]
            }
        }
    }

    $selectedProps = $gridViewData | Out-GridView -Title "$Title (Review and Click OK to proceed)" -PassThru

    # Also display in console with colors
    Write-Host "`n$Title" -ForegroundColor Cyan
    foreach ($section in $properties.Keys) {
        Write-Host "`n$section" -ForegroundColor Yellow
        foreach ($prop in $properties[$section].Keys) {
            if ($properties[$section][$prop]) {
                Write-Host "  $prop : " -NoNewline -ForegroundColor Gray
                Write-Host "$($properties[$section][$prop])" -ForegroundColor White
            }
        }
    }

    # Export to XML and CSV
    $timestamp = Get-Date -Format "yyyyMMddHHmmss"
    $baseFilename = "$($UserObject.SamAccountName)_Properties_$timestamp"
    
    # Export to XML
    $properties | Export-Clixml -Path "$baseFilename.xml"
    Write-Host "`nProperties exported to: $baseFilename.xml" -ForegroundColor Green
    
    # Export to CSV
    $gridViewData | Export-Csv -Path "$baseFilename.csv" -NoTypeInformation
    Write-Host "Properties exported to: $baseFilename.csv" -ForegroundColor Green

    return $properties
}


function Set-UserProperties {
    param (
        [string]$NewUserName,
        [object]$Properties
    )

    try {
        # Get the new user
        $newUser = Get-ADUser -Identity $NewUserName -Properties *
        
        if ($null -eq $newUser) {
            Write-Host "New user not found: $NewUserName" -ForegroundColor Red
            return $false
        }

        # Create hashtable for splatting
        $profileSettings = @{}

        # Add properties with correct Set-ADUser parameter names
        # Profile Settings
        if ($Properties.'Profile Settings'.'Home Directory') { 
            $profileSettings['HomeDirectory'] = $Properties.'Profile Settings'.'Home Directory' 
        }
        if ($Properties.'Profile Settings'.'Home Drive') { 
            $profileSettings['HomeDrive'] = $Properties.'Profile Settings'.'Home Drive' 
        }
        if ($Properties.'Profile Settings'.'Profile Path') { 
            $profileSettings['ProfilePath'] = $Properties.'Profile Settings'.'Profile Path' 
        }
        if ($Properties.'Profile Settings'.'Script Path') { 
            $profileSettings['ScriptPath'] = $Properties.'Profile Settings'.'Script Path' 
        }
        
        # Employment Details
        if ($Properties.'Employment Details'.'Title') { 
            $profileSettings['Title'] = $Properties.'Employment Details'.'Title' 
        }
        if ($Properties.'Employment Details'.'Department') { 
            $profileSettings['Department'] = $Properties.'Employment Details'.'Department' 
        }
        if ($Properties.'Employment Details'.'Company') { 
            $profileSettings['Company'] = $Properties.'Employment Details'.'Company' 
        }
        if ($Properties.'Employment Details'.'Manager') { 
            $profileSettings['Manager'] = $Properties.'Employment Details'.'Manager' 
        }
        
        # Contact Information
        if ($Properties.'Contact Information'.'Street Address') { 
            $profileSettings['StreetAddress'] = $Properties.'Contact Information'.'Street Address' 
        }
        if ($Properties.'Contact Information'.'City') { 
            $profileSettings['City'] = $Properties.'Contact Information'.'City' 
        }
        if ($Properties.'Contact Information'.'State/Province') { 
            $profileSettings['State'] = $Properties.'Contact Information'.'State/Province' 
        }
        if ($Properties.'Contact Information'.'Postal Code') { 
            $profileSettings['PostalCode'] = $Properties.'Contact Information'.'Postal Code' 
        }
        if ($Properties.'Contact Information'.'Country') { 
            $profileSettings['Country'] = $Properties.'Contact Information'.'Country' 
        }
        if ($Properties.'Contact Information'.'Phone Number') { 
            $profileSettings['OfficePhone'] = $Properties.'Contact Information'.'Phone Number' 
        }
        if ($Properties.'Contact Information'.'Mobile') { 
            $profileSettings['MobilePhone'] = $Properties.'Contact Information'.'Mobile' 
        }
        if ($Properties.'Contact Information'.'Fax') { 
            $profileSettings['Fax'] = $Properties.'Contact Information'.'Fax' 
        }
        
        # User Information
        if ($Properties.'User Information'.'Description') { 
            $profileSettings['Description'] = $Properties.'User Information'.'Description' 
        }
        if ($Properties.'User Information'.'Email') { 
            $profileSettings['EmailAddress'] = $Properties.'User Information'.'Email' 
        }

        Write-Host "Updating user properties..." -ForegroundColor Yellow
        Write-Host "Properties being set:" -ForegroundColor Yellow
        $profileSettings.GetEnumerator() | ForEach-Object {
            Write-Host "  $($_.Key): $($_.Value)" -ForegroundColor Gray
        }
        
        # Update AD user properties using splatting
        Set-ADUser -Identity $NewUserName @profileSettings
        Write-Host "User properties updated successfully." -ForegroundColor Green

        # Add user to groups
        $groups = $Properties.Groups.'Member Of'
        if ($groups) {
            Write-Host "Adding user to groups..." -ForegroundColor Yellow
            foreach ($group in $groups) {
                try {
                    Add-ADGroupMember -Identity $group -Members $NewUserName -ErrorAction Stop
                    Write-Host "Added to group: $group" -ForegroundColor Green
                }
                catch {
                    Write-Host "Failed to add to group $group : $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
        }

        Write-Host "Successfully updated user properties" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "Error setting user properties: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}


# Main script
try {
    Write-Host "AD User Property Migration Tool" -ForegroundColor Cyan
    Write-Host "This tool will help you migrate properties from a deleted user to a new user." -ForegroundColor Yellow
    
    # Get deleted user details
    $lastName = Read-Host "`nEnter the last name of the deleted user"
    $deletedUser = Get-DeletedUserDetails -LastName $lastName

    if ($null -eq $deletedUser) {
        throw "No deleted user found or selected."
    }

    # Show properties and get confirmation
    $properties = Show-UserProperties -UserObject $deletedUser -Title "Deleted User Properties"
    
    $proceed = Read-Host "`nWould you like to proceed with migrating these properties? (Y/N)"
    if ($proceed -ne 'Y') {
        Write-Host "Operation cancelled by user." -ForegroundColor Yellow
        exit
    }

    # Get new user
    $newUserSam = Read-Host "`nEnter the SAM account name of the new user"
    
    # Set properties on new user
    $success = Set-UserProperties -NewUserName $newUserSam -Properties $properties
    
    if ($success) {
        # Show new user properties
        Write-Host "`nNew user properties after migration:" -ForegroundColor Cyan
        $newUser = Get-ADUser -Identity $newUserSam -Properties *
        Show-UserProperties -UserObject $newUser -Title "New User Properties"
    }
}
catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
}
finally {
    Write-Host "`nScript completed." -ForegroundColor Cyan
}