# Script to create high-level admin user with specific permissions
# Warning: This script creates a user with extensive administrative privileges

function New-HighPrivilegeUser {
    # Parameters for the new user
    $userParams = @{
        GivenName = "Alaaeldin"
        Surname = "Mukhtar"
        Name = "Alaaeldin Mukhtar"
        DisplayName = "Alaaeldin Mukhtar"
        SamAccountName = "amukhtar"
        UserPrincipalName = "amukhtar@tunngavik.com"
        AccountPassword = (ConvertTo-SecureString "NOVAnti321!" -AsPlainText -Force)
        Enabled = $true
        PasswordNeverExpires = $true
        CannotChangePassword = $false
        Path = "CN=Users,$(([ADSI]'LDAP://RootDSE').defaultNamingContext)"
    }

    try {
        # Check if user already exists
        $existingUser = Get-ADUser -Filter {SamAccountName -eq "amukhtar"} -ErrorAction SilentlyContinue
        
        if ($existingUser) {
            Write-Host "User 'amukhtar' already exists. Checking group memberships..." -ForegroundColor Yellow
            $user = $existingUser
        } else {
            Write-Host "Creating new user 'amukhtar'..." -ForegroundColor Green
            $user = New-ADUser @userParams -PassThru
            Write-Host "User created successfully." -ForegroundColor Green
        }

        # Array of admin groups to add the user to
        $adminGroups = @(
            "Domain Admins",
            "Schema Admins",
            "Enterprise Admins"
        )

        # Add user to admin groups
        foreach ($group in $adminGroups) {
            try {
                $groupMember = Get-ADGroupMember -Identity $group -Recursive | 
                    Where-Object {$_.SamAccountName -eq "amukhtar"}
                
                if (-not $groupMember) {
                    Add-ADGroupMember -Identity $group -Members $user
                    Write-Host "Added to $group successfully." -ForegroundColor Green
                } else {
                    Write-Host "Already a member of $group." -ForegroundColor Yellow
                }
            }
            catch {
                $errorMessage = $_.Exception.Message
                Write-Host "Error adding to $group`: $errorMessage" -ForegroundColor Red
            }
        }

        # Verify final settings
        $verifyUser = Get-ADUser -Identity "amukhtar" -Properties *
        
        Write-Host "`nFinal User Configuration:" -ForegroundColor Cyan
        Write-Host "Username: $($verifyUser.SamAccountName)"
        Write-Host "Display Name: $($verifyUser.DisplayName)"
        Write-Host "UPN: $($verifyUser.UserPrincipalName)"
        Write-Host "Enabled: $($verifyUser.Enabled)"
        Write-Host "Password Never Expires: $($verifyUser.PasswordNeverExpires)"
        
        Write-Host "`nGroup Memberships:" -ForegroundColor Cyan
        Get-ADPrincipalGroupMembership -Identity "amukhtar" | 
            Select-Object -ExpandProperty Name |
            ForEach-Object { Write-Host "- $_" }

    }
    catch {
        $errorMessage = $_.Exception.Message
        Write-Host "Error occurred: $errorMessage" -ForegroundColor Red
        Write-Host "Please ensure you have sufficient permissions to perform these actions." -ForegroundColor Red
    }
}

# Run the function
Write-Host "=== Creating High-Privilege Admin User ===" -ForegroundColor Cyan
Write-Host "This script will create a user with high-level administrative privileges." -ForegroundColor Yellow
Write-Host "Please ensure you want to proceed with this action.`n" -ForegroundColor Yellow

$confirm = Read-Host "Do you want to proceed? (y/n)"
if ($confirm -eq 'y') {
    New-HighPrivilegeUser
} else {
    Write-Host "Operation cancelled." -ForegroundColor Yellow
}