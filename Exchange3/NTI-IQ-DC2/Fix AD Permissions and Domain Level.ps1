# 1. Enable inheritance for the user
function Enable-UserInheritance {
    param($UserDN)
    
    try {
        $acl = Get-Acl -Path "AD:\$UserDN"
        Write-Host "Current inheritance status: $(-not $acl.AreAccessRulesProtected)" -ForegroundColor Yellow
        
        # Enable inheritance and preserve existing explicit permissions
        $acl.SetAccessRuleProtection($false, $true)
        Set-Acl -Path "AD:\$UserDN" -AclObject $acl
        
        Write-Host "Successfully enabled inheritance" -ForegroundColor Green
    } catch {
        Write-Host "Error enabling inheritance: $_" -ForegroundColor Red
    }
}

# 2. Check Domain and Forest Functional Levels
function Get-DomainInfo {
    $domain = Get-ADDomain
    $forest = Get-ADForest
    
    Write-Host "`nCurrent Domain Level: $($domain.DomainMode)" -ForegroundColor Yellow
    Write-Host "Current Forest Level: $($forest.ForestMode)" -ForegroundColor Yellow
    
    # Check for 2016 DCs
    $dcs = Get-ADDomainController -Filter *
    $has2016DC = $false
    
    Write-Host "`nDomain Controllers:" -ForegroundColor Cyan
    foreach($dc in $dcs) {
        $os = Get-ADObject $dc.ComputerObjectDN -Properties OperatingSystem
        Write-Host "$($dc.Name) - $($os.OperatingSystem)" -ForegroundColor Yellow
        if($os.OperatingSystem -like "*2016*" -or $os.OperatingSystem -like "*2019*" -or $os.OperatingSystem -like "*2022*") {
            $has2016DC = $true
        }
    }
    
    return @{
        Has2016DC = $has2016DC
        DomainLevel = $domain.DomainMode
        ForestLevel = $forest.ForestMode
    }
}

# Main execution
$userDN = "CN=Nathaniel Alexander,OU=Users,OU=Information Technology and Systems,OU=2-DepartmentalUnits,DC=iq,DC=nti,DC=local"

Write-Host "Step 1: Enabling Inheritance" -ForegroundColor Cyan
Enable-UserInheritance -UserDN $userDN

Write-Host "`nStep 2: Checking Domain Configuration" -ForegroundColor Cyan
$domainInfo = Get-DomainInfo

if (-not $domainInfo.Has2016DC) {
    Write-Host "`nAction Required:" -ForegroundColor Red
    Write-Host "1. Add a Windows Server 2016 (or newer) Domain Controller" -ForegroundColor Yellow
}

if ($domainInfo.DomainLevel -lt "Windows2016Domain") {
    Write-Host "`nTo raise domain and forest functional levels, use these commands:" -ForegroundColor Yellow
    Write-Host "Set-ADDomainMode -Identity $((Get-ADDomain).DNSRoot) -DomainMode Windows2016Domain" -ForegroundColor Gray
    Write-Host "Set-ADForestMode -Identity $((Get-ADForest).Name) -ForestMode Windows2016Forest" -ForegroundColor Gray
}

Write-Host "`nAfter raising functional levels and having a 2016 DC, run these commands:" -ForegroundColor Yellow