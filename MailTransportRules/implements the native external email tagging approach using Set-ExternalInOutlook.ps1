# I'll help create a PowerShell function that implements the native external email tagging approach using `Set-ExternalInOutlook`.

# ```powershell

#Requires -Modules ExchangeOnlineManagement, PSWriteHTML

function Connect-ExchangeOnlineWithReport {
    [CmdletBinding()]
    param()
    
    try {
        $null = Get-OrganizationConfig -ErrorAction Stop
        Write-Host "Already connected to Exchange Online" -ForegroundColor Green
    }
    catch {
        Write-Host "Connecting to Exchange Online..." -ForegroundColor Yellow
        Connect-ExchangeOnline -ErrorAction Stop
    }
}


#Requires -Modules ExchangeOnlineManagement, PSWriteHTML

function Set-ExternalEmailTag {
    [CmdletBinding()]
    param (
        [Parameter()]
        [bool]$Enabled = $true,
        
        [Parameter()]
        [string[]]$AllowList,
        
        [Parameter()]
        [string]$ReportPath = ".\ExternalEmailTag"
    )
    
    $null = New-Item -ItemType Directory -Force -Path $ReportPath
    $results = [System.Collections.ArrayList]::new()

    try {
        Connect-ExchangeOnlineWithReport
        $currentConfig = Get-ExternalInOutlook
        
        if ($AllowList) {
            $config = Set-ExternalInOutlook -Enabled $Enabled -AllowList $AllowList
        }
        else {
            $config = Set-ExternalInOutlook -Enabled $Enabled
        }
        
        $resultObject = [PSCustomObject]@{
            Status = 'Success'
            Enabled = $config.Enabled
            AllowList = $config.AllowList -join '; '
            PreviousState = $currentConfig.Enabled
            ErrorMessage = ''
        }
    }
    catch {
        $resultObject = [PSCustomObject]@{
            Status = 'Failed'
            Enabled = $null
            AllowList = ''
            PreviousState = $null
            ErrorMessage = $_.Exception.Message
        }
    }
    
    $null = $results.Add($resultObject)
    Export-ExternalTagReport -Results $results -ReportPath $ReportPath
    return $resultObject
}

function Export-ExternalTagReport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $Results,
        
        [Parameter(Mandatory)]
        [string]$ReportPath
    )
    
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $htmlPath = Join-Path $ReportPath "ExternalEmailTag_$timestamp.html"
    
    New-HTML -Title "External Email Tag Configuration Report" -FilePath $htmlPath -ShowHTML {
        New-HTMLSection -HeaderText "Configuration Summary" {
            New-HTMLTable -DataTable $Results -ScrollX {
                New-TableCondition -Name 'Status' -ComparisonType string -Operator eq -Value 'Failed' -BackgroundColor Salmon -Color Black
                New-TableCondition -Name 'Status' -ComparisonType string -Operator eq -Value 'Success' -BackgroundColor LightGreen -Color Black
            }
        }
    }
    
    Write-Host "`nReport generated at: $htmlPath" -ForegroundColor Green
}

# Example usage:
$params = @{
    Enabled = $true
    AllowList = @("trusted-domain.com", "partner@example.com")
    ReportPath = "C:\ExchangeReports"
}

Set-ExternalEmailTag @params
# ```

# This implementation:
# 1. Uses native Exchange Online cmdlets for external tagging
# 2. Supports allow-listing trusted domains/emails
# 3. Generates HTML reports using PSWriteHTML
# 4. Includes error handling and connection management
# 5. Reports previous and current configuration states

# Note: Changes may take 24-48 hours to propagate fully across the organization.