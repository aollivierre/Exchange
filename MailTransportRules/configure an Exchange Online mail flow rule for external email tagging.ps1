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

function New-ExternalEmailTag {
    [CmdletBinding()]
    param (
        [Parameter()]
        [string]$RuleName = "External Email Warning Banner",
        
        [Parameter()]
        [string]$HeaderText = "EXTERNAL EMAIL",
        
        [Parameter()]
        [string]$WarningText = "This email originated from outside of the organization. Do not click links or open attachments unless you recognize the sender and know the content is safe.",
        
        [Parameter()]
        [string]$ReportPath = ".\ExternalEmailTag"
    )
    
    $null = New-Item -ItemType Directory -Force -Path $ReportPath
    
    $htmlTemplate = @"
<div style="background-color:#FFEB9C; border:1px solid #9C6500; padding:5px;">
<span style="color:#9C6500; font-weight:bold;">${HeaderText}: </span>
${WarningText}
</div>
"@

    $results = [System.Collections.ArrayList]::new()

    try {
        Connect-ExchangeOnlineWithReport

        $ruleParams = @{
            Name                              = $RuleName
            FromScope                         = 'NotInOrganization'
            SentToScope                       = 'InOrganization'
            ApplyHtmlDisclaimerText           = $htmlTemplate
            ApplyHtmlDisclaimerLocation       = 'Prepend'
            ApplyHtmlDisclaimerFallbackAction = 'Wrap'
            ErrorAction                       = 'Stop'
        }
        
        $rule = New-TransportRule @ruleParams
        
        $resultObject = [PSCustomObject]@{
            RuleName         = $rule.Name
            Status           = 'Success'
            CreationTime     = $rule.WhenCreated
            LastModifiedTime = $rule.WhenChanged
            ErrorMessage     = ''
        }
    }
    catch {
        $resultObject = [PSCustomObject]@{
            RuleName         = $RuleName
            Status           = 'Failed'
            CreationTime     = $null
            LastModifiedTime = $null
            ErrorMessage     = $_.Exception.Message
        }
    }
    
    $null = $results.Add($resultObject)
    
    # Generate report
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $htmlPath = Join-Path $ReportPath "ExternalEmailTag_$timestamp.html"
    
    New-HTML -Title "External Email Tag Configuration Report" -FilePath $htmlPath -ShowHTML {
        New-HTMLSection -HeaderText "Configuration Summary" {
            New-HTMLPanel {
                New-HTMLText -Text @"
                <h3>Rule Details</h3>
                <ul>
                    <li>Rule Name: $($resultObject.RuleName)</li>
                    <li>Status: $($resultObject.Status)</li>
                    <li>Created: $($resultObject.CreationTime)</li>
                    <li>Last Modified: $($resultObject.LastModifiedTime)</li>
                </ul>
"@
            }
        }
        
        New-HTMLSection -HeaderText "Configuration Results" {
            New-HTMLTable -DataTable $results -ScrollX {
                New-TableCondition -Name 'Status' -ComparisonType string -Operator eq -Value 'Failed' -BackgroundColor Salmon -Color Black
                New-TableCondition -Name 'Status' -ComparisonType string -Operator eq -Value 'Success' -BackgroundColor LightGreen -Color Black
            }
        }
    }
    
    Write-Host "`nReport generated at: $htmlPath" -ForegroundColor Green
    
    return $resultObject
}

# Example usage:
$params = @{
    RuleName    = "External Email Warning Banner"
    HeaderText  = "EXTERNAL EMAIL"
    WarningText = "This email originated from outside of the organization. Exercise caution."
    ReportPath  = "C:\ExchangeReports"
}

New-ExternalEmailTag @params