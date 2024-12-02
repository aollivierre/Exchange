[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Function to ensure PSWriteHTML module is installed
function Install-RequiredModule {
    if (-not (Get-Module -ListAvailable -Name PSWriteHTML)) {
        Write-Host "Installing PSWriteHTML module..." -ForegroundColor Yellow
        try {
            Install-Module -Name PSWriteHTML -Force -AllowClobber -Scope CurrentUser
            Import-Module PSWriteHTML
            Write-Host "PSWriteHTML module installed successfully." -ForegroundColor Green
        }
        catch {
            Write-Host "Error installing PSWriteHTML module: $_" -ForegroundColor Red
            Write-Host "Please run PowerShell as administrator and try again." -ForegroundColor Yellow
            exit
        }
    }
    else {
        Import-Module PSWriteHTML
    }
}

# Install and import required module
Install-RequiredModule



function Get-MailContactsRoutingAudit {
    [CmdletBinding()]
    param ()

    try {
        Write-Host "Auditing internal mail flow contacts..." -ForegroundColor Cyan
        $results = [System.Collections.ArrayList]::new()
        
        $contacts = Get-CBMailContact -ResultSize Unlimited
        
        foreach ($contact in $contacts) {
            $emailAddressStrings = $contact.EmailAddresses | ForEach-Object { $_.ToString() }
            
            # Check if any email address contains tunngavik.com
            $isInternalMailFlow = $emailAddressStrings | Where-Object { $_ -match "tunngavik\.com" }
            
            $routingAddress = $emailAddressStrings | 
                Where-Object { $_ -match "smtp:.*?@tunngavik\.mail\.onmicrosoft\.com$" } | 
                Select-Object -First 1

            if ($isInternalMailFlow) {
                [void]$results.Add([PSCustomObject]@{
                    DisplayName = $contact.DisplayName
                    PrimarySmtpAddress = $contact.PrimarySmtpAddress
                    ContactType = "Internal Mail Flow"
                    HasCorrectRoutingAddress = if ($routingAddress) { "Yes" } else { "No" }
                    CurrentRoutingAddress = if ($routingAddress) { 
                        $routingAddress.Replace("smtp:", "") 
                    } else { 
                        "Missing tunngavik.mail.onmicrosoft.com" 
                    }
                    EmailAddresses = [string]::Join(", ", $emailAddressStrings)
                })
            }
        }

        $missingRouting = ($results | Where-Object { $_.HasCorrectRoutingAddress -eq "No" }).Count
        $totalInternalContacts = $results.Count

        Write-Host "`nAudit Summary:" -ForegroundColor Green
        Write-Host "Total Internal Mail Flow Contacts: $totalInternalContacts" -ForegroundColor Yellow
        Write-Host "Missing tunngavik.mail.onmicrosoft.com routing: $missingRouting" -ForegroundColor Red

        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $results | Sort-Object HasCorrectRoutingAddress, DisplayName | 
            Out-HtmlView -Title "Internal Mail Flow Contacts Routing Audit - $timestamp"
    }
    catch {
        Write-Error "Error in audit: $_"
        $_.Exception | Format-List -Force
    }
}

Get-MailContactsRoutingAudit