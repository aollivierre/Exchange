function Get-MailContactsOUAudit {
    [CmdletBinding()]
    param()

    try {
        $results = [System.Collections.ArrayList]::new()
        $contacts = Get-MailContact -ResultSize Unlimited
        
        foreach ($contact in $contacts) {
            # Get distinguishedName and parse OU path
            $dnParts = $contact.DistinguishedName.Split(',')
            $ouPath = ($dnParts | Select-Object -Skip 1) -join ','
            
            # Determine if internal mail flow
            $emailAddressStrings = $contact.EmailAddresses | ForEach-Object { $_.ToString() }
            $isInternal = if ($emailAddressStrings | Where-Object { $_ -match "tunngavik\.com" }) { "Yes" } else { "No" }

            [void]$results.Add([PSCustomObject]@{
                DisplayName = $contact.DisplayName
                PrimarySmtpAddress = $contact.PrimarySmtpAddress
                OrganizationalUnit = $ouPath
                InternalMailFlow = $isInternal
            })
        }

        # Get OU statistics
        $ouStats = $results | Group-Object OrganizationalUnit | Select-Object Name, Count

        Write-Host "`nOU Placement Summary:" -ForegroundColor Green
        $ouStats | ForEach-Object {
            Write-Host "$($_.Name): $($_.Count) contacts" -ForegroundColor Yellow
        }

        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $results | Sort-Object OrganizationalUnit, DisplayName | 
            Out-HtmlView -Title "Mail Contacts OU Audit - $timestamp"
    }
    catch {
        Write-Error "Error in OU audit: $_"
    }
}

Get-MailContactsOUAudit