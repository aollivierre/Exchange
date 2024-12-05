function Get-MailContactsOUAudit {
    [CmdletBinding()]
    param()

    try {
        $results = [System.Collections.ArrayList]::new(1000)
        $contacts = Get-MailContact -ResultSize Unlimited | Select-Object DisplayName, PrimarySmtpAddress, DistinguishedName, EmailAddresses
        
        $tunngavikPattern = [regex]"tunngavik\.com"
        
        foreach ($contact in $contacts) {
            $dnParts = $contact.DistinguishedName.Split(',', 2)
            $ouPath = $dnParts[1]
            $isInternal = if ($contact.EmailAddresses -match $tunngavikPattern) { "Yes" } else { "No" }

            [void]$results.Add([PSCustomObject]@{
                DisplayName = $contact.DisplayName
                PrimarySmtpAddress = $contact.PrimarySmtpAddress
                OrganizationalUnit = $ouPath
                InternalMailFlow = $isInternal
            })
        }

        $ouStats = $results | Group-Object OrganizationalUnit -NoElement | Select-Object Name, Count

        Write-Host "`nOU Placement Summary:" -ForegroundColor Green
        $ouStats | ForEach-Object { Write-Host "$($_.Name): $($_.Count) contacts" -ForegroundColor Yellow }

        $results | Sort-Object OrganizationalUnit, DisplayName | 
            Out-HtmlView -Title "Mail Contacts OU Audit - $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    }
    catch {
        Write-Error "Error in OU audit: $_"
    }
}

Get-MailContactsOUAudit