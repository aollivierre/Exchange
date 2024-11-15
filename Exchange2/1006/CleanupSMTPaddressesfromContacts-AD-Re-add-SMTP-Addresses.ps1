# Define the SMTP addresses to be added
$smtpAddressesToAdd = @(
    "smtp:LisaBerting@chfcanada365.mail.onmicrosoft.com",
    "smtp:LBerting@chfcanada.coop",
    "smtp:LBerting@fhcc.ca",
    "smtp:LBerting@chfc.coop",
    "smtp:LBerting@fhcc.coop",
    "smtp:LisaBerting@chfcanada.coop"
)

# Fetch the AD user for Lisa Berting
$lisaBertingUser = Get-ADUser -Filter { Name -eq "Lisa Berting" } -Properties proxyAddresses

if ($null -ne $lisaBertingUser) {
    Write-Host "Adding SMTP addresses to $($lisaBertingUser.Name):" -ForegroundColor Magenta

    # Iterate through each SMTP address and add it to the user's proxyAddresses
    foreach ($smtpAddress in $smtpAddressesToAdd) {
        # Check if the SMTP address already exists to avoid duplicates
        if ($lisaBertingUser.proxyAddresses -notcontains $smtpAddress) {
            Write-Host "Adding Address: $smtpAddress" -ForegroundColor Yellow

            # Add the SMTP address to the user's proxyAddresses
            Set-ADUser -Identity $lisaBertingUser -Add @{proxyAddresses = $smtpAddress}
        } else {
            Write-Host "Address already exists: $smtpAddress" -ForegroundColor Cyan
        }
    }

    Write-Host "Finished adding SMTP addresses for $($lisaBertingUser.Name)." -ForegroundColor Green
} else {
    Write-Host "Lisa Berting user not found in AD" -ForegroundColor Red
}
