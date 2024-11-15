# Define a list of Exchange-related service names
$exchangeServices = @(
    "MSExchangeADTopology", 
    "MSExchangeIS", 
    "MSExchangeTransport", 
    "MSExchangeUM", 
    "MSExchangeMailboxAssistants",
    "MSExchangeRPC",
    "W3Svc" # For IIS
)

# Function to start a service if it's stopped and display status
function CheckAndStartService {
    param (
        [string]$serviceName
    )

    # Get the service status
    $service = Get-Service -Name $serviceName -ErrorAction SilentlyContinue

    if ($null -eq $service) {
        Write-Host "Service $serviceName not found on this server." -ForegroundColor Yellow
        return
    }

    switch ($service.Status) {
        'Running' {
            Write-Host "Service $serviceName is already running." -ForegroundColor Green
        }
        'Stopped' {
            Write-Host "Service $serviceName is stopped. Attempting to start..." -ForegroundColor Yellow
            Start-Service $serviceName
            Write-Host "Service $serviceName started." -ForegroundColor Green
        }
        Default {
            Write-Host "Service $serviceName is in a non-standard state: $($service.Status)" -ForegroundColor Red
        }
    }
}

# Iterate over each service and check/start them
foreach ($service in $exchangeServices) {
    CheckAndStartService -serviceName $service
}
