# Input & Output CSV Files
# $inputCsv = "path_to_input.csv"
$inputCsv = "C:\Code\Imports\2023_07_25_12_26_54_RicohPrinters_ARH_CSV_Input.csv"
# $outputCsv = "path_to_output.csv"
$outputCsv = "C:\Code\Exports\PingTest_ARH_July_25_2023.csv"

# Create an array to hold the results
$results = @()

# Read the IP addresses from the CSV file
$ipAddresses = Import-Csv -Path $inputCsv

foreach ($ip in $ipAddresses) {
    $pingable = Test-Connection -ComputerName $ip.IP -Count 1 -Quiet
    $httpPort = Test-NetConnection -ComputerName $ip.IP -Port 80 -WarningAction SilentlyContinue
    $httpsPort = Test-NetConnection -ComputerName $ip.IP -Port 443 -WarningAction SilentlyContinue

    # Check results and prepare output message
    $httpResult = if ($httpPort.TcpTestSucceeded) { "Success" } else { "Fail" }
    $httpsResult = if ($httpsPort.TcpTestSucceeded) { "Success" } else { "Fail" }

    $results += [PSCustomObject]@{
        'IP'       = $ip.IP
        'Pingable' = $pingable
        'HTTP'     = $httpResult
        'HTTPS'    = $httpsResult
    }

    # Output to console
    Write-Host "IP: $($ip.IP)"
    Write-Host "Pingable: $pingable"
    Write-Host "HTTP (Port 80): $httpResult" -ForegroundColor $(if ($httpResult -eq 'Success') { 'Green' } else { 'Red' })
    Write-Host "HTTPS (Port 443): $httpsResult" -ForegroundColor $(if ($httpsResult -eq 'Success') { 'Green' } else { 'Red' })
    Write-Host "-----------------------------------"
}

# Output results to CSV
$results | Export-Csv -Path $outputCsv -NoTypeInformation