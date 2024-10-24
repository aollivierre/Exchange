# Input & Output CSV Files
# $inputCsv = "path_to_input.csv"
$inputCsv = "C:\Code\Imports\2023_07_25_12_26_54_RicohPrinters_ARH_CSV_Input.csv"
# $outputCsv = "path_to_output.csv"
$outputCsv = "C:\Code\Exports\PingTest_ARH_Oct_30_2023_v3.csv"

# Create an array to hold the results
$results = @()

# Counters
$totalSuccess = 0
$totalFailure = 0
$totalTests = 0
$totalIPsInput = 0
$totalIPsOutput = 0

# Read the IP addresses from the CSV file
$ipAddresses = Import-Csv -Path $inputCsv

$totalIPsInput = $ipAddresses.Count

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

    # Update counters
    $totalTests += 2 # For each IP, we test 2 ports: HTTP & HTTPS
    if ($httpResult -eq 'Success') { $totalSuccess++ }
    if ($httpsResult -eq 'Success') { $totalSuccess++ }
    if ($httpResult -eq 'Fail') { $totalFailure++ }
    if ($httpsResult -eq 'Fail') { $totalFailure++ }

    # Output to console
    Write-Host "IP: $($ip.IP)"
    Write-Host "Pingable: $pingable"
    Write-Host "HTTP (Port 80): $httpResult" -ForegroundColor $(if ($httpResult -eq 'Success') { 'Green' } else { 'Red' })
    Write-Host "HTTPS (Port 443): $httpsResult" -ForegroundColor $(if ($httpsResult -eq 'Success') { 'Green' } else { 'Red' })
    Write-Host "-----------------------------------"
}

# Output results to CSV
$results | Export-Csv -Path $outputCsv -NoTypeInformation

# Update total IPs in the output
$totalIPsOutput = $results.Count

# Output aggregated statistics
Write-Host "-----------------------------------"
Write-Host "Total number of tests: $totalTests"
Write-Host "Total number of success: $totalSuccess"
Write-Host "Total number of failures: $totalFailure"
Write-Host "Total number of IPs in the input: $totalIPsInput"
Write-Host "Total number of IPs in the output: $totalIPsOutput"
