# Exchange
Scripts for Exchange On-prem, Hybrid and Online



Exchange Migration Script Documentation
Overview
This PowerShell script automates the process of creating and starting a migration batch in an Exchange Server environment. The script accepts a CSV file as input, containing the users to be migrated, and then creates a new migration batch with the defined parameters.

Usage
CSV File: The script begins by reading a CSV file that contains the list of users to be migrated. You need to replace "C:\path\to\file.csv" with the actual path to your CSV file.

Migration Batch Name: You must also define a name for your migration batch. Replace "[June172023]-[Prod]-[7Prod]-[Batch8]-[Phase6]" with the name you prefer for your migration batch.

Source Endpoint, Target Delivery Domain, Notification Emails, and Complete After parameters: You need to set these parameters as per your requirements in the $params hashtable.

Finally, run the script in PowerShell.

Here's an example of the script in action:


# Output the total number of users in the CSV file
$csvFilePath = "C:\path\to\file.csv"
$csvFile = @(Import-Csv -Path $csvFilePath)
$totalUsers = $csvFile.Count
Write-Host "$(Get-Date) - [INFO] Total number of users in the CSV file: $totalUsers" -ForegroundColor Green

# Then, you can proceed with creating the migration batch
$newMigrationBatchName = "[June172023]-[Prod]-[7Prod]-[Batch8]-[Phase6]"

# Define parameters in a hashtable (splatting)
$params = @{
    Name = $newMigrationBatchName
    SourceEndpoint = "Hybrid Migration Endpoint - EWS (Default Web Site)"
    TargetDeliveryDomain = "orgname.mail.onmicrosoft.com"
    CSVData = [System.IO.File]::ReadAllBytes($csvFilePath)
    NotificationEmails = "Youremailaddress@somedomain.onmicrosoft.com"
    CompleteAfter = "2023-06-17 11:59:00 PM"
}

# Create the new migration batch with the defined parameters
New-MigrationBatch @params

# Start the migration batch
Start-MigrationBatch -Identity $newMigrationBatchName
Prerequisites
To run this script, you need to have the Exchange Server module installed and be connected to your Exchange Server environment.

Disclaimer
Ensure to test this script in a controlled environment before using it in a production setting. Always verify the CSV file path and the parameters for the migration batch before running the script.

For more information, refer to the official Exchange Server PowerShell documentation.
