# Check if the script is running with administrative privileges
function IsAdmin {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = [Security.Principal.WindowsPrincipal] $identity
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
}

if (-not (IsAdmin)) {
    Write-Host "Script needs to be run as Administrator. Relaunching with admin privileges." -ForegroundColor Yellow
    Start-Process powershell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}


function Initialize-ScriptAndLogging {
    $ErrorActionPreference = 'SilentlyContinue'
    $deploymentName = "CompareExchangeRemoteMailboxes" # Replace this with your actual deployment name
    $scriptPath = "C:\code\$deploymentName"
    # $hadError = $false

    try {
        if (-not (Test-Path -Path $scriptPath)) {
            New-Item -ItemType Directory -Path $scriptPath -Force | Out-Null
            Write-Host "Created directory: $scriptPath"
        }

        $computerName = $env:COMPUTERNAME
        $Filename = "CompareExchangeRemoteMailboxes"
        $logDir = Join-Path -Path $scriptPath -ChildPath "exports\Logs\$computerName"
        $logPath = Join-Path -Path $logDir -ChildPath "$(Get-Date -Format 'yyyy-MM-dd-HH-mm-ss')"
        
        if (!(Test-Path $logPath)) {
            Write-Host "Did not find log file at $logPath" -ForegroundColor Yellow
            Write-Host "Creating log file at $logPath" -ForegroundColor Yellow
            $createdLogDir = New-Item -ItemType Directory -Path $logPath -Force -ErrorAction Stop
            Write-Host "Created log file at $logPath" -ForegroundColor Green
        }
        
        $logFile = Join-Path -Path $logPath -ChildPath "$Filename-Transcript.log"
        Start-Transcript -Path $logFile -ErrorAction Stop | Out-Null

        $CSVDir = Join-Path -Path $scriptPath -ChildPath "exports\CSV"
        $CSVFilePath = Join-Path -Path $CSVDir -ChildPath "$computerName"
        
        if (!(Test-Path $CSVFilePath)) {
            Write-Host "Did not find CSV file at $CSVFilePath" -ForegroundColor Yellow
            Write-Host "Creating CSV file at $CSVFilePath" -ForegroundColor Yellow
            $createdCSVDir = New-Item -ItemType Directory -Path $CSVFilePath -Force -ErrorAction Stop
            Write-Host "Created CSV file at $CSVFilePath" -ForegroundColor Green
        }

        return @{
            ScriptPath  = $scriptPath
            Filename    = $Filename
            LogPath     = $logPath
            LogFile     = $logFile
            CSVFilePath = $CSVFilePath
        }

    }
    catch {
        Write-Error "An error occurred while initializing script and logging: $_"
    }
}
$initializationInfo = Initialize-ScriptAndLogging



# Script Execution and Variable Assignment
# After the function Initialize-ScriptAndLogging is called, its return values (in the form of a hashtable) are stored in the variable $initializationInfo.

# Then, individual elements of this hashtable are extracted into separate variables for ease of use:

# $ScriptPath: The path of the script's main directory.
# $Filename: The base name used for log files.
# $logPath: The full path of the directory where logs are stored.
# $logFile: The full path of the transcript log file.
# $CSVFilePath: The path of the directory where CSV files are stored.
# This structure allows the script to have a clear organization regarding where logs and other files are stored, making it easier to manage and maintain, especially for logging purposes. It also encapsulates the setup logic in a function, making the main script cleaner and more focused on its primary tasks.


$ScriptPath = $initializationInfo['ScriptPath']
$Filename = $initializationInfo['Filename']
$logPath = $initializationInfo['LogPath']
$logFile = $initializationInfo['LogFile']
$CSVFilePath = $initializationInfo['CSVFilePath']




function AppendCSVLog {
    param (
        [string]$Message,
        [string]$CSVFilePath
       
    )

    $csvData = [PSCustomObject]@{
        TimeStamp    = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        ComputerName = $env:COMPUTERNAME
        Message      = $Message
    }

    $csvData | Export-Csv -Path $CSVFilePath -Append -NoTypeInformation -Force
}



function CreateEventSourceAndLog {
    param (
        [string]$LogName,
        [string]$EventSource
    )


    # Validate parameters
    if (-not $LogName) {
        Write-Warning "LogName is required."
        return
    }
    if (-not $EventSource) {
        Write-Warning "Source is required."
        return
    }

    # Function to create event log and source
    function CreateEventLogSource($logName, $EventSource) {
        try {
            if ($PSVersionTable.PSVersion.Major -lt 6) {
                New-EventLog -LogName $logName -Source $EventSource
            }
            else {
                [System.Diagnostics.EventLog]::CreateEventSource($EventSource, $logName)
            }
            Write-Host "Event source '$EventSource' created in log '$logName'" -ForegroundColor Green
        }
        catch {
            Write-Warning "Error creating the event log. Make sure you run PowerShell as an Administrator."
        }
    }

    # Check if the event log exists
    if (-not (Get-WinEvent -ListLog $LogName -ErrorAction SilentlyContinue)) {
        CreateEventLogSource $LogName $EventSource
    }
    # Check if the event source exists
    elseif (-not ([System.Diagnostics.EventLog]::SourceExists($EventSource))) {
        # Unregister the source if it's registered with a different log
        $existingLogName = (Get-WinEvent -ListLog * | Where-Object { $_.LogName -contains $EventSource }).LogName
        if ($existingLogName -ne $LogName) {
            Remove-EventLog -Source $EventSource -ErrorAction SilentlyContinue
        }
        CreateEventLogSource $LogName $EventSource
    }
    else {
        Write-Host "Event source '$EventSource' already exists in log '$LogName'" -ForegroundColor Yellow
    }
}

$LogName = (Get-Date -Format "HHmmss") + "_CompareExchangeRemoteMailboxes"
$EventSource = (Get-Date -Format "HHmmss") + "_CompareExchangeRemoteMailboxes"

# Call the Create-EventSourceAndLog function
CreateEventSourceAndLog -LogName $LogName -EventSource $EventSource

# Call the Write-CustomEventLog function with custom parameters and level
# Write-CustomEventLog -LogName $LogName -EventSource $EventSource -EventMessage "Outlook Signature Restore completed with warnings." -EventID 1001 -Level 'WARNING'




function Write-EventLogMessage {
    param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Message,

        [string]$LogName = 'CompareExchangeRemoteMailboxes',
        [string]$EventSource,

        [int]$EventID = 1000  # Default event ID
    )

    $ErrorActionPreference = 'SilentlyContinue'
    $hadError = $false

    try {
        if (-not $EventSource) {
            throw "EventSource is required."
        }

        if ($PSVersionTable.PSVersion.Major -lt 6) {
            # PowerShell version is less than 6, use Write-EventLog
            Write-EventLog -LogName $logName -Source $EventSource -EntryType Information -EventId $EventID -Message $Message
        }
        else {
            # PowerShell version is 6 or greater, use System.Diagnostics.EventLog
            $eventLog = New-Object System.Diagnostics.EventLog($logName)
            $eventLog.Source = $EventSource
            $eventLog.WriteEntry($Message, [System.Diagnostics.EventLogEntryType]::Information, $EventID)
        }

        # Write-host "Event log entry created: $Message" 
    }
    catch {
        Write-host "Error creating event log entry: $_" 
        $hadError = $true
    }

    if (-not $hadError) {
        # Write-host "Event log message writing completed successfully."
    }
}




function Write-EnhancedLog {
    param (
        [string]$Message,
        [string]$Level = 'INFO',
        [ConsoleColor]$ForegroundColor = [ConsoleColor]::White,
        [string]$CSVFilePath = "$scriptPath\exports\CSV\$(Get-Date -Format 'yyyy-MM-dd')-Log.csv",
        [string]$CentralCSVFilePath = "$scriptPath\exports\CSV\$Filename.csv",
        [switch]$UseModule = $false,
        [string]$Caller = (Get-PSCallStack)[0].Command
    )

    # Add timestamp, computer name, and log level to the message
    $formattedMessage = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') $($env:COMPUTERNAME): [$Level] [$Caller] $Message"

    # Set foreground color based on log level
    switch ($Level) {
        'INFO' { $ForegroundColor = [ConsoleColor]::Green }
        'WARNING' { $ForegroundColor = [ConsoleColor]::Yellow }
        'ERROR' { $ForegroundColor = [ConsoleColor]::Red }
    }

    # Write the message with the specified colors
    $currentForegroundColor = $Host.UI.RawUI.ForegroundColor
    $Host.UI.RawUI.ForegroundColor = $ForegroundColor
    # Write-output $formattedMessage
    Write-host $formattedMessage
    $Host.UI.RawUI.ForegroundColor = $currentForegroundColor

    # Append to CSV file
    AppendCSVLog -Message $formattedMessage -CSVFilePath $CSVFilePath
    AppendCSVLog -Message $formattedMessage -CSVFilePath $CentralCSVFilePath

    # Write to event log (optional)
    # Write-CustomEventLog -EventMessage $formattedMessage -Level $Level


    # Write-CustomEventLog -LogName $LogName -EventSource $EventSource -EventMessage $formattedMessage -EventID 1001 -Level $Level
    # Write-EventLogMessage -LogName $LogName -EventSource $EventSource -EventMessage $formattedMessage -EventID 1001 -Level $Level

    # Adjust this line in your script where you call the function
    Write-EventLogMessage -LogName $LogName -EventSource $EventSource -Message $formattedMessage -EventID 1001

}

function Export-EventLog {
    param (
        [Parameter(Mandatory = $true)]
        [string]$LogName,
        [Parameter(Mandatory = $true)]
        [string]$ExportPath
    )

    try {
        wevtutil epl $LogName $ExportPath

        if (Test-Path $ExportPath) {
            Write-EnhancedLog -Message "Event log '$LogName' exported to '$ExportPath'" -Level "INFO" -ForegroundColor ([ConsoleColor]::Green)
        }
        else {
            Write-EnhancedLog -Message "Event log '$LogName' not exported: File does not exist at '$ExportPath'" -Level "WARNING" -ForegroundColor ([ConsoleColor]::Yellow)
        }
    }
    catch {
        Write-EnhancedLog -Message "Error exporting event log '$LogName': $($_.Exception.Message)" -Level "ERROR" -ForegroundColor ([ConsoleColor]::Red)
    }
}

# # Example usage
# $LogName = 'CompareExchangeRemoteMailboxesLog'
# # $ExportPath = 'Path\to\your\exported\eventlog.evtx'
# $ExportPath = "C:\code\CompareExchangeRemoteMailboxes\exports\Logs\$logname.evtx"
# Export-EventLog -LogName $LogName -ExportPath $ExportPath






#################################################################################################################################
################################################# END LOGGING ###################################################################
#################################################################################################################################



Write-EnhancedLog -Message "Logging works" -Level "INFO" -ForegroundColor ([ConsoleColor]::Green)


#################################################################################################################################
################################################# END LOGGING ###################################################################
#################################################################################################################################

function Install-RequiredModules {

    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

    # $requiredModules = @("Microsoft.Graph", "Microsoft.Graph.Authentication")
    $requiredModules = @("ExchangeOnlineManagement")

    foreach ($module in $requiredModules) {
        if (!(Get-Module -ListAvailable -Name $module)) {

            Write-EnhancedLog -Message "Installing module: $module" -Level "INFO" -ForegroundColor ([ConsoleColor]::Cyan)
            Install-Module -Name $module -Force
            Write-EnhancedLog -Message "Module: $module has been installed" -Level "INFO" -ForegroundColor ([ConsoleColor]::Cyan)
        }
        else {
            Write-EnhancedLog -Message "Module $module is already installed" -Level "INFO" -ForegroundColor ([ConsoleColor]::Cyan)
        }
    }


    $ImportedModules = @("ExchangeOnlineManagement")
    
    foreach ($Importedmodule in $ImportedModules) {
        if ((Get-Module -ListAvailable -Name $Importedmodule)) {
            Write-EnhancedLog -Message "Importing module: $Importedmodule" -Level "INFO" -ForegroundColor ([ConsoleColor]::Cyan)
            Import-Module -Name $Importedmodule
            Write-EnhancedLog -Message "Module: $Importedmodule has been Imported" -Level "INFO" -ForegroundColor ([ConsoleColor]::Cyan)
        }
    }


}
# Call the function to install the required modules and dependencies
# Install-RequiredModules
Write-EnhancedLog -Message "All modules installed" -Level "INFO" -ForegroundColor ([ConsoleColor]::Green)

#################################################################################################################################
################################################# END Importing Modules ######################################################## 
#################################################################################################################################



#################################################################################################################################
################################################# END Connecting to MG Graph ################################################## 
#################################################################################################################################

function ARHCompareMailboxes {
    # Logging the start of the function
    Write-EnhancedLog -Message "Starting ARHCompareMailboxes function." -Level "INFO" -ForegroundColor ([ConsoleColor]::Green)

    # Retrieving all remote mailboxes
    $remoteMailboxes = Get-RemoteMailbox
    $remoteMailboxCount = $remoteMailboxes.Count
    Write-EnhancedLog -Message "Retrieved a total of $remoteMailboxCount remote mailboxes." -Level "INFO" -ForegroundColor ([ConsoleColor]::Green)

    # Connecting to Exchange Online and retrieving all mailboxes
    # Assuming the session for Exchange Online is already established
    $onlineMailboxes = Get-EXOMailbox
    $onlineMailboxCount = $onlineMailboxes.Count
    Write-EnhancedLog -Message "Retrieved a total of $onlineMailboxCount Exchange Online mailboxes." -Level "INFO" -ForegroundColor ([ConsoleColor]::Green)

    # Comparing the lists to find out which mailboxes are not in Exchange Online
    $notInOnline = $remoteMailboxes | Where-Object { $_.PrimarySmtpAddress -notin $onlineMailboxes.PrimarySmtpAddress }
    $notInOnlineCount = $notInOnline.Count
    Write-EnhancedLog -Message "Identified $notInOnlineCount mailboxes that are in Remote but not in Exchange Online." -Level "INFO" -ForegroundColor ([ConsoleColor]::Green)

    # Return the results
    return $notInOnline

    # Logging the end of the function
    Write-EnhancedLog -Message "Completed ARHCompareMailboxes function." -Level "INFO" -ForegroundColor ([ConsoleColor]::Green)
}

# Example usage of the function
$missingMailboxes = ARHCompareMailboxes

$missingMailboxes | Out-GridView -Title "Mailboxes in Remote but not in Exchange Online"
$missingMailboxes | export-csv -path $CSVFilePath\missingMailboxes.csv -NoTypeInformation -Force


#################################################################################################################################
################################################# END OF MAIN FUNCTION ####################################################### 
#################################################################################################################################



#################################################################################################################################
################################################# Stopping Transcription ####################################################### 
#################################################################################################################################


# Stop transcript logging
Stop-Transcript


# Example usage
$EvenlogExportPath = Join-Path -Path $logPath -ChildPath "$LogName-Transcript.evtx"
Export-EventLog -LogName $LogName -ExportPath $EvenlogExportPath