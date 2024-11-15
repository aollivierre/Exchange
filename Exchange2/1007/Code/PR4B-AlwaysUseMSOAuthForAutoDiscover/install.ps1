# Read configuration from the JSON file
# Assign values from JSON to variables


<#
.SYNOPSIS
Dot-sources all PowerShell scripts in the 'private' folder relative to the script root.

.DESCRIPTION
This function finds all PowerShell (.ps1) scripts in a 'private' folder located in the script root directory and dot-sources them. It logs the process, including any errors encountered, with optional color coding.

.EXAMPLE
Dot-SourcePrivateScripts

Dot-sources all scripts in the 'private' folder and logs the process.

.NOTES
Ensure the Write-EnhancedLog function is defined before using this function for logging purposes.
#>

# Read configuration from the JSON file
$configPath = Join-Path -Path $PSScriptRoot -ChildPath "config.json"
$env:MYMODULE_CONFIG_PATH = $configPath

$config = Get-Content -Path $configPath -Raw | ConvertFrom-Json

# Assign values from JSON to variables
$PackageName = $config.PackageName
$PackageUniqueGUID = $config.PackageUniqueGUID
$Version = $config.Version
$PackageExecutionContext = $config.PackageExecutionContext
$RepetitionInterval = $config.RepetitionInterval
$ScriptMode = $config.ScriptMode




<#
.SYNOPSIS
Dot-sources all PowerShell scripts in the 'private' folder relative to the script root.

.DESCRIPTION
This function finds all PowerShell (.ps1) scripts in a 'private' folder located in the script root directory and dot-sources them. It logs the process, including any errors encountered, with optional color coding.

.EXAMPLE
Dot-SourcePrivateScripts

Dot-sources all scripts in the 'private' folder and logs the process.

.NOTES
Ensure the Write-EnhancedLog function is defined before using this function for logging purposes.
#>

function Get-PrivateScriptPathsAndVariables {
    param (
        [string]$BaseDirectory
    )

    try {
        $privateFolderPath = Join-Path -Path $BaseDirectory -ChildPath "private"
        
        if (-not (Test-Path -Path $privateFolderPath)) {
            throw "Private folder path does not exist: $privateFolderPath"
        }

        # Construct and return a PSCustomObject
        return [PSCustomObject]@{
            BaseDirectory     = $BaseDirectory
            PrivateFolderPath = $privateFolderPath
        }
    }
    catch {
        Write-Host "Error in finding private script files: $_" -ForegroundColor Red
        # Optionally, you could return a PSCustomObject indicating an error state
        # return [PSCustomObject]@{ Error = $_.Exception.Message }
    }
}



# Retrieve script paths and related variables
$DotSourcinginitializationInfo = Get-PrivateScriptPathsAndVariables -BaseDirectory $PSScriptRoot

# $DotSourcinginitializationInfo
$DotSourcinginitializationInfo | Format-List


function Import-ModuleWithRetry {

    <#
.SYNOPSIS
Imports a PowerShell module with retries on failure.

.DESCRIPTION
This function attempts to import a specified PowerShell module, retrying the import process up to a specified number of times upon failure. It waits for a specified delay between retries. The function uses advanced logging to provide detailed feedback about the import process.

.PARAMETER ModulePath
The path to the PowerShell module file (.psm1) that should be imported.

.PARAMETER MaxRetries
The maximum number of retries to attempt if importing the module fails. Default is 30.

.PARAMETER WaitTimeSeconds
The number of seconds to wait between retry attempts. Default is 2 seconds.

.EXAMPLE
$modulePath = "C:\Modules\MyPowerShellModule.psm1"
Import-ModuleWithRetry -ModulePath $modulePath

Tries to import the module located at "C:\Modules\MyPowerShellModule.psm1", with up to 30 retries, waiting 2 seconds between each retry.

.NOTES
This function requires the `Write-EnhancedLog` function to be defined in the script for logging purposes.

.LINK
Write-EnhancedLog

#>

    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$ModulePath,

        [int]$MaxRetries = 30,

        [int]$WaitTimeSeconds = 2
    )

    Begin {
        $retryCount = 0
        $isModuleLoaded = $false
        # Write-EnhancedLog -Message "Starting to import module from path: $ModulePath" -Level "INFO"
        Write-Host "Starting to import module from path: $ModulePath"
    }

    Process {
        while (-not $isModuleLoaded -and $retryCount -lt $MaxRetries) {
            try {
                Import-Module $ModulePath -ErrorAction Stop
                $isModuleLoaded = $true
                Write-EnhancedLog -Message "Module: $ModulePath imported successfully." -Level "INFO"
            }
            catch {
                # Write-EnhancedLog -Message "Attempt $retryCount to load module failed. Waiting $WaitTimeSeconds seconds before retrying." -Level "WARNING"
                Write-Host "Attempt $retryCount to load module failed. Waiting $WaitTimeSeconds seconds before retrying."
                Start-Sleep -Seconds $WaitTimeSeconds
            }
            finally {
                $retryCount++
            }

            if ($retryCount -eq $MaxRetries -and -not $isModuleLoaded) {
                # Write-EnhancedLog -Message "Failed to import module after $MaxRetries retries." -Level "ERROR"
                Write-Host "Failed to import module after $MaxRetries retries."
                break
            }
        }
    }

    End {
        if ($isModuleLoaded) {
            Write-EnhancedLog -Message "Module: $ModulePath loaded successfully." -Level "INFO"
        }
        else {
            # Write-EnhancedLog -Message "Failed to load module $ModulePath within the maximum retry limit." -Level "CRITICAL"
            Write-Host "Failed to load module $ModulePath within the maximum retry limit."
        }
    }
}

# Example of how to use the function
# $PSScriptRoot = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
$modulePath = Join-Path -Path $PSScriptRoot -ChildPath "Private\EnhancedLoggingAO\2.0.0\EnhancedLoggingAO.psm1"

# Call the function to import the module with retry logic
Import-ModuleWithRetry -ModulePath $modulePath




# ################################################################################################################################
# ################################################ END MODULE LOADING ############################################################
# ################################################################################################################################



function Ensure-LoggingFunctionExists {
    if (Get-Command Write-EnhancedLog -ErrorAction SilentlyContinue) {
        Write-EnhancedLog -Message "Logging works" -Level "INFO" -ForegroundColor ([ConsoleColor]::Green)
    }
    else {
        throw "Write-EnhancedLog function not found. Terminating script."
    }
}

# Usage
try {
    Ensure-LoggingFunctionExists
    # Continue with the rest of the script here
    # exit
}
catch {
    Write-Host "Critical error: $_" -ForegroundColor Red
    exit
}


# ################################################################################################################################
# ################################################ END MODULE CHECKING ###########################################################
# ################################################################################################################################




function Test-RunningAsSystem {
    $systemSid = New-Object System.Security.Principal.SecurityIdentifier "S-1-5-18"
    $currentSid = [System.Security.Principal.WindowsIdentity]::GetCurrent().User

    return $currentSid -eq $systemSid
}


function CheckAndElevate {

    <#
.SYNOPSIS
Elevates the script to run with administrative privileges if not already running as an administrator.

.DESCRIPTION
The CheckAndElevate function checks if the current PowerShell session is running with administrative privileges. If it is not, the function attempts to restart the script with elevated privileges using the 'RunAs' verb. This is useful for scripts that require administrative privileges to perform their tasks.

.EXAMPLE
CheckAndElevate

Checks the current session for administrative privileges and elevates if necessary.

.NOTES
This function will cause the script to exit and restart if it is not already running with administrative privileges. Ensure that any state or data required after elevation is managed appropriately.
#>
    [CmdletBinding()]
    param (
        # Advanced parameters could be added here if needed. For this function, parameters aren't strictly necessary,
        # but you could, for example, add parameters to control logging behavior or to specify a different method of elevation.
        # [switch]$Elevated
    )

    begin {
        try {
            $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
            $isAdmin = $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

            Write-EnhancedLog -Message "Checking for administrative privileges..." -Level "INFO" -ForegroundColor ([ConsoleColor]::Blue)
        }
        catch {
            Write-EnhancedLog -Message "Error determining administrative status: $_" -Level "ERROR" -ForegroundColor ([ConsoleColor]::Red)
            throw $_
        }
    }

    process {
        if (-not $isAdmin) {
            try {
                Write-EnhancedLog -Message "The script is not running with administrative privileges. Attempting to elevate..." -Level "WARNING" -ForegroundColor ([ConsoleColor]::Yellow)
                
                $arguments = "-NoProfile -ExecutionPolicy Bypass -NoExit -File `"$PSCommandPath`" $args"
                Start-Process PowerShell -Verb RunAs -ArgumentList $arguments

                # Invoke-AsSystem -PsExec64Path $PsExec64Path
                
                Write-EnhancedLog -Message "Script re-launched with administrative privileges. Exiting current session." -Level "INFO" -ForegroundColor ([ConsoleColor]::Green)
                exit
            }
            catch {
                Write-EnhancedLog -Message "Failed to elevate privileges: $_" -Level "ERROR" -ForegroundColor ([ConsoleColor]::Red)
                throw $_
            }
        }
        else {
            Write-EnhancedLog -Message "Script is already running with administrative privileges." -Level "INFO" -ForegroundColor ([ConsoleColor]::Green)
        }
    }

    end {
        # This block is typically used for cleanup. In this case, there's nothing to clean up,
        # but it's useful to know about this structure for more complex functions.
    }
}



function Remove-ExistingPsExec {
    [CmdletBinding()]
    param(
        [string]$TargetFolder = "$PSScriptRoot\private"
    )

    # Full path for PsExec64.exe
    $PsExec64Path = Join-Path -Path $TargetFolder -ChildPath "PsExec64.exe"

    try {
        # Check if PsExec64.exe exists
        if (Test-Path -Path $PsExec64Path) {
            Write-EnhancedLog -Message "Removing existing PsExec64.exe from: $TargetFolder"
            # Remove PsExec64.exe
            Remove-Item -Path $PsExec64Path -Force
            Write-Output "PsExec64.exe has been removed from: $TargetFolder"
        }
        else {
            Write-EnhancedLog -Message "No PsExec64.exe file found in: $TargetFolder"
        }
    }
    catch {
        # Handle any errors during the removal
        Write-Error "An error occurred while trying to remove PsExec64.exe: $_"
    }
}






function Download-PsExec {
    [CmdletBinding()]
    param(
        [string]$TargetFolder = "$PSScriptRoot\private"
    )

    Begin {

        Remove-ExistingPsExec
    }



    process {

        # Define the URL for PsExec download
        $url = "https://download.sysinternals.com/files/PSTools.zip"
    
        # Ensure the target folder exists
        if (-Not (Test-Path -Path $TargetFolder)) {
            New-Item -Path $TargetFolder -ItemType Directory
        }
  
        # Full path for the downloaded file
        $zipPath = Join-Path -Path $TargetFolder -ChildPath "PSTools.zip"
  
        try {
            # Download the PSTools.zip file containing PsExec
            Write-EnhancedLog -Message "Downloading PSTools.zip from: $url to: $zipPath"
            Invoke-WebRequest -Uri $url -OutFile $zipPath
  
            # Extract PsExec64.exe from the zip file
            Expand-Archive -Path $zipPath -DestinationPath "$TargetFolder\PStools" -Force
  
            # Specific extraction of PsExec64.exe
            $extractedFolderPath = Join-Path -Path $TargetFolder -ChildPath "PSTools"
            $PsExec64Path = Join-Path -Path $extractedFolderPath -ChildPath "PsExec64.exe"
            $finalPath = Join-Path -Path $TargetFolder -ChildPath "PsExec64.exe"
  
            # Move PsExec64.exe to the desired location
            if (Test-Path -Path $PsExec64Path) {
  
                Write-EnhancedLog -Message "Moving PSExec64.exe from: $PsExec64Path to: $finalPath"
                Move-Item -Path $PsExec64Path -Destination $finalPath
  
                # Remove the downloaded zip file and extracted folder
                Remove-Item -Path $zipPath -Force
                Remove-Item -Path $extractedFolderPath -Recurse -Force
  
                Write-EnhancedLog -Message "PsExec64.exe has been successfully downloaded and moved to: $finalPath"
            }
        }
        catch {
            # Handle any errors during the process
            Write-Error "An error occurred: $_"
        }
    }


  

}




function Invoke-AsSystem {
    <#
.SYNOPSIS
Executes a PowerShell script under the SYSTEM context, similar to Intune's execution context.

.DESCRIPTION
The Invoke-AsSystem function executes a PowerShell script using PsExec64.exe to run under the SYSTEM context. This method is useful for scenarios requiring elevated privileges beyond the current user's capabilities.

.PARAMETER PsExec64Path
Specifies the full path to PsExec64.exe. If not provided, it assumes PsExec64.exe is in the same directory as the script.

.EXAMPLE
Invoke-AsSystem -PsExec64Path "C:\Tools\PsExec64.exe"

Executes PowerShell as SYSTEM using PsExec64.exe located at "C:\Tools\PsExec64.exe".

.NOTES
Ensure PsExec64.exe is available and the script has the necessary permissions to execute it.

.LINK
https://docs.microsoft.com/en-us/sysinternals/downloads/psexec
#>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$PsExec64Path,
        [string]$ScriptPathAsSYSTEM  # Path to the PowerShell script you want to run as SYSTEM
    )

    begin {
        CheckAndElevate
        # Define the arguments for PsExec64.exe to run PowerShell as SYSTEM with the script
        $argList = "-accepteula -i -s -d powershell.exe -NoExit -ExecutionPolicy Bypass -File `"$ScriptPathAsSYSTEM`""
        Write-EnhancedLog -Message "Preparing to execute PowerShell as SYSTEM using PsExec64 with the script: $ScriptPathAsSYSTEM" -Level "INFO" -ForegroundColor ([ConsoleColor]::Green)

        Download-PsExec
    }

    process {
        try {
            # Ensure PsExec64Path exists
            if (-not (Test-Path -Path $PsExec64Path)) {
                $errorMessage = "PsExec64.exe not found at path: $PsExec64Path"
                Write-EnhancedLog -Message $errorMessage -Level "ERROR" -ForegroundColor ([ConsoleColor]::Red)
                throw $errorMessage
            }

            # Run PsExec64.exe with the defined arguments to execute the script as SYSTEM
            $executingMessage = "Executing PsExec64.exe to start PowerShell as SYSTEM running script: $ScriptPathAsSYSTEM"
            Write-EnhancedLog -Message $executingMessage -Level "INFO" -ForegroundColor ([ConsoleColor]::Green)
            Start-Process -FilePath "$PsExec64Path" -ArgumentList $argList -Wait -NoNewWindow
            
            Write-EnhancedLog -Message "SYSTEM session started. Closing elevated session..." -Level "INFO" -ForegroundColor ([ConsoleColor]::Green)
            exit

        }
        catch {
            Write-EnhancedLog -Message "An error occurred: $_" -Level "ERROR" -ForegroundColor ([ConsoleColor]::Red)
        }
    }
}




# Assuming Invoke-AsSystem and Write-EnhancedLog are already defined
# Update the path to your actual location of PsExec64.exe

# Write-EnhancedLog -Message "calling Test-RunningAsSystem" -Level "INFO" -ForegroundColor ([ConsoleColor]::Green)
# if (-not (Test-RunningAsSystem)) {
#     $privateFolderPath = Join-Path -Path $PSScriptRoot -ChildPath "private"
#     $PsExec64Path = Join-Path -Path $privateFolderPath -ChildPath "PsExec64.exe"

#     Write-EnhancedLog -Message "Current session is not running as SYSTEM. Attempting to invoke as SYSTEM..." -Level "INFO" -ForegroundColor ([ConsoleColor]::Yellow)

#     $ScriptToRunAsSystem = $MyInvocation.MyCommand.Path
#     Invoke-AsSystem -PsExec64Path $PsExec64Path -ScriptPath $ScriptToRunAsSystem

# }
# else {
#     Write-EnhancedLog -Message "Session is already running as SYSTEM." -Level "INFO" -ForegroundColor ([ConsoleColor]::Green)
# }


    
    
#################################################################################################################################
################################################# END LOGGING ###################################################################
#################################################################################################################################





# ## Download the latest release of PSAppDeployToolkit from GitHub
# $githubRepository = "psappdeploytoolkit/psappdeploytoolkit"
# $filenamePatternMatch = "PSAppDeployToolkit*.zip"
# $psadtReleaseUri = "https://api.github.com/repos/$githubRepository/releases/latest"
# $psadtDownloadUri = ((Invoke-RestMethod -Method GET -Uri $psadtReleaseUri).assets | Where-Object name -Like $filenamePatternMatch ).browser_download_url
# # $zipExtractionPath = Join-Path $env:USERPROFILE "Downloads" "PSAppDeployToolkit"
# $zipExtractionPath = Join-Path "$PSScriptRoot\private" "PSAppDeployToolkit"
# $zipTempDownloadPath = Join-Path -Path $([System.IO.Path]::GetTempPath()) -ChildPath $(Split-Path -Path $psadtDownloadUri -Leaf)

# ## Download to a temporary folder
# Invoke-WebRequest -Uri $psadtDownloadUri -Out $zipTempDownloadPath
# ## Remove any Zone.Identifier alternate data streams to unblock the file (if required)
# Unblock-File -Path $zipTempDownloadPath
# ## Extract the contents of the zip file to a subfolder under the current user's Downloads folder
# Expand-Archive -Path $zipTempDownloadPath -DestinationPath $zipExtractionPath -Force

# Write-Host ("File: {0} extracted to Path: {1}" -f $psadtDownloadUri, $zipExtractionPath) -ForegroundColor Yellow



function Download-PSAppDeployToolkit {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$GithubRepository,

        [Parameter(Mandatory)]
        [string]$FilenamePatternMatch,

        [Parameter(Mandatory)]
        [string]$ZipExtractionPath
    )

    begin {
        try {
            # Log the beginning of the function
            Write-EnhancedLog -Message "Starting Download-PSAppDeployToolkit function." -Level "INFO"

            # Set the URI to get the latest release information from the GitHub repository
            $psadtReleaseUri = "https://api.github.com/repos/$GithubRepository/releases/latest"
            Write-EnhancedLog -Message "GitHub release URI: $psadtReleaseUri" -Level "INFO"
        }
        catch {
            Write-EnhancedLog -Message "Error in begin block: $_" -Level "ERROR"
            throw $_
        }
    }

    process {
        try {
            # Get the download URL for the matching filename pattern
            Write-EnhancedLog -Message "Fetching the latest release information from GitHub." -Level "INFO"
            $psadtDownloadUri = (Invoke-RestMethod -Method GET -Uri $psadtReleaseUri).assets | Where-Object { $_.name -like $FilenamePatternMatch } | Select-Object -ExpandProperty browser_download_url
            
            if (-not $psadtDownloadUri) {
                throw "No matching file found for pattern: $FilenamePatternMatch"
            }
            Write-EnhancedLog -Message "Found matching download URL: $psadtDownloadUri" -Level "INFO"
            
            # Set the path for the temporary download location
            $zipTempDownloadPath = Join-Path -Path $([System.IO.Path]::GetTempPath()) -ChildPath (Split-Path -Path $psadtDownloadUri -Leaf)
            Write-EnhancedLog -Message "Temporary download path: $zipTempDownloadPath" -Level "INFO"

            # Download the file to the temporary location
            Write-EnhancedLog -Message "Downloading file from $psadtDownloadUri to $zipTempDownloadPath" -Level "INFO"
            Invoke-WebRequest -Uri $psadtDownloadUri -OutFile $zipTempDownloadPath

            # Unblock the downloaded file if necessary
            Write-EnhancedLog -Message "Unblocking file at $zipTempDownloadPath" -Level "INFO"
            Unblock-File -Path $zipTempDownloadPath

            # Extract the contents of the zip file to the specified extraction path
            Write-EnhancedLog -Message "Extracting file from $zipTempDownloadPath to $ZipExtractionPath" -Level "INFO"
            Expand-Archive -Path $zipTempDownloadPath -DestinationPath $ZipExtractionPath -Force
        }
        catch {
            Write-EnhancedLog -Message "Error in process block: $_" -Level "ERROR"
            throw $_
        }
    }

    end {
        try {
            Write-Host ("File: {0} extracted to Path: {1}" -f $psadtDownloadUri, $ZipExtractionPath) -ForegroundColor Yellow
            Write-EnhancedLog -Message "File extracted successfully to $ZipExtractionPath" -Level "INFO"
        }
        catch {
            Write-EnhancedLog -Message "Error in end block: $_" -Level "ERROR"
            throw $_
        }
    }
}

# Define the variables to be used for the function
$PSADTdownloadParams = @{
    GithubRepository     = "psappdeploytoolkit/psappdeploytoolkit"
    FilenamePatternMatch = "PSAppDeployToolkit*.zip"
    ZipExtractionPath    = Join-Path "$PSScriptRoot\private" "PSAppDeployToolkit"
}

# Call the function with the variables
# Download-PSAppDeployToolkit @PSADTdownloadParams



#################################################################################################################################
################################################# END DOWNLOADING PSADT #########################################################
#################################################################################################################################




###########################################################################################################################
###########################################################################################################################
###########################################################################################################################
###########################################################################################################################
###########################################################################################################################
###########################################################################################################################
#############################################STARTING THE MAIN FUNCTION LOGIC HERE#########################################
###########################################################################################################################
###########################################################################################################################
###########################################################################################################################



Write-EnhancedLog -Message "calling Test-RunningAsSystem" -Level "INFO" -ForegroundColor ([ConsoleColor]::Green)


function Set-LocalPathBasedOnContext {
    Write-EnhancedLog -Message "Checking running context..." -Level "INFO" -ForegroundColor ([System.ConsoleColor]::Cyan)
    if (Test-RunningAsSystem) {
        Write-EnhancedLog -Message "Running as system, setting path to Program Files" -Level "INFO" -ForegroundColor ([System.ConsoleColor]::Yellow)
        # return "$ENV:Programfiles\_MEM"
        return "C:\_MEM"
    }
    else {
        Write-EnhancedLog -Message "Running as user, setting path to Local AppData" -Level "INFO" -ForegroundColor ([System.ConsoleColor]::Yellow)
        return "$ENV:LOCALAPPDATA\_MEM"
    }
}


$global:Path_local = Set-LocalPathBasedOnContext



Write-EnhancedLog -Message "calling Initialize-ScriptVariables" -Level "INFO" -ForegroundColor ([ConsoleColor]::Green)












<#
.SYNOPSIS
Initializes global script variables and defines the path for storing related files.

.DESCRIPTION
This function initializes global script variables such as PackageName, PackageUniqueGUID, Version, and ScriptMode. Additionally, it constructs the path where related files will be stored based on the provided parameters.

.PARAMETER PackageName
The name of the package being processed.

.PARAMETER PackageUniqueGUID
The unique identifier for the package being processed.

.PARAMETER Version
The version of the package being processed.

.PARAMETER ScriptMode
The mode in which the script is being executed (e.g., "Remediation", "PackageName").

.EXAMPLE
Initialize-ScriptVariables -PackageName "MyPackage" -PackageUniqueGUID "1234-5678" -Version 1 -ScriptMode "Remediation"

This example initializes the script variables with the specified values.

#>
function Initialize-ScriptVariables {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$PackageName,

        [Parameter(Mandatory = $true)]
        [string]$PackageUniqueGUID,

        [Parameter(Mandatory = $true)]
        [int]$Version,

        [Parameter(Mandatory = $true)]
        [string]$ScriptMode
    )

    # Assuming Set-LocalPathBasedOnContext and Test-RunningAsSystem are defined elsewhere
    # $global:Path_local = Set-LocalPathBasedOnContext

    # Default logic for $Path_local if not set by Set-LocalPathBasedOnContext
    if (-not $Path_local) {
        if (Test-RunningAsSystem) {
            # $Path_local = "$ENV:ProgramFiles\_MEM"
            $Path_local = "c:\_MEM"
        }
        else {
            $Path_local = "$ENV:LOCALAPPDATA\_MEM"
        }
    }

    $Path_PR = "$Path_local\Data\$PackageName-$PackageUniqueGUID"
    $schtaskName = "$PackageName - $PackageUniqueGUID"
    $schtaskDescription = "Version $Version"

    try {
        # Assuming Write-EnhancedLog is defined elsewhere
        Write-EnhancedLog -Message "Initializing script variables..." -Level "INFO" -ForegroundColor ([System.ConsoleColor]::Green)

        # Returning a hashtable of all the important variables
        return @{
            PackageName        = $PackageName
            PackageUniqueGUID  = $PackageUniqueGUID
            Version            = $Version
            ScriptMode         = $ScriptMode
            Path_local         = $Path_local
            Path_PR            = $Path_PR
            schtaskName        = $schtaskName
            schtaskDescription = $schtaskDescription
        }
    }
    catch {
        Write-Error "An error occurred while initializing script variables: $_"
    }
}

# Invocation of the function and storing returned hashtable in a variable
# $initializationInfo = Initialize-ScriptVariables -PackageName "YourPackageName" -PackageUniqueGUID "YourGUID" -Version 1 -ScriptMode "YourMode"

# Call Initialize-ScriptVariables with splatting
$InitializeScriptVariablesParams = @{
    PackageName       = $PackageName
    PackageUniqueGUID = $PackageUniqueGUID
    Version           = $Version
    ScriptMode        = $ScriptMode
}

$initializationInfo = Initialize-ScriptVariables @InitializeScriptVariablesParams

$initializationInfo

$global:PackageName = $initializationInfo['PackageName']
$global:PackageUniqueGUID = $initializationInfo['PackageUniqueGUID']
$global:Version = $initializationInfo['Version']
$global:ScriptMode = $initializationInfo['ScriptMode']
$global:Path_local = $initializationInfo['Path_local']
$global:Path_PR = $initializationInfo['Path_PR']
$global:schtaskName = $initializationInfo['schtaskName']
$global:schtaskDescription = $initializationInfo['schtaskDescription']














<#
.SYNOPSIS
Ensures that all necessary script paths exist, creating them if they do not.

.DESCRIPTION
This function checks for the existence of essential script paths and creates them if they are not found. It is designed to be called after initializing script variables to ensure the environment is correctly prepared for the script's operations.

.PARAMETER Path_local
The local path where the script's data will be stored. This path varies based on the execution context (system vs. user).

.PARAMETER Path_PR
The specific path for storing package-related files, constructed based on the package name and unique GUID.

.EXAMPLE
Ensure-ScriptPathsExist -Path_local $global:Path_local -Path_PR $global:Path_PR

This example ensures that the paths stored in the global variables $Path_local and $Path_PR exist, creating them if necessary.
#>
function Ensure-ScriptPathsExist {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$Path_local,

        [Parameter(Mandatory = $true)]
        [string]$Path_PR
    )

    try {
        # Ensure Path_local exists
        if (-not (Test-Path -Path $Path_local)) {
            New-Item -Path $Path_local -ItemType Directory -Force | Out-Null
            Write-EnhancedLog -Message "Created directory: $Path_local" -Level "INFO" -ForegroundColor ([System.ConsoleColor]::Green)
        }

        # Ensure Path_PR exists
        if (-not (Test-Path -Path $Path_PR)) {
            New-Item -Path $Path_PR -ItemType Directory -Force | Out-Null
            Write-EnhancedLog -Message "Created directory: $Path_PR" -Level "INFO" -ForegroundColor ([System.ConsoleColor]::Green)
        }
    }
    catch {
        Write-EnhancedLog -Message "An error occurred while ensuring script paths exist: $_" -Level "ERROR" -ForegroundColor ([System.ConsoleColor]::Red)
    }
}



# Assuming $global:Path_local and $global:Path_PR are set from previous initialization
Ensure-ScriptPathsExist -Path_local $global:Path_local -Path_PR $global:Path_PR

if (-not (Test-Path -Path $global:Path_PR -PathType Container)) {
    Write-EnhancedLog -Message "Failed to create $global:Path_PR. Please check permissions and path validity." -Level "ERROR" -ForegroundColor ([System.ConsoleColor]::Red)
}
else {
    Write-EnhancedLog -Message "$global:Path_PR exists." -Level "INFO" -ForegroundColor ([System.ConsoleColor]::Green)
}




<#
.SYNOPSIS
Copies all files and folders in the same directory as the script to the specified destination path.

.DESCRIPTION
This function copies all files and folders located in the same directory as the script to the specified destination path. It can be used to bundle necessary files and folders with the script for distribution or deployment.

.PARAMETER DestinationPath
The destination path where the files and folders will be copied.

.EXAMPLE
Copy-FilesToPath -DestinationPath "C:\Temp"

This example copies all files and folders in the same directory as the script to the "C:\Temp" directory.
#>
function Copy-FilesToPath {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$DestinationPath
    )

    begin {
        Write-EnhancedLog -Message "Starting the copy process from the Script Root $PSScriptRoot to $DestinationPath" -Level "INFO" -ForegroundColor ([System.ConsoleColor]::Cyan)
        # Ensure the destination directory exists
        if (-not (Test-Path -Path $DestinationPath)) {
            New-Item -Path $DestinationPath -ItemType Directory | Out-Null
        }
    }

    process {
        try {
            # Get the path of the script directory
            # $ScriptDirectory = Split-Path -Path $script:MyInvocation.MyCommand.Definition

            # Copy all items from the script's directory to the destination, including subdirectories
            Copy-Item -Path "$PSScriptRoot\*" -Destination $DestinationPath -Recurse -Force -ErrorAction Stop

            Write-EnhancedLog -Message "All items copied successfully from the Script Root $PSScriptRoot to $DestinationPath." -Level "INFO" -ForegroundColor ([System.ConsoleColor]::Green)
        }
        catch {
            Write-EnhancedLog -Message "Error occurred during the copy process: $_" -Level "ERROR" -ForegroundColor ([System.ConsoleColor]::Red)
            throw $_
        }
    }

    end {
        Write-EnhancedLog -Message "Copy process completed." -Level "INFO" -ForegroundColor ([System.ConsoleColor]::Cyan)
    }
}

Copy-FilesToPath -DestinationPath $global:Path_PR



function Verify-CopyOperation {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$DestinationPath
    )

    begin {
        Write-EnhancedLog -Message "Verifying copy operation..." -Level "INFO" -ForegroundColor ([System.ConsoleColor]::Cyan)
        $sourceItems = Get-ChildItem -Path $PSScriptRoot -Recurse
        $destinationItems = Get-ChildItem -Path $DestinationPath -Recurse

        # Use a generic list for better performance compared to using an array with +=
        $verificationResults = New-Object System.Collections.Generic.List[Object]
    }

    process {
        try {
            foreach ($item in $sourceItems) {
                $relativePath = $item.FullName.Substring($PSScriptRoot.Length)
                $correspondingPath = Join-Path -Path $DestinationPath -ChildPath $relativePath

                if (-not (Test-Path -Path $correspondingPath)) {
                    $verificationResults.Add([PSCustomObject]@{
                            Status       = "Missing"
                            SourcePath   = $item.FullName
                            ExpectedPath = $correspondingPath
                        })
                }
            }

            foreach ($item in $destinationItems) {
                $relativePath = $item.FullName.Substring($DestinationPath.Length)
                $correspondingPath = Join-Path -Path $PSScriptRoot -ChildPath $relativePath

                if (-not (Test-Path -Path $correspondingPath)) {
                    $verificationResults.Add([PSCustomObject]@{
                            Status     = "Extra"
                            SourcePath = $correspondingPath
                            ActualPath = $item.FullName
                        })
                }
            }
        }
        catch {
            Write-EnhancedLog -Message "Error during verification process: $_" -Level "ERROR" -ForegroundColor ([System.ConsoleColor]::Red)
        }
    }

    end {
        if ($verificationResults.Count -gt 0) {
            Write-EnhancedLog -Message "Discrepancies found. See detailed log." -Level "WARNING" -ForegroundColor ([System.ConsoleColor]::Yellow)
            $verificationResults | Format-Table -AutoSize | Out-String | ForEach-Object { Write-EnhancedLog -Message $_ -Level "INFO" }
        }
        else {
            Write-EnhancedLog -Message "All items verified successfully. No discrepancies found." -Level "INFO" -ForegroundColor ([System.ConsoleColor]::Green)
        }

        Write-EnhancedLog -Message ("Total items in source: " + $sourceItems.Count) -Level "INFO" -ForegroundColor ([System.ConsoleColor]::Cyan)
        Write-EnhancedLog -Message ("Total items in destination: " + $destinationItems.Count) -Level "INFO" -ForegroundColor ([System.ConsoleColor]::Cyan)
    }
}
# Assuming $global:Path_PR is set to your destination path and Write-EnhancedLog is defined
Verify-CopyOperation -DestinationPath $global:Path_PR






<#
.SYNOPSIS
Creates a VBScript file to run a PowerShell script hidden from the user interface.

.DESCRIPTION
This function generates a VBScript (.vbs) file designed to execute a PowerShell script without displaying the PowerShell window. It's particularly useful for running background tasks or scripts that do not require user interaction. The path to the PowerShell script is taken as an argument, and the VBScript is created in a specified directory within the global path variable.

.EXAMPLE
$Path_VBShiddenPS = Create-VBShiddenPS

This example creates the VBScript file and returns its path.
#>


function Create-VBShiddenPS {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$Path_local,

        [string]$DataFolder = "Data",

        [string]$FileName = "run-ps-hidden.vbs"
    )

    try {
        # Construct the full path for DataFolder and validate it manually
        $fullDataFolderPath = Join-Path -Path $Path_local -ChildPath $DataFolder
        if (-not (Test-Path -Path $fullDataFolderPath -PathType Container)) {
            throw "DataFolder does not exist or is not a directory: $fullDataFolderPath"
        }

        # Log message about creating VBScript
        Write-EnhancedLog -Message "Creating VBScript to hide PowerShell window..." -Level "INFO" -ForegroundColor Magenta

        $scriptBlock = @"
Dim shell,fso,file

Set shell=CreateObject("WScript.Shell")
Set fso=CreateObject("Scripting.FileSystemObject")

strPath=WScript.Arguments.Item(0)

If fso.FileExists(strPath) Then
    set file=fso.GetFile(strPath)
    strCMD="powershell -nologo -executionpolicy ByPass -command " & Chr(34) & "&{" & file.ShortPath & "}" & Chr(34)
    shell.Run strCMD,0
End If
"@

        # Combine paths to construct the full path for the VBScript
        $folderPath = $fullDataFolderPath
        $Path_VBShiddenPS = Join-Path -Path $folderPath -ChildPath $FileName

        # Write the script block to the VBScript file
        $scriptBlock | Out-File -FilePath (New-Item -Path $Path_VBShiddenPS -Force) -Force

        # Validate the VBScript file creation
        if (Test-Path -Path $Path_VBShiddenPS) {
            Write-EnhancedLog -Message "VBScript created successfully at $Path_VBShiddenPS" -Level "INFO" -ForegroundColor Green
        }
        else {
            throw "Failed to create VBScript at $Path_VBShiddenPS"
        }

        return $Path_VBShiddenPS
    }
    catch {
        Write-EnhancedLog -Message "An error occurred while creating VBScript: $_" -Level "ERROR" -ForegroundColor Red
        throw $_
    }
}

# Ensure the script runs with administrative privileges
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-EnhancedLog -Message "Script requires administrative privileges to write to $Path_local." -Level "ERROR" -ForegroundColor ([System.ConsoleColor]::Red)
    exit
}

# Pre-defined paths
$Path_local = "C:\Program Files\_MEM"
$DataFolder = "Data"
$DataFolderPath = Join-Path -Path $Path_local -ChildPath $DataFolder

# Ensure the Data folder exists
if (-not (Test-Path -Path $DataFolderPath -PathType Container)) {
    New-Item -ItemType Directory -Path $DataFolderPath -Force | Out-Null
    Write-EnhancedLog -Message "Data folder created at $DataFolderPath" -Level "INFO" -ForegroundColor ([System.ConsoleColor]::Green)
}
else {
    Write-EnhancedLog -Message "Data folder already exists at $DataFolderPath" -Level "INFO" -ForegroundColor ([System.ConsoleColor]::Yellow)
}

# Then call Create-VBShiddenPS
$FileName = "run-ps-hidden.vbs"
try {
    $global:Path_VBShiddenPS = Create-VBShiddenPS -Path_local $Path_local -DataFolder $DataFolder -FileName $FileName
    # Validation of the VBScript file creation
    if (Test-Path -Path $global:Path_VBShiddenPS) {
        Write-EnhancedLog -Message "Validation successful: VBScript file exists at $global:Path_VBShiddenPS" -Level "INFO" -ForegroundColor ([System.ConsoleColor]::Green)
    }
    else {
        Write-EnhancedLog -Message "Validation failed: VBScript file does not exist at $global:Path_VBShiddenPS. Check script execution and permissions." -Level "WARNING" -ForegroundColor ([System.ConsoleColor]::Red)
    }
}
catch {
    Write-EnhancedLog -Message "An error occurred: $_" -Level "ERROR" -ForegroundColor ([System.ConsoleColor]::Red)
}





<#
.SYNOPSIS
Checks for the existence of a specified scheduled task.

.DESCRIPTION
This function searches for a scheduled task by name and optionally filters it by version. It returns $true if a task matching the specified criteria exists, otherwise $false.

.PARAMETER taskName
The name of the scheduled task to search for.

.PARAMETER version
The version of the scheduled task to match. The task's description must start with "Version" followed by this parameter value.

.EXAMPLE
$exists = Check-ExistingTask -taskName "MyTask" -version "1"
This example checks if a scheduled task named "MyTask" with a description starting with "Version 1" exists.

#>
function Check-ExistingTask {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$taskName,

        [string]$version
    )

    try {
        Write-EnhancedLog -Message "Checking for existing scheduled task: $taskName" -Level "INFO" -ForegroundColor Magenta
        $task_existing = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
        if ($null -eq $task_existing) {
            Write-EnhancedLog -Message "No existing task named $taskName found." -Level "INFO" -ForegroundColor Yellow
            return $false
        }

        if ($null -ne $version) {
            $versionMatch = $task_existing.Description -like "Version $version*"
            if ($versionMatch) {
                Write-EnhancedLog -Message "Found matching task with version: $version" -Level "INFO" -ForegroundColor Green
            }
            else {
                Write-EnhancedLog -Message "No matching version found for task: $taskName" -Level "INFO" -ForegroundColor Yellow
            }
            return $versionMatch
        }

        return $true
    }
    catch {
        Write-EnhancedLog -Message "An error occurred while checking for the scheduled task: $_" -Level "ERROR" -ForegroundColor Red
        throw $_
    }
}







<#
.SYNOPSIS
Executes detection and remediation scripts located in a specified directory.

.DESCRIPTION
This function navigates to the specified directory and executes the detection script. If the detection script exits with a non-zero exit code, indicating a positive detection, the remediation script is then executed. The function uses enhanced logging for status messages and error handling to manage any issues that arise during execution.

.PARAMETER Path_PR
The path to the directory containing the detection and remediation scripts.

.EXAMPLE
Execute-DetectionAndRemediation -Path_PR "C:\Scripts\MyTask"
This example executes the detection and remediation scripts located in "C:\Scripts\MyTask".
#>
function Execute-DetectionAndRemediation {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        # [ValidateScript({Test-Path $_ -PathType 'Container'})]
        [string]$Path_PR
    )

    try {
        Write-EnhancedLog -Message "Executing detection and remediation scripts in $Path_PR..." -Level "INFO" -ForegroundColor Magenta
        Set-Location -Path $Path_PR

        # Execution of the detection script
        & .\detection.ps1
        if ($LASTEXITCODE -ne 0) {
            Write-EnhancedLog -Message "Detection positive, remediation starts now." -Level "INFO" -ForegroundColor Green
            & .\remediation.ps1
        }
        else {
            Write-EnhancedLog -Message "Detection negative, no further action needed." -Level "INFO" -ForegroundColor Yellow
        }
    }
    catch {
        Write-EnhancedLog -Message "An error occurred during detection and remediation execution: $_" -Level "ERROR" -ForegroundColor Red
        throw $_
    }
}







<#
.SYNOPSIS
Registers a scheduled task with the system.

.DESCRIPTION
This function creates a new scheduled task with the specified parameters, including the name, description, VBScript path, and PowerShell script path. It sets up a basic daily trigger and runs the task as the SYSTEM account with the highest privileges. Enhanced logging is used for status messages and error handling to manage potential issues.

.PARAMETER schtaskName
The name of the scheduled task to register.

.PARAMETER schtaskDescription
A description for the scheduled task.

.PARAMETER Path_vbs
The path to the VBScript file used to run the PowerShell script.

.PARAMETER Path_PSscript
The path to the PowerShell script to execute.

.EXAMPLE
MyRegisterScheduledTask -schtaskName "MyTask" -schtaskDescription "Performs automated checks" -Path_vbs "C:\Scripts\run-hidden.vbs" -Path_PSscript "C:\Scripts\myScript.ps1"

This example registers a new scheduled task named "MyTask" that executes "myScript.ps1" using "run-hidden.vbs".
#>
function MyRegisterScheduledTask {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$schtaskName,

        [Parameter(Mandatory = $true)]
        [string]$schtaskDescription,

        [Parameter(Mandatory = $true)]
        # [ValidateScript({Test-Path $_ -File})]
        [string]$Path_vbs,

        [Parameter(Mandatory = $true)]
        # [ValidateScript({Test-Path $_ -File})]
        [string]$Path_PSscript
    )

    try {
        Write-EnhancedLog -Message "Registering scheduled task: $schtaskName" -Level "INFO" -ForegroundColor Magenta

        $startTime = (Get-Date).AddMinutes(1).ToString("HH:mm")

        
        # $action = New-ScheduledTaskAction -Execute "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe" -Argument "`"$Path_vbs`" `"$Path_PSscript`""
        # $argList = "-NoExit -ExecutionPolicy Bypass -File"
        # $action = New-ScheduledTaskAction -Execute "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe" -Argument "`"$argList`" `"$Path_PSscript`""



        # Define the path to the PowerShell script
        # $Path_PSscript = "C:\Path\To\Your\Script.ps1"

        # Define the arguments for the PowerShell executable
        # $argList = "-NoExit -ExecutionPolicy Bypass -File `"$Path_PSscript`""

        # # Create the scheduled task action
        # $action = New-ScheduledTaskAction -Execute "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe" -Argument $argList








        # # Load the configuration from config.json
        # $configPath = Join-Path -Path $PSScriptRoot -ChildPath "config.json"
        # $config = Get-Content -Path $configPath -Raw | ConvertFrom-Json

        # # Define the principal for the task
        # $principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
        # Write-EnhancedLog -Message "Principal for the task defined." -Level "INFO"

        # Define the action based on the provided options in the config.json
        if ($config.UsePSADT) {
            Write-EnhancedLog -Message "setting up Schedule Task action for Service UI and PSADT" -Level "INFO" -ForegroundColor Magenta

            # Define the path to the PowerShell Application Deployment Toolkit executable
            # $ToolkitExecutable = "$PSScriptRoot\Private\PSAppDeployToolkit\Toolkit\Deploy-Application.exe"
            $ToolkitExecutable = "$Path_PR\Private\PSAppDeployToolkit\Toolkit\Deploy-Application.exe"

            # Define the path to the ServiceUI executable
            # $ServiceUIExecutable = "$PSScriptRoot\Private\ServiceUI.exe"
            $ServiceUIExecutable = "$Path_PR\Private\ServiceUI.exe"

            # Define the deployment type
            $DeploymentType = "install"

            # Define the arguments for ServiceUI.exe
            $argList = "-process:explorer.exe `"$ToolkitExecutable`" -DeploymentType $DeploymentType"

            # Create the scheduled task action
            $action = New-ScheduledTaskAction -Execute $ServiceUIExecutable -Argument $argList
        }
        else {
            Write-EnhancedLog -Message "Setting up Scheduled Task action for wscript and VBS" -Level "INFO" -ForegroundColor Magenta

            # Define the arguments for wscript.exe
            $argList = "`"$Path_vbs`" `"$Path_PSscript`""

            # Create the scheduled task action for wscript and VBS
            $action = New-ScheduledTaskAction -Execute "C:\Windows\System32\wscript.exe" -Argument $argList
        }


        # Write-EnhancedLog -Message "Scheduled Task '$($config.TaskName)' created successfully." -Level "INFO"

        


















        #option 1 - NO PSADT but rather Wscript and VBS

        # $action = New-ScheduledTaskAction -Execute "C:\Windows\System32\wscript.exe" -Argument "`"$Path_vbs`" `"$Path_PSscript`""




        # #option 2 - ServiceUI calling PSADT in the SYSTEM context
        # Write-EnhancedLog -Message "setting up Schedule Task action for Service UI and PSADT" -Level "INFO" -ForegroundColor Magenta

        # # Define the path to the PowerShell Application Deployment Toolkit executable
        # # $ToolkitExecutable = "$PSScriptRoot\Private\PSAppDeployToolkit\Toolkit\Deploy-Application.exe"
        # $ToolkitExecutable = "$Path_PR\Private\PSAppDeployToolkit\Toolkit\Deploy-Application.exe"

        # # Define the path to the ServiceUI executable
        # # $ServiceUIExecutable = "$PSScriptRoot\Private\ServiceUI.exe"
        # $ServiceUIExecutable = "$Path_PR\Private\ServiceUI.exe"

        # # Define the deployment type
        # $DeploymentType = "install"

        # # Define the arguments for ServiceUI.exe
        # $argList = "-process:explorer.exe `"$ToolkitExecutable`" -DeploymentType $DeploymentType"

        # # Create the scheduled task action
        # $action = New-ScheduledTaskAction -Execute $ServiceUIExecutable -Argument $argList



        #option 1: Trigger - Daily Frequency

        # $trigger = New-ScheduledTaskTrigger -Daily -At $startTime

        #option 2: Trigger On logon of user defaultuser0 (OOBE)




        # Load the configuration from config.json
        # $configPath = Join-Path -Path $PSScriptRoot -ChildPath "config.json"
        # $config = Get-Content -Path $configPath -Raw | ConvertFrom-Json

        # Define the trigger based on the TriggerType
        if ($config.TriggerType -eq "Daily") {
            $trigger = New-ScheduledTaskTrigger -Daily -At $startTime
            Write-EnhancedLog -Message "Trigger set to Daily at $startTime" -Level "INFO"
        }
        elseif ($config.TriggerType -eq "Logon") {
            if (-not $config.LogonUserId) {
                throw "LogonUserId must be specified for Logon trigger type."
            }
            # $trigger = New-ScheduledTaskTrigger -AtLogOn -User $config.LogonUserId
            $trigger = New-ScheduledTaskTrigger -AtLogOn
            Write-EnhancedLog -Message "Trigger set to logon of user $($config.LogonUserId)" -Level "INFO"
        }
        else {
            throw "Invalid TriggerType specified in the configuration."
        }

        $principal = New-ScheduledTaskPrincipal -UserId "NT AUTHORITY\SYSTEM" -LogonType ServiceAccount -RunLevel Highest

        # $task = Register-ScheduledTask -TaskName $schtaskName -Action $action -Trigger $trigger -Principal $principal -Description $schtaskDescription -Force


  

        # Check if the task should run on demand (Zero triggers defined)
        if ($config.RunOnDemand -eq $true) {
            Write-EnhancedLog -Message "calling Register-ScheduledTask with RunOnDemand set to $($config.RunOnDemand)"
            # Task to run on demand; no trigger defined
            $task = Register-ScheduledTask -TaskName $schtaskName -Action $action -Principal $principal -Description $schtaskDescription -Force

            $task = Get-ScheduledTask -TaskName $schtaskName
        }
        else {
            # Define your trigger here
            Write-EnhancedLog -Message "calling Register-ScheduledTask with RunOnDemand set to $($config.RunOnDemand)"
            $task = Register-ScheduledTask -TaskName $schtaskName -Action $action -Trigger $trigger -Principal $principal -Description $schtaskDescription -Force
            # $DBG

            $task = Get-ScheduledTask -TaskName $schtaskName
            $task.Triggers[0].Repetition.Interval = $RepetitionInterval
            $task | Set-ScheduledTask
        }



        # Updating the task to include repetition with a 5-minute interval
        

        # Check the execution context specified in the config
        if ($PackageExecutionContext -eq "User") {
            # This code block will only execute if ExecutionContext is set to "User"

            # Connect to the Task Scheduler service
            $ShedService = New-Object -ComObject 'Schedule.Service'
            $ShedService.Connect()

            # Get the folder where the task is stored (root folder in this case)
            $taskFolder = $ShedService.GetFolder("\")
    
            # Get the existing task by name
            $Task = $taskFolder.GetTask("$schtaskName")

            # Update the task with a new definition
            $taskFolder.RegisterTaskDefinition("$schtaskName", $Task.Definition, 6, 'Users', $null, 4)  # 6 is TASK_CREATE_OR_UPDATE
        }
        else {
            Write-Host "Execution context is not set to 'User', skipping this block."
        }



        Write-EnhancedLog -Message "Scheduled task $schtaskName registered successfully." -Level "INFO" -ForegroundColor Green
    }
    catch {
        Write-EnhancedLog -Message "An error occurred while registering the scheduled task: $_" -Level "ERROR" -ForegroundColor Red
        throw $_
    }
}







<#
.SYNOPSIS
Sets up a new task environment for scheduled task execution.

.DESCRIPTION
This function prepares the environment for a new scheduled task. It creates a specified directory, determines the PowerShell script path based on the script mode, generates a VBScript to run the PowerShell script hidden, and finally registers the scheduled task with the provided parameters. It utilizes enhanced logging for feedback and error handling to manage potential issues.

.PARAMETER Path_PR
The path where the task's scripts and support files will be stored.

.PARAMETER schtaskName
The name of the scheduled task to be created.

.PARAMETER schtaskDescription
A description for the scheduled task.

.PARAMETER ScriptMode
Determines the script type to be executed ("Remediation" or "PackageName").

.EXAMPLE
SetupNewTaskEnvironment -Path_PR "C:\Tasks\MyTask" -schtaskName "MyScheduledTask" -schtaskDescription "This task does something important" -ScriptMode "Remediation"

This example sets up the environment for a scheduled task named "MyScheduledTask" with a specific description, intended for remediation purposes.
#>
function SetupNewTaskEnvironment {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        # [ValidateScript({Test-Path $_ -PathType 'Container'})]
        [string]$Path_PR,

        [Parameter(Mandatory = $true)]
        [string]$schtaskName,

        [Parameter(Mandatory = $true)]
        [string]$schtaskDescription,

        [Parameter(Mandatory = $true)]
        [ValidateSet("Remediation", "PackageName")]
        [string]$ScriptMode
    )

    try {
        Write-EnhancedLog -Message "Setting up new task environment at $Path_PR." -Level "INFO" -ForegroundColor Cyan

        # New-Item -Path $Path_PR -ItemType Directory -Force | Out-Null
        # Write-EnhancedLog -Message "Created new directory at $Path_PR" -Level "INFO" -ForegroundColor Green

        $Path_PSscript = switch ($ScriptMode) {
            "Remediation" { Join-Path $Path_PR "remediation.ps1" }
            "PackageName" { Join-Path $Path_PR "$PackageName.ps1" }
            Default { throw "Invalid ScriptMode: $ScriptMode. Expected 'Remediation' or 'PackageName'." }
        }

        # $Path_vbs = Create-VBShiddenPS -Path_local $Path_PR
        $Path_vbs = $global:Path_VBShiddenPS

        $scheduledTaskParams = @{
            schtaskName        = $schtaskName
            schtaskDescription = $schtaskDescription
            Path_vbs           = $Path_vbs
            Path_PSscript      = $Path_PSscript
        }

        MyRegisterScheduledTask @scheduledTaskParams

        Write-EnhancedLog -Message "Scheduled task $schtaskName with description '$schtaskDescription' registered successfully." -Level "INFO" -ForegroundColor Green
    }
    catch {
        Write-EnhancedLog -Message "An error occurred during setup of new task environment: $_" -Level "ERROR" -ForegroundColor Red
        throw $_
    }
}



<#
.SYNOPSIS
Checks for an existing scheduled task and executes tasks based on conditions.

.DESCRIPTION
This function checks if a scheduled task with the specified name and version exists. If it does, it proceeds to execute detection and remediation scripts. If not, it sets up a new task environment and registers the task. It uses enhanced logging for status messages and error handling to manage potential issues.

.PARAMETER schtaskName
The name of the scheduled task to check and potentially execute.

.PARAMETER Version
The version of the task to check for. This is used to verify if the correct task version is already scheduled.

.PARAMETER Path_PR
The path to the directory containing the detection and remediation scripts, used if the task needs to be executed.

.EXAMPLE
CheckAndExecuteTask -schtaskName "MyScheduledTask" -Version 1 -Path_PR "C:\Tasks\MyTask"

This example checks for an existing scheduled task named "MyScheduledTask" of version 1. If it exists, it executes the associated tasks; otherwise, it sets up a new environment and registers the task.
#>
function CheckAndExecuteTask {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$schtaskName,

        [Parameter(Mandatory = $true)]
        [int]$Version,

        [Parameter(Mandatory = $true)]
        [string]$Path_PR,

        [Parameter(Mandatory = $true)]
        [string]$ScriptMode # Adding ScriptMode as a parameter
    )

    try {
        Write-EnhancedLog -Message "Checking for existing task: $schtaskName" -Level "INFO" -ForegroundColor Cyan

        $taskExists = Check-ExistingTask -taskName $schtaskName -version $Version
        if ($taskExists) {
            Write-EnhancedLog -Message "Existing task found. Executing detection and remediation scripts." -Level "INFO" -ForegroundColor Green
            Execute-DetectionAndRemediation -Path_PR $Path_PR
        }
        else {
            Write-EnhancedLog -Message "No existing task found. Setting up new task environment." -Level "INFO" -ForegroundColor Yellow
            SetupNewTaskEnvironment -Path_PR $Path_PR -schtaskName $schtaskName -schtaskDescription $schtaskDescription -ScriptMode $ScriptMode
        }
    }
    catch {
        Write-EnhancedLog -Message "An error occurred while checking and executing the task: $_" -Level "ERROR" -ForegroundColor Red
        throw $_
    }
}

# Ensure global variables are initialized correctly beforehand

# Define the parameters in a hashtable using global variables, including ScriptMode
$CheckAndExecuteTaskparams = @{
    schtaskName = $global:schtaskName
    Version     = $global:Version
    Path_PR     = $global:Path_PR
    ScriptMode  = $global:ScriptMode # Assuming you have this variable defined globally
}

# Call the function using splatting with dynamically set global variables
CheckAndExecuteTask @CheckAndExecuteTaskparams
