<#
.SYNOPSIS
    A short one-line action-based description, e.g. 'Tests if a function is valid'
.DESCRIPTION
    A longer description of the function, its purpose, common use cases, etc.
.NOTES
    Information or caveats about the function e.g. 'This function is not supported in Linux'
.LINK
    Specify a URI to a help page, this will show when Get-Help -Online is used.
.EXAMPLE


Running as Administrator
Connecting to "AGH-DC01.AGH.com"
Logging in as current user using SSPI
Importing directory from file "C:\Code\AD\SSO_LDAP\OpenSSL_LDAP_SSO_DC2\LDAPS\Artifcats\enable_ldaps.txt"
Loading entries..
1 entry modified successfully.

The command has completed successfully

#>




# Check if running as Administrator
$IsAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $IsAdmin) {
    # Get the ID and security principal of the current user account
    $myWindowsID=[System.Security.Principal.WindowsIdentity]::GetCurrent()
    $myWindowsPrincipal=new-object System.Security.Principal.WindowsPrincipal($myWindowsID)

    # Get the security principal for the administrator role
    $adminRole=[System.Security.Principal.WindowsBuiltInRole]::Administrator

    # Check to see if we are currently running "as Administrator"
    if ($myWindowsPrincipal.IsInRole($adminRole))
    {
        # We are running "as Administrator" - so change the title and background color to indicate this
        $Host.UI.RawUI.WindowTitle = $myInvocation.MyCommand.Definition + "(Elevated)"
        $Host.UI.RawUI.BackgroundColor = "DarkBlue"
        clear-host
    }
    else {
        # We are not running "as Administrator" - so relaunch as administrator
        # Create a new process object that starts PowerShell
        $newProcess = new-object System.Diagnostics.ProcessStartInfo "PowerShell";
        # Specify the current script path and name as a parameter
        $newProcess.Arguments = $myInvocation.MyCommand.Definition;
        # Indicate that the process should be elevated
        $newProcess.Verb = "runas";
        # Start the new process
        [System.Diagnostics.Process]::Start($newProcess);
        # Exit from the current, unelevated, process
        exit
    }
}

# Your script code goes here
Write-Output "Running as Administrator"

$scriptDir = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
function Invoke-LdifdeImport {
    param (
        [string]$fileName = 'enable_ldaps.txt'
    )

    $filePath = Join-Path -Path $scriptDir -ChildPath $fileName

    ldifde -i -f $filePath
}

Invoke-LdifdeImport