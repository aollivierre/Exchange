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

function Export-LdapsPfx {
    param (
        [string]$thumbprint = '9B1715B6E26464080A795D43A35E59E6D3CA4277',
        [string]$pfxFileName = 'LDAPS_PRIVATEKEY.pfx',
        [string]$password = 'WfOO$ouPCL>%XcME<bS#(QkfRY+7kXg?'
    )

    $pfxFilePath = Join-Path -Path $scriptDir -ChildPath $pfxFileName
    $pfxPass = (ConvertTo-SecureString -AsPlainText -Force -String $password)
    
    Get-ChildItem "Cert:\LocalMachine\My\$thumbprint" | Export-PfxCertificate -FilePath $pfxFilePath -Password $pfxPass
}

Export-LdapsPfx
