[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
#set variables for Entra credentials


# Check if the required modules are installed
# $requiredModules = @("Microsoft.Graph", "Microsoft.Graph.Authentication", "Microsoft.Graph.Intune")
$requiredModules = @("Microsoft.Graph", "Microsoft.Graph.Authentication")
# $requiredModules = @("Microsoft.Graph.Authentication")

foreach ($module in $requiredModules) {
    if (!(Get-Module -ListAvailable -Name $module)) {
        Write-Host "Installing module: $module" -ForegroundColor Yellow
        # Install-Module -Name $module -Force -Verbose
        Install-Module -Name $module -Force
    } else {
        Write-Host "Module $module is already installed" -ForegroundColor Green
    }
}


# Import required modules
# Import-Module -Name Microsoft.Graph.Authentication, Microsoft.Graph.Intune -Force
Import-Module -Name Microsoft.Graph.Authentication
Import-Module -Name Microsoft.Graph.Identity.DirectoryManagement


# $clientId = "YOUR CLIENT ID" # Replace with your app registration client ID
# $clientsecret = "'YOUR CLIENT SECRET'" # Replace with your app registration client ID
# $tenantId = "YOUR TENANT ID" # Replace with your Entra tenant ID

Disconnect-MgGraph

# 
# $clientId = "401add67-6c6e-494c-8501-bee62c16c924" # Replace with your app registration client ID
# $clientsecret = "'YOUR CLIENT SECRET'" # Replace with your app registration client ID
# $tenantId = "7bafa247-244f-431e-9bb1-739c81d19fb2" # Replace with your Entra tenant ID

# Authenticate to Entra using the provided credentials
# Connect-MgGraph -ClientId $clientId -TenantId $tenantId
connect-mgGraph -scopes "User.Read.All", "User.ReadWrite.All", "Directory.ReadWrite.All", "Directory.AccessAsUser.All", "Group.ReadWrite.All", "GroupMember.ReadWrite.All", "Application.ReadWrite.All", "Mail.Send", "Mail.ReadWrite", "MailboxSettings.ReadWrite", "MailboxSettings.Read", "MailboxSettings.ReadWrite"