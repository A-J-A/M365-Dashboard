<#
.SYNOPSIS
    Automates Entra ID App Registration for M365 Dashboard

.DESCRIPTION
    This script creates and configures an Entra ID (Azure AD) App Registration
    with APPLICATION permissions for the M365 Dashboard. This allows helpdesk
    users to view tenant-wide data without needing admin rights themselves.

.PARAMETER AppName
    Name for the App Registration (default: "M365 Dashboard")

.PARAMETER RedirectUri
    Redirect URI for the application (default: "https://localhost:5001")

.PARAMETER CreateClientSecret
    Whether to create a client secret (default: $true)

.PARAMETER SecretExpiryMonths
    Number of months until the client secret expires (default: 24)

.EXAMPLE
    .\Register-M365DashboardApp.ps1
    
.EXAMPLE
    .\Register-M365DashboardApp.ps1 -AppName "My M365 Dashboard" -RedirectUri "https://myapp.azurewebsites.net"

.NOTES
    Requires Microsoft.Graph PowerShell module
    Run as a user with Application Administrator or Global Administrator role
    
    PERMISSION MODEL:
    - Delegated: User.Read (for user's own profile in frontend)
    - Application: All others (backend accesses tenant data as the app)
#>

[CmdletBinding()]
param(
    [Parameter()]
    [string]$AppName = "M365 Dashboard",
    
    [Parameter()]
    [string]$RedirectUri = "https://localhost:5001",
    
    [Parameter()]
    [bool]$CreateClientSecret = $true,
    
    [Parameter()]
    [int]$SecretExpiryMonths = 24
)

$ErrorActionPreference = "Stop"

# Colors for output
function Write-Success { param($Message) Write-Host "[OK] $Message" -ForegroundColor Green }
function Write-Info { param($Message) Write-Host "[..] $Message" -ForegroundColor Cyan }
function Write-Warn { param($Message) Write-Host "[!!] $Message" -ForegroundColor Yellow }
function Write-Step { param($Message) Write-Host "`n$Message" -ForegroundColor White }

Write-Host ""
Write-Host "================================================================" -ForegroundColor Blue
Write-Host "       M365 Dashboard - Entra ID App Registration              " -ForegroundColor Blue
Write-Host "                  (Application Permissions)                    " -ForegroundColor Blue
Write-Host "================================================================" -ForegroundColor Blue
Write-Host ""

Write-Host "Permission Model:" -ForegroundColor White
Write-Host "  - Frontend: Delegated (User.Read) - user signs in"
Write-Host "  - Backend:  Application - accesses tenant data as app"
Write-Host "  - Benefit:  Helpdesk users see all data without admin rights"
Write-Host ""

# Check for Microsoft.Graph module
Write-Step "Step 1: Checking prerequisites..."

$graphModule = Get-Module -ListAvailable -Name Microsoft.Graph.Applications
if (-not $graphModule) {
    Write-Info "Microsoft.Graph module not found. Installing..."
    try {
        Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber
        Write-Success "Microsoft.Graph module installed"
    }
    catch {
        Write-Error "Failed to install Microsoft.Graph module. Please run: Install-Module Microsoft.Graph -Scope CurrentUser"
        exit 1
    }
}
else {
    Write-Success "Microsoft.Graph module found (v$($graphModule.Version))"
}

# Thoroughly disconnect any existing sessions
Write-Step "Step 2: Clearing existing sessions and connecting to Microsoft Graph..."

Write-Info "Disconnecting any existing Microsoft Graph sessions..."
try {
    # Disconnect from Microsoft Graph
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
}
catch {
    # Ignore errors
}

# Clear any cached tokens
try {
    # Remove cached Graph contexts
    $env:MSAL_CACHE = $null
    
    # Clear the Graph SDK context
    $mgContext = Get-MgContext -ErrorAction SilentlyContinue
    if ($mgContext) {
        Write-Info "Found existing context for: $($mgContext.Account) - disconnecting..."
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
        Start-Sleep -Seconds 2
    }
}
catch {
    # Ignore errors
}

Write-Success "Cleared existing sessions"
Write-Info "A browser window will open for authentication..."
Write-Info "Please sign in with an account that has Application Administrator or Global Administrator role"
Write-Host ""

try {
    # Connect with required scopes
    Connect-MgGraph -Scopes @(
        "Application.ReadWrite.All",
        "Directory.ReadWrite.All",
        "AppRoleAssignment.ReadWrite.All",
        "DelegatedPermissionGrant.ReadWrite.All"
    ) -NoWelcome
    
    $context = Get-MgContext
    Write-Success "Connected to tenant: $($context.TenantId)"
    Write-Info "Signed in as: $($context.Account)"
}
catch {
    Write-Error "Failed to connect to Microsoft Graph: $_"
    exit 1
}

# Microsoft Graph App ID
$graphAppId = "00000003-0000-0000-c000-000000000000"

# ============================================================================
# PERMISSION DEFINITIONS
# ============================================================================

# Delegated permissions (for frontend user sign-in)
$delegatedPermissions = @(
    @{ Id = "e1fe6dd8-ba31-4d61-89e7-88639da4683d"; Type = "Scope" }  # User.Read
    @{ Id = "37f7f235-527c-4136-accd-4a02d197296e"; Type = "Scope" }  # openid
    @{ Id = "14dad69e-099b-42c9-810b-d002981feec1"; Type = "Scope" }  # profile
    @{ Id = "64a6cdd6-aab1-4aaf-94b8-3cc8405e90d0"; Type = "Scope" }  # email
)

# Application permissions (for backend to access tenant data)
$applicationPermissions = @(
    @{ Id = "df021288-bdef-4463-88db-98f22de89214"; Type = "Role" }  # User.Read.All
    @{ Id = "5b567255-7703-4780-807c-7be8301ae99b"; Type = "Role" }  # Group.Read.All
    @{ Id = "230c1aed-a721-4c5d-9cb4-a90514e508ef"; Type = "Role" }  # Reports.Read.All
    @{ Id = "b0afded3-3588-46d8-8b3d-9842eff778da"; Type = "Role" }  # AuditLog.Read.All
    @{ Id = "7ab1d382-f21e-4acd-a863-ba3e13f7da61"; Type = "Role" }  # Directory.Read.All
    @{ Id = "2f51be20-0bb4-4fed-bf7b-db946066c75e"; Type = "Role" }  # DeviceManagementManagedDevices.Read.All
    @{ Id = "483bed4a-2ad3-4361-a73b-c83ccdbdc53c"; Type = "Role" }  # RoleManagement.Read.Directory
    @{ Id = "246dd0d5-5bd0-4def-940b-0421030a5b68"; Type = "Role" }  # Policy.Read.All
    @{ Id = "78145de6-330d-4800-a6ce-494ff2d33d07"; Type = "Role" }  # UserAuthenticationMethod.Read.All
    @{ Id = "498476ce-e0fe-48b0-b801-37ba7e2685c6"; Type = "Role" }  # GroupMember.Read.All
    @{ Id = "dc5007c0-2d7d-4c42-879c-2dab87571379"; Type = "Role" }  # IdentityRiskyUser.Read.All
    @{ Id = "bf394140-e372-4bf9-a898-299cfc7564e5"; Type = "Role" }  # SecurityEvents.Read.All
    @{ Id = "810c84a8-4a9e-49e6-bf7d-12d183f40d01"; Type = "Role" }  # Mail.Read
    @{ Id = "b633e1c5-b582-4048-a93e-9f11b44c7e96"; Type = "Role" }  # Mail.Send
)

# Combined for RequiredResourceAccess
$allPermissions = $delegatedPermissions + $applicationPermissions

# Check if app already exists
Write-Step "Step 3: Creating App Registration..."

$existingApp = Get-MgApplication -Filter "displayName eq '$AppName'" -ErrorAction SilentlyContinue

if ($existingApp) {
    Write-Warn "An app with name '$AppName' already exists!"
    Write-Host "  Existing App ID: $($existingApp.AppId)"
    Write-Host ""
    $response = Read-Host "Do you want to DELETE and recreate it? (y/n)"
    if ($response -eq 'y') {
        Write-Info "Deleting existing app registration..."
        try {
            # First delete any service principal
            $existingSp = Get-MgServicePrincipal -Filter "appId eq '$($existingApp.AppId)'" -ErrorAction SilentlyContinue
            if ($existingSp) {
                Remove-MgServicePrincipal -ServicePrincipalId $existingSp.Id -ErrorAction SilentlyContinue
                Write-Info "Deleted existing service principal"
            }
            # Then delete the app
            Remove-MgApplication -ApplicationId $existingApp.Id
            Write-Success "Deleted existing app registration"
            Start-Sleep -Seconds 3  # Wait for deletion to propagate
            $existingApp = $null
        }
        catch {
            Write-Warn "Could not delete existing app: $_"
            Write-Info "Continuing with update instead..."
        }
    }
    else {
        Write-Info "Exiting without changes"
        Disconnect-MgGraph | Out-Null
        exit 0
    }
}

# Define App Roles - must be created with the app
$adminRoleId = [Guid]::NewGuid().ToString()
$readerRoleId = [Guid]::NewGuid().ToString()

$appRoles = @(
    @{
        Id = $adminRoleId
        DisplayName = "Dashboard Admin"
        Description = "Full access to dashboard, settings, and cache management"
        Value = "Dashboard.Admin"
        AllowedMemberTypes = @("User")
        IsEnabled = $true
    },
    @{
        Id = $readerRoleId
        DisplayName = "Dashboard Reader"
        Description = "Read-only access to dashboard data"
        Value = "Dashboard.Reader"
        AllowedMemberTypes = @("User")
        IsEnabled = $true
    }
)

# Prepare redirect URIs - ensure no duplicates and proper format
$webRedirectUris = @(
    "https://localhost:5001/signin-oidc",
    "https://localhost:44477/signin-oidc"
) | Select-Object -Unique

if ($RedirectUri -ne "https://localhost:5001" -and $RedirectUri -ne "https://localhost:44477") {
    $webRedirectUris += "$RedirectUri/signin-oidc"
}

$spaRedirectUris = @(
    "https://localhost:5001",
    "https://localhost:44477"
) | Select-Object -Unique

if ($RedirectUri -ne "https://localhost:5001" -and $RedirectUri -ne "https://localhost:44477") {
    $spaRedirectUris += $RedirectUri
}

# Create the application
Write-Info "Creating new app registration..."

try {
    $appParams = @{
        DisplayName = $AppName
        SignInAudience = "AzureADMyOrg"
        Web = @{
            RedirectUris = $webRedirectUris
            ImplicitGrantSettings = @{
                EnableAccessTokenIssuance = $true
                EnableIdTokenIssuance = $true
            }
        }
        Spa = @{
            RedirectUris = $spaRedirectUris
        }
        RequiredResourceAccess = @(
            @{
                ResourceAppId = $graphAppId
                ResourceAccess = $allPermissions
            }
        )
        AppRoles = $appRoles
    }
    
    $app = New-MgApplication -BodyParameter $appParams
    Write-Success "Created app registration"
    Write-Info "Application ID (Client ID): $($app.AppId)"
    Write-Info "Object ID: $($app.Id)"
}
catch {
    Write-Host ""
    Write-Host "ERROR DETAILS:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    if ($_.ErrorDetails) {
        Write-Host $_.ErrorDetails.Message -ForegroundColor Red
    }
    Write-Error "Failed to create app registration: $_"
    Disconnect-MgGraph | Out-Null
    exit 1
}

# Create Service Principal
Write-Step "Step 4: Creating Service Principal..."

Start-Sleep -Seconds 2  # Wait for app to propagate

try {
    $sp = New-MgServicePrincipal -AppId $app.AppId
    Write-Success "Created service principal"
    Write-Info "Service Principal ID: $($sp.Id)"
}
catch {
    Write-Warn "Could not create service principal: $_"
    Write-Info "Trying to get existing service principal..."
    $sp = Get-MgServicePrincipal -Filter "appId eq '$($app.AppId)'" -ErrorAction SilentlyContinue
    if ($sp) {
        Write-Success "Found existing service principal"
    }
    else {
        Write-Error "Failed to create or find service principal"
        exit 1
    }
}

# Create API scope (for frontend to call backend)
Write-Step "Step 5: Configuring API scope..."

$apiScopeId = [Guid]::NewGuid().ToString()
$identifierUri = "api://$($app.AppId)"

try {
    Update-MgApplication -ApplicationId $app.Id -BodyParameter @{
        IdentifierUris = @($identifierUri)
        Api = @{
            Oauth2PermissionScopes = @(
                @{
                    Id = $apiScopeId
                    AdminConsentDescription = "Allow the application to access M365 Dashboard API on behalf of the signed-in user."
                    AdminConsentDisplayName = "Access M365 Dashboard API"
                    UserConsentDescription = "Allow the application to access M365 Dashboard on your behalf."
                    UserConsentDisplayName = "Access M365 Dashboard"
                    IsEnabled = $true
                    Type = "User"
                    Value = "access_as_user"
                }
            )
            PreAuthorizedApplications = @(
                @{
                    AppId = $app.AppId
                    DelegatedPermissionIds = @($apiScopeId)
                }
            )
        }
    }
    Write-Success "Configured API scope: $identifierUri/access_as_user"
    Write-Success "Added client application as pre-authorized (no consent required)"
}
catch {
    Write-Warn "Could not configure API scope: $_"
}

# Create client secret
$clientSecret = $null
if ($CreateClientSecret) {
    Write-Step "Step 6: Creating client secret..."
    
    try {
        $secretParams = @{
            PasswordCredential = @{
                DisplayName = "M365 Dashboard Secret"
                EndDateTime = (Get-Date).AddMonths($SecretExpiryMonths)
            }
        }
        
        $secret = Add-MgApplicationPassword -ApplicationId $app.Id -BodyParameter $secretParams
        $clientSecret = $secret.SecretText
        Write-Success "Created client secret (expires: $($secret.EndDateTime.ToString('yyyy-MM-dd')))"
        Write-Warn "IMPORTANT: Save this secret now - it won't be shown again!"
    }
    catch {
        Write-Warn "Could not create client secret: $_"
    }
}

# Grant admin consent for APPLICATION permissions
Write-Step "Step 7: Granting admin consent for Application permissions..."

Start-Sleep -Seconds 2  # Wait for SP to propagate

try {
    # Get Microsoft Graph service principal
    $graphSp = Get-MgServicePrincipal -Filter "appId eq '$graphAppId'"
    
    if (-not $graphSp) {
        Write-Warn "Could not find Microsoft Graph service principal"
    }
    else {
        # Application permission IDs we need to grant
        $appPermissionIds = @(
            "df021288-bdef-4463-88db-98f22de89214"  # User.Read.All
            "5b567255-7703-4780-807c-7be8301ae99b"  # Group.Read.All
            "230c1aed-a721-4c5d-9cb4-a90514e508ef"  # Reports.Read.All
            "b0afded3-3588-46d8-8b3d-9842eff778da"  # AuditLog.Read.All
            "7ab1d382-f21e-4acd-a863-ba3e13f7da61"  # Directory.Read.All
            "2f51be20-0bb4-4fed-bf7b-db946066c75e"  # DeviceManagementManagedDevices.Read.All
            "483bed4a-2ad3-4361-a73b-c83ccdbdc53c"  # RoleManagement.Read.Directory
            "246dd0d5-5bd0-4def-940b-0421030a5b68"  # Policy.Read.All
            "78145de6-330d-4800-a6ce-494ff2d33d07"  # UserAuthenticationMethod.Read.All
            "498476ce-e0fe-48b0-b801-37ba7e2685c6"  # GroupMember.Read.All
            "dc5007c0-2d7d-4c42-879c-2dab87571379"  # IdentityRiskyUser.Read.All
            "bf394140-e372-4bf9-a898-299cfc7564e5"  # SecurityEvents.Read.All
            "810c84a8-4a9e-49e6-bf7d-12d183f40d01"  # Mail.Read
            "b633e1c5-b582-4048-a93e-9f11b44c7e96"  # Mail.Send
        )
        
        $grantedCount = 0
        $failedCount = 0
        
        foreach ($permId in $appPermissionIds) {
            try {
                $bodyParam = @{
                    PrincipalId = $sp.Id
                    ResourceId = $graphSp.Id
                    AppRoleId = $permId
                }
                New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $sp.Id -BodyParameter $bodyParam -ErrorAction Stop | Out-Null
                $grantedCount++
            }
            catch {
                if ($_.Exception.Message -like "*already exists*" -or $_.Exception.Message -like "*Permission being assigned already exists*") {
                    # Already granted, that's fine
                    $grantedCount++
                }
                else {
                    Write-Warn "Could not grant permission: $_"
                    $failedCount++
                }
            }
        }
        
        Write-Success "Granted $grantedCount of $($appPermissionIds.Count) application permissions"
        
        if ($failedCount -gt 0) {
            Write-Warn "$failedCount permissions need manual consent in Azure Portal"
        }
        
        # Grant delegated permissions (OAuth2 permission grant)
        Write-Info "Granting delegated permissions..."
        
        try {
            $oauth2Grant = @{
                ClientId = $sp.Id
                ConsentType = "AllPrincipals"
                ResourceId = $graphSp.Id
                Scope = "User.Read openid profile email"
            }
            New-MgOauth2PermissionGrant -BodyParameter $oauth2Grant -ErrorAction Stop | Out-Null
            Write-Success "Granted delegated permissions"
        }
        catch {
            if ($_.Exception.Message -like "*already exists*") {
                Write-Success "Delegated permissions already granted"
            }
            else {
                Write-Warn "Could not grant delegated permissions: $_"
            }
        }
    }
}
catch {
    Write-Warn "Could not auto-grant admin consent: $_"
    Write-Info "You may need to grant admin consent manually in the Azure Portal"
    Write-Info "Go to: Azure Portal -> App Registrations -> $AppName -> API Permissions -> Grant admin consent"
}

# Get tenant ID
$tenantId = $context.TenantId

# Output summary
Write-Host ""
Write-Host "================================================================" -ForegroundColor Green
Write-Host "                    Registration Complete!                      " -ForegroundColor Green
Write-Host "================================================================" -ForegroundColor Green
Write-Host ""

Write-Host "Application Details:" -ForegroundColor White
Write-Host "----------------------------------------------------------------"
Write-Host "  App Name:        $AppName"
Write-Host "  Tenant ID:       $tenantId"
Write-Host "  Client ID:       $($app.AppId)"
if ($clientSecret) {
    Write-Host "  Client Secret:   $clientSecret" -ForegroundColor Yellow
}
Write-Host "  API Scope:       $identifierUri/access_as_user"
Write-Host "----------------------------------------------------------------"

Write-Host ""
Write-Host "Permissions Configured:" -ForegroundColor White
Write-Host "  Delegated (Frontend - user context):" -ForegroundColor Cyan
Write-Host "    - User.Read, openid, profile, email"
Write-Host "  Application (Backend - app context):" -ForegroundColor Cyan
Write-Host "    - User.Read.All"
Write-Host "    - Group.Read.All"
Write-Host "    - GroupMember.Read.All"
Write-Host "    - Reports.Read.All"
Write-Host "    - AuditLog.Read.All"
Write-Host "    - Directory.Read.All"
Write-Host "    - DeviceManagementManagedDevices.Read.All"
Write-Host "    - RoleManagement.Read.Directory"
Write-Host "    - Policy.Read.All"
Write-Host "    - UserAuthenticationMethod.Read.All"
Write-Host "    - IdentityRiskyUser.Read.All"
Write-Host "    - SecurityEvents.Read.All"
Write-Host "    - Mail.Read"
Write-Host "    - Mail.Send"
Write-Host "----------------------------------------------------------------"

# Update configuration files
Write-Step "Step 8: Updating configuration files..."

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectRoot = Split-Path -Parent $scriptDir

# Update appsettings.Development.json
$appSettingsPath = Join-Path $projectRoot "src\M365Dashboard.Api\appsettings.Development.json"
if (Test-Path $appSettingsPath) {
    try {
        $appSettings = Get-Content $appSettingsPath -Raw | ConvertFrom-Json
        $appSettings.AzureAd.TenantId = $tenantId
        $appSettings.AzureAd.ClientId = $app.AppId
        if ($clientSecret) {
            $appSettings.AzureAd.ClientSecret = $clientSecret
        }
        $appSettings | ConvertTo-Json -Depth 10 | Set-Content $appSettingsPath
        Write-Success "Updated appsettings.Development.json"
    }
    catch {
        Write-Warn "Could not update appsettings.Development.json: $_"
    }
}
else {
    Write-Warn "appsettings.Development.json not found at: $appSettingsPath"
}

# Update .env.local for frontend
$envPath = Join-Path $projectRoot "src\M365Dashboard.Api\ClientApp\.env.local"
$envContent = @"
# M365 Dashboard Frontend Configuration
# Auto-generated by Register-M365DashboardApp.ps1
# Generated: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")

VITE_AZURE_TENANT_ID=$tenantId
VITE_AZURE_CLIENT_ID=$($app.AppId)
"@

try {
    # Ensure directory exists
    $envDir = Split-Path -Parent $envPath
    if (-not (Test-Path $envDir)) {
        New-Item -ItemType Directory -Path $envDir -Force | Out-Null
    }
    $envContent | Set-Content $envPath
    Write-Success "Created .env.local for frontend"
}
catch {
    Write-Warn "Could not create .env.local: $_"
}

# Create a config summary file
$configSummaryPath = Join-Path $projectRoot "app-registration-config.txt"
$secretExpiry = (Get-Date).AddMonths($SecretExpiryMonths).ToString("yyyy-MM-dd")
$configSummary = @"
M365 Dashboard - App Registration Configuration
Generated: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
================================================================

TENANT CONFIGURATION
--------------------
Tenant ID:     $tenantId
Client ID:     $($app.AppId)
Client Secret: $clientSecret
Secret Expiry: $secretExpiry

API CONFIGURATION  
--------------------
API Scope:     $identifierUri/access_as_user

Redirect URIs (Web): 
  - https://localhost:5001/signin-oidc
  - https://localhost:44477/signin-oidc

Redirect URIs (SPA):
  - https://localhost:5001
  - https://localhost:44477

APP ROLES
--------------------
Dashboard.Admin  - Full access to dashboard, settings, cache management
Dashboard.Reader - Read-only access to dashboard data

NEXT STEPS
--------------------
1. Assign users to app roles in Azure Portal:
   Enterprise Applications -> $AppName -> Users and groups
   -> Add user/group -> Select role (Dashboard.Admin or Dashboard.Reader)

2. Run the application:
   cd src\M365Dashboard.Api
   dotnet run

3. Access the dashboard:
   https://localhost:5001

================================================================
"@

try {
    $configSummary | Set-Content $configSummaryPath
    Write-Success "Saved configuration summary to: app-registration-config.txt"
}
catch {
    Write-Warn "Could not save configuration summary"
}

# Disconnect
Disconnect-MgGraph | Out-Null

Write-Host ""
Write-Host "Next Steps:" -ForegroundColor White
Write-Host "----------------------------------------------------------------"
Write-Host "1. " -NoNewline; Write-Host "Assign yourself the 'Dashboard.Admin' role:" -ForegroundColor Cyan
Write-Host "   -> Azure Portal -> Enterprise Applications -> $AppName"
Write-Host "   -> Users and groups -> Add user/group -> Select 'Dashboard Admin'"
Write-Host ""
Write-Host "2. " -NoNewline; Write-Host "Run the application:" -ForegroundColor Cyan
Write-Host "   cd src\M365Dashboard.Api"
Write-Host "   dotnet run"
Write-Host ""
Write-Host "3. " -NoNewline; Write-Host "Access the dashboard at:" -ForegroundColor Cyan
Write-Host "   https://localhost:5001"
Write-Host "----------------------------------------------------------------"
Write-Host ""

if ($clientSecret) {
    Write-Host "================================================================" -ForegroundColor Yellow
    Write-Host "  IMPORTANT: Copy and save the Client Secret shown above!      " -ForegroundColor Yellow
    Write-Host "  It will NOT be displayed again.                              " -ForegroundColor Yellow
    Write-Host "================================================================" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "Press any key to exit..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
