<#
.SYNOPSIS
    Register M365 Dashboard App in Entra ID (Azure AD)
.DESCRIPTION
    This script automates the creation of the Entra ID App Registration with all
    required permissions for the M365 Dashboard. Run this BEFORE deploying infrastructure.
.PARAMETER AppName
    Name for the App Registration (default: M365 Dashboard)
.EXAMPLE
    .\Register-EntraApp.ps1
.EXAMPLE
    .\Register-EntraApp.ps1 -AppName "My M365 Dashboard"
#>

param(
    [string]$AppName
)

$ErrorActionPreference = "Stop"

Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "M365 Dashboard - Entra ID App Registration" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""

# Check Azure CLI is installed and logged in
Write-Host "Checking Azure CLI..." -ForegroundColor Yellow
try {
    $accountJson = az account show 2>$null
    if (-not $accountJson) {
        Write-Host "Please run 'az login' first" -ForegroundColor Red
        exit 1
    }
    $account = $accountJson | ConvertFrom-Json
    Write-Host "Logged in as: $($account.user.name)" -ForegroundColor Green
    Write-Host "Tenant: $($account.tenantId)" -ForegroundColor Green
} catch {
    Write-Host "Azure CLI not found or not logged in. Please install Azure CLI and run 'az login'" -ForegroundColor Red
    exit 1
}

$tenantId = $account.tenantId

# Prompt for app name if not provided
if (-not $AppName) {
    Write-Host ""
    Write-Host "App Registration Name" -ForegroundColor Cyan
    Write-Host "---------------------" -ForegroundColor Cyan
    Write-Host "Enter a name for the Entra ID App Registration."
    Write-Host "This name will appear in Azure Portal and when users sign in."
    Write-Host ""
    Write-Host "Examples: 'M365 Dashboard', 'Contoso M365 Dashboard', 'IT Admin Portal'"
    Write-Host ""
    $AppName = Read-Host "App name (default: M365 Dashboard)"
    
    if ([string]::IsNullOrWhiteSpace($AppName)) {
        $AppName = "M365 Dashboard"
    }
    Write-Host "  Using: $AppName" -ForegroundColor Green
}

# Microsoft Graph App ID (constant)
$graphAppId = "00000003-0000-0000-c000-000000000000"

Write-Host ""
Write-Host "Creating App Registration: $AppName" -ForegroundColor Yellow

# Step 1: Create the app registration using cmd to avoid PowerShell stderr issues
$appJson = cmd /c "az ad app create --display-name `"$AppName`" --sign-in-audience AzureADMyOrg --enable-access-token-issuance true --enable-id-token-issuance true --web-redirect-uris https://localhost:5001/authentication/login-callback http://localhost:5173/authentication/login-callback 2>nul"

if (-not $appJson) {
    # App might already exist, try to get it
    Write-Host "  Checking for existing app..." -ForegroundColor Gray
    $appJson = cmd /c "az ad app list --display-name `"$AppName`" --query `"[0]`" 2>nul"
}

if (-not $appJson -or $appJson -eq "null") {
    Write-Host "Failed to create or find app registration" -ForegroundColor Red
    exit 1
}

$app = $appJson | ConvertFrom-Json
$appId = $app.appId
$objectId = $app.id

Write-Host "App ready!" -ForegroundColor Green
Write-Host "  Client ID: $appId" -ForegroundColor White

# Step 2: Create Service Principal (Enterprise Application)
Write-Host "Creating Service Principal..." -ForegroundColor Yellow
$spJson = cmd /c "az ad sp create --id $appId 2>nul"
if (-not $spJson) {
    # Already exists, get it
    $spJson = cmd /c "az ad sp show --id $appId 2>nul"
}
Write-Host "  Service Principal ready" -ForegroundColor Green

# Step 3: Add Microsoft Graph API permissions
Write-Host "Adding Microsoft Graph permissions..." -ForegroundColor Yellow

# Microsoft Graph permissions
$graphPermissions = @(
    # Core
    @{ id = "df021288-bdef-4463-88db-98f22de89214"; name = "User.Read.All" }
    @{ id = "5b567255-7703-4780-807c-7be8301ae99b"; name = "Group.Read.All" }
    @{ id = "7ab1d382-f21e-4acd-a863-ba3e13f7da61"; name = "Directory.Read.All" }
    @{ id = "498476ce-e0fe-48b0-b801-37ba7e2685c6"; name = "Organization.Read.All" }
    @{ id = "246dd0d5-5bd0-4def-940b-0421030a5b68"; name = "Policy.Read.All" }
    @{ id = "dbb9058a-0e50-45d7-ae91-66909b5d4664"; name = "Domain.Read.All" }
    # Devices & Intune
    @{ id = "7438b122-aefc-4978-80ed-43db9fcc7715"; name = "Device.Read.All" }
    @{ id = "2f51be20-0bb4-4fed-bf7b-db946066c75e"; name = "DeviceManagementManagedDevices.Read.All" }
    @{ id = "dc377aa6-52d8-4e23-b271-2a7ae04cedf3"; name = "DeviceManagementConfiguration.Read.All" }
    @{ id = "7a6ee1e7-141e-4cec-ae74-d9db155731ff"; name = "DeviceManagementApps.Read.All" }
    @{ id = "06a5fe6d-c49d-46a7-b082-56b1b14103c7"; name = "DeviceManagementServiceConfig.Read.All" }
    # Mail & Reports
    @{ id = "810c84a8-4a9e-49e6-bf7d-12d183f40d01"; name = "Mail.Read" }
    @{ id = "b633e1c5-b582-4048-a93e-9f11b44c7e96"; name = "Mail.Send" }
    @{ id = "230c1aed-a721-4c5d-9cb4-a90514e508ef"; name = "Reports.Read.All" }
    # Security
    @{ id = "bf394140-e372-4bf9-a898-299cfc7564e5"; name = "SecurityEvents.Read.All" }
    @{ id = "dc5007c0-2d7d-4c42-879c-2dab87571379"; name = "IdentityRiskyUser.Read.All" }
    @{ id = "6e472fd1-ad78-48da-a0f0-97ab2c6b769e"; name = "IdentityRiskEvent.Read.All" }
    @{ id = "b0afded3-3588-46d8-8b3d-9842eff778da"; name = "AuditLog.Read.All" }
    @{ id = "e0b77adb-e790-44a3-b0a0-257d06303687"; name = "UserAuthenticationMethod.Read.All" }
    @{ id = "93283d0a-6322-4fa8-966b-8c121624760d"; name = "AttackSimulation.Read.All" }
    # SharePoint
    @{ id = "332a536c-c7ef-4017-ab91-336970924f0d"; name = "Sites.Read.All" }
    # Teams
    @{ id = "45bbb07e-7321-4fd7-a8f6-3ff27e6a81c8"; name = "CallRecords.Read.All" }
)

foreach ($perm in $graphPermissions) {
    cmd /c "az ad app permission add --id $appId --api $graphAppId --api-permissions $($perm.id)=Role 2>nul" | Out-Null
    Write-Host "  Added: $($perm.name)" -ForegroundColor Gray
}

# Exchange Online permissions
Write-Host "Adding Exchange Online permissions..." -ForegroundColor Yellow
$exchangeAppId = "00000002-0000-0ff1-ce00-000000000000"
$exchangePermId = "dc50a0fb-09a3-484d-be87-e023b12c6440" # Exchange.ManageAsApp
cmd /c "az ad app permission add --id $appId --api $exchangeAppId --api-permissions $exchangePermId=Role 2>nul" | Out-Null
Write-Host "  Added: Exchange.ManageAsApp" -ForegroundColor Gray

# Step 4: Add App Roles
Write-Host "Adding App Roles..." -ForegroundColor Yellow

$adminRoleId = [guid]::NewGuid().ToString()
$readerRoleId = [guid]::NewGuid().ToString()

$appRolesJson = @"
[
    {
        "id": "$adminRoleId",
        "allowedMemberTypes": ["User"],
        "description": "Full administrative access to M365 Dashboard",
        "displayName": "Dashboard Admin",
        "isEnabled": true,
        "value": "Dashboard.Admin"
    },
    {
        "id": "$readerRoleId",
        "allowedMemberTypes": ["User"],
        "description": "Read-only access to M365 Dashboard",
        "displayName": "Dashboard Reader",
        "isEnabled": true,
        "value": "Dashboard.Reader"
    }
]
"@

# Write to temp file for az cli
$tempFile = [System.IO.Path]::GetTempFileName()
$appRolesJson | Out-File -FilePath $tempFile -Encoding UTF8

cmd /c "az ad app update --id $appId --app-roles @$tempFile 2>nul" | Out-Null
Remove-Item $tempFile -Force -ErrorAction SilentlyContinue

Write-Host "  Added: Dashboard Admin" -ForegroundColor Gray
Write-Host "  Added: Dashboard Reader" -ForegroundColor Gray

# Step 4b: Expose access_as_user scope
Write-Host "Exposing access_as_user scope..." -ForegroundColor Yellow
$appScopeId = [guid]::NewGuid().ToString()
$scopeBody = "{`"api`":{`"oauth2PermissionScopes`":[{`"adminConsentDescription`":`"Allow the application to access M365 Dashboard on behalf of the signed-in user`",`"adminConsentDisplayName`":`"Access M365 Dashboard`",`"id`":`"$appScopeId`",`"isEnabled`":true,`"type`":`"User`",`"userConsentDescription`":`"Allow the application to access M365 Dashboard on your behalf`",`"userConsentDisplayName`":`"Access M365 Dashboard`",`"value`":`"access_as_user`"} ] }}"
$scopeFile = [System.IO.Path]::GetTempFileName() + ".json"
[System.IO.File]::WriteAllText($scopeFile, $scopeBody, [System.Text.Encoding]::UTF8)
cmd /c "az rest --method PATCH --uri `"https://graph.microsoft.com/v1.0/applications/$objectId`" --body @`"$scopeFile`" --headers Content-Type=application/json 2>nul" | Out-Null
Remove-Item $scopeFile -ErrorAction SilentlyContinue
Write-Host "  access_as_user scope exposed" -ForegroundColor Green

# Step 5: Create client secret
Write-Host "Creating client secret..." -ForegroundColor Yellow
$secretJson = cmd /c "az ad app credential reset --id $appId --append --display-name M365Dashboard-Secret --years 2 2>nul"
if (-not $secretJson) {
    Write-Host "  Failed to create secret" -ForegroundColor Red
    exit 1
}
$secret = $secretJson | ConvertFrom-Json
$clientSecret = $secret.password

Write-Host "  Secret created (valid for 2 years)" -ForegroundColor Green

# Step 6: Grant admin consent
Write-Host "Granting admin consent..." -ForegroundColor Yellow
Write-Host "  (This requires Global Administrator or Privileged Role Administrator)" -ForegroundColor Gray

cmd /c "az ad app permission admin-consent --id $appId 2>nul"
if ($LASTEXITCODE -eq 0) {
    Write-Host "  Admin consent granted!" -ForegroundColor Green
} else {
    Write-Host "  Could not auto-grant consent. Please grant manually in Azure Portal:" -ForegroundColor Yellow
    Write-Host "  Azure Portal > App registrations > $AppName > API permissions > Grant admin consent" -ForegroundColor Gray
}

# Save configuration
$configOutput = @{
    TenantId = $tenantId
    ClientId = $appId
    ClientSecret = $clientSecret
    AppName = $AppName
    CreatedAt = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
}

$configPath = Join-Path (Join-Path $PSScriptRoot "..") "entra-app-config.json"
$configOutput | ConvertTo-Json | Out-File $configPath -Encoding UTF8

# Auto-generate appsettings.Development.json so local dev works immediately
# This file is in .gitignore - secrets stay off disk and out of source control
Write-Host "Updating appsettings.Development.json for local development..." -ForegroundColor Yellow

$devSettingsPath = Join-Path (Join-Path $PSScriptRoot "..") "src\M365Dashboard.Api\appsettings.Development.json"

# Preserve any existing non-secret values if the file already exists
$existingSettings = @{}
if (Test-Path $devSettingsPath) {
    try {
        $existingSettings = Get-Content $devSettingsPath -Raw | ConvertFrom-Json -AsHashtable
    } catch {
        $existingSettings = @{}
    }
}

$devSettings = @{
    Logging = @{
        LogLevel = @{
            Default = "Debug"
            "Microsoft.AspNetCore" = "Warning"
            "Microsoft.EntityFrameworkCore" = "Warning"
            M365Dashboard = "Debug"
        }
    }
    AzureAd = @{
        Instance = "https://login.microsoftonline.com/"
        TenantId = $tenantId
        ClientId = $appId
        ClientSecret = $clientSecret
        Audience = "api://$appId"
    }
    KeyVault = @{
        # Leave empty locally - dev uses inline values above
        # Set this to your Key Vault URI to test Key Vault auth locally
        Uri = if ($existingSettings.KeyVault.Uri) { $existingSettings.KeyVault.Uri } else { "" }
    }
    ConnectionStrings = @{
        # Populated automatically by Deploy-M365Dashboard.ps1 after infrastructure is created
        DefaultConnection = if ($existingSettings.ConnectionStrings.DefaultConnection) { $existingSettings.ConnectionStrings.DefaultConnection } else { "" }
    }
    Cache = @{
        DefaultTtlMinutes = 5
        SignInDataTtlMinutes = 2
        LicenseDataTtlMinutes = 15
        ReportDataTtlMinutes = 10
    }
}

$devSettings | ConvertTo-Json -Depth 10 | Out-File $devSettingsPath -Encoding UTF8
Write-Host "  appsettings.Development.json updated" -ForegroundColor Green
Write-Host "  (This file is gitignored - secrets are safe)" -ForegroundColor Gray

# Output results
Write-Host ""
Write-Host "============================================" -ForegroundColor Green
Write-Host "App Registration Complete!" -ForegroundColor Green
Write-Host "============================================" -ForegroundColor Green
Write-Host ""
Write-Host "Configuration saved to: entra-app-config.json" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Tenant ID:     $tenantId" -ForegroundColor White
Write-Host "  Client ID:     $appId" -ForegroundColor White
Write-Host "  Client Secret: $clientSecret" -ForegroundColor White
Write-Host ""
Write-Host "  !!! SAVE THE CLIENT SECRET - You won't see it again !!!" -ForegroundColor Red
Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "Next Steps:" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "1. Run the deployment script:" -ForegroundColor White
Write-Host "   .\scripts\Deploy-M365Dashboard.ps1" -ForegroundColor Gray
Write-Host ""
Write-Host "That's it! The deployment script handles everything else automatically." -ForegroundColor White
Write-Host ""
