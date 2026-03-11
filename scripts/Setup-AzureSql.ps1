<#
.SYNOPSIS
    Creates Azure SQL Database for M365 Dashboard

.DESCRIPTION
    This script provisions an Azure SQL Server and Database for the M365 Dashboard.
    It prompts for resource group and region, or uses defaults.

.EXAMPLE
    .\Setup-AzureSql.ps1

.NOTES
    Requires Azure CLI (az) to be installed and logged in
#>

[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"

function Write-Success { param($Message) Write-Host "[OK] $Message" -ForegroundColor Green }
function Write-Info { param($Message) Write-Host "[..] $Message" -ForegroundColor Cyan }
function Write-Warn { param($Message) Write-Host "[!!] $Message" -ForegroundColor Yellow }
function Write-Step { param($Message) Write-Host "`n$Message" -ForegroundColor White }

Write-Host ""
Write-Host "================================================================" -ForegroundColor Blue
Write-Host "       M365 Dashboard - Azure SQL Database Setup               " -ForegroundColor Blue
Write-Host "================================================================" -ForegroundColor Blue
Write-Host ""

# Check for Azure CLI
Write-Step "Step 1: Checking prerequisites..."

$azVersion = az version 2>$null | ConvertFrom-Json
if (-not $azVersion) {
    Write-Host "Azure CLI is not installed." -ForegroundColor Red
    Write-Host "Please install it from: https://docs.microsoft.com/en-us/cli/azure/install-azure-cli" -ForegroundColor Yellow
    exit 1
}
Write-Success "Azure CLI found (v$($azVersion.'azure-cli'))"

# Check if logged in
$account = az account show 2>$null | ConvertFrom-Json
if (-not $account) {
    Write-Info "Not logged in to Azure. Opening browser for login..."
    az login
    $account = az account show | ConvertFrom-Json
}
Write-Success "Logged in as: $($account.user.name)"
Write-Info "Subscription: $($account.name) ($($account.id))"

# Prompt for resource group
Write-Step "Step 2: Resource Group Configuration"
Write-Host ""
Write-Host "Available options:"
Write-Host "  1. Create new resource group (default: 'm365-dashboard-rg')"
Write-Host "  2. Use existing resource group"
Write-Host ""

$rgChoice = Read-Host "Enter choice (1 or 2) [1]"
if ([string]::IsNullOrWhiteSpace($rgChoice)) { $rgChoice = "1" }

if ($rgChoice -eq "1") {
    $defaultRgName = "m365-dashboard-rg"
    $rgName = Read-Host "Enter resource group name [$defaultRgName]"
    if ([string]::IsNullOrWhiteSpace($rgName)) { $rgName = $defaultRgName }
    $createRg = $true
}
else {
    # List existing resource groups
    Write-Info "Fetching existing resource groups..."
    $resourceGroups = az group list --query "[].name" -o tsv
    Write-Host ""
    Write-Host "Existing resource groups:" -ForegroundColor White
    $i = 1
    $rgList = @()
    foreach ($rg in $resourceGroups) {
        Write-Host "  $i. $rg"
        $rgList += $rg
        $i++
    }
    Write-Host ""
    $rgIndex = Read-Host "Enter number of resource group to use"
    $rgName = $rgList[[int]$rgIndex - 1]
    $createRg = $false
}

# Prompt for region
Write-Step "Step 3: Azure Region Selection"
Write-Host ""
Write-Host "Common regions:"
Write-Host "  1. uksouth (UK South)"
Write-Host "  2. ukwest (UK West)"
Write-Host "  3. northeurope (North Europe - Ireland)"
Write-Host "  4. westeurope (West Europe - Netherlands)"
Write-Host "  5. eastus (East US)"
Write-Host "  6. westus2 (West US 2)"
Write-Host "  7. Enter custom region"
Write-Host ""

$regionChoice = Read-Host "Enter choice (1-7) [1]"
if ([string]::IsNullOrWhiteSpace($regionChoice)) { $regionChoice = "1" }

$regionMap = @{
    "1" = "uksouth"
    "2" = "ukwest"
    "3" = "northeurope"
    "4" = "westeurope"
    "5" = "eastus"
    "6" = "westus2"
}

if ($regionChoice -eq "7") {
    $location = Read-Host "Enter Azure region name (e.g., australiaeast)"
}
else {
    $location = $regionMap[$regionChoice]
}

# Generate unique names
$uniqueSuffix = -join ((48..57) + (97..122) | Get-Random -Count 6 | ForEach-Object { [char]$_ })
$sqlServerName = "m365dash-sql-$uniqueSuffix"
$sqlDbName = "M365Dashboard"

# Prompt for SQL admin credentials
Write-Step "Step 4: SQL Server Credentials"
Write-Host ""
$defaultAdminUser = "sqladmin"
$sqlAdminUser = Read-Host "Enter SQL admin username [$defaultAdminUser]"
if ([string]::IsNullOrWhiteSpace($sqlAdminUser)) { $sqlAdminUser = $defaultAdminUser }

# Generate secure password
$passwordChars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!@#$%^&*"
$sqlAdminPassword = -join (1..20 | ForEach-Object { $passwordChars[(Get-Random -Maximum $passwordChars.Length)] })
# Ensure password meets complexity requirements
$sqlAdminPassword = "P@ss" + $sqlAdminPassword + "1!"

Write-Info "A secure password will be generated automatically"

# Confirm settings
Write-Step "Step 5: Confirm Settings"
Write-Host ""
Write-Host "The following resources will be created:" -ForegroundColor White
Write-Host "----------------------------------------------------------------"
Write-Host "  Resource Group:    $rgName $(if ($createRg) { '(new)' } else { '(existing)' })"
Write-Host "  Location:          $location"
Write-Host "  SQL Server:        $sqlServerName.database.windows.net"
Write-Host "  Database:          $sqlDbName"
Write-Host "  Admin Username:    $sqlAdminUser"
Write-Host "  SKU:               Basic (5 DTU) - approx. £3.77/month"
Write-Host "----------------------------------------------------------------"
Write-Host ""

$confirm = Read-Host "Proceed with creation? (y/n) [y]"
if ([string]::IsNullOrWhiteSpace($confirm)) { $confirm = "y" }
if ($confirm -ne "y") {
    Write-Host "Cancelled by user" -ForegroundColor Yellow
    exit 0
}

# Create resources
Write-Step "Step 6: Creating Azure Resources..."

# Create resource group if needed
if ($createRg) {
    Write-Info "Creating resource group '$rgName'..."
    az group create --name $rgName --location $location | Out-Null
    Write-Success "Resource group created"
}

# Create SQL Server
Write-Info "Creating SQL Server '$sqlServerName' (this may take a few minutes)..."
az sql server create `
    --name $sqlServerName `
    --resource-group $rgName `
    --location $location `
    --admin-user $sqlAdminUser `
    --admin-password $sqlAdminPassword | Out-Null
Write-Success "SQL Server created"

# Configure firewall - allow Azure services
Write-Info "Configuring firewall rules..."
az sql server firewall-rule create `
    --resource-group $rgName `
    --server $sqlServerName `
    --name "AllowAzureServices" `
    --start-ip-address 0.0.0.0 `
    --end-ip-address 0.0.0.0 | Out-Null

# Get current public IP and add firewall rule
try {
    $publicIp = (Invoke-RestMethod -Uri "https://api.ipify.org?format=json").ip
    Write-Info "Adding firewall rule for your IP ($publicIp)..."
    az sql server firewall-rule create `
        --resource-group $rgName `
        --server $sqlServerName `
        --name "ClientIP-$(Get-Date -Format 'yyyyMMdd')" `
        --start-ip-address $publicIp `
        --end-ip-address $publicIp | Out-Null
    Write-Success "Firewall rules configured"
}
catch {
    Write-Warn "Could not detect public IP. You may need to add firewall rule manually."
}

# Create database
Write-Info "Creating database '$sqlDbName'..."
az sql db create `
    --resource-group $rgName `
    --server $sqlServerName `
    --name $sqlDbName `
    --edition Basic `
    --capacity 5 `
    --max-size 2GB | Out-Null
Write-Success "Database created"

# Build connection string
$connectionString = "Server=tcp:$sqlServerName.database.windows.net,1433;Initial Catalog=$sqlDbName;Persist Security Info=False;User ID=$sqlAdminUser;Password=$sqlAdminPassword;MultipleActiveResultSets=True;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;"

# Update appsettings
Write-Step "Step 7: Updating configuration files..."

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectRoot = Split-Path -Parent $scriptDir
$appSettingsPath = Join-Path $projectRoot "src\M365Dashboard.Api\appsettings.Development.json"

if (Test-Path $appSettingsPath) {
    try {
        $appSettings = Get-Content $appSettingsPath -Raw | ConvertFrom-Json
        
        # Ensure ConnectionStrings object exists
        if (-not $appSettings.ConnectionStrings) {
            $appSettings | Add-Member -NotePropertyName "ConnectionStrings" -NotePropertyValue @{} -Force
        }
        
        $appSettings.ConnectionStrings.DefaultConnection = $connectionString
        $appSettings | ConvertTo-Json -Depth 10 | Set-Content $appSettingsPath
        Write-Success "Updated appsettings.Development.json"
    }
    catch {
        Write-Warn "Could not update appsettings.Development.json: $_"
    }
}

# Save credentials to file
$credentialsPath = Join-Path $projectRoot "azure-sql-credentials.txt"
$credentialsContent = @"
Azure SQL Database Credentials
Generated: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
================================================================

IMPORTANT: Keep this file secure and do not commit to source control!

Resource Group:     $rgName
Location:           $location
SQL Server:         $sqlServerName.database.windows.net
Database:           $sqlDbName
Admin Username:     $sqlAdminUser
Admin Password:     $sqlAdminPassword

Connection String:
$connectionString

Azure Portal:
https://portal.azure.com/#resource/subscriptions/$($account.id)/resourceGroups/$rgName/providers/Microsoft.Sql/servers/$sqlServerName

================================================================
"@

$credentialsContent | Set-Content $credentialsPath
Write-Success "Saved credentials to: azure-sql-credentials.txt"

# Apply migrations
Write-Step "Step 8: Applying database migrations..."

$apiPath = Join-Path $projectRoot "src\M365Dashboard.Api"
Push-Location $apiPath

try {
    Write-Info "Running Entity Framework migrations..."
    $env:ConnectionStrings__DefaultConnection = $connectionString
    dotnet ef database update 2>&1 | ForEach-Object { Write-Host $_ }
    Write-Success "Database migrations applied"
}
catch {
    Write-Warn "Could not apply migrations automatically: $_"
    Write-Info "Run manually: dotnet ef database update"
}
finally {
    Pop-Location
}

# Summary
Write-Host ""
Write-Host "================================================================" -ForegroundColor Green
Write-Host "                    Setup Complete!                             " -ForegroundColor Green
Write-Host "================================================================" -ForegroundColor Green
Write-Host ""
Write-Host "Azure SQL Database Details:" -ForegroundColor White
Write-Host "----------------------------------------------------------------"
Write-Host "  Server:       $sqlServerName.database.windows.net"
Write-Host "  Database:     $sqlDbName"
Write-Host "  Username:     $sqlAdminUser"
Write-Host "  Password:     $sqlAdminPassword" -ForegroundColor Yellow
Write-Host "----------------------------------------------------------------"
Write-Host ""
Write-Host "Configuration files updated:" -ForegroundColor White
Write-Host "  - appsettings.Development.json"
Write-Host "  - azure-sql-credentials.txt (keep secure!)"
Write-Host ""
Write-Host "Estimated monthly cost: ~£3.77 (Basic tier, 5 DTU)" -ForegroundColor Cyan
Write-Host ""
Write-Host "Next steps:" -ForegroundColor White
Write-Host "  1. Restart the backend: dotnet run"
Write-Host "  2. Refresh the dashboard in your browser"
Write-Host ""

Write-Host "================================================================" -ForegroundColor Yellow
Write-Host "  IMPORTANT: Save the password above - it won't be shown again! " -ForegroundColor Yellow
Write-Host "  Credentials also saved to: azure-sql-credentials.txt         " -ForegroundColor Yellow
Write-Host "================================================================" -ForegroundColor Yellow
Write-Host ""

Write-Host "Press any key to exit..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
