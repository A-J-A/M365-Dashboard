<#
.SYNOPSIS
    Creates an Azure Maps account and configures the M365 Dashboard to use it.

.DESCRIPTION
    This script automates the setup of Azure Maps for the M365 Dashboard Sign-ins Map feature.
    It creates an Azure Maps account (Gen2, free tier), retrieves the subscription key,
    and updates the application configuration.

.PARAMETER ResourceGroupName
    The Azure resource group name to create the Azure Maps account in.

.PARAMETER Location
    The Azure region for the Azure Maps account. If not specified, you will be prompted to select.

.PARAMETER AccountName
    The name for the Azure Maps account. Default is 'm365dashboard-maps'.

.EXAMPLE
    .\Setup-AzureMaps.ps1 -ResourceGroupName 'rg-m365dashboard'

.EXAMPLE
    .\Setup-AzureMaps.ps1 -ResourceGroupName 'rg-m365dashboard' -Location 'uksouth'

.NOTES
    Requires Azure CLI to be installed and logged in.
#>

param(
    [Parameter(Mandatory = $false)]
    [string]$ResourceGroupName = 'm365-dashboard-maps-rg',

    [Parameter(Mandatory = $false)]
    [string]$Location,

    [Parameter(Mandatory = $false)]
    [string]$AccountName = 'm365dashboard-maps'
)

$ErrorActionPreference = 'Stop'

Write-Host '=== Azure Maps Setup for M365 Dashboard ===' -ForegroundColor Cyan
Write-Host ''

# Check if Azure CLI is installed
try {
    $azVersion = az version --output json | ConvertFrom-Json
    Write-Host "[OK] Azure CLI version: $($azVersion.'azure-cli')" -ForegroundColor Green
}
catch {
    Write-Host '[X] Azure CLI is not installed or not in PATH' -ForegroundColor Red
    Write-Host '    Install from: https://docs.microsoft.com/en-us/cli/azure/install-azure-cli' -ForegroundColor Yellow
    exit 1
}

# Check if logged in
Write-Host 'Checking Azure login status...' -ForegroundColor Gray
$account = az account show --output json 2>$null | ConvertFrom-Json
if (-not $account) {
    Write-Host 'Not logged in to Azure. Initiating login...' -ForegroundColor Yellow
    az login
    $account = az account show --output json | ConvertFrom-Json
}
Write-Host "[OK] Logged in as: $($account.user.name)" -ForegroundColor Green
Write-Host "     Subscription: $($account.name)" -ForegroundColor Gray

# If location not specified, prompt user to select
if (-not $Location) {
    Write-Host ''
    Write-Host 'Select Azure region for the Maps account:' -ForegroundColor Yellow
    Write-Host ''
    
    $locations = @(
        @{ Index = 1; Name = 'uksouth'; Display = 'UK South' }
        @{ Index = 2; Name = 'ukwest'; Display = 'UK West' }
        @{ Index = 3; Name = 'westeurope'; Display = 'West Europe (Netherlands)' }
        @{ Index = 4; Name = 'northeurope'; Display = 'North Europe (Ireland)' }
        @{ Index = 5; Name = 'eastus'; Display = 'East US' }
        @{ Index = 6; Name = 'westus'; Display = 'West US' }
        @{ Index = 7; Name = 'centralus'; Display = 'Central US' }
        @{ Index = 8; Name = 'australiaeast'; Display = 'Australia East' }
        @{ Index = 9; Name = 'southeastasia'; Display = 'Southeast Asia' }
        @{ Index = 10; Name = 'japaneast'; Display = 'Japan East' }
    )
    
    foreach ($loc in $locations) {
        Write-Host "  [$($loc.Index)] $($loc.Display) ($($loc.Name))" -ForegroundColor Gray
    }
    
    Write-Host ''
    $selection = Read-Host 'Enter selection (1-10)'
    
    $selectedLoc = $locations | Where-Object { $_.Index -eq [int]$selection }
    if ($selectedLoc) {
        $Location = $selectedLoc.Name
        Write-Host "[OK] Selected: $($selectedLoc.Display)" -ForegroundColor Green
    }
    else {
        Write-Host '[!] Invalid selection, defaulting to UK South' -ForegroundColor Yellow
        $Location = 'uksouth'
    }
}

# Check if resource group exists
Write-Host ''
Write-Host "Checking resource group '$ResourceGroupName'..." -ForegroundColor Gray
$rgExists = az group exists --name $ResourceGroupName
if ($rgExists -eq 'false') {
    Write-Host "Creating resource group '$ResourceGroupName' in '$Location'..." -ForegroundColor Yellow
    az group create --name $ResourceGroupName --location $Location --output none
    Write-Host '[OK] Resource group created' -ForegroundColor Green
}
else {
    Write-Host '[OK] Resource group exists' -ForegroundColor Green
}

# Check if Azure Maps account already exists (suppress expected error)
Write-Host ''
Write-Host 'Checking for existing Azure Maps account...' -ForegroundColor Gray
$existingAccount = $null
try {
    $existingAccount = az maps account show --name $AccountName --resource-group $ResourceGroupName --output json 2>$null | ConvertFrom-Json
}
catch {
    # Account doesn't exist - this is expected
}

if ($existingAccount) {
    Write-Host "[OK] Azure Maps account '$AccountName' already exists" -ForegroundColor Green
}
else {
    # Create Azure Maps account
    Write-Host "Creating Azure Maps account '$AccountName'..." -ForegroundColor Yellow
    Write-Host '     SKU: G2 (Gen2 - includes free tier)' -ForegroundColor Gray
    
    az maps account create --name $AccountName --resource-group $ResourceGroupName --sku G2 --kind Gen2 --accept-tos --output none
    
    if ($LASTEXITCODE -eq 0) {
        Write-Host '[OK] Azure Maps account created' -ForegroundColor Green
    }
    else {
        Write-Host '[X] Failed to create Azure Maps account' -ForegroundColor Red
        exit 1
    }
}

# Get the subscription key
Write-Host ''
Write-Host 'Retrieving subscription key...' -ForegroundColor Gray
$keys = az maps account keys list --name $AccountName --resource-group $ResourceGroupName --output json | ConvertFrom-Json
$subscriptionKey = $keys.primaryKey

Write-Host '[OK] Subscription key retrieved' -ForegroundColor Green

# Update appsettings.json
Write-Host ''
Write-Host 'Updating application configuration...' -ForegroundColor Gray

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$appSettingsPath = Join-Path $scriptDir 'appsettings.json'
$appSettingsDevPath = Join-Path $scriptDir 'appsettings.Development.json'

$updated = $false

foreach ($settingsPath in @($appSettingsPath, $appSettingsDevPath)) {
    if (Test-Path $settingsPath) {
        $appSettings = Get-Content $settingsPath -Raw | ConvertFrom-Json
        
        if (-not $appSettings.PSObject.Properties['AzureMaps']) {
            $appSettings | Add-Member -NotePropertyName 'AzureMaps' -NotePropertyValue ([PSCustomObject]@{ SubscriptionKey = '' })
        }
        
        $appSettings.AzureMaps.SubscriptionKey = $subscriptionKey
        
        $appSettings | ConvertTo-Json -Depth 10 | Set-Content $settingsPath -Encoding UTF8
        Write-Host "[OK] Updated $(Split-Path $settingsPath -Leaf)" -ForegroundColor Green
        $updated = $true
    }
}

if (-not $updated) {
    Write-Host "[!] No appsettings.json found at: $scriptDir" -ForegroundColor Yellow
    Write-Host '    Please manually add the following to your appsettings.json:' -ForegroundColor Yellow
    Write-Host ''
    Write-Host '  "AzureMaps": {' -ForegroundColor Cyan
    Write-Host "    `"SubscriptionKey`": `"$subscriptionKey`"" -ForegroundColor Cyan
    Write-Host '  }' -ForegroundColor Cyan
}

# Summary
Write-Host ''
Write-Host '=== Setup Complete ===' -ForegroundColor Cyan
Write-Host ''
Write-Host 'Azure Maps Account Details:' -ForegroundColor White
Write-Host "  Name: $AccountName" -ForegroundColor Gray
Write-Host "  Resource Group: $ResourceGroupName" -ForegroundColor Gray
Write-Host "  Location: $Location" -ForegroundColor Gray
Write-Host '  SKU: G2 (Gen2)' -ForegroundColor Gray
Write-Host ''
Write-Host 'Subscription Key:' -ForegroundColor White
Write-Host "  $subscriptionKey" -ForegroundColor Green
Write-Host ''
Write-Host 'Free Tier Limits (per month):' -ForegroundColor White
Write-Host '  - 1,000 free map tile requests' -ForegroundColor Gray
Write-Host '  - 5,000 free geolocation requests' -ForegroundColor Gray
Write-Host '  - 25,000 free search requests' -ForegroundColor Gray
Write-Host ''
Write-Host 'Next Steps:' -ForegroundColor White
Write-Host '  1. Restart the M365 Dashboard application' -ForegroundColor Gray
Write-Host '  2. Navigate to Sign-ins Map page' -ForegroundColor Gray
Write-Host '  3. The map should now display sign-in locations' -ForegroundColor Gray
