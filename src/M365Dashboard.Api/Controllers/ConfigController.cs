using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;

namespace M365Dashboard.Api.Controllers;

[ApiController]
[Route("api/[controller]")]
[Authorize]
public class ConfigController : ControllerBase
{
    private readonly IConfiguration _configuration;
    private readonly ILogger<ConfigController> _logger;
    private readonly IWebHostEnvironment _environment;

    public ConfigController(
        IConfiguration configuration, 
        ILogger<ConfigController> logger,
        IWebHostEnvironment environment)
    {
        _configuration = configuration;
        _logger = logger;
        _environment = environment;
    }

    /// <summary>
    /// Get Azure Maps subscription key for frontend map rendering
    /// </summary>
    [HttpGet("azure-maps-key")]
    public IActionResult GetAzureMapsKey()
    {
        var key = _configuration["AzureMaps:SubscriptionKey"];
        
        if (string.IsNullOrEmpty(key))
        {
            _logger.LogWarning("Azure Maps subscription key not configured");
            return Ok(new { key = (string?)null, configured = false });
        }

        return Ok(new { key, configured = true });
    }

    /// <summary>
    /// Download setup script for external services
    /// </summary>
    [HttpGet("setup-script/{service}")]
    [AllowAnonymous]
    public IActionResult GetSetupScript(string service)
    {
        if (service.ToLower() != "azure-maps")
        {
            return NotFound(new { error = "Unknown service" });
        }

        // Try to find the script in the application directory
        var scriptPath = Path.Combine(_environment.ContentRootPath, "Setup-AzureMaps.ps1");
        
        if (System.IO.File.Exists(scriptPath))
        {
            var content = System.IO.File.ReadAllText(scriptPath);
            return File(System.Text.Encoding.UTF8.GetBytes(content), "application/octet-stream", "Setup-AzureMaps.ps1");
        }

        // If script doesn't exist, generate it inline
        var script = GenerateAzureMapsSetupScript();
        return File(System.Text.Encoding.UTF8.GetBytes(script), "application/octet-stream", "Setup-AzureMaps.ps1");
    }

    private static string GenerateAzureMapsSetupScript()
    {
        return @"<#
.SYNOPSIS
    Creates an Azure Maps account and configures the M365 Dashboard to use it.

.DESCRIPTION
    This script automates the setup of Azure Maps for the M365 Dashboard Sign-ins Map feature.
    It creates an Azure Maps account (Gen2, free tier), retrieves the subscription key,
    and updates the application configuration.

.PARAMETER ResourceGroupName
    The Azure resource group name to create the Azure Maps account in.

.PARAMETER Location
    The Azure region for the Azure Maps account. Default is 'westeurope'.

.PARAMETER AccountName
    The name for the Azure Maps account. Default is 'm365dashboard-maps'.

.EXAMPLE
    .\Setup-AzureMaps.ps1 -ResourceGroupName 'rg-m365dashboard'

.NOTES
    Requires Azure CLI to be installed and logged in.
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$ResourceGroupName,

    [Parameter(Mandatory = $false)]
    [string]$Location = 'westeurope',

    [Parameter(Mandatory = $false)]
    [string]$AccountName = 'm365dashboard-maps'
)

$ErrorActionPreference = 'Stop'

Write-Host '=== Azure Maps Setup for M365 Dashboard ===' -ForegroundColor Cyan
Write-Host ''

# Check if Azure CLI is installed
try {
    $azVersion = az version --output json | ConvertFrom-Json
    Write-Host ""[OK] Azure CLI version: $($azVersion.'azure-cli')"" -ForegroundColor Green
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
Write-Host ""[OK] Logged in as: $($account.user.name)"" -ForegroundColor Green
Write-Host ""     Subscription: $($account.name)"" -ForegroundColor Gray

# Check if resource group exists
Write-Host ''
Write-Host ""Checking resource group '$ResourceGroupName'..."" -ForegroundColor Gray
$rgExists = az group exists --name $ResourceGroupName
if ($rgExists -eq 'false') {
    Write-Host ""Creating resource group '$ResourceGroupName' in '$Location'..."" -ForegroundColor Yellow
    az group create --name $ResourceGroupName --location $Location --output none
    Write-Host '[OK] Resource group created' -ForegroundColor Green
}
else {
    Write-Host '[OK] Resource group exists' -ForegroundColor Green
}

# Check if Azure Maps account already exists
Write-Host ''
Write-Host 'Checking for existing Azure Maps account...' -ForegroundColor Gray
$existingAccount = az maps account show --name $AccountName --resource-group $ResourceGroupName --output json 2>$null | ConvertFrom-Json

if ($existingAccount) {
    Write-Host ""[OK] Azure Maps account '$AccountName' already exists"" -ForegroundColor Green
}
else {
    # Create Azure Maps account
    Write-Host ""Creating Azure Maps account '$AccountName'..."" -ForegroundColor Yellow
    Write-Host '     SKU: G2 (Gen2 - includes free tier)' -ForegroundColor Gray
    
    az maps account create --name $AccountName --resource-group $ResourceGroupName --sku G2 --kind Gen2 --accept-tos --output none
    
    Write-Host '[OK] Azure Maps account created' -ForegroundColor Green
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
        Write-Host ""[OK] Updated $(Split-Path $settingsPath -Leaf)"" -ForegroundColor Green
        $updated = $true
    }
}

if (-not $updated) {
    Write-Host ""[!] No appsettings.json found at: $scriptDir"" -ForegroundColor Yellow
    Write-Host '    Please manually add the following to your appsettings.json:' -ForegroundColor Yellow
    Write-Host ''
    Write-Host '  ""AzureMaps"": {' -ForegroundColor Cyan
    Write-Host ""    \""SubscriptionKey\"": \""$subscriptionKey\"""" -ForegroundColor Cyan
    Write-Host '  }' -ForegroundColor Cyan
}

# Summary
Write-Host ''
Write-Host '=== Setup Complete ===' -ForegroundColor Cyan
Write-Host ''
Write-Host 'Azure Maps Account Details:' -ForegroundColor White
Write-Host ""  Name: $AccountName"" -ForegroundColor Gray
Write-Host ""  Resource Group: $ResourceGroupName"" -ForegroundColor Gray
Write-Host ""  Location: $Location"" -ForegroundColor Gray
Write-Host '  SKU: G2 (Gen2)' -ForegroundColor Gray
Write-Host ''
Write-Host 'Subscription Key:' -ForegroundColor White
Write-Host ""  $subscriptionKey"" -ForegroundColor Green
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
";
    }
}
