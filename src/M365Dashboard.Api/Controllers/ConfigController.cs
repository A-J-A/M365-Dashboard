using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Azure.Identity;
using Azure.Core;
using System.Text;
using System.Text.Json;

namespace M365Dashboard.Api.Controllers;

[ApiController]
[Route("api/update")]
[Authorize]
public class UpdateController : ControllerBase
{
    private readonly IConfiguration _configuration;
    private readonly ILogger<UpdateController> _logger;
    private readonly IHttpClientFactory _httpClientFactory;

    private static readonly string VersionFilePath = "/app/version.txt";
    private const string FallbackVersion = "unknown";

    public UpdateController(
        IConfiguration configuration,
        ILogger<UpdateController> logger,
        IHttpClientFactory httpClientFactory)
    {
        _configuration = configuration;
        _logger = logger;
        _httpClientFactory = httpClientFactory;
    }

    [HttpGet("check")]
    public async Task<IActionResult> CheckForUpdates()
    {
        var ghcrRepo = _configuration["ContainerApp:GhcrRepo"] ?? "Alex-C1/m365-dashboard";
        var currentVersion = GetCurrentVersion();

        try
        {
            var client = _httpClientFactory.CreateClient();
            client.DefaultRequestHeaders.Add("User-Agent", "M365Dashboard-UpdateCheck/1.0");
            client.DefaultRequestHeaders.Add("Accept", "application/vnd.github+json");

            var response = await client.GetAsync(
                $"https://api.github.com/repos/{ghcrRepo}/releases/latest");

            if (!response.IsSuccessStatusCode)
            {
                return Ok(new
                {
                    currentVersion,
                    latestVersion    = (string?)null,
                    updateAvailable  = false,
                    releaseNotes     = (string?)null,
                    releaseUrl       = (string?)null,
                    publishedAt      = (string?)null,
                    error            = $"GitHub API returned {(int)response.StatusCode}",
                    updateConfigured = IsUpdateConfigured(),
                });
            }

            var json    = await response.Content.ReadAsStringAsync();
            var release = JsonDocument.Parse(json).RootElement;

            var latestVersion = release.GetProperty("tag_name").GetString() ?? "unknown";
            var releaseNotes  = release.GetProperty("body").GetString();
            var releaseUrl    = release.GetProperty("html_url").GetString();
            var publishedAt   = release.GetProperty("published_at").GetString();
            var updateAvailable = IsNewerVersion(latestVersion, currentVersion);

            return Ok(new
            {
                currentVersion,
                latestVersion,
                updateAvailable,
                releaseNotes,
                releaseUrl,
                publishedAt,
                error            = (string?)null,
                updateConfigured = IsUpdateConfigured(),
            });
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Failed to check for updates from GitHub");
            return Ok(new
            {
                currentVersion,
                latestVersion    = (string?)null,
                updateAvailable  = false,
                releaseNotes     = (string?)null,
                releaseUrl       = (string?)null,
                publishedAt      = (string?)null,
                error            = "Could not reach GitHub. Check your internet connection.",
                updateConfigured = IsUpdateConfigured(),
            });
        }
    }

    [HttpPost("apply")]
    [Authorize(Policy = "RequireAdminRole")]
    public async Task<IActionResult> ApplyUpdate([FromBody] ApplyUpdateRequest request)
    {
        if (string.IsNullOrWhiteSpace(request.Version))
            return BadRequest(new { error = "Version is required" });

        var subscriptionId = _configuration["ContainerApp:SubscriptionId"];
        var resourceGroup  = _configuration["ContainerApp:ResourceGroup"];
        var appName        = _configuration["ContainerApp:Name"];
        var ghcrRepo       = _configuration["ContainerApp:GhcrRepo"] ?? "Alex-C1/m365-dashboard";

        if (string.IsNullOrEmpty(subscriptionId) || string.IsNullOrEmpty(resourceGroup) || string.IsNullOrEmpty(appName))
            return BadRequest(new { error = "Container App update is not configured. Set ContainerApp:SubscriptionId, ContainerApp:ResourceGroup and ContainerApp:Name in configuration." });

        var imageName   = $"ghcr.io/{ghcrRepo.ToLowerInvariant()}:{request.Version}";
        var requestedBy = User.FindFirst("preferred_username")?.Value
                       ?? User.FindFirst("upn")?.Value
                       ?? User.Identity?.Name ?? "unknown";

        _logger.LogInformation("Update requested by {User}: applying image {Image}", requestedBy, imageName);

        try
        {
            var credential   = new DefaultAzureCredential();
            var tokenContext = new TokenRequestContext(new[] { "https://management.azure.com/.default" });
            var token        = await credential.GetTokenAsync(tokenContext);

            var mgmtClient = _httpClientFactory.CreateClient();
            mgmtClient.DefaultRequestHeaders.Add("Authorization", $"Bearer {token.Token}");

            var appUrl = $"https://management.azure.com/subscriptions/{subscriptionId}" +
                         $"/resourceGroups/{resourceGroup}" +
                         $"/providers/Microsoft.App/containerApps/{appName}" +
                         $"?api-version=2023-05-01";

            var getResponse = await mgmtClient.GetAsync(appUrl);
            if (!getResponse.IsSuccessStatusCode)
            {
                var err = await getResponse.Content.ReadAsStringAsync();
                return StatusCode(500, new { error = "Could not read current Container App configuration", detail = err });
            }

            var appJson = await getResponse.Content.ReadAsStringAsync();
            var appRoot = JsonDocument.Parse(appJson).RootElement;

            var containerName = appRoot
                .GetProperty("properties")
                .GetProperty("template")
                .GetProperty("containers")[0]
                .GetProperty("name").GetString();

            var patchBody = new
            {
                properties = new
                {
                    template = new
                    {
                        containers = new[] { new { name = containerName, image = imageName } }
                    }
                }
            };

            token = await credential.GetTokenAsync(tokenContext);
            mgmtClient.DefaultRequestHeaders.Remove("Authorization");
            mgmtClient.DefaultRequestHeaders.Add("Authorization", $"Bearer {token.Token}");

            var patchContent  = new StringContent(JsonSerializer.Serialize(patchBody), Encoding.UTF8, "application/json");
            var patchResponse = await mgmtClient.PatchAsync(appUrl, patchContent);

            if (patchResponse.IsSuccessStatusCode)
            {
                _logger.LogInformation("Container App update initiated: {Image} by {User}", imageName, requestedBy);
                return Ok(new
                {
                    message   = $"Update to {request.Version} initiated. The app will restart in 30–60 seconds.",
                    image     = imageName,
                    appliedBy = requestedBy,
                    appliedAt = DateTime.UtcNow,
                });
            }

            var errBody = await patchResponse.Content.ReadAsStringAsync();
            _logger.LogError("Container App PATCH failed ({Status}): {Body}", patchResponse.StatusCode, errBody);
            return StatusCode(500, new { error = "Failed to update Container App", detail = errBody });
        }
        catch (CredentialUnavailableException)
        {
            return StatusCode(503, new
            {
                error = "Managed Identity is not available. Ensure this Container App has a system-assigned Managed Identity with Contributor role on the resource group."
            });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Unexpected error applying update");
            return StatusCode(500, new { error = "Unexpected error applying update", detail = ex.Message });
        }
    }

    private static string GetCurrentVersion()
    {
        try { if (System.IO.File.Exists(VersionFilePath)) return System.IO.File.ReadAllText(VersionFilePath).Trim(); }
        catch { /* ignore */ }
        return FallbackVersion;
    }

    private bool IsUpdateConfigured() =>
        !string.IsNullOrEmpty(_configuration["ContainerApp:SubscriptionId"]) &&
        !string.IsNullOrEmpty(_configuration["ContainerApp:ResourceGroup"]) &&
        !string.IsNullOrEmpty(_configuration["ContainerApp:Name"]);

    private static bool IsNewerVersion(string latest, string current)
    {
        static System.Version? Parse(string v)
        {
            v = v.TrimStart('v');
            return System.Version.TryParse(v, out var ver) ? ver : null;
        }
        var l = Parse(latest);
        var c = Parse(current);
        if (l == null || c == null) return latest != current && current == FallbackVersion;
        return l > c;
    }
}

public record ApplyUpdateRequest(string Version);

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
