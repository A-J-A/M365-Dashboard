using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Azure.Identity;
using System.Text;
using System.Text.Json;

namespace M365Dashboard.Api.Controllers;

/// <summary>
/// Handles application self-update via GHCR image pull + Container App revision update.
/// 
/// Flow:
///   1. GET /api/update/check  — compares current version against latest GHCR release tag
///   2. POST /api/update/apply — pulls the new image into the Container App via Azure REST API
///
/// The Container App's managed identity must have the "Contributor" role on the Container App
/// resource (or at minimum the "Azure Container Apps Contributor" role) for apply to work.
/// This is granted automatically by the deploy script.
/// </summary>
[ApiController]
[Route("api/[controller]")]
[Authorize(Policy = "RequireAdminRole")]
public class UpdateController : ControllerBase
{
    private readonly IConfiguration _configuration;
    private readonly ILogger<UpdateController> _logger;
    private readonly IWebHostEnvironment _environment;

    // GHCR image — published by GitHub Actions on every tagged release
    private const string GhcrOwner = "Alex-C1";
    private const string GhcrRepo  = "m365-dashboard";
    private const string GhcrImage = "ghcr.io/alex-c1/m365-dashboard";

    public UpdateController(
        IConfiguration configuration,
        ILogger<UpdateController> logger,
        IWebHostEnvironment environment)
    {
        _configuration = configuration;
        _logger = logger;
        _environment = environment;
    }

    // ── GET /api/update/check ─────────────────────────────────────────────

    [HttpGet("check")]
    public async Task<IActionResult> CheckForUpdates()
    {
        try
        {
            var currentVersion = GetCurrentVersion();
            var latestRelease  = await GetLatestGhcrReleaseAsync();

            // updateConfigured = we have enough config to actually apply an update
            var subscriptionId = _configuration["ContainerApp:SubscriptionId"];
            var resourceGroup  = _configuration["ContainerApp:ResourceGroup"];
            var appName        = _configuration["ContainerApp:Name"];
            var updateConfigured = !string.IsNullOrEmpty(subscriptionId)
                                && !string.IsNullOrEmpty(resourceGroup)
                                && !string.IsNullOrEmpty(appName);

            bool updateAvailable = false;
            if (latestRelease != null && currentVersion != "unknown" && currentVersion != "dev")
            {
                // Strip leading 'v' for comparison
                var current = currentVersion.TrimStart('v');
                var latest  = latestRelease.TrimStart('v');
                updateAvailable = !string.Equals(current, latest, StringComparison.OrdinalIgnoreCase)
                               && IsNewerVersion(latest, current);
            }

            return Ok(new
            {
                currentVersion,
                latestVersion      = latestRelease,
                updateAvailable,
                updateConfigured,
                ghcrImage          = GhcrImage,
                checkedAt          = DateTime.UtcNow
            });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error checking for updates");
            return StatusCode(500, new { error = "Failed to check for updates", message = ex.Message });
        }
    }

    // ── POST /api/update/apply ────────────────────────────────────────────

    [HttpPost("apply")]
    public async Task<IActionResult> ApplyUpdate([FromBody] ApplyUpdateRequest request)
    {
        if (string.IsNullOrWhiteSpace(request.Version))
            return BadRequest(new { error = "Version is required" });

        var subscriptionId = _configuration["ContainerApp:SubscriptionId"];
        var resourceGroup  = _configuration["ContainerApp:ResourceGroup"];
        var appName        = _configuration["ContainerApp:Name"];

        if (string.IsNullOrEmpty(subscriptionId) || string.IsNullOrEmpty(resourceGroup) || string.IsNullOrEmpty(appName))
            return BadRequest(new { error = "Container App configuration is not set. Re-run the deployment script to configure auto-updates." });

        try
        {
            var version    = request.Version.TrimStart('v');
            var imageTag   = $"{GhcrImage}:{version}";

            _logger.LogInformation("Applying update to {Version} via image {Image}", version, imageTag);

            // Use the Container App's managed identity to call the Azure Container Apps REST API
            var credential = new DefaultAzureCredential();
            var tokenResult = await credential.GetTokenAsync(
                new Azure.Core.TokenRequestContext(new[] { "https://management.azure.com/.default" }),
                CancellationToken.None);

            // GET current Container App to preserve its template
            using var http = new HttpClient();
            http.DefaultRequestHeaders.Authorization =
                new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", tokenResult.Token);

            var apiBase = $"https://management.azure.com/subscriptions/{subscriptionId}" +
                          $"/resourceGroups/{resourceGroup}/providers/Microsoft.App" +
                          $"/containerApps/{appName}";
            var apiVersion = "?api-version=2023-05-01";

            var getResp = await http.GetAsync(apiBase + apiVersion);
            if (!getResp.IsSuccessStatusCode)
            {
                var errBody = await getResp.Content.ReadAsStringAsync();
                _logger.LogError("Failed to GET Container App: {Status} {Body}", getResp.StatusCode, errBody);
                return StatusCode(500, new { error = "Could not read Container App configuration", detail = errBody });
            }

            var appJson  = await getResp.Content.ReadAsStringAsync();
            var appDoc   = JsonDocument.Parse(appJson);
            var appRoot  = appDoc.RootElement;

            // Patch just the container image — preserve everything else
            // We update the first container in the template (which is always m365dashboard)
            var patchBody = BuildImagePatchBody(appRoot, imageTag);

            var patchContent = new StringContent(
                JsonSerializer.Serialize(patchBody),
                Encoding.UTF8,
                "application/json");

            var patchResp = await http.PatchAsync(apiBase + apiVersion, patchContent);
            if (!patchResp.IsSuccessStatusCode)
            {
                var errBody = await patchResp.Content.ReadAsStringAsync();
                _logger.LogError("Failed to PATCH Container App: {Status} {Body}", patchResp.StatusCode, errBody);
                return StatusCode(500, new { error = "Failed to update Container App image", detail = errBody });
            }

            _logger.LogInformation("Container App update initiated successfully to {Version}", version);

            return Ok(new
            {
                message  = $"Update to v{version} initiated. The dashboard will restart and be available again in about 60 seconds.",
                version,
                image    = imageTag,
                appliedAt = DateTime.UtcNow
            });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error applying update to {Version}", request.Version);
            return StatusCode(500, new { error = "Failed to apply update", message = ex.Message });
        }
    }

    // ── Helpers ───────────────────────────────────────────────────────────

    private string GetCurrentVersion()
    {
        // Version is written to version.txt by the Dockerfile at build time from BUILD_VERSION arg
        try
        {
            var versionFile = Path.Combine(_environment.ContentRootPath, "version.txt");
            if (System.IO.File.Exists(versionFile))
            {
                var v = System.IO.File.ReadAllText(versionFile).Trim();
                if (!string.IsNullOrEmpty(v)) return v;
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Could not read version.txt");
        }
        return "unknown";
    }

    private async Task<string?> GetLatestGhcrReleaseAsync()
    {
        try
        {
            // GitHub Releases API — no auth needed for public repos
            using var http = new HttpClient();
            http.DefaultRequestHeaders.Add("User-Agent", "M365Dashboard-UpdateCheck/1.0");
            http.DefaultRequestHeaders.Add("Accept", "application/vnd.github+json");

            var url  = $"https://api.github.com/repos/{GhcrOwner}/{GhcrRepo}/releases/latest";
            var resp = await http.GetAsync(url);

            if (!resp.IsSuccessStatusCode)
            {
                _logger.LogWarning("GitHub releases API returned {Status}", resp.StatusCode);
                return null;
            }

            var json = await resp.Content.ReadAsStringAsync();
            var doc  = JsonDocument.Parse(json);

            if (doc.RootElement.TryGetProperty("tag_name", out var tag))
                return tag.GetString()?.TrimStart('v');

            return null;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Could not fetch latest release from GitHub");
            return null;
        }
    }

    /// <summary>
    /// Builds a minimal PATCH body that only changes the container image,
    /// preserving the existing Container App configuration.
    /// </summary>
    private static object BuildImagePatchBody(JsonElement existing, string newImage)
    {
        // Extract existing properties we need to preserve
        var location   = existing.TryGetProperty("location", out var loc) ? loc.GetString() : "uksouth";
        var properties = existing.GetProperty("properties");
        var config     = properties.GetProperty("configuration");
        var template   = properties.GetProperty("template");

        // Clone containers array, replacing only the image on the first container
        var containers = template.GetProperty("containers").EnumerateArray().ToList();
        var patchedContainers = containers.Select((c, i) =>
        {
            if (i == 0)
            {
                // Replace image, keep everything else
                return new
                {
                    name  = c.TryGetProperty("name", out var n) ? n.GetString() : "m365dashboard",
                    image = newImage,
                    resources = c.TryGetProperty("resources", out var r)
                        ? (object)JsonSerializer.Deserialize<object>(r.GetRawText())!
                        : new { cpu = 0.5, memory = "1Gi" },
                    env = c.TryGetProperty("env", out var e)
                        ? (object)JsonSerializer.Deserialize<object>(e.GetRawText())!
                        : new object[] { }
                };
            }
            return (object)JsonSerializer.Deserialize<object>(c.GetRawText())!;
        }).ToList();

        return new
        {
            location,
            properties = new
            {
                configuration = JsonSerializer.Deserialize<object>(config.GetRawText()),
                template = new
                {
                    containers = patchedContainers,
                    scale = template.TryGetProperty("scale", out var sc)
                        ? (object)JsonSerializer.Deserialize<object>(sc.GetRawText())!
                        : null
                }
            }
        };
    }

    private static bool IsNewerVersion(string latest, string current)
    {
        // Simple semver comparison — handles v1.2.3 style tags
        // Returns true if latest > current
        try
        {
            var l = ParseVersion(latest);
            var c = ParseVersion(current);
            return l > c;
        }
        catch
        {
            // If parsing fails, assume update is available if versions differ
            return !string.Equals(latest, current, StringComparison.OrdinalIgnoreCase);
        }
    }

    private static Version ParseVersion(string v)
    {
        // Strip pre-release suffixes like -dev-2026.01.01-abc1234
        var clean = v.Split('-')[0];
        // Handle date-style versions like 2026.01.01
        if (Version.TryParse(clean, out var parsed))
            return parsed;
        return new Version(0, 0, 0);
    }
}

public record ApplyUpdateRequest(string Version);
