using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Azure.Identity;
using System.Text;
using System.Text.Json;

namespace M365Dashboard.Api.Controllers;

[ApiController]
[Route("api/update")]
[Authorize(Policy = "RequireAdminRole")]
public class AppUpdateController : ControllerBase
{
    private readonly IConfiguration _configuration;
    private readonly ILogger<AppUpdateController> _logger;
    private readonly IWebHostEnvironment _environment;

    private const string GhcrOwner = "alex-c1";
    private const string GhcrRepo  = "m365-dashboard";
    private const string GhcrImage = "ghcr.io/alex-c1/m365-dashboard";

    public AppUpdateController(
        IConfiguration configuration,
        ILogger<AppUpdateController> logger,
        IWebHostEnvironment environment)
    {
        _configuration = configuration;
        _logger = logger;
        _environment = environment;
    }

    [HttpGet("check")]
    public async Task<IActionResult> CheckForUpdates()
    {
        try
        {
            var currentVersion = GetCurrentVersion();
            var latestRelease  = await GetLatestGhcrReleaseAsync();

            var subscriptionId   = _configuration["ContainerApp:SubscriptionId"];
            var resourceGroup    = _configuration["ContainerApp:ResourceGroup"];
            var appName          = _configuration["ContainerApp:Name"];
            var updateConfigured = !string.IsNullOrEmpty(subscriptionId)
                                && !string.IsNullOrEmpty(resourceGroup)
                                && !string.IsNullOrEmpty(appName);

            bool updateAvailable = false;
            if (latestRelease != null && currentVersion != "unknown" && currentVersion != "dev")
            {
                var current = currentVersion.TrimStart('v');
                var latest  = latestRelease.TrimStart('v');
                updateAvailable = !string.Equals(current, latest, StringComparison.OrdinalIgnoreCase)
                               && IsNewerVersion(latest, current);
            }

            return Ok(new
            {
                currentVersion,
                latestVersion    = latestRelease,
                updateAvailable,
                updateConfigured,
                ghcrImage        = GhcrImage,
                checkedAt        = DateTime.UtcNow
            });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error checking for updates");
            return StatusCode(500, new { error = "Failed to check for updates", message = ex.Message });
        }
    }

    [HttpPost("apply")]
    public async Task<IActionResult> ApplyUpdate([FromBody] AppUpdateRequest request)
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
            var version  = request.Version.TrimStart('v');
            var imageTag = $"{GhcrImage}:{version}";

            _logger.LogInformation("Applying update to {Version} via image {Image}", version, imageTag);

            var credential  = new DefaultAzureCredential();
            var tokenResult = await credential.GetTokenAsync(
                new Azure.Core.TokenRequestContext(new[] { "https://management.azure.com/.default" }),
                CancellationToken.None);

            using var http = new HttpClient();
            http.DefaultRequestHeaders.Authorization =
                new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", tokenResult.Token);

            var apiBase    = $"https://management.azure.com/subscriptions/{subscriptionId}" +
                             $"/resourceGroups/{resourceGroup}/providers/Microsoft.App" +
                             $"/containerApps/{appName}";
            var apiVersion = "?api-version=2023-05-01";

            var getResp = await http.GetAsync(apiBase + apiVersion);
            if (!getResp.IsSuccessStatusCode)
            {
                var body = await getResp.Content.ReadAsStringAsync();
                _logger.LogError("Failed to GET Container App: {Status} {Body}", getResp.StatusCode, body);
                return StatusCode(500, new { error = "Could not read Container App configuration", detail = body });
            }

            var appDoc    = JsonDocument.Parse(await getResp.Content.ReadAsStringAsync());
            var patchBody = BuildImagePatchBody(appDoc.RootElement, imageTag);

            var patchResp = await http.PatchAsync(apiBase + apiVersion,
                new StringContent(JsonSerializer.Serialize(patchBody), Encoding.UTF8, "application/json"));

            if (!patchResp.IsSuccessStatusCode)
            {
                var body = await patchResp.Content.ReadAsStringAsync();
                _logger.LogError("Failed to PATCH Container App: {Status} {Body}", patchResp.StatusCode, body);
                return StatusCode(500, new { error = "Failed to update Container App image", detail = body });
            }

            _logger.LogInformation("Container App update initiated to {Version}", version);
            return Ok(new
            {
                message   = $"Update to v{version} initiated. The dashboard will restart and be available again in about 60 seconds.",
                version,
                image     = imageTag,
                appliedAt = DateTime.UtcNow
            });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error applying update to {Version}", request.Version);
            return StatusCode(500, new { error = "Failed to apply update", message = ex.Message });
        }
    }

    private string GetCurrentVersion()
    {
        try
        {
            var path = Path.Combine(_environment.ContentRootPath, "version.txt");
            if (System.IO.File.Exists(path))
            {
                var v = System.IO.File.ReadAllText(path).Trim();
                if (!string.IsNullOrEmpty(v)) return v;
            }
        }
        catch (Exception ex) { _logger.LogWarning(ex, "Could not read version.txt"); }
        return "unknown";
    }

    private async Task<string?> GetLatestGhcrReleaseAsync()
    {
        try
        {
            using var http = new HttpClient();
            http.DefaultRequestHeaders.Add("User-Agent", "M365Dashboard-UpdateCheck/1.0");
            http.DefaultRequestHeaders.Add("Accept", "application/vnd.github+json");

            var resp = await http.GetAsync($"https://api.github.com/repos/{GhcrOwner}/{GhcrRepo}/releases/latest");
            if (!resp.IsSuccessStatusCode) return null;

            var doc = JsonDocument.Parse(await resp.Content.ReadAsStringAsync());
            if (doc.RootElement.TryGetProperty("tag_name", out var tag))
                return tag.GetString()?.TrimStart('v');
        }
        catch (Exception ex) { _logger.LogWarning(ex, "Could not fetch latest release from GitHub"); }
        return null;
    }

    private static object BuildImagePatchBody(JsonElement existing, string newImage)
    {
        var location   = existing.TryGetProperty("location", out var loc) ? loc.GetString() : "uksouth";
        var properties = existing.GetProperty("properties");
        var config     = properties.GetProperty("configuration");
        var template   = properties.GetProperty("template");

        var patchedContainers = template.GetProperty("containers").EnumerateArray()
            .Select((c, i) => i == 0
                ? (object)new
                {
                    name      = c.TryGetProperty("name", out var n) ? n.GetString() : "m365dashboard",
                    image     = newImage,
                    resources = c.TryGetProperty("resources", out var r)
                        ? (object)JsonSerializer.Deserialize<object>(r.GetRawText())!
                        : new { cpu = 0.5, memory = "1Gi" },
                    env = c.TryGetProperty("env", out var e)
                        ? (object)JsonSerializer.Deserialize<object>(e.GetRawText())!
                        : new object[] { }
                }
                : (object)JsonSerializer.Deserialize<object>(c.GetRawText())!)
            .ToList();

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
        try { return ParseVersion(latest) > ParseVersion(current); }
        catch { return !string.Equals(latest, current, StringComparison.OrdinalIgnoreCase); }
    }

    private static Version ParseVersion(string v)
    {
        var clean = v.Split('-')[0];
        return Version.TryParse(clean, out var parsed) ? parsed : new Version(0, 0, 0);
    }
}

public record AppUpdateRequest(string Version);
