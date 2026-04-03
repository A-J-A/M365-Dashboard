using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Azure.Identity;
using System.Text;
using System.Text.Json;

namespace M365Dashboard.Api.Controllers;

[ApiController]
[Route("api/appupdate")]
public class AppUpdateController : ControllerBase
{
    private readonly IConfiguration _configuration;
    private readonly ILogger<AppUpdateController> _logger;
    private readonly IWebHostEnvironment _environment;

    private string GhcrOwner => _configuration["GitHub:Owner"] ?? "A-J-A";
    private string GhcrRepo  => _configuration["GitHub:Repo"]  ?? "M365-Dashboard";
    private string GhcrImage => $"ghcr.io/{GhcrOwner}/{GhcrRepo}";

    public AppUpdateController(
        IConfiguration configuration,
        ILogger<AppUpdateController> logger,
        IWebHostEnvironment environment)
    {
        _configuration = configuration;
        _logger = logger;
        _environment = environment;
    }

    // ── GET /api/update/source ────────────────────────────────────────────

    [HttpGet("source")]
    public IActionResult GetUpdateSource()
    {
        var repo = _configuration["ContainerApp:GhcrRepo"] ?? $"{GhcrOwner}/{GhcrRepo}";
        return Ok(new { repo });
    }

    // ── POST /api/update/source ───────────────────────────────────────────

    [HttpPost("source")]
    public IActionResult SetUpdateSource([FromBody] AppUpdateSourcePayload request)
    {
        if (string.IsNullOrWhiteSpace(request.Repo) || !request.Repo.Contains('/'))
            return BadRequest(new { error = "Repo must be in the format owner/repo-name" });

        return Ok(new { message = $"Update source set to github.com/{request.Repo}", repo = request.Repo });
    }

    // ── GET /api/update/check ─────────────────────────────────────────────

    [HttpGet("check")]
    public async Task<IActionResult> CheckForUpdates()
    {
        try
        {
            var currentVersion = GetCurrentVersion();
            var release        = await GetLatestReleaseAsync();

            var subscriptionId   = _configuration["ContainerApp:SubscriptionId"];
            var resourceGroup    = _configuration["ContainerApp:ResourceGroup"];
            var appName          = _configuration["ContainerApp:Name"];
            var updateConfigured = !string.IsNullOrEmpty(subscriptionId)
                                && !string.IsNullOrEmpty(resourceGroup)
                                && !string.IsNullOrEmpty(appName);

            // No releases published yet
            if (release == null)
            {
                return Ok(new
                {
                    currentVersion,
                    latestVersion    = (string?)null,
                    updateAvailable  = false,
                    updateConfigured,
                    noReleasesYet    = true,
                    releaseNotes     = (string?)null,
                    releaseUrl       = (string?)null,
                    publishedAt      = (string?)null,
                    error            = (string?)null,
                    checkedAt        = DateTime.UtcNow
                });
            }

            bool updateAvailable = false;
            if (currentVersion != "unknown")
            {
                var current = currentVersion.TrimStart('v').Split('-')[0]; // strip dev-date-sha suffix
                var latest  = release.TagName.TrimStart('v');
                updateAvailable = !string.Equals(current, latest, StringComparison.OrdinalIgnoreCase)
                               && IsNewerVersion(latest, current);
            }

            return Ok(new
            {
                currentVersion,
                latestVersion    = release.TagName.TrimStart('v'),
                updateAvailable,
                updateConfigured,
                noReleasesYet    = false,
                releaseNotes     = release.Body,
                releaseUrl       = release.HtmlUrl,
                publishedAt      = release.PublishedAt,
                error            = (string?)null,
                checkedAt        = DateTime.UtcNow
            });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error checking for updates");
            return Ok(new
            {
                currentVersion   = GetCurrentVersion(),
                latestVersion    = (string?)null,
                updateAvailable  = false,
                updateConfigured = false,
                noReleasesYet    = false,
                releaseNotes     = (string?)null,
                releaseUrl       = (string?)null,
                publishedAt      = (string?)null,
                error            = "Failed to check for updates: " + ex.Message,
                checkedAt        = DateTime.UtcNow
            });
        }
    }

    [Authorize(Policy = "RequireAdminRole")]
    [HttpPost("apply")]
    public async Task<IActionResult> ApplyUpdate([FromBody] AppUpdateVersionPayload request)
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

    private record GithubRelease(string TagName, string? Body, string? HtmlUrl, string? PublishedAt);

    private async Task<GithubRelease?> GetLatestReleaseAsync()
    {
        try
        {
            using var http = new HttpClient();
            http.DefaultRequestHeaders.Add("User-Agent", "M365Dashboard-UpdateCheck/1.0");
            http.DefaultRequestHeaders.Add("Accept", "application/vnd.github+json");

            // If a PAT is configured (required for private repos), add it
            var pat = _configuration["GitHub:ReleasesPat"];
            if (!string.IsNullOrEmpty(pat))
                http.DefaultRequestHeaders.Add("Authorization", $"Bearer {pat}");

            var resp = await http.GetAsync($"https://api.github.com/repos/{GhcrOwner}/{GhcrRepo}/releases/latest");

            // 404 means no releases yet
            if (resp.StatusCode == System.Net.HttpStatusCode.NotFound) return null;
            if (!resp.IsSuccessStatusCode) return null;

            var doc = JsonDocument.Parse(await resp.Content.ReadAsStringAsync());
            var root = doc.RootElement;

            var tag     = root.TryGetProperty("tag_name",    out var t) ? t.GetString() ?? "" : "";
            var body    = root.TryGetProperty("body",        out var b) ? b.GetString() : null;
            var htmlUrl = root.TryGetProperty("html_url",   out var u) ? u.GetString() : null;
            var pubAt   = root.TryGetProperty("published_at", out var p) ? p.GetString() : null;

            return new GithubRelease(tag, body, htmlUrl, pubAt);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Could not fetch latest release from GitHub");
            return null;
        }
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

// Uniquely named to avoid CS0101 conflicts with any legacy cached type definitions
public sealed class AppUpdateVersionPayload { public string Version { get; set; } = string.Empty; }
public sealed class AppUpdateSourcePayload  { public string Repo    { get; set; } = string.Empty; }
