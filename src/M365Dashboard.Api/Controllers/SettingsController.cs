using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using M365Dashboard.Api.Models;
using M365Dashboard.Api.Services;
using System.Text.Json;

namespace M365Dashboard.Api.Controllers;

[ApiController]
[Route("api/[controller]")]
[Authorize]
public class SettingsController : ControllerBase
{
    private readonly ILogger<SettingsController> _logger;
    private readonly IWebHostEnvironment _environment;
    private readonly ITenantSettingsService _tenantSettingsService;
    private readonly IConfiguration _configuration;
    private static readonly string SettingsFileName = "report-settings.json";

    public SettingsController(
        ILogger<SettingsController> logger,
        IWebHostEnvironment environment,
        ITenantSettingsService tenantSettingsService,
        IConfiguration configuration)
    {
        _logger = logger;
        _environment = environment;
        _tenantSettingsService = tenantSettingsService;
        _configuration = configuration;
    }

    private string GetTenantId() =>
        _configuration["AzureAd:TenantId"] ?? "default";

    private string GetSettingsFilePath()
    {
        var dataPath = Path.Combine(_environment.ContentRootPath, "App_Data");
        if (!Directory.Exists(dataPath))
        {
            Directory.CreateDirectory(dataPath);
        }
        return Path.Combine(dataPath, SettingsFileName);
    }

    /// <summary>
    /// Get current report settings
    /// </summary>
    [HttpGet("report")]
    public IActionResult GetReportSettings()
    {
        try
        {
            var filePath = GetSettingsFilePath();
            
            if (!System.IO.File.Exists(filePath))
            {
                return Ok(new ReportSettings());
            }

            var json = System.IO.File.ReadAllText(filePath);
            var settings = JsonSerializer.Deserialize<ReportSettings>(json, new JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true
            });

            return Ok(settings ?? new ReportSettings());
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error loading report settings");
            return Ok(new ReportSettings());
        }
    }

    /// <summary>
    /// Save report settings
    /// </summary>
    [HttpPost("report")]
    public IActionResult SaveReportSettings([FromBody] ReportSettings settings)
    {
        try
        {
            settings.UpdatedAt = DateTime.UtcNow;
            
            var filePath = GetSettingsFilePath();
            var json = JsonSerializer.Serialize(settings, new JsonSerializerOptions
            {
                WriteIndented = true
            });
            
            System.IO.File.WriteAllText(filePath, json);
            
            _logger.LogInformation("Report settings saved successfully");
            return Ok(settings);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error saving report settings");
            return StatusCode(500, new { error = "Failed to save settings", message = ex.Message });
        }
    }

    /// <summary>
    /// Upload logo for reports
    /// </summary>
    [HttpPost("report/logo")]
    public async Task<IActionResult> UploadLogo(IFormFile file)
    {
        try
        {
            if (file == null || file.Length == 0)
            {
                return BadRequest(new { error = "No file provided" });
            }

            // Validate file type
            var allowedTypes = new[] { "image/png", "image/jpeg", "image/gif", "image/svg+xml" };
            if (!allowedTypes.Contains(file.ContentType.ToLower()))
            {
                return BadRequest(new { error = "Invalid file type. Please upload PNG, JPEG, GIF, or SVG." });
            }

            // Validate file size (max 2MB)
            if (file.Length > 2 * 1024 * 1024)
            {
                return BadRequest(new { error = "File too large. Maximum size is 2MB." });
            }

            // Read file to base64
            using var memoryStream = new MemoryStream();
            await file.CopyToAsync(memoryStream);
            var base64 = Convert.ToBase64String(memoryStream.ToArray());

            // Load existing settings
            var filePath = GetSettingsFilePath();
            ReportSettings settings;
            
            if (System.IO.File.Exists(filePath))
            {
                var json = System.IO.File.ReadAllText(filePath);
                settings = JsonSerializer.Deserialize<ReportSettings>(json, new JsonSerializerOptions
                {
                    PropertyNameCaseInsensitive = true
                }) ?? new ReportSettings();
            }
            else
            {
                settings = new ReportSettings();
            }

            // Update logo
            settings.LogoBase64 = base64;
            settings.LogoContentType = file.ContentType;
            settings.UpdatedAt = DateTime.UtcNow;

            // Save settings
            var updatedJson = JsonSerializer.Serialize(settings, new JsonSerializerOptions
            {
                WriteIndented = true
            });
            System.IO.File.WriteAllText(filePath, updatedJson);

            _logger.LogInformation("Logo uploaded successfully");
            return Ok(new { 
                message = "Logo uploaded successfully",
                logoBase64 = base64,
                contentType = file.ContentType
            });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error uploading logo");
            return StatusCode(500, new { error = "Failed to upload logo", message = ex.Message });
        }
    }

    /// <summary>
    /// Remove logo from reports
    /// </summary>
    [HttpDelete("report/logo")]
    public IActionResult RemoveLogo()
    {
        try
        {
            var filePath = GetSettingsFilePath();
            
            if (!System.IO.File.Exists(filePath))
            {
                return Ok(new { message = "No logo to remove" });
            }

            var json = System.IO.File.ReadAllText(filePath);
            var settings = JsonSerializer.Deserialize<ReportSettings>(json, new JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true
            }) ?? new ReportSettings();

            settings.LogoBase64 = null;
            settings.LogoContentType = null;
            settings.UpdatedAt = DateTime.UtcNow;

            var updatedJson = JsonSerializer.Serialize(settings, new JsonSerializerOptions
            {
                WriteIndented = true
            });
            System.IO.File.WriteAllText(filePath, updatedJson);

            _logger.LogInformation("Logo removed successfully");
            return Ok(new { message = "Logo removed successfully" });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error removing logo");
            return StatusCode(500, new { error = "Failed to remove logo", message = ex.Message });
        }
    }

    /// <summary>
    /// Fetch the banner logo from Entra organizational branding and import it as the report logo
    /// </summary>
    [HttpPost("report/logo/entra")]
    public async Task<IActionResult> ImportEntraBrandingLogo(
        [FromServices] Microsoft.Graph.GraphServiceClient graphClient,
        [FromServices] IConfiguration configuration)
    {
        try
        {
            // Get the tenant/org ID
            var tenantId = configuration["AzureAd:TenantId"];

            // Fetch the banner logo stream from Entra branding
            // Uses localizations/0 which is the default locale
            byte[]? logoBytes = null;
            string contentType = "image/png";

            try
            {
                var orgs = await graphClient.Organization.GetAsync();
                var orgId = orgs?.Value?.FirstOrDefault()?.Id ?? tenantId;

                var logoStream = await graphClient.Organization[orgId]
                    .Branding.Localizations["0"]
                    .BannerLogo
                    .GetAsync();

                if (logoStream != null)
                {
                    using var ms = new MemoryStream();
                    await logoStream.CopyToAsync(ms);
                    logoBytes = ms.ToArray();

                    // Try to detect content type from magic bytes
                    if (logoBytes.Length >= 4)
                    {
                        if (logoBytes[0] == 0x89 && logoBytes[1] == 0x50) contentType = "image/png";
                        else if (logoBytes[0] == 0xFF && logoBytes[1] == 0xD8) contentType = "image/jpeg";
                        else if (logoBytes[0] == 0x47 && logoBytes[1] == 0x49) contentType = "image/gif";
                        else if (logoBytes[0] == 0x3C) contentType = "image/svg+xml";
                    }
                }
            }
            catch (Microsoft.Graph.Models.ODataErrors.ODataError ex) when (ex.ResponseStatusCode == 404)
            {
                return NotFound(new { error = "No Entra branding configured for this tenant", message = "Set up organisational branding in the Entra admin centre first." });
            }

            if (logoBytes == null || logoBytes.Length == 0)
            {
                return NotFound(new { error = "No banner logo found in Entra branding", message = "Upload a banner logo in Entra admin centre under Company Branding." });
            }

            // Save it as the report logo
            var filePath = GetSettingsFilePath();
            ReportSettings settings;
            if (System.IO.File.Exists(filePath))
            {
                var json = System.IO.File.ReadAllText(filePath);
                settings = JsonSerializer.Deserialize<ReportSettings>(json) ?? new ReportSettings();
            }
            else
            {
                settings = new ReportSettings();
            }

            settings.LogoBase64 = Convert.ToBase64String(logoBytes);
            settings.LogoContentType = contentType;
            settings.UpdatedAt = DateTime.UtcNow;

            System.IO.File.WriteAllText(filePath, JsonSerializer.Serialize(settings, new JsonSerializerOptions { WriteIndented = true }));

            _logger.LogInformation("Entra branding logo imported successfully ({Bytes} bytes, {Type})", logoBytes.Length, contentType);

            return Ok(new
            {
                logoBase64 = settings.LogoBase64,
                contentType = settings.LogoContentType,
                message = "Entra branding logo imported successfully"
            });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error importing Entra branding logo");
            return StatusCode(500, new { error = "Failed to import Entra branding logo", message = ex.Message });
        }
    }

    // -------------------------------------------------------------------------
    // Break Glass Accounts — stored in SQL via TenantSettingsService
    // -------------------------------------------------------------------------

    /// <summary>
    /// Get break glass account settings from SQL, resolving each UPN against the directory
    /// </summary>
    [HttpGet("breakglass")]
    public async Task<IActionResult> GetBreakGlassSettings()
    {
        try
        {
            var result = await _tenantSettingsService.GetBreakGlassSettingsAsync(GetTenantId());
            return Ok(result);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error loading break glass settings");
            return StatusCode(500, new { error = "Failed to load break glass settings" });
        }
    }

    /// <summary>
    /// Save break glass accounts to SQL
    /// </summary>
    [HttpPut("breakglass")]
    public async Task<IActionResult> SaveBreakGlassSettings(
        [FromBody] M365Dashboard.Api.Models.Dtos.UpdateBreakGlassSettingsRequest request)
    {
        try
        {
            var currentUser = User.FindFirst("preferred_username")?.Value
                           ?? User.FindFirst("upn")?.Value
                           ?? User.Identity?.Name
                           ?? "unknown";

            var result = await _tenantSettingsService.UpdateBreakGlassSettingsAsync(
                GetTenantId(),
                request.UserPrincipalNames,
                currentUser);

            return Ok(result);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error saving break glass settings");
            return StatusCode(500, new { error = "Failed to save break glass accounts", message = ex.Message });
        }
    }
}
