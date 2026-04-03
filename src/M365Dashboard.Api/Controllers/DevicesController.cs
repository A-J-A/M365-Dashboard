using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Graph;
using M365Dashboard.Api.Services;

namespace M365Dashboard.Api.Controllers;

[ApiController]
[Route("api/[controller]")]
[Authorize]
public class DevicesController : ControllerBase
{
    private readonly IGraphService _graphService;
    private readonly GraphServiceClient _graphClient;
    private readonly ILogger<DevicesController> _logger;

    public DevicesController(IGraphService graphService, GraphServiceClient graphClient, ILogger<DevicesController> logger)
    {
        _graphService = graphService;
        _graphClient = graphClient;
        _logger = logger;
    }

    /// <summary>
    /// Get all Intune managed devices
    /// </summary>
    [HttpGet]
    public async Task<IActionResult> GetDevices(
        [FromQuery] string? filter = null,
        [FromQuery] string? orderBy = "deviceName",
        [FromQuery] bool ascending = true,
        [FromQuery] int take = 200)
    {
        try
        {
            var result = await _graphService.GetDevicesAsync(filter, orderBy, ascending, take);
            return Ok(result);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching devices");
            return StatusCode(500, new { error = "Failed to fetch devices", message = ex.Message });
        }
    }

    /// <summary>
    /// Get device statistics
    /// </summary>
    [HttpGet("stats")]
    public async Task<IActionResult> GetDeviceStats()
    {
        try
        {
            var stats = await _graphService.GetDeviceStatsAsync();
            return Ok(stats);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching device statistics");
            return StatusCode(500, new { error = "Failed to fetch device statistics", message = ex.Message });
        }
    }

    /// <summary>
    /// Get Apple Push Notification certificate information
    /// </summary>
    [HttpGet("apple-push-certificate")]
    public async Task<IActionResult> GetApplePushCertificate()
    {
        try
        {
            _logger.LogInformation("Fetching Apple Push Notification certificate");
            
            var certificate = await _graphClient.DeviceManagement.ApplePushNotificationCertificate
                .GetAsync();

            if (certificate == null)
            {
                _logger.LogInformation("Apple Push certificate not configured");
                return Ok(new
                {
                    isConfigured = false,
                    message = "Apple Push Notification certificate is not configured",
                    lastUpdated = DateTime.UtcNow
                });
            }

            var expirationDateTime = certificate.ExpirationDateTime;
            var now = DateTimeOffset.UtcNow;
            var daysUntilExpiry = expirationDateTime.HasValue 
                ? (int)(expirationDateTime.Value - now).TotalDays 
                : 0;

            string status;
            if (!expirationDateTime.HasValue)
            {
                status = "Unknown";
            }
            else if (daysUntilExpiry < 0)
            {
                status = "Expired";
            }
            else if (daysUntilExpiry <= 30)
            {
                status = "Critical";
            }
            else if (daysUntilExpiry <= 60)
            {
                status = "Warning";
            }
            else
            {
                status = "Healthy";
            }

            _logger.LogInformation("Apple Push certificate found, expires in {Days} days, status: {Status}", daysUntilExpiry, status);

            return Ok(new
            {
                isConfigured = true,
                appleIdentifier = certificate.AppleIdentifier,
                topicIdentifier = certificate.TopicIdentifier,
                expirationDateTime = certificate.ExpirationDateTime,
                lastModifiedDateTime = certificate.LastModifiedDateTime,
                certificateSerialNumber = certificate.CertificateSerialNumber,
                certificateUploadStatus = certificate.CertificateUploadStatus,
                daysUntilExpiry,
                status,
                lastUpdated = DateTime.UtcNow
            });
        }
        catch (Microsoft.Graph.Models.ODataErrors.ODataError ex) when (ex.ResponseStatusCode == 404)
        {
            _logger.LogInformation("Apple Push certificate not found (404)");
            return Ok(new
            {
                isConfigured = false,
                message = "Apple Push Notification certificate is not configured",
                lastUpdated = DateTime.UtcNow
            });
        }
        catch (Microsoft.Graph.Models.ODataErrors.ODataError ex) when (ex.ResponseStatusCode == 403)
        {
            _logger.LogWarning("Insufficient permissions to read Apple Push certificate: {Message}", ex.Message);
            return Ok(new
            {
                isConfigured = false,
                permissionRequired = true,
                error = "Missing permission: DeviceManagementServiceConfig.Read.All. Please add this permission in Azure AD and grant admin consent.",
                lastUpdated = DateTime.UtcNow
            });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching Apple Push certificate");
            
            // Check if it's a permission error
            if (ex.Message.Contains("DeviceManagementServiceConfig") || ex.Message.Contains("not authorized"))
            {
                return Ok(new
                {
                    isConfigured = false,
                    permissionRequired = true,
                    error = "Missing permission: DeviceManagementServiceConfig.Read.All. Please add this permission in Azure AD and grant admin consent.",
                    lastUpdated = DateTime.UtcNow
                });
            }
            
            return Ok(new
            {
                isConfigured = false,
                error = ex.Message,
                lastUpdated = DateTime.UtcNow
            });
        }
    }

    /// <summary>
    /// Get detailed information about a specific device
    /// </summary>
    [HttpGet("details/{deviceId}")]
    public async Task<IActionResult> GetDeviceDetails(string deviceId)
    {
        try
        {
            var details = await _graphService.GetDeviceDetailsAsync(deviceId);
            return Ok(details);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching device details for {DeviceId}", deviceId);
            return StatusCode(500, new { error = "Failed to fetch device details", message = ex.Message });
        }
    }

    /// <summary>
    /// Get Apple Enrollment Program (DEP/ADE) tokens
    /// </summary>
    [HttpGet("dep-tokens")]
    public async Task<IActionResult> GetDepTokens()
    {
        try
        {
            _logger.LogInformation("Fetching Apple Enrollment Program tokens");
            
            // DEP onboarding settings is a beta API only
            // Create a beta Graph client using the same credentials
            var config = HttpContext.RequestServices.GetRequiredService<IConfiguration>();
            var tenantId = config["AzureAd:TenantId"];
            var clientId = config["AzureAd:ClientId"];
            var clientSecret = config["AzureAd:ClientSecret"];
            
            var credential = new Azure.Identity.ClientSecretCredential(tenantId, clientId, clientSecret);
            var betaClient = new GraphServiceClient(credential, new[] { "https://graph.microsoft.com/.default" }, "https://graph.microsoft.com/beta");
            
            // Use the beta client's request adapter
            var requestInfo = new Microsoft.Kiota.Abstractions.RequestInformation
            {
                HttpMethod = Microsoft.Kiota.Abstractions.Method.GET,
                URI = new Uri("https://graph.microsoft.com/beta/deviceManagement/depOnboardingSettings")
            };
            
            var response = await betaClient.RequestAdapter.SendPrimitiveAsync<System.IO.Stream>(requestInfo);
            
            if (response == null)
            {
                _logger.LogInformation("No DEP tokens configured - null response");
                return Ok(new
                {
                    tokens = Array.Empty<object>(),
                    totalCount = 0,
                    message = "Apple Enrollment Program is not configured",
                    lastUpdated = DateTime.UtcNow
                });
            }
            
            using var reader = new System.IO.StreamReader(response);
            var json = await reader.ReadToEndAsync();
            var jsonDoc = System.Text.Json.JsonDocument.Parse(json);
            
            var tokenList = new List<object>();
            var now = DateTimeOffset.UtcNow;

            if (jsonDoc.RootElement.TryGetProperty("value", out var valueArray) && valueArray.ValueKind == System.Text.Json.JsonValueKind.Array)
            {
                foreach (var token in valueArray.EnumerateArray())
                {
                    DateTimeOffset? expirationDateTime = null;
                    if (token.TryGetProperty("tokenExpirationDateTime", out var expProp) && expProp.ValueKind != System.Text.Json.JsonValueKind.Null)
                    {
                        if (DateTimeOffset.TryParse(expProp.GetString(), out var parsed))
                        {
                            expirationDateTime = parsed;
                        }
                    }

                    var daysUntilExpiry = expirationDateTime.HasValue
                        ? (int)(expirationDateTime.Value - now).TotalDays
                        : 0;

                    string status;
                    if (!expirationDateTime.HasValue)
                    {
                        status = "Unknown";
                    }
                    else if (daysUntilExpiry < 0)
                    {
                        status = "Expired";
                    }
                    else if (daysUntilExpiry <= 30)
                    {
                        status = "Critical";
                    }
                    else if (daysUntilExpiry <= 60)
                    {
                        status = "Warning";
                    }
                    else
                    {
                        status = "Healthy";
                    }

                    DateTimeOffset? lastModified = null;
                    if (token.TryGetProperty("lastModifiedDateTime", out var modProp) && modProp.ValueKind != System.Text.Json.JsonValueKind.Null)
                    {
                        DateTimeOffset.TryParse(modProp.GetString(), out var parsed);
                        lastModified = parsed;
                    }
                    
                    DateTimeOffset? lastSync = null;
                    if (token.TryGetProperty("lastSuccessfulSyncDateTime", out var syncProp) && syncProp.ValueKind != System.Text.Json.JsonValueKind.Null)
                    {
                        DateTimeOffset.TryParse(syncProp.GetString(), out var parsed);
                        lastSync = parsed;
                    }
                    
                    DateTimeOffset? lastTrig = null;
                    if (token.TryGetProperty("lastSyncTriggeredDateTime", out var trigProp) && trigProp.ValueKind != System.Text.Json.JsonValueKind.Null)
                    {
                        DateTimeOffset.TryParse(trigProp.GetString(), out var parsed);
                        lastTrig = parsed;
                    }

                    tokenList.Add(new
                    {
                        id = token.TryGetProperty("id", out var idProp) ? idProp.GetString() : null,
                        tokenName = token.TryGetProperty("tokenName", out var nameProp) ? nameProp.GetString() : null,
                        appleIdentifier = token.TryGetProperty("appleIdentifier", out var appleProp) ? appleProp.GetString() : null,
                        tokenExpirationDateTime = expirationDateTime,
                        lastModifiedDateTime = lastModified,
                        lastSuccessfulSyncDateTime = lastSync,
                        lastSyncTriggeredDateTime = lastTrig,
                        lastSyncErrorCode = token.TryGetProperty("lastSyncErrorCode", out var errProp) && errProp.ValueKind == System.Text.Json.JsonValueKind.Number ? errProp.GetInt32() : 0,
                        dataSharingConsentGranted = token.TryGetProperty("dataSharingConsentGranted", out var consentProp) && consentProp.ValueKind == System.Text.Json.JsonValueKind.True,
                        tokenType = token.TryGetProperty("tokenType", out var typeProp) ? typeProp.GetString() : null,
                        daysUntilExpiry = daysUntilExpiry,
                        status = status
                    });
                }
            }

            _logger.LogInformation("Found {Count} DEP tokens", tokenList.Count);

            return Ok(new
            {
                tokens = tokenList,
                totalCount = tokenList.Count,
                lastUpdated = DateTime.UtcNow
            });
        }
        catch (Microsoft.Graph.Models.ODataErrors.ODataError ex) when (ex.ResponseStatusCode == 404)
        {
            _logger.LogInformation("DEP tokens endpoint not found (404)");
            return Ok(new
            {
                tokens = Array.Empty<object>(),
                totalCount = 0,
                message = "Apple Enrollment Program is not configured",
                lastUpdated = DateTime.UtcNow
            });
        }
        catch (Microsoft.Graph.Models.ODataErrors.ODataError ex) when (ex.ResponseStatusCode == 403)
        {
            _logger.LogWarning(ex, "Insufficient permissions to read DEP tokens");
            return Ok(new
            {
                tokens = Array.Empty<object>(),
                totalCount = 0,
                permissionRequired = true,
                error = "Missing permission: DeviceManagementServiceConfig.Read.All. Please add this permission in Azure AD and grant admin consent.",
                lastUpdated = DateTime.UtcNow
            });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching DEP tokens");
            
            if (ex.Message.Contains("DeviceManagementServiceConfig") || ex.Message.Contains("not authorized") || ex.Message.Contains("403") || ex.Message.Contains("Forbidden"))
            {
                return Ok(new
                {
                    tokens = Array.Empty<object>(),
                    totalCount = 0,
                    permissionRequired = true,
                    error = "Missing permission: DeviceManagementServiceConfig.Read.All. Please add this permission in Azure AD and grant admin consent.",
                    lastUpdated = DateTime.UtcNow
                });
            }
            
            return Ok(new
            {
                tokens = Array.Empty<object>(),
                totalCount = 0,
                error = ex.Message,
                lastUpdated = DateTime.UtcNow
            });
        }
    }
}
