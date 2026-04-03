using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using M365Dashboard.Api.Services;
using M365Dashboard.Api.Configuration;

namespace M365Dashboard.Api.Controllers;

[Authorize]
[ApiController]
[Route("api/[controller]")]
public class DashboardController : ControllerBase
{
    private readonly IWidgetDataService _widgetDataService;
    private readonly ILogger<DashboardController> _logger;

    public DashboardController(
        IWidgetDataService widgetDataService,
        ILogger<DashboardController> logger)
    {
        _widgetDataService = widgetDataService;
        _logger = logger;
    }

    /// <summary>
    /// Get dashboard summary with key metrics
    /// </summary>
    [HttpGet("summary")]
    public async Task<IActionResult> GetSummary()
    {
        try
        {
            var summary = await _widgetDataService.GetDashboardSummaryAsync();
            return Ok(summary);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error retrieving dashboard summary");
            return StatusCode(500, new { error = "Failed to retrieve dashboard summary" });
        }
    }

    /// <summary>
    /// Get active users data
    /// </summary>
    [HttpGet("widgets/active-users")]
    public async Task<IActionResult> GetActiveUsers([FromQuery] string dateRange = "last30days")
    {
        try
        {
            var data = await _widgetDataService.GetActiveUsersDataAsync(dateRange);
            return Ok(data);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error retrieving active users data");
            return StatusCode(500, new { error = "Failed to retrieve active users data" });
        }
    }

    /// <summary>
    /// Get sign-in analytics data
    /// </summary>
    [HttpGet("widgets/sign-in-analytics")]
    public async Task<IActionResult> GetSignInAnalytics([FromQuery] string dateRange = "last7days")
    {
        try
        {
            var data = await _widgetDataService.GetSignInAnalyticsDataAsync(dateRange);
            return Ok(data);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error retrieving sign-in analytics");
            return StatusCode(500, new { error = "Failed to retrieve sign-in analytics" });
        }
    }

    /// <summary>
    /// Get license usage data
    /// </summary>
    [HttpGet("widgets/license-usage")]
    public async Task<IActionResult> GetLicenseUsage()
    {
        try
        {
            var data = await _widgetDataService.GetLicenseUsageDataAsync();
            return Ok(data);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error retrieving license usage");
            return StatusCode(500, new { error = "Failed to retrieve license usage" });
        }
    }

    /// <summary>
    /// Get device compliance data
    /// </summary>
    [HttpGet("widgets/device-compliance")]
    public async Task<IActionResult> GetDeviceCompliance()
    {
        try
        {
            var data = await _widgetDataService.GetDeviceComplianceDataAsync();
            return Ok(data);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error retrieving device compliance");
            return StatusCode(500, new { error = "Failed to retrieve device compliance" });
        }
    }

    /// <summary>
    /// Get mail activity data
    /// </summary>
    [HttpGet("widgets/mail-activity")]
    public async Task<IActionResult> GetMailActivity([FromQuery] string dateRange = "last30days")
    {
        try
        {
            var data = await _widgetDataService.GetMailActivityDataAsync(dateRange);
            return Ok(data);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error retrieving mail activity");
            return StatusCode(500, new { error = "Failed to retrieve mail activity" });
        }
    }

    /// <summary>
    /// Get Teams activity data
    /// </summary>
    [HttpGet("widgets/teams-activity")]
    public async Task<IActionResult> GetTeamsActivity([FromQuery] string dateRange = "last30days")
    {
        try
        {
            var data = await _widgetDataService.GetTeamsActivityDataAsync(dateRange);
            return Ok(data);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error retrieving Teams activity");
            return StatusCode(500, new { error = "Failed to retrieve Teams activity" });
        }
    }

    /// <summary>
    /// Refresh cache for specific metric type (Admin only)
    /// </summary>
    [HttpPost("cache/refresh")]
    [Authorize(Policy = "RequireAdminRole")]
    public async Task<IActionResult> RefreshCache([FromQuery] string metricType = "all")
    {
        try
        {
            await _widgetDataService.RefreshCacheAsync(metricType);
            return Ok(new { message = $"Cache refreshed for: {metricType}" });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error refreshing cache");
            return StatusCode(500, new { error = "Failed to refresh cache" });
        }
    }

    /// <summary>
    /// Get available widget definitions
    /// </summary>
    [HttpGet("widgets/definitions")]
    public IActionResult GetWidgetDefinitions()
    {
        var definitions = WidgetTypes.Definitions.Values.Select(d => new
        {
            d.Type,
            d.Name,
            d.Description,
            d.Category,
            d.RequiredPermissions,
            d.DefaultWidth,
            d.DefaultHeight
        });

        return Ok(definitions);
    }
}
