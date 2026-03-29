using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using M365Dashboard.Api.Services;
using M365Dashboard.Api.Models.Dtos;

namespace M365Dashboard.Api.Controllers;

[ApiController]
[Route("api/[controller]")]
[Authorize]
public class ReportsController : ControllerBase
{
    private readonly IGraphService _graphService;
    private readonly IReportService _reportService;
    private readonly IExecutiveReportService _executiveReportService;
    private readonly ITenantSettingsService _tenantSettingsService;
    private readonly IConfiguration _configuration;
    private readonly ILogger<ReportsController> _logger;

    public ReportsController(
        IGraphService graphService, 
        IReportService reportService,
        IExecutiveReportService executiveReportService,
        ITenantSettingsService tenantSettingsService,
        IConfiguration configuration,
        ILogger<ReportsController> logger)
    {
        _graphService = graphService;
        _reportService = reportService;
        _executiveReportService = executiveReportService;
        _tenantSettingsService = tenantSettingsService;
        _configuration = configuration;
        _logger = logger;
    }

    /// <summary>
    /// Gets the user's object ID from the JWT claims, trying multiple claim types
    /// to handle ASP.NET Core's automatic claim remapping.
    /// </summary>
    private string? GetUserId() =>
        User.FindFirst("oid")?.Value
        ?? User.FindFirst("http://schemas.microsoft.com/identity/claims/objectidentifier")?.Value
        ?? User.FindFirst("sub")?.Value;

    private string? GetUserEmail() =>
        User.FindFirst("preferred_username")?.Value
        ?? User.FindFirst("http://schemas.xmlsoap.org/ws/2005/05/identity/claims/emailaddress")?.Value
        ?? User.FindFirst("email")?.Value
        ?? User.FindFirst("upn")?.Value;

    /// <summary>
    /// Get available report definitions
    /// </summary>
    [HttpGet("definitions")]
    public IActionResult GetReportDefinitions()
    {
        var definitions = _reportService.GetReportDefinitions();
        return Ok(definitions);
    }

    /// <summary>
    /// Generate a report
    /// </summary>
    [HttpPost("generate")]
    public async Task<IActionResult> GenerateReport([FromBody] GenerateReportRequest request)
    {
        try
        {
            var result = await _reportService.GenerateReportAsync(request);
            return Ok(result);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error generating report {ReportType}", request.ReportType);
            return StatusCode(500, new { error = "Failed to generate report", message = ex.Message });
        }
    }

    /// <summary>
    /// Export a report to CSV
    /// </summary>
    [HttpPost("export")]
    public async Task<IActionResult> ExportReport([FromBody] GenerateReportRequest request)
    {
        try
        {
            if (request.Format?.ToLower() == "html")
            {
                var htmlData = await _reportService.ExportReportToHtmlAsync(request);
                var fileName = $"{request.ReportType}_{DateTime.UtcNow:yyyyMMdd_HHmmss}.html";
                return File(System.Text.Encoding.UTF8.GetBytes(htmlData), "text/html", fileName);
            }
            else
            {
                var csvData = await _reportService.ExportReportToCsvAsync(request);
                var fileName = $"{request.ReportType}_{DateTime.UtcNow:yyyyMMdd_HHmmss}.csv";
                return File(System.Text.Encoding.UTF8.GetBytes(csvData), "text/csv", fileName);
            }
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error exporting report {ReportType}", request.ReportType);
            return StatusCode(500, new { error = "Failed to export report", message = ex.Message });
        }
    }

    /// <summary>
    /// Get all scheduled reports for the current user
    /// </summary>
    [HttpGet("schedules")]
    public async Task<IActionResult> GetScheduledReports()
    {
        var userId = GetUserId();
        if (string.IsNullOrEmpty(userId))
        {
            return Unauthorized();
        }

        var schedules = await _reportService.GetScheduledReportsAsync(userId);
        return Ok(schedules);
    }

    /// <summary>
    /// Create a new scheduled report
    /// </summary>
    [HttpPost("schedules")]
    public async Task<IActionResult> CreateScheduledReport([FromBody] CreateScheduledReportRequest request)
    {
        var userId = GetUserId();
        var userEmail = GetUserEmail();

        if (string.IsNullOrEmpty(userId))
        {
            return Unauthorized();
        }

        try
        {
            var schedule = await _reportService.CreateScheduledReportAsync(userId, userEmail, request);
            return Ok(schedule);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error creating scheduled report");
            return StatusCode(500, new { error = "Failed to create scheduled report", message = ex.Message });
        }
    }

    /// <summary>
    /// Update a scheduled report
    /// </summary>
    [HttpPut("schedules/{scheduleId}")]
    public async Task<IActionResult> UpdateScheduledReport(string scheduleId, [FromBody] UpdateScheduledReportRequest request)
    {
        var userId = GetUserId();
        if (string.IsNullOrEmpty(userId))
        {
            return Unauthorized();
        }

        try
        {
            var schedule = await _reportService.UpdateScheduledReportAsync(userId, scheduleId, request);
            if (schedule == null)
            {
                return NotFound();
            }
            return Ok(schedule);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error updating scheduled report {ScheduleId}", scheduleId);
            return StatusCode(500, new { error = "Failed to update scheduled report", message = ex.Message });
        }
    }

    /// <summary>
    /// Delete a scheduled report
    /// </summary>
    [HttpDelete("schedules/{scheduleId}")]
    public async Task<IActionResult> DeleteScheduledReport(string scheduleId)
    {
        var userId = GetUserId();
        if (string.IsNullOrEmpty(userId))
        {
            return Unauthorized();
        }

        try
        {
            var success = await _reportService.DeleteScheduledReportAsync(userId, scheduleId);
            if (!success)
            {
                return NotFound();
            }
            return NoContent();
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error deleting scheduled report {ScheduleId}", scheduleId);
            return StatusCode(500, new { error = "Failed to delete scheduled report", message = ex.Message });
        }
    }

    /// <summary>
    /// Download the executive summary as a PDF file
    /// </summary>
    [HttpPost("download")]
    public async Task<IActionResult> DownloadReport([FromBody] GenerateReportRequest request)
    {
        try
        {
            if (request.ReportType != "executive-summary-pdf")
                return BadRequest(new { error = "Only executive-summary-pdf supports PDF download" });

            var tenantId = _configuration["AzureAd:TenantId"] ?? "default";
            var settings = await _tenantSettingsService.GetReportSettingsAsync(tenantId);

            // Build a clean slug from the company name, skipping the default placeholder
            var company = settings.CompanyName?.Trim();
            var isDefault = string.IsNullOrEmpty(company)
                || company.Equals("M365 Dashboard", StringComparison.OrdinalIgnoreCase)
                || company.Equals("My Organisation", StringComparison.OrdinalIgnoreCase);
            var companySlug = isDefault ? "" : company!.Replace(" ", "_") + "_";

            var pdfBytes = await _executiveReportService.GeneratePdfAsync();
            var fileName = $"{companySlug}Executive_Summary_{DateTime.UtcNow:yyyy-MM-dd}.pdf";
            return File(pdfBytes, "application/pdf", fileName);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error generating executive summary PDF");
            return StatusCode(500, new { error = "Failed to generate PDF", message = ex.Message });
        }
    }

    /// <summary>
    /// Get report history
    /// </summary>
    [HttpGet("history")]
    public async Task<IActionResult> GetReportHistory([FromQuery] int take = 20)
    {
        var userId = GetUserId();
        if (string.IsNullOrEmpty(userId))
        {
            return Unauthorized();
        }

        var history = await _reportService.GetReportHistoryAsync(userId, take);
        return Ok(history);
    }
}
