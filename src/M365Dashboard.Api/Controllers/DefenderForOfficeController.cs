using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using M365Dashboard.Api.Services;

namespace M365Dashboard.Api.Controllers;

[ApiController]
[Route("api/defender-office")]
[Authorize]
public class DefenderForOfficeController : ControllerBase
{
    private readonly IDefenderForOfficeService _defenderService;
    private readonly ILogger<DefenderForOfficeController> _logger;

    public DefenderForOfficeController(IDefenderForOfficeService defenderService, ILogger<DefenderForOfficeController> logger)
    {
        _defenderService = defenderService;
        _logger = logger;
    }

    /// <summary>
    /// Get all Defender for Office 365 policy data in one call
    /// </summary>
    [HttpGet("overview")]
    public async Task<IActionResult> GetOverview()
    {
        try
        {
            _logger.LogInformation("Fetching Defender for Office 365 overview");
            var result = await _defenderService.GetOverviewAsync();
            return Ok(result);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching Defender for Office overview");
            return StatusCode(500, new { error = "Failed to fetch Defender overview", message = ex.Message });
        }
    }

    [HttpGet("anti-phish")]
    public async Task<IActionResult> GetAntiPhishPolicies()
    {
        try
        {
            var result = await _defenderService.GetAntiPhishPoliciesAsync();
            return Ok(result);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching anti-phish policies");
            return StatusCode(500, new { error = "Failed to fetch anti-phish policies", message = ex.Message });
        }
    }

    [HttpGet("anti-malware")]
    public async Task<IActionResult> GetAntiMalwarePolicies()
    {
        try
        {
            var result = await _defenderService.GetAntiMalwarePoliciesAsync();
            return Ok(result);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching anti-malware policies");
            return StatusCode(500, new { error = "Failed to fetch anti-malware policies", message = ex.Message });
        }
    }

    [HttpGet("anti-spam")]
    public async Task<IActionResult> GetAntiSpamPolicies()
    {
        try
        {
            var result = await _defenderService.GetAntiSpamPoliciesAsync();
            return Ok(result);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching anti-spam policies");
            return StatusCode(500, new { error = "Failed to fetch anti-spam policies", message = ex.Message });
        }
    }

    [HttpGet("outbound-spam")]
    public async Task<IActionResult> GetOutboundSpamPolicies()
    {
        try
        {
            var result = await _defenderService.GetOutboundSpamPoliciesAsync();
            return Ok(result);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching outbound spam policies");
            return StatusCode(500, new { error = "Failed to fetch outbound spam policies", message = ex.Message });
        }
    }

    [HttpGet("safe-attachments")]
    public async Task<IActionResult> GetSafeAttachmentsPolicies()
    {
        try
        {
            var result = await _defenderService.GetSafeAttachmentsPoliciesAsync();
            return Ok(result);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching Safe Attachments policies");
            return StatusCode(500, new { error = "Failed to fetch Safe Attachments policies", message = ex.Message });
        }
    }

    [HttpGet("safe-links")]
    public async Task<IActionResult> GetSafeLinksPolicies()
    {
        try
        {
            var result = await _defenderService.GetSafeLinksPoliciesAsync();
            return Ok(result);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching Safe Links policies");
            return StatusCode(500, new { error = "Failed to fetch Safe Links policies", message = ex.Message });
        }
    }
}
