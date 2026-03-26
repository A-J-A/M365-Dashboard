using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using M365Dashboard.Api.Services;

namespace M365Dashboard.Api.Controllers;

[ApiController]
[Route("api/[controller]")]
[Authorize]
public class SecurityController : ControllerBase
{
    private readonly IGraphService _graphService;
    private readonly ILogger<SecurityController> _logger;

    public SecurityController(IGraphService graphService, ILogger<SecurityController> logger)
    {
        _graphService = graphService;
        _logger = logger;
    }

    /// <summary>
    /// Get security overview including secure score, risky users, and risky sign-ins
    /// </summary>
    [HttpGet("overview")]
    public async Task<IActionResult> GetSecurityOverview()
    {
        try
        {
            var overview = await _graphService.GetSecurityOverviewAsync();
            return Ok(overview);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching security overview");
            return StatusCode(500, new { error = "Failed to fetch security overview", message = ex.Message });
        }
    }

    /// <summary>
    /// Get Microsoft Secure Score
    /// </summary>
    [HttpGet("securescore")]
    public async Task<IActionResult> GetSecureScore()
    {
        try
        {
            var score = await _graphService.GetSecureScoreAsync();
            if (score == null)
            {
                return Ok(new { message = "Secure Score not available. This may require additional permissions." });
            }
            return Ok(score);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching secure score");
            return StatusCode(500, new { error = "Failed to fetch secure score", message = ex.Message });
        }
    }

    /// <summary>
    /// Get risky users from Identity Protection
    /// </summary>
    [HttpGet("riskyusers")]
    public async Task<IActionResult> GetRiskyUsers()
    {
        try
        {
            var users = await _graphService.GetRiskyUsersAsync();
            return Ok(users);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching risky users");
            return StatusCode(500, new { error = "Failed to fetch risky users", message = ex.Message });
        }
    }

    /// <summary>
    /// Get risky sign-ins from the specified time period
    /// </summary>
    [HttpGet("riskysignins")]
    public async Task<IActionResult> GetRiskySignIns([FromQuery] int hours = 24)
    {
        try
        {
            var signIns = await _graphService.GetRiskySignInsAsync(hours);
            return Ok(signIns);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching risky sign-ins");
            return StatusCode(500, new { error = "Failed to fetch risky sign-ins", message = ex.Message });
        }
    }

    /// <summary>
    /// Get security statistics
    /// </summary>
    [HttpGet("stats")]
    public async Task<IActionResult> GetSecurityStats()
    {
        try
        {
            var stats = await _graphService.GetSecurityStatsAsync();
            return Ok(stats);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching security statistics");
            return StatusCode(500, new { error = "Failed to fetch security statistics", message = ex.Message });
        }
    }

    /// <summary>
    /// Get MFA registration details for all users
    /// </summary>
    [HttpGet("mfa")]
    public async Task<IActionResult> GetMfaRegistrationDetails()
    {
        try
        {
            var details = await _graphService.GetMfaRegistrationDetailsAsync();
            return Ok(details);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching MFA registration details");
            return StatusCode(500, new { error = "Failed to fetch MFA registration details", message = ex.Message });
        }
    }

    /// <summary>
    /// Get app registration credential status (expiring and expired secrets/certificates)
    /// </summary>
    [HttpGet("app-credentials")]
    public async Task<IActionResult> GetAppCredentialStatus([FromQuery] int thresholdDays = 45)
    {
        try
        {
            var status = await _graphService.GetAppCredentialStatusAsync(thresholdDays);
            return Ok(status);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching app credential status");
            return StatusCode(500, new { error = "Failed to fetch app credential status", message = ex.Message });
        }
    }
}
