using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using M365Dashboard.Api.Services;

namespace M365Dashboard.Api.Controllers;

[Authorize]
[ApiController]
[Route("api/[controller]")]
public class UsersController : ControllerBase
{
    private readonly IGraphService _graphService;
    private readonly ILogger<UsersController> _logger;

    public UsersController(
        IGraphService graphService,
        ILogger<UsersController> logger)
    {
        _graphService = graphService;
        _logger = logger;
    }

    /// <summary>
    /// Get all users in the tenant with details
    /// </summary>
    [HttpGet]
    public async Task<IActionResult> GetUsers(
        [FromQuery] string? filter = null,
        [FromQuery] string? orderBy = "displayName",
        [FromQuery] bool ascending = true,
        [FromQuery] int take = 100)
    {
        try
        {
            _logger.LogInformation("Fetching users with filter: {Filter}, orderBy: {OrderBy}", filter, orderBy);
            var result = await _graphService.GetUsersAsync(filter, orderBy, ascending, take);
            return Ok(result);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error retrieving users");
            return StatusCode(500, new { error = "Failed to retrieve users" });
        }
    }

    /// <summary>
    /// Get a specific user's details
    /// </summary>
    [HttpGet("{userId}")]
    public async Task<IActionResult> GetUser(string userId)
    {
        try
        {
            _logger.LogInformation("Fetching user details for {UserId}", userId);
            var user = await _graphService.GetUserDetailsAsync(userId);
            return Ok(user);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error retrieving user {UserId}", userId);
            return StatusCode(500, new { error = "Failed to retrieve user details" });
        }
    }

    /// <summary>
    /// Get user statistics summary
    /// </summary>
    [HttpGet("stats")]
    public async Task<IActionResult> GetUserStats()
    {
        try
        {
            var stats = await _graphService.GetUserStatsAsync();
            return Ok(stats);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error retrieving user statistics");
            return StatusCode(500, new { error = "Failed to retrieve user statistics" });
        }
    }
}
