using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using M365Dashboard.Api.Services;

namespace M365Dashboard.Api.Controllers;

[ApiController]
[Route("api/[controller]")]
[Authorize]
public class GroupsController : ControllerBase
{
    private readonly IGraphService _graphService;
    private readonly ILogger<GroupsController> _logger;

    public GroupsController(IGraphService graphService, ILogger<GroupsController> logger)
    {
        _graphService = graphService;
        _logger = logger;
    }

    /// <summary>
    /// Get all groups in the tenant
    /// </summary>
    [HttpGet]
    public async Task<IActionResult> GetGroups(
        [FromQuery] string? filter = null,
        [FromQuery] string? orderBy = "displayName",
        [FromQuery] bool ascending = true,
        [FromQuery] int take = 100)
    {
        try
        {
            var result = await _graphService.GetGroupsAsync(filter, orderBy, ascending, take);
            return Ok(result);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching groups");
            return StatusCode(500, new { error = "Failed to fetch groups", message = ex.Message });
        }
    }

    /// <summary>
    /// Get group statistics
    /// </summary>
    [HttpGet("stats")]
    public async Task<IActionResult> GetGroupStats()
    {
        try
        {
            var stats = await _graphService.GetGroupStatsAsync();
            return Ok(stats);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching group statistics");
            return StatusCode(500, new { error = "Failed to fetch group statistics", message = ex.Message });
        }
    }

    /// <summary>
    /// Get detailed information about a specific group
    /// </summary>
    [HttpGet("{groupId}")]
    public async Task<IActionResult> GetGroupDetails(string groupId)
    {
        try
        {
            var details = await _graphService.GetGroupDetailsAsync(groupId);
            return Ok(details);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching group details for {GroupId}", groupId);
            return StatusCode(500, new { error = "Failed to fetch group details", message = ex.Message });
        }
    }

    /// <summary>
    /// Get distribution lists specifically (mail-enabled, non-security groups)
    /// </summary>
    [HttpGet("distribution-lists")]
    public async Task<IActionResult> GetDistributionLists([FromQuery] int take = 200)
    {
        try
        {
            var result = await _graphService.GetDistributionListsAsync(take);
            return Ok(result);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching distribution lists");
            return StatusCode(500, new { error = "Failed to fetch distribution lists", message = ex.Message });
        }
    }
}
