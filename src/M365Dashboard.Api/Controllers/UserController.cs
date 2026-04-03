using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using M365Dashboard.Api.Services;
using System.Security.Claims;

namespace M365Dashboard.Api.Controllers;

[Authorize]
[ApiController]
[Route("api/[controller]")]
public class UserController : ControllerBase
{
    private readonly IGraphService _graphService;
    private readonly ILogger<UserController> _logger;

    public UserController(
        IGraphService graphService,
        ILogger<UserController> logger)
    {
        _graphService = graphService;
        _logger = logger;
    }

    private string GetUserId()
    {
        // Try oid first (Azure AD Object ID) - this is the correct ID for Graph API
        var oid = User.FindFirst("http://schemas.microsoft.com/identity/claims/objectidentifier")?.Value
            ?? User.FindFirst("oid")?.Value;
        
        if (!string.IsNullOrEmpty(oid))
        {
            _logger.LogDebug("Using oid claim: {Oid}", oid);
            return oid;
        }
        
        // Fallback to sub (but this won't work with Graph API)
        var sub = User.FindFirst(ClaimTypes.NameIdentifier)?.Value
            ?? User.FindFirst("sub")?.Value;
        
        if (!string.IsNullOrEmpty(sub))
        {
            _logger.LogWarning("Using sub claim instead of oid - Graph API calls may fail. Sub: {Sub}", sub);
            return sub;
        }
        
        _logger.LogError("No user ID found in token. Available claims: {Claims}",
            string.Join(", ", User.Claims.Select(c => $"{c.Type}={c.Value}")));
        throw new UnauthorizedAccessException("User ID not found in token");
    }

    /// <summary>
    /// Get current user's profile from Microsoft Graph
    /// Uses Application permissions to read the user's profile
    /// </summary>
    [HttpGet("profile")]
    public async Task<IActionResult> GetProfile()
    {
        try
        {
            var userId = GetUserId();
            _logger.LogInformation("Fetching profile for user {UserId}", userId);
            
            var profile = await _graphService.GetUserProfileAsync(userId);
            
            // Add roles from the token (app roles assigned to this user)
            var roles = User.Claims
                .Where(c => c.Type == ClaimTypes.Role || c.Type == "roles")
                .Select(c => c.Value)
                .ToList();

            var profileWithRoles = profile with { Roles = roles };
            
            return Ok(profileWithRoles);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error retrieving user profile");
            return StatusCode(500, new { error = "Failed to retrieve user profile" });
        }
    }

    /// <summary>
    /// Get current user's roles from the token
    /// </summary>
    [HttpGet("roles")]
    public IActionResult GetRoles()
    {
        var roles = User.Claims
            .Where(c => c.Type == ClaimTypes.Role || c.Type == "roles")
            .Select(c => c.Value)
            .ToList();

        var isAdmin = roles.Contains("Dashboard.Admin");
        var isReader = roles.Contains("Dashboard.Reader") || isAdmin;

        return Ok(new
        {
            Roles = roles,
            IsAdmin = isAdmin,
            IsReader = isReader
        });
    }

    /// <summary>
    /// Check if the current user is authenticated and has access
    /// </summary>
    [HttpGet("check")]
    public IActionResult CheckAccess()
    {
        var userId = GetUserId();
        var name = User.FindFirst("name")?.Value ?? User.Identity?.Name ?? "Unknown";
        var email = User.FindFirst("preferred_username")?.Value ?? 
                    User.FindFirst(ClaimTypes.Email)?.Value ?? "";
        
        var roles = User.Claims
            .Where(c => c.Type == ClaimTypes.Role || c.Type == "roles")
            .Select(c => c.Value)
            .ToList();

        return Ok(new
        {
            Authenticated = true,
            UserId = userId,
            Name = name,
            Email = email,
            Roles = roles,
            HasAccess = roles.Any(r => r == "Dashboard.Admin" || r == "Dashboard.Reader")
        });
    }
}
