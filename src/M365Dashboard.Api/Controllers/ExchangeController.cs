using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using M365Dashboard.Api.Services;

namespace M365Dashboard.Api.Controllers;

[ApiController]
[Route("api/exchange")]
[Authorize]
public class ExchangeController : ControllerBase
{
    private readonly IExchangeOnlineService _exchangeService;
    private readonly ILogger<ExchangeController> _logger;

    public ExchangeController(IExchangeOnlineService exchangeService, ILogger<ExchangeController> logger)
    {
        _exchangeService = exchangeService;
        _logger = logger;
    }

    /// <summary>
    /// Get all Exchange distribution lists
    /// </summary>
    [HttpGet("distribution-lists")]
    public async Task<IActionResult> GetDistributionLists([FromQuery] int take = 100)
    {
        try
        {
            _logger.LogInformation("ExchangeController: Fetching distribution lists");
            var result = await _exchangeService.GetDistributionListsAsync(take);
            _logger.LogInformation("ExchangeController: Found {Count} distribution lists", result.TotalCount);
            return Ok(result);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching Exchange distribution lists");
            return StatusCode(500, new { error = "Failed to fetch distribution lists", message = ex.Message });
        }
    }

    /// <summary>
    /// Test Exchange connection
    /// </summary>
    [HttpGet("test")]
    public async Task<IActionResult> TestConnection()
    {
        try
        {
            _logger.LogInformation("Testing Exchange connection...");
            var result = await _exchangeService.GetDistributionListsAsync(5);
            return Ok(new { 
                success = true, 
                message = $"Exchange connection successful. Found {result.TotalCount} distribution lists.",
                count = result.TotalCount,
                lists = result.DistributionLists.Select(d => new { d.DisplayName, d.PrimarySmtpAddress })
            });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Exchange connection test failed");
            return Ok(new { 
                success = false, 
                message = ex.Message,
                innerException = ex.InnerException?.Message
            });
        }
    }

    /// <summary>
    /// Debug: Get all mail-enabled recipients
    /// </summary>
    [HttpGet("debug/recipients")]
    public async Task<IActionResult> DebugGetRecipients()
    {
        try
        {
            var result = await _exchangeService.DebugGetRecipientsAsync();
            return Ok(result);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Debug recipients failed");
            return Ok(new { success = false, message = ex.Message });
        }
    }

    /// <summary>
    /// Get a specific distribution list with details
    /// </summary>
    [HttpGet("distribution-lists/{identity}")]
    public async Task<IActionResult> GetDistributionList(string identity)
    {
        try
        {
            var result = await _exchangeService.GetDistributionListAsync(identity);
            if (result == null)
            {
                return NotFound(new { error = "Distribution list not found" });
            }
            return Ok(result);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching distribution list: {Identity}", identity);
            return StatusCode(500, new { error = "Failed to fetch distribution list", message = ex.Message });
        }
    }

    /// <summary>
    /// Get members of a distribution list
    /// </summary>
    [HttpGet("distribution-lists/{identity}/members")]
    public async Task<IActionResult> GetDistributionListMembers(string identity)
    {
        try
        {
            var result = await _exchangeService.GetDistributionListMembersAsync(identity);
            return Ok(new { members = result, count = result.Count });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching distribution list members: {Identity}", identity);
            return StatusCode(500, new { error = "Failed to fetch distribution list members", message = ex.Message });
        }
    }

    /// <summary>
    /// Get mailboxes with forwarding enabled (Exchange-level forwarding)
    /// </summary>
    [HttpGet("mailbox-forwarding")]
    public async Task<IActionResult> GetMailboxesWithForwarding([FromQuery] int take = 500)
    {
        try
        {
            _logger.LogInformation("ExchangeController: Fetching mailboxes with forwarding");
            var result = await _exchangeService.GetMailboxesWithForwardingAsync(take);
            _logger.LogInformation("ExchangeController: Found {Count} mailboxes with forwarding", result.TotalCount);
            return Ok(result);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching mailboxes with forwarding");
            return StatusCode(500, new { error = "Failed to fetch mailbox forwarding data", message = ex.Message });
        }
    }

    /// <summary>
    /// Get all mailboxes that a user has been granted access to
    /// </summary>
    [HttpGet("mailbox-access/by-user")]
    public async Task<IActionResult> GetMailboxAccessForUser([FromQuery] string email)
    {
        if (string.IsNullOrWhiteSpace(email) || !email.Contains('@'))
            return BadRequest(new { error = "A valid email address is required" });

        try
        {
            _logger.LogInformation("Checking mailbox access for user: {Email}", email);
            var result = await _exchangeService.GetMailboxAccessForUserAsync(email.Trim().ToLower());
            return Ok(result);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error checking mailbox access for {Email}", email);
            return StatusCode(500, new { error = "Failed to check mailbox access", message = ex.Message });
        }
    }

    /// <summary>
    /// Get all delegates that have been granted access to a specific mailbox
    /// </summary>
    [HttpGet("mailbox-access/delegates")]
    public async Task<IActionResult> GetMailboxDelegates([FromQuery] string email)
    {
        if (string.IsNullOrWhiteSpace(email) || !email.Contains('@'))
            return BadRequest(new { error = "A valid email address is required" });

        try
        {
            _logger.LogInformation("Checking delegates for mailbox: {Email}", email);
            var result = await _exchangeService.GetMailboxDelegatesAsync(email.Trim().ToLower());
            return Ok(result);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error checking delegates for {Email}", email);
            return StatusCode(500, new { error = "Failed to check mailbox delegates", message = ex.Message });
        }
    }

    /// <summary>
    /// Get inbox rules with forwarding actions (via Exchange PowerShell)
    /// </summary>
    [HttpGet("inbox-rules-forwarding")]
    public async Task<IActionResult> GetInboxRulesWithForwarding([FromQuery] int take = 100)
    {
        try
        {
            _logger.LogInformation("ExchangeController: Fetching inbox rules with forwarding for up to {Take} mailboxes", take);
            var result = await _exchangeService.GetInboxRulesWithForwardingAsync(take);
            _logger.LogInformation("ExchangeController: Found {Count} forwarding rules in {Mailboxes} mailboxes", 
                result.TotalForwardingRules, result.MailboxesWithForwarding);
            return Ok(result);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching inbox rules with forwarding");
            return StatusCode(500, new { error = "Failed to fetch inbox rules forwarding data", message = ex.Message });
        }
    }
}
