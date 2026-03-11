using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Graph;
using Microsoft.Graph.Models;

namespace M365Dashboard.Api.Controllers;

[ApiController]
[Route("api/[controller]")]
[Authorize]
public class TeamsPhoneController : ControllerBase
{
    private readonly GraphServiceClient _graphClient;
    private readonly ILogger<TeamsPhoneController> _logger;

    public TeamsPhoneController(GraphServiceClient graphClient, ILogger<TeamsPhoneController> logger)
    {
        _graphClient = graphClient;
        _logger = logger;
    }

    /// <summary>
    /// Get PSTN call records summary
    /// </summary>
    [HttpGet("pstn-calls")]
    public async Task<IActionResult> GetPstnCalls([FromQuery] int days = 30)
    {
        try
        {
            _logger.LogInformation("Fetching PSTN call records for last {Days} days", days);

            var fromDate = DateTime.UtcNow.AddDays(-days);
            var toDate = DateTime.UtcNow;

            // Get PSTN call records using the Call Records API
            var callRecordsResponse = await _graphClient.Communications.CallRecords
                .MicrosoftGraphCallRecordsGetPstnCallsWithFromDateTimeWithToDateTime(fromDate, toDate)
                .GetAsync();

            var calls = callRecordsResponse?.Value ?? new List<Microsoft.Graph.Models.CallRecords.PstnCallLogRow>();

            // Calculate summary statistics
            var totalCalls = calls.Count;
            var inboundCalls = calls.Count(c => c.CallType?.Contains("inbound", StringComparison.OrdinalIgnoreCase) == true);
            var outboundCalls = calls.Count(c => c.CallType?.Contains("outbound", StringComparison.OrdinalIgnoreCase) == true);
            
            var totalDurationMinutes = calls.Sum(c => c.Duration ?? 0) / 60.0;
            var answeredCalls = calls.Count(c => (c.Duration ?? 0) > 0);
            var missedCalls = totalCalls - answeredCalls;

            // Group by date for trend
            var dailyTrend = calls
                .Where(c => c.StartDateTime.HasValue)
                .GroupBy(c => c.StartDateTime!.Value.Date)
                .OrderBy(g => g.Key)
                .Select(g => new
                {
                    date = g.Key,
                    totalCalls = g.Count(),
                    inbound = g.Count(c => c.CallType?.Contains("inbound", StringComparison.OrdinalIgnoreCase) == true),
                    outbound = g.Count(c => c.CallType?.Contains("outbound", StringComparison.OrdinalIgnoreCase) == true),
                    durationMinutes = g.Sum(c => c.Duration ?? 0) / 60.0
                })
                .ToList();

            // Top callers
            var topCallers = calls
                .Where(c => !string.IsNullOrEmpty(c.UserPrincipalName))
                .GroupBy(c => c.UserPrincipalName)
                .OrderByDescending(g => g.Count())
                .Take(10)
                .Select(g => new
                {
                    user = g.Key,
                    displayName = g.First().UserDisplayName ?? g.Key,
                    callCount = g.Count(),
                    totalMinutes = Math.Round(g.Sum(c => c.Duration ?? 0) / 60.0, 1)
                })
                .ToList();

            // Calls by hour of day
            var callsByHour = calls
                .Where(c => c.StartDateTime.HasValue)
                .GroupBy(c => c.StartDateTime!.Value.Hour)
                .OrderBy(g => g.Key)
                .Select(g => new { hour = g.Key, count = g.Count() })
                .ToList();

            // Individual call records (limit to most recent 200)
            var callRecords = calls
                .OrderByDescending(c => c.StartDateTime)
                .Take(200)
                .Select(c => new
                {
                    id = c.Id,
                    userPrincipalName = c.UserPrincipalName,
                    userDisplayName = c.UserDisplayName,
                    startDateTime = c.StartDateTime,
                    endDateTime = c.EndDateTime,
                    duration = c.Duration ?? 0,
                    callType = c.CallType,
                    calleeNumber = c.CalleeNumber,
                    callerNumber = c.CallerNumber,
                    destinationName = c.DestinationName,
                    destinationContext = c.DestinationContext,
                    isAnswered = (c.Duration ?? 0) > 0,
                    charge = c.Charge,
                    currency = c.Currency,
                    connectionCharge = c.ConnectionCharge,
                    licenseCapability = c.LicenseCapability,
                    inventoryType = c.InventoryType
                })
                .ToList();

            return Ok(new
            {
                summary = new
                {
                    totalCalls,
                    inboundCalls,
                    outboundCalls,
                    answeredCalls,
                    missedCalls,
                    totalDurationMinutes = Math.Round(totalDurationMinutes, 1),
                    averageCallDurationSeconds = totalCalls > 0 ? Math.Round(calls.Average(c => c.Duration ?? 0), 0) : 0,
                    answerRate = totalCalls > 0 ? Math.Round((double)answeredCalls / totalCalls * 100, 1) : 0
                },
                dailyTrend,
                topCallers,
                callsByHour,
                calls = callRecords,
                period = new { fromDate, toDate },
                lastUpdated = DateTime.UtcNow
            });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching PSTN call records");
            // Return empty data structure if API fails (might not have Teams Phone license)
            return Ok(new
            {
                summary = new
                {
                    totalCalls = 0,
                    inboundCalls = 0,
                    outboundCalls = 0,
                    answeredCalls = 0,
                    missedCalls = 0,
                    totalDurationMinutes = 0.0,
                    averageCallDurationSeconds = 0.0,
                    answerRate = 0.0
                },
                dailyTrend = new List<object>(),
                topCallers = new List<object>(),
                callsByHour = new List<object>(),
                period = new { fromDate = DateTime.UtcNow.AddDays(-days), toDate = DateTime.UtcNow },
                lastUpdated = DateTime.UtcNow,
                error = "Unable to fetch PSTN call data. Ensure you have Teams Phone System licenses and appropriate permissions.",
                errorDetail = ex.Message
            });
        }
    }

    /// <summary>
    /// Get Direct Routing call records
    /// </summary>
    [HttpGet("direct-routing-calls")]
    public async Task<IActionResult> GetDirectRoutingCalls([FromQuery] int days = 30)
    {
        try
        {
            _logger.LogInformation("Fetching Direct Routing call records for last {Days} days", days);

            var fromDate = DateTime.UtcNow.AddDays(-days);
            var toDate = DateTime.UtcNow;

            var callRecordsResponse = await _graphClient.Communications.CallRecords
                .MicrosoftGraphCallRecordsGetDirectRoutingCallsWithFromDateTimeWithToDateTime(fromDate, toDate)
                .GetAsync();

            var calls = callRecordsResponse?.Value ?? new List<Microsoft.Graph.Models.CallRecords.DirectRoutingLogRow>();

            var totalCalls = calls.Count;
            var successfulCalls = calls.Count(c => c.FinalSipCode == 200 || c.FinalSipCode == 0);
            var failedCalls = totalCalls - successfulCalls;
            var totalDurationMinutes = calls.Sum(c => c.Duration ?? 0) / 60.0;

            // Group by SBC for analysis
            var callsBySbc = calls
                .Where(c => !string.IsNullOrEmpty(c.TrunkFullyQualifiedDomainName))
                .GroupBy(c => c.TrunkFullyQualifiedDomainName)
                .Select(g => new
                {
                    sbc = g.Key,
                    totalCalls = g.Count(),
                    successfulCalls = g.Count(c => c.FinalSipCode == 200 || c.FinalSipCode == 0),
                    failedCalls = g.Count(c => c.FinalSipCode != 200 && c.FinalSipCode != 0),
                    totalMinutes = Math.Round(g.Sum(c => c.Duration ?? 0) / 60.0, 1)
                })
                .OrderByDescending(s => s.totalCalls)
                .ToList();

            // Daily trend
            var dailyTrend = calls
                .Where(c => c.StartDateTime.HasValue)
                .GroupBy(c => c.StartDateTime!.Value.Date)
                .OrderBy(g => g.Key)
                .Select(g => new
                {
                    date = g.Key,
                    totalCalls = g.Count(),
                    successful = g.Count(c => c.FinalSipCode == 200 || c.FinalSipCode == 0),
                    failed = g.Count(c => c.FinalSipCode != 200 && c.FinalSipCode != 0)
                })
                .ToList();

            // Common failure codes
            var failureCodes = calls
                .Where(c => c.FinalSipCode != 200 && c.FinalSipCode != 0 && c.FinalSipCode.HasValue)
                .GroupBy(c => c.FinalSipCode)
                .OrderByDescending(g => g.Count())
                .Take(10)
                .Select(g => new
                {
                    sipCode = g.Key,
                    description = GetSipCodeDescription(g.Key ?? 0),
                    count = g.Count()
                })
                .ToList();

            return Ok(new
            {
                summary = new
                {
                    totalCalls,
                    successfulCalls,
                    failedCalls,
                    totalDurationMinutes = Math.Round(totalDurationMinutes, 1),
                    successRate = totalCalls > 0 ? Math.Round((double)successfulCalls / totalCalls * 100, 1) : 0
                },
                callsBySbc,
                dailyTrend,
                failureCodes,
                period = new { fromDate, toDate },
                lastUpdated = DateTime.UtcNow
            });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching Direct Routing call records");
            return Ok(new
            {
                summary = new
                {
                    totalCalls = 0,
                    successfulCalls = 0,
                    failedCalls = 0,
                    totalDurationMinutes = 0.0,
                    successRate = 0.0
                },
                callsBySbc = new List<object>(),
                dailyTrend = new List<object>(),
                failureCodes = new List<object>(),
                period = new { fromDate = DateTime.UtcNow.AddDays(-days), toDate = DateTime.UtcNow },
                lastUpdated = DateTime.UtcNow,
                error = "Unable to fetch Direct Routing data. Ensure you have Direct Routing configured and appropriate permissions.",
                errorDetail = ex.Message
            });
        }
    }

    /// <summary>
    /// Get Auto Attendant call analytics
    /// </summary>
    [HttpGet("auto-attendant")]
    public async Task<IActionResult> GetAutoAttendantAnalytics([FromQuery] int days = 30)
    {
        try
        {
            _logger.LogInformation("Fetching Auto Attendant analytics for last {Days} days", days);

            // Auto Attendant data comes from Teams Analytics reports
            // We need to use the getTeamsUserActivityUserDetail report or call quality dashboard APIs
            
            // For now, we'll try to get this from PSTN calls that went to auto attendants
            var fromDate = DateTime.UtcNow.AddDays(-days);
            var toDate = DateTime.UtcNow;

            var callRecordsResponse = await _graphClient.Communications.CallRecords
                .MicrosoftGraphCallRecordsGetPstnCallsWithFromDateTimeWithToDateTime(fromDate, toDate)
                .GetAsync();

            var calls = callRecordsResponse?.Value ?? new List<Microsoft.Graph.Models.CallRecords.PstnCallLogRow>();

            // Filter for auto attendant calls (these typically have specific indicators)
            var autoAttendantCalls = calls
                .Where(c => c.CallType?.Contains("auto", StringComparison.OrdinalIgnoreCase) == true ||
                            c.CalleeNumber?.Contains("aa", StringComparison.OrdinalIgnoreCase) == true ||
                            c.CallerNumber?.Contains("aa", StringComparison.OrdinalIgnoreCase) == true)
                .ToList();

            var totalCalls = autoAttendantCalls.Count;
            var handledCalls = autoAttendantCalls.Count(c => (c.Duration ?? 0) > 30); // Calls longer than 30 seconds
            var transferredCalls = autoAttendantCalls.Count(c => (c.Duration ?? 0) > 10 && (c.Duration ?? 0) <= 60);
            var abandonedCalls = autoAttendantCalls.Count(c => (c.Duration ?? 0) <= 10);

            return Ok(new
            {
                summary = new
                {
                    totalCalls,
                    handledCalls,
                    transferredCalls,
                    abandonedCalls,
                    averageHandleTimeSeconds = totalCalls > 0 ? Math.Round(autoAttendantCalls.Average(c => c.Duration ?? 0), 0) : 0,
                    abandonmentRate = totalCalls > 0 ? Math.Round((double)abandonedCalls / totalCalls * 100, 1) : 0
                },
                period = new { fromDate, toDate },
                lastUpdated = DateTime.UtcNow,
                note = "Auto Attendant analytics are estimated from PSTN call data. For detailed analytics, use the Teams Admin Center or Call Quality Dashboard."
            });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching Auto Attendant analytics");
            return Ok(new
            {
                summary = new
                {
                    totalCalls = 0,
                    handledCalls = 0,
                    transferredCalls = 0,
                    abandonedCalls = 0,
                    averageHandleTimeSeconds = 0.0,
                    abandonmentRate = 0.0
                },
                period = new { fromDate = DateTime.UtcNow.AddDays(-days), toDate = DateTime.UtcNow },
                lastUpdated = DateTime.UtcNow,
                error = "Unable to fetch Auto Attendant data.",
                errorDetail = ex.Message
            });
        }
    }

    /// <summary>
    /// Get Call Queue analytics
    /// </summary>
    [HttpGet("call-queues")]
    public async Task<IActionResult> GetCallQueueAnalytics([FromQuery] int days = 30)
    {
        try
        {
            _logger.LogInformation("Fetching Call Queue analytics for last {Days} days", days);

            var fromDate = DateTime.UtcNow.AddDays(-days);
            var toDate = DateTime.UtcNow;

            var callRecordsResponse = await _graphClient.Communications.CallRecords
                .MicrosoftGraphCallRecordsGetPstnCallsWithFromDateTimeWithToDateTime(fromDate, toDate)
                .GetAsync();

            var calls = callRecordsResponse?.Value ?? new List<Microsoft.Graph.Models.CallRecords.PstnCallLogRow>();

            // Filter for call queue calls
            var queueCalls = calls
                .Where(c => c.CallType?.Contains("queue", StringComparison.OrdinalIgnoreCase) == true ||
                            c.CalleeNumber?.Contains("cq", StringComparison.OrdinalIgnoreCase) == true)
                .ToList();

            var totalCalls = queueCalls.Count;
            var answeredCalls = queueCalls.Count(c => (c.Duration ?? 0) > 0);
            var abandonedCalls = totalCalls - answeredCalls;
            var averageWaitTime = queueCalls.Any() ? queueCalls.Average(c => c.Duration ?? 0) : 0;

            return Ok(new
            {
                summary = new
                {
                    totalCalls,
                    answeredCalls,
                    abandonedCalls,
                    averageWaitTimeSeconds = Math.Round(averageWaitTime, 0),
                    answerRate = totalCalls > 0 ? Math.Round((double)answeredCalls / totalCalls * 100, 1) : 0,
                    abandonmentRate = totalCalls > 0 ? Math.Round((double)abandonedCalls / totalCalls * 100, 1) : 0
                },
                period = new { fromDate, toDate },
                lastUpdated = DateTime.UtcNow,
                note = "Call Queue analytics are estimated from PSTN call data. For detailed analytics, use the Teams Admin Center or Call Quality Dashboard."
            });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching Call Queue analytics");
            return Ok(new
            {
                summary = new
                {
                    totalCalls = 0,
                    answeredCalls = 0,
                    abandonedCalls = 0,
                    averageWaitTimeSeconds = 0.0,
                    answerRate = 0.0,
                    abandonmentRate = 0.0
                },
                period = new { fromDate = DateTime.UtcNow.AddDays(-days), toDate = DateTime.UtcNow },
                lastUpdated = DateTime.UtcNow,
                error = "Unable to fetch Call Queue data.",
                errorDetail = ex.Message
            });
        }
    }

    /// <summary>
    /// Get combined Teams Phone dashboard data
    /// </summary>
    [HttpGet("dashboard")]
    public async Task<IActionResult> GetDashboard([FromQuery] int days = 30)
    {
        try
        {
            _logger.LogInformation("Fetching Teams Phone dashboard data for last {Days} days", days);

            var fromDate = DateTime.UtcNow.AddDays(-days);
            var toDate = DateTime.UtcNow;

            // Fetch PSTN calls
            List<Microsoft.Graph.Models.CallRecords.PstnCallLogRow> pstnCalls;
            try
            {
                var pstnResponse = await _graphClient.Communications.CallRecords
                    .MicrosoftGraphCallRecordsGetPstnCallsWithFromDateTimeWithToDateTime(fromDate, toDate)
                    .GetAsync();
                pstnCalls = pstnResponse?.Value?.ToList() ?? new List<Microsoft.Graph.Models.CallRecords.PstnCallLogRow>();
            }
            catch
            {
                pstnCalls = new List<Microsoft.Graph.Models.CallRecords.PstnCallLogRow>();
            }

            // Fetch Direct Routing calls
            List<Microsoft.Graph.Models.CallRecords.DirectRoutingLogRow> drCalls;
            try
            {
                var drResponse = await _graphClient.Communications.CallRecords
                    .MicrosoftGraphCallRecordsGetDirectRoutingCallsWithFromDateTimeWithToDateTime(fromDate, toDate)
                    .GetAsync();
                drCalls = drResponse?.Value?.ToList() ?? new List<Microsoft.Graph.Models.CallRecords.DirectRoutingLogRow>();
            }
            catch
            {
                drCalls = new List<Microsoft.Graph.Models.CallRecords.DirectRoutingLogRow>();
            }

            // Calculate PSTN summary
            var pstnTotalCalls = pstnCalls.Count;
            var pstnInbound = pstnCalls.Count(c => c.CallType?.Contains("inbound", StringComparison.OrdinalIgnoreCase) == true);
            var pstnOutbound = pstnCalls.Count(c => c.CallType?.Contains("outbound", StringComparison.OrdinalIgnoreCase) == true);
            var pstnAnswered = pstnCalls.Count(c => (c.Duration ?? 0) > 0);
            var pstnMissed = pstnTotalCalls - pstnAnswered;
            var pstnTotalMinutes = pstnCalls.Sum(c => c.Duration ?? 0) / 60.0;

            // Calculate Direct Routing summary
            var drTotalCalls = drCalls.Count;
            var drSuccessful = drCalls.Count(c => c.FinalSipCode == 200 || c.FinalSipCode == 0);
            var drFailed = drTotalCalls - drSuccessful;
            var drTotalMinutes = drCalls.Sum(c => c.Duration ?? 0) / 60.0;

            // Combined daily trend
            var pstnDaily = pstnCalls
                .Where(c => c.StartDateTime.HasValue)
                .GroupBy(c => c.StartDateTime!.Value.Date)
                .Select(g => new { date = g.Key, pstn = g.Count(), dr = 0 });

            var drDaily = drCalls
                .Where(c => c.StartDateTime.HasValue)
                .GroupBy(c => c.StartDateTime!.Value.Date)
                .Select(g => new { date = g.Key, pstn = 0, dr = g.Count() });

            var combinedDaily = pstnDaily.Concat(drDaily)
                .GroupBy(d => d.date)
                .OrderBy(g => g.Key)
                .Select(g => new
                {
                    date = g.Key,
                    pstnCalls = g.Sum(x => x.pstn),
                    directRoutingCalls = g.Sum(x => x.dr),
                    totalCalls = g.Sum(x => x.pstn) + g.Sum(x => x.dr)
                })
                .ToList();

            // Top users by call volume
            var topUsers = pstnCalls
                .Where(c => !string.IsNullOrEmpty(c.UserPrincipalName))
                .GroupBy(c => new { c.UserPrincipalName, c.UserDisplayName })
                .OrderByDescending(g => g.Count())
                .Take(10)
                .Select(g => new
                {
                    userPrincipalName = g.Key.UserPrincipalName,
                    displayName = g.Key.UserDisplayName ?? g.Key.UserPrincipalName,
                    callCount = g.Count(),
                    totalMinutes = Math.Round(g.Sum(c => c.Duration ?? 0) / 60.0, 1)
                })
                .ToList();

            // Calls by hour
            var callsByHour = pstnCalls
                .Where(c => c.StartDateTime.HasValue)
                .GroupBy(c => c.StartDateTime!.Value.Hour)
                .OrderBy(g => g.Key)
                .Select(g => new { hour = g.Key, count = g.Count() })
                .ToList();

            return Ok(new
            {
                pstn = new
                {
                    totalCalls = pstnTotalCalls,
                    inboundCalls = pstnInbound,
                    outboundCalls = pstnOutbound,
                    answeredCalls = pstnAnswered,
                    missedCalls = pstnMissed,
                    totalMinutes = Math.Round(pstnTotalMinutes, 1),
                    answerRate = pstnTotalCalls > 0 ? Math.Round((double)pstnAnswered / pstnTotalCalls * 100, 1) : 0
                },
                directRouting = new
                {
                    totalCalls = drTotalCalls,
                    successfulCalls = drSuccessful,
                    failedCalls = drFailed,
                    totalMinutes = Math.Round(drTotalMinutes, 1),
                    successRate = drTotalCalls > 0 ? Math.Round((double)drSuccessful / drTotalCalls * 100, 1) : 0
                },
                combined = new
                {
                    totalCalls = pstnTotalCalls + drTotalCalls,
                    totalMinutes = Math.Round(pstnTotalMinutes + drTotalMinutes, 1)
                },
                dailyTrend = combinedDaily,
                topUsers,
                callsByHour,
                period = new { fromDate, toDate },
                lastUpdated = DateTime.UtcNow
            });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching Teams Phone dashboard");
            return StatusCode(500, new { error = "Failed to fetch Teams Phone data", message = ex.Message });
        }
    }

    private static string GetSipCodeDescription(int sipCode)
    {
        return sipCode switch
        {
            400 => "Bad Request",
            401 => "Unauthorized",
            403 => "Forbidden",
            404 => "Not Found",
            408 => "Request Timeout",
            480 => "Temporarily Unavailable",
            486 => "Busy Here",
            487 => "Request Terminated",
            488 => "Not Acceptable Here",
            500 => "Server Internal Error",
            502 => "Bad Gateway",
            503 => "Service Unavailable",
            504 => "Server Time-out",
            _ => $"SIP Code {sipCode}"
        };
    }
}
