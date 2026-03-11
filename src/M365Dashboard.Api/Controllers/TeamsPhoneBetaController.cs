using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Identity.Web;
using Azure.Identity;
using System.Net.Http.Headers;
using System.Text.Json;

namespace M365Dashboard.Api.Controllers;

[ApiController]
[Route("api/teamsphone/beta")]
[Authorize]
public class TeamsPhoneBetaController : ControllerBase
{
    private readonly ITokenAcquisition _tokenAcquisition;
    private readonly IHttpClientFactory _httpClientFactory;
    private readonly ILogger<TeamsPhoneBetaController> _logger;
    private readonly IConfiguration _configuration;

    public TeamsPhoneBetaController(
        ITokenAcquisition tokenAcquisition,
        IHttpClientFactory httpClientFactory,
        ILogger<TeamsPhoneBetaController> logger,
        IConfiguration configuration)
    {
        _tokenAcquisition = tokenAcquisition;
        _httpClientFactory = httpClientFactory;
        _logger = logger;
        _configuration = configuration;
    }

    /// <summary>
    /// Get detailed call records using Beta API
    /// </summary>
    [HttpGet("call-records")]
    public async Task<IActionResult> GetCallRecords([FromQuery] int days = 30)
    {
        try
        {
            _logger.LogInformation("Fetching call records from Beta API for last {Days} days", days);

            var fromDate = DateTime.UtcNow.AddDays(-days);
            var toDate = DateTime.UtcNow;

            // Get access token for Graph API
            var scopes = new[] { "https://graph.microsoft.com/.default" };
            var accessToken = await _tokenAcquisition.GetAccessTokenForAppAsync(scopes[0]);

            var client = _httpClientFactory.CreateClient();
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

            // Use Beta endpoint for call records with more details
            var url = $"https://graph.microsoft.com/beta/communications/callRecords?" +
                      $"$filter=startDateTime ge {fromDate:yyyy-MM-ddTHH:mm:ssZ} and startDateTime le {toDate:yyyy-MM-ddTHH:mm:ssZ}" +
                      $"&$orderby=startDateTime desc" +
                      $"&$top=100";

            var response = await client.GetAsync(url);
            
            if (!response.IsSuccessStatusCode)
            {
                var error = await response.Content.ReadAsStringAsync();
                _logger.LogWarning("Beta API call failed: {StatusCode} - {Error}", response.StatusCode, error);
                
                return Ok(new
                {
                    callRecords = new List<object>(),
                    totalCount = 0,
                    period = new { fromDate, toDate },
                    lastUpdated = DateTime.UtcNow,
                    error = "Unable to fetch call records from Beta API",
                    errorDetail = error
                });
            }

            var content = await response.Content.ReadAsStringAsync();
            var jsonDoc = JsonDocument.Parse(content);
            var records = new List<object>();

            if (jsonDoc.RootElement.TryGetProperty("value", out var valueArray))
            {
                foreach (var record in valueArray.EnumerateArray())
                {
                    records.Add(ParseCallRecord(record));
                }
            }

            return Ok(new
            {
                callRecords = records,
                totalCount = records.Count,
                period = new { fromDate, toDate },
                lastUpdated = DateTime.UtcNow
            });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching call records from Beta API");
            return Ok(new
            {
                callRecords = new List<object>(),
                totalCount = 0,
                period = new { fromDate = DateTime.UtcNow.AddDays(-days), toDate = DateTime.UtcNow },
                lastUpdated = DateTime.UtcNow,
                error = "Failed to fetch call records",
                errorDetail = ex.Message
            });
        }
    }

    /// <summary>
    /// Get a specific call record with full session details
    /// </summary>
    [HttpGet("call-records/{callId}")]
    public async Task<IActionResult> GetCallRecordDetails(string callId)
    {
        try
        {
            _logger.LogInformation("Fetching call record details for {CallId}", callId);

            var scopes = new[] { "https://graph.microsoft.com/.default" };
            var accessToken = await _tokenAcquisition.GetAccessTokenForAppAsync(scopes[0]);

            var client = _httpClientFactory.CreateClient();
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

            // Get call record with expanded sessions and segments
            var url = $"https://graph.microsoft.com/beta/communications/callRecords/{callId}?$expand=sessions($expand=segments)";

            var response = await client.GetAsync(url);
            
            if (!response.IsSuccessStatusCode)
            {
                var error = await response.Content.ReadAsStringAsync();
                _logger.LogWarning("Beta API call failed: {StatusCode} - {Error}", response.StatusCode, error);
                return NotFound(new { error = "Call record not found", errorDetail = error });
            }

            var content = await response.Content.ReadAsStringAsync();
            var jsonDoc = JsonDocument.Parse(content);
            var record = ParseDetailedCallRecord(jsonDoc.RootElement);

            return Ok(record);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching call record details for {CallId}", callId);
            return StatusCode(500, new { error = "Failed to fetch call record details", errorDetail = ex.Message });
        }
    }

    /// <summary>
    /// Get PSTN blocked users report (Beta)
    /// </summary>
    [HttpGet("pstn-blocked-users")]
    public async Task<IActionResult> GetPstnBlockedUsers()
    {
        try
        {
            var scopes = new[] { "https://graph.microsoft.com/.default" };
            var accessToken = await _tokenAcquisition.GetAccessTokenForAppAsync(scopes[0]);

            var client = _httpClientFactory.CreateClient();
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

            var url = "https://graph.microsoft.com/beta/communications/callRecords/getPstnBlockedUsersLog(fromDateTime=2024-01-01T00:00:00Z,toDateTime=2024-12-31T23:59:59Z)";

            var response = await client.GetAsync(url);
            
            if (!response.IsSuccessStatusCode)
            {
                return Ok(new { blockedUsers = new List<object>(), error = "Unable to fetch blocked users" });
            }

            var content = await response.Content.ReadAsStringAsync();
            return Ok(JsonDocument.Parse(content).RootElement);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching PSTN blocked users");
            return Ok(new { blockedUsers = new List<object>(), error = ex.Message });
        }
    }

    /// <summary>
    /// Get call quality metrics (Beta) - SMS logs
    /// </summary>
    [HttpGet("sms-logs")]
    public async Task<IActionResult> GetSmsLogs([FromQuery] int days = 30)
    {
        try
        {
            var fromDate = DateTime.UtcNow.AddDays(-days);
            var toDate = DateTime.UtcNow;

            var scopes = new[] { "https://graph.microsoft.com/.default" };
            var accessToken = await _tokenAcquisition.GetAccessTokenForAppAsync(scopes[0]);

            var client = _httpClientFactory.CreateClient();
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

            var url = $"https://graph.microsoft.com/beta/communications/callRecords/getSmsLog(fromDateTime={fromDate:yyyy-MM-ddTHH:mm:ssZ},toDateTime={toDate:yyyy-MM-ddTHH:mm:ssZ})";

            var response = await client.GetAsync(url);
            
            if (!response.IsSuccessStatusCode)
            {
                var error = await response.Content.ReadAsStringAsync();
                return Ok(new { 
                    smsLogs = new List<object>(), 
                    totalCount = 0,
                    period = new { fromDate, toDate },
                    error = "Unable to fetch SMS logs. This feature requires Teams Phone with SMS capability.",
                    errorDetail = error
                });
            }

            var content = await response.Content.ReadAsStringAsync();
            var jsonDoc = JsonDocument.Parse(content);
            var logs = new List<object>();

            if (jsonDoc.RootElement.TryGetProperty("value", out var valueArray))
            {
                foreach (var log in valueArray.EnumerateArray())
                {
                    logs.Add(new
                    {
                        id = GetStringProperty(log, "id"),
                        sentDateTime = GetStringProperty(log, "sentDateTime"),
                        userPrincipalName = GetStringProperty(log, "userPrincipalName"),
                        userDisplayName = GetStringProperty(log, "userDisplayName"),
                        destinationNumber = GetStringProperty(log, "destinationNumber"),
                        sourceNumber = GetStringProperty(log, "sourceNumber"),
                        smsType = GetStringProperty(log, "smsType"),
                        callCharge = GetDecimalProperty(log, "callCharge"),
                        currency = GetStringProperty(log, "currency"),
                        destinationContext = GetStringProperty(log, "destinationContext"),
                        destinationName = GetStringProperty(log, "destinationName")
                    });
                }
            }

            return Ok(new
            {
                smsLogs = logs,
                totalCount = logs.Count,
                period = new { fromDate, toDate },
                lastUpdated = DateTime.UtcNow
            });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching SMS logs");
            return Ok(new { 
                smsLogs = new List<object>(), 
                totalCount = 0,
                error = ex.Message 
            });
        }
    }

    /// <summary>
    /// Get enhanced PSTN call logs with more details (Beta)
    /// </summary>
    [HttpGet("pstn-calls-enhanced")]
    public async Task<IActionResult> GetEnhancedPstnCalls([FromQuery] int days = 30)
    {
        try
        {
            var fromDate = DateTime.UtcNow.AddDays(-days);
            var toDate = DateTime.UtcNow;

            var scopes = new[] { "https://graph.microsoft.com/.default" };
            var accessToken = await _tokenAcquisition.GetAccessTokenForAppAsync(scopes[0]);

            var client = _httpClientFactory.CreateClient();
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

            // Beta endpoint with more properties
            var url = $"https://graph.microsoft.com/beta/communications/callRecords/getPstnCalls(fromDateTime={fromDate:yyyy-MM-ddTHH:mm:ssZ},toDateTime={toDate:yyyy-MM-ddTHH:mm:ssZ})";

            var response = await client.GetAsync(url);
            
            if (!response.IsSuccessStatusCode)
            {
                var error = await response.Content.ReadAsStringAsync();
                return Ok(new
                {
                    calls = new List<object>(),
                    summary = new { totalCalls = 0 },
                    period = new { fromDate, toDate },
                    error = "Unable to fetch enhanced PSTN calls",
                    errorDetail = error
                });
            }

            var content = await response.Content.ReadAsStringAsync();
            var jsonDoc = JsonDocument.Parse(content);
            var calls = new List<object>();

            if (jsonDoc.RootElement.TryGetProperty("value", out var valueArray))
            {
                foreach (var call in valueArray.EnumerateArray())
                {
                    calls.Add(new
                    {
                        id = GetStringProperty(call, "id"),
                        callId = GetStringProperty(call, "callId"),
                        userId = GetStringProperty(call, "userId"),
                        userPrincipalName = GetStringProperty(call, "userPrincipalName"),
                        userDisplayName = GetStringProperty(call, "userDisplayName"),
                        startDateTime = GetStringProperty(call, "startDateTime"),
                        endDateTime = GetStringProperty(call, "endDateTime"),
                        duration = GetIntProperty(call, "duration"),
                        charge = GetDecimalProperty(call, "charge"),
                        callType = GetStringProperty(call, "callType"),
                        currency = GetStringProperty(call, "currency"),
                        calleeNumber = GetStringProperty(call, "calleeNumber"),
                        callerNumber = GetStringProperty(call, "callerNumber"),
                        destinationContext = GetStringProperty(call, "destinationContext"),
                        destinationName = GetStringProperty(call, "destinationName"),
                        conferenceId = GetStringProperty(call, "conferenceId"),
                        licenseCapability = GetStringProperty(call, "licenseCapability"),
                        inventoryType = GetStringProperty(call, "inventoryType"),
                        operator_ = GetStringProperty(call, "operator"),
                        callDurationSource = GetStringProperty(call, "callDurationSource"),
                        // Beta-specific fields
                        tenantCountryCode = GetStringProperty(call, "tenantCountryCode"),
                        usageCountryCode = GetStringProperty(call, "usageCountryCode"),
                        connectionCharge = GetDecimalProperty(call, "connectionCharge"),
                        otherPartyCountryCode = GetStringProperty(call, "otherPartyCountryCode"),
                        administrativeUnit = GetStringProperty(call, "administrativeUnit"),
                        clientLocalIpV4Address = GetStringProperty(call, "clientLocalIpV4Address"),
                        clientLocalIpV6Address = GetStringProperty(call, "clientLocalIpV6Address"),
                        clientPublicIpV4Address = GetStringProperty(call, "clientPublicIpV4Address"),
                        clientPublicIpV6Address = GetStringProperty(call, "clientPublicIpV6Address"),
                        isAnswered = GetIntProperty(call, "duration") > 0
                    });
                }
            }

            // Calculate summary
            var totalCalls = calls.Count;
            var answeredCalls = calls.Count(c => ((dynamic)c).isAnswered);
            var totalDuration = calls.Sum(c => (int)((dynamic)c).duration);
            var totalCharge = calls.Sum(c => 
            {
                var chargeValue = ((dynamic)c).charge;
                return chargeValue is decimal d ? d : 0m;
            });

            return Ok(new
            {
                calls,
                summary = new
                {
                    totalCalls,
                    answeredCalls,
                    missedCalls = totalCalls - answeredCalls,
                    totalDurationSeconds = totalDuration,
                    totalDurationMinutes = Math.Round(totalDuration / 60.0, 1),
                    totalCharge,
                    answerRate = totalCalls > 0 ? Math.Round((double)answeredCalls / totalCalls * 100, 1) : 0
                },
                period = new { fromDate, toDate },
                lastUpdated = DateTime.UtcNow
            });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching enhanced PSTN calls");
            return Ok(new
            {
                calls = new List<object>(),
                summary = new { totalCalls = 0 },
                period = new { fromDate = DateTime.UtcNow.AddDays(-days), toDate = DateTime.UtcNow },
                error = ex.Message
            });
        }
    }

    /// <summary>
    /// Get Direct Routing calls with enhanced details (Beta)
    /// </summary>
    [HttpGet("direct-routing-enhanced")]
    public async Task<IActionResult> GetEnhancedDirectRoutingCalls([FromQuery] int days = 30)
    {
        try
        {
            var fromDate = DateTime.UtcNow.AddDays(-days);
            var toDate = DateTime.UtcNow;

            var scopes = new[] { "https://graph.microsoft.com/.default" };
            var accessToken = await _tokenAcquisition.GetAccessTokenForAppAsync(scopes[0]);

            var client = _httpClientFactory.CreateClient();
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

            var url = $"https://graph.microsoft.com/beta/communications/callRecords/getDirectRoutingCalls(fromDateTime={fromDate:yyyy-MM-ddTHH:mm:ssZ},toDateTime={toDate:yyyy-MM-ddTHH:mm:ssZ})";

            var response = await client.GetAsync(url);
            
            if (!response.IsSuccessStatusCode)
            {
                var error = await response.Content.ReadAsStringAsync();
                return Ok(new
                {
                    calls = new List<object>(),
                    summary = new { totalCalls = 0 },
                    period = new { fromDate, toDate },
                    error = "Unable to fetch Direct Routing calls",
                    errorDetail = error
                });
            }

            var content = await response.Content.ReadAsStringAsync();
            var jsonDoc = JsonDocument.Parse(content);
            var calls = new List<object>();

            if (jsonDoc.RootElement.TryGetProperty("value", out var valueArray))
            {
                foreach (var call in valueArray.EnumerateArray())
                {
                    calls.Add(new
                    {
                        id = GetStringProperty(call, "id"),
                        correlationId = GetStringProperty(call, "correlationId"),
                        userId = GetStringProperty(call, "userId"),
                        userPrincipalName = GetStringProperty(call, "userPrincipalName"),
                        userDisplayName = GetStringProperty(call, "userDisplayName"),
                        startDateTime = GetStringProperty(call, "startDateTime"),
                        endDateTime = GetStringProperty(call, "endDateTime"),
                        inviteDateTime = GetStringProperty(call, "inviteDateTime"),
                        failureDateTime = GetStringProperty(call, "failureDateTime"),
                        duration = GetIntProperty(call, "duration"),
                        callType = GetStringProperty(call, "callType"),
                        calleeNumber = GetStringProperty(call, "calleeNumber"),
                        callerNumber = GetStringProperty(call, "callerNumber"),
                        // SBC/Trunk details
                        trunkFullyQualifiedDomainName = GetStringProperty(call, "trunkFullyQualifiedDomainName"),
                        mediaBypassEnabled = GetBoolProperty(call, "mediaBypassEnabled"),
                        // SIP details
                        finalSipCode = GetIntProperty(call, "finalSipCode"),
                        finalSipCodePhrase = GetStringProperty(call, "finalSipCodePhrase"),
                        sipResponseSubCode = GetIntProperty(call, "sipResponseSubCode"),
                        // Call quality
                        mediaPathLocation = GetStringProperty(call, "mediaPathLocation"),
                        signalingLocation = GetStringProperty(call, "signalingLocation"),
                        // Additional Beta fields
                        userCountryCode = GetStringProperty(call, "userCountryCode"),
                        otherPartyCountryCode = GetStringProperty(call, "otherPartyCountryCode"),
                        administrativeUnit = GetStringProperty(call, "administrativeUnit"),
                        isSuccess = GetIntProperty(call, "finalSipCode") == 200 || GetIntProperty(call, "finalSipCode") == 0
                    });
                }
            }

            // Calculate summary
            var totalCalls = calls.Count;
            var successfulCalls = calls.Count(c => ((dynamic)c).isSuccess);
            var totalDuration = calls.Sum(c => (int)((dynamic)c).duration);

            // Group by SBC
            var sbcStats = calls
                .GroupBy(c => ((dynamic)c).trunkFullyQualifiedDomainName ?? "Unknown")
                .Select(g => new
                {
                    sbc = g.Key,
                    totalCalls = g.Count(),
                    successfulCalls = g.Count(c => ((dynamic)c).isSuccess),
                    failedCalls = g.Count(c => !((dynamic)c).isSuccess),
                    totalMinutes = Math.Round(g.Sum(c => (int)((dynamic)c).duration) / 60.0, 1)
                })
                .OrderByDescending(s => s.totalCalls)
                .ToList();

            // Common SIP failure codes
            var sipFailures = calls
                .Where(c => !((dynamic)c).isSuccess && ((dynamic)c).finalSipCode > 0)
                .GroupBy(c => new { code = ((dynamic)c).finalSipCode, phrase = ((dynamic)c).finalSipCodePhrase })
                .Select(g => new
                {
                    sipCode = g.Key.code,
                    sipCodePhrase = g.Key.phrase ?? GetSipCodeDescription((int)g.Key.code),
                    count = g.Count()
                })
                .OrderByDescending(s => s.count)
                .Take(10)
                .ToList();

            return Ok(new
            {
                calls,
                summary = new
                {
                    totalCalls,
                    successfulCalls,
                    failedCalls = totalCalls - successfulCalls,
                    totalDurationSeconds = totalDuration,
                    totalDurationMinutes = Math.Round(totalDuration / 60.0, 1),
                    successRate = totalCalls > 0 ? Math.Round((double)successfulCalls / totalCalls * 100, 1) : 0
                },
                sbcStats,
                sipFailures,
                period = new { fromDate, toDate },
                lastUpdated = DateTime.UtcNow
            });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching enhanced Direct Routing calls");
            return Ok(new
            {
                calls = new List<object>(),
                summary = new { totalCalls = 0 },
                period = new { fromDate = DateTime.UtcNow.AddDays(-days), toDate = DateTime.UtcNow },
                error = ex.Message
            });
        }
    }

    // Helper methods for parsing JSON
    private object ParseCallRecord(JsonElement element)
    {
        return new
        {
            id = GetStringProperty(element, "id"),
            version = GetIntProperty(element, "version"),
            type = GetStringProperty(element, "type"),
            modalities = GetArrayProperty(element, "modalities"),
            startDateTime = GetStringProperty(element, "startDateTime"),
            endDateTime = GetStringProperty(element, "endDateTime"),
            joinWebUrl = GetStringProperty(element, "joinWebUrl"),
            organizer = GetParticipant(element, "organizer"),
            participants = GetParticipants(element)
        };
    }

    private object ParseDetailedCallRecord(JsonElement element)
    {
        var sessions = new List<object>();
        if (element.TryGetProperty("sessions", out var sessionsArray))
        {
            foreach (var session in sessionsArray.EnumerateArray())
            {
                var segments = new List<object>();
                if (session.TryGetProperty("segments", out var segmentsArray))
                {
                    foreach (var segment in segmentsArray.EnumerateArray())
                    {
                        segments.Add(new
                        {
                            id = GetStringProperty(segment, "id"),
                            startDateTime = GetStringProperty(segment, "startDateTime"),
                            endDateTime = GetStringProperty(segment, "endDateTime"),
                            caller = GetEndpoint(segment, "caller"),
                            callee = GetEndpoint(segment, "callee"),
                            failureInfo = GetFailureInfo(segment),
                            media = GetMediaStats(segment)
                        });
                    }
                }

                sessions.Add(new
                {
                    id = GetStringProperty(session, "id"),
                    modalities = GetArrayProperty(session, "modalities"),
                    startDateTime = GetStringProperty(session, "startDateTime"),
                    endDateTime = GetStringProperty(session, "endDateTime"),
                    caller = GetEndpoint(session, "caller"),
                    callee = GetEndpoint(session, "callee"),
                    failureInfo = GetFailureInfo(session),
                    segments
                });
            }
        }

        return new
        {
            id = GetStringProperty(element, "id"),
            version = GetIntProperty(element, "version"),
            type = GetStringProperty(element, "type"),
            modalities = GetArrayProperty(element, "modalities"),
            startDateTime = GetStringProperty(element, "startDateTime"),
            endDateTime = GetStringProperty(element, "endDateTime"),
            joinWebUrl = GetStringProperty(element, "joinWebUrl"),
            organizer = GetParticipant(element, "organizer"),
            participants = GetParticipants(element),
            sessions
        };
    }

    private object? GetParticipant(JsonElement element, string propertyName)
    {
        if (!element.TryGetProperty(propertyName, out var participant))
            return null;

        if (participant.TryGetProperty("user", out var user))
        {
            return new
            {
                id = GetStringProperty(user, "id"),
                displayName = GetStringProperty(user, "displayName"),
                tenantId = GetStringProperty(user, "tenantId")
            };
        }
        return null;
    }

    private List<object> GetParticipants(JsonElement element)
    {
        var result = new List<object>();
        if (element.TryGetProperty("participants", out var participants))
        {
            foreach (var participant in participants.EnumerateArray())
            {
                if (participant.TryGetProperty("user", out var user))
                {
                    result.Add(new
                    {
                        id = GetStringProperty(user, "id"),
                        displayName = GetStringProperty(user, "displayName"),
                        tenantId = GetStringProperty(user, "tenantId")
                    });
                }
            }
        }
        return result;
    }

    private object? GetEndpoint(JsonElement element, string propertyName)
    {
        if (!element.TryGetProperty(propertyName, out var endpoint))
            return null;

        return new
        {
            userAgent = GetUserAgent(endpoint),
            identity = GetIdentity(endpoint)
        };
    }

    private object? GetUserAgent(JsonElement element)
    {
        if (!element.TryGetProperty("userAgent", out var userAgent))
            return null;

        return new
        {
            headerValue = GetStringProperty(userAgent, "headerValue"),
            applicationVersion = GetStringProperty(userAgent, "applicationVersion"),
            platform = GetStringProperty(userAgent, "platform"),
            productFamily = GetStringProperty(userAgent, "productFamily")
        };
    }

    private object? GetIdentity(JsonElement element)
    {
        if (!element.TryGetProperty("identity", out var identity))
            return null;

        if (identity.TryGetProperty("user", out var user))
        {
            return new
            {
                id = GetStringProperty(user, "id"),
                displayName = GetStringProperty(user, "displayName"),
                tenantId = GetStringProperty(user, "tenantId")
            };
        }
        return null;
    }

    private object? GetFailureInfo(JsonElement element)
    {
        if (!element.TryGetProperty("failureInfo", out var failureInfo))
            return null;

        return new
        {
            reason = GetStringProperty(failureInfo, "reason"),
            stage = GetStringProperty(failureInfo, "stage")
        };
    }

    private List<object> GetMediaStats(JsonElement element)
    {
        var result = new List<object>();
        if (element.TryGetProperty("media", out var mediaArray))
        {
            foreach (var media in mediaArray.EnumerateArray())
            {
                result.Add(new
                {
                    label = GetStringProperty(media, "label"),
                    callerNetwork = GetNetworkInfo(media, "callerNetwork"),
                    calleeNetwork = GetNetworkInfo(media, "calleeNetwork"),
                    streams = GetStreams(media)
                });
            }
        }
        return result;
    }

    private object? GetNetworkInfo(JsonElement element, string propertyName)
    {
        if (!element.TryGetProperty(propertyName, out var network))
            return null;

        return new
        {
            connectionType = GetStringProperty(network, "connectionType"),
            wifiBand = GetStringProperty(network, "wifiBand"),
            basicServiceSetIdentifier = GetStringProperty(network, "basicServiceSetIdentifier"),
            wifiRadioType = GetStringProperty(network, "wifiRadioType"),
            wifiSignalStrength = GetIntProperty(network, "wifiSignalStrength"),
            bandwidthLowEventRatio = GetDecimalProperty(network, "bandwidthLowEventRatio")
        };
    }

    private List<object> GetStreams(JsonElement element)
    {
        var result = new List<object>();
        if (element.TryGetProperty("streams", out var streams))
        {
            foreach (var stream in streams.EnumerateArray())
            {
                result.Add(new
                {
                    streamId = GetStringProperty(stream, "streamId"),
                    streamDirection = GetStringProperty(stream, "streamDirection"),
                    averageAudioDegradation = GetDecimalProperty(stream, "averageAudioDegradation"),
                    averageJitter = GetStringProperty(stream, "averageJitter"),
                    maxJitter = GetStringProperty(stream, "maxJitter"),
                    averagePacketLossRate = GetDecimalProperty(stream, "averagePacketLossRate"),
                    maxPacketLossRate = GetDecimalProperty(stream, "maxPacketLossRate"),
                    averageRoundTripTime = GetStringProperty(stream, "averageRoundTripTime"),
                    maxRoundTripTime = GetStringProperty(stream, "maxRoundTripTime"),
                    packetUtilization = GetIntProperty(stream, "packetUtilization"),
                    averageBandwidthEstimate = GetIntProperty(stream, "averageBandwidthEstimate"),
                    wasMediaBypassed = GetBoolProperty(stream, "wasMediaBypassed")
                });
            }
        }
        return result;
    }

    private static string? GetStringProperty(JsonElement element, string propertyName)
    {
        if (element.TryGetProperty(propertyName, out var property) && property.ValueKind == JsonValueKind.String)
            return property.GetString();
        return null;
    }

    private static int GetIntProperty(JsonElement element, string propertyName)
    {
        if (element.TryGetProperty(propertyName, out var property) && property.ValueKind == JsonValueKind.Number)
            return property.GetInt32();
        return 0;
    }

    private static decimal GetDecimalProperty(JsonElement element, string propertyName)
    {
        if (element.TryGetProperty(propertyName, out var property) && property.ValueKind == JsonValueKind.Number)
            return property.GetDecimal();
        return 0m;
    }

    private static bool GetBoolProperty(JsonElement element, string propertyName)
    {
        if (element.TryGetProperty(propertyName, out var property))
        {
            if (property.ValueKind == JsonValueKind.True) return true;
            if (property.ValueKind == JsonValueKind.False) return false;
        }
        return false;
    }

    private static List<string> GetArrayProperty(JsonElement element, string propertyName)
    {
        var result = new List<string>();
        if (element.TryGetProperty(propertyName, out var array) && array.ValueKind == JsonValueKind.Array)
        {
            foreach (var item in array.EnumerateArray())
            {
                if (item.ValueKind == JsonValueKind.String)
                    result.Add(item.GetString() ?? "");
            }
        }
        return result;
    }

    private static string GetSipCodeDescription(int sipCode)
    {
        return sipCode switch
        {
            200 => "OK",
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
