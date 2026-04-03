using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using SecurityAlert = Microsoft.Graph.Models.Security.Alert;

namespace M365Dashboard.Api.Controllers;

[ApiController]
[Route("api/[controller]")]
[Authorize]
public class ThreatIntelligenceController : ControllerBase
{
    private readonly GraphServiceClient _graphClient;
    private readonly ILogger<ThreatIntelligenceController> _logger;

    public ThreatIntelligenceController(GraphServiceClient graphClient, ILogger<ThreatIntelligenceController> logger)
    {
        _graphClient = graphClient;
        _logger = logger;
    }

    /// <summary>
    /// Get security alerts from Microsoft Graph Security API
    /// </summary>
    [HttpGet("alerts")]
    public async Task<IActionResult> GetSecurityAlerts([FromQuery] int days = 30)
    {
        try
        {
            var fromDate = DateTime.UtcNow.AddDays(-days);

            var alerts = await _graphClient.Security.Alerts_v2
                .GetAsync(config =>
                {
                    config.QueryParameters.Filter = $"createdDateTime ge {fromDate:yyyy-MM-ddTHH:mm:ssZ}";
                    config.QueryParameters.Top = 200;
                    config.QueryParameters.Orderby = new[] { "createdDateTime desc" };
                });

            var alertList = alerts?.Value ?? new List<SecurityAlert>();

            var result = alertList.Select(a => new
            {
                id = a.Id,
                title = a.Title,
                description = a.Description,
                severity = a.Severity?.ToString(),
                status = a.Status?.ToString(),
                classification = a.Classification?.ToString(),
                determination = a.Determination?.ToString(),
                category = a.Category,
                serviceSources = a.ServiceSource?.ToString(),
                detectorId = a.DetectorId,
                createdDateTime = a.CreatedDateTime,
                lastUpdateDateTime = a.LastUpdateDateTime,
                resolvedDateTime = a.ResolvedDateTime,
                firstActivityDateTime = a.FirstActivityDateTime,
                lastActivityDateTime = a.LastActivityDateTime,
                assignedTo = a.AssignedTo,
                incidentId = a.IncidentId,
                tenantId = a.TenantId,
                threatDisplayName = a.ThreatDisplayName,
                threatFamilyName = a.ThreatFamilyName,
                evidence = a.Evidence?.Take(5).Select(e => new
                {
                    type = e.OdataType,
                    createdDateTime = e.CreatedDateTime,
                    verdict = e.Verdict?.ToString(),
                    remediationStatus = e.RemediationStatus?.ToString()
                }).ToList()
            }).ToList();

            // Summary by severity
            var severitySummary = result
                .GroupBy(a => a.severity ?? "Unknown")
                .Select(g => new { severity = g.Key, count = g.Count() })
                .OrderByDescending(x => x.severity == "High" ? 3 : x.severity == "Medium" ? 2 : x.severity == "Low" ? 1 : 0)
                .ToList();

            // Summary by status
            var statusSummary = result
                .GroupBy(a => a.status ?? "Unknown")
                .Select(g => new { status = g.Key, count = g.Count() })
                .ToList();

            // Summary by category
            var categorySummary = result
                .GroupBy(a => a.category ?? "Unknown")
                .Select(g => new { category = g.Key, count = g.Count() })
                .OrderByDescending(x => x.count)
                .Take(10)
                .ToList();

            // Summary by service source
            var serviceSummary = result
                .GroupBy(a => a.serviceSources ?? "Unknown")
                .Select(g => new { service = g.Key, count = g.Count() })
                .OrderByDescending(x => x.count)
                .ToList();

            return Ok(new
            {
                alerts = result.Take(100),
                summary = new
                {
                    totalAlerts = result.Count,
                    highSeverity = result.Count(a => a.severity == "High"),
                    mediumSeverity = result.Count(a => a.severity == "Medium"),
                    lowSeverity = result.Count(a => a.severity == "Low"),
                    newAlerts = result.Count(a => a.status == "New"),
                    inProgressAlerts = result.Count(a => a.status == "InProgress"),
                    resolvedAlerts = result.Count(a => a.status == "Resolved")
                },
                severitySummary,
                statusSummary,
                categorySummary,
                serviceSummary,
                period = new { fromDate, toDate = DateTime.UtcNow },
                lastUpdated = DateTime.UtcNow
            });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching security alerts");
            return Ok(new
            {
                alerts = new List<object>(),
                summary = new { totalAlerts = 0 },
                error = ex.Message
            });
        }
    }

    /// <summary>
    /// Get risky users from Identity Protection
    /// </summary>
    [HttpGet("risky-users")]
    public async Task<IActionResult> GetRiskyUsers()
    {
        try
        {
            var riskyUsers = await _graphClient.IdentityProtection.RiskyUsers
                .GetAsync(config =>
                {
                    config.QueryParameters.Top = 100;
                    config.QueryParameters.Orderby = new[] { "riskLastUpdatedDateTime desc" };
                });

            var userList = riskyUsers?.Value ?? new List<RiskyUser>();

            var result = userList.Select(u => new
            {
                id = u.Id,
                userPrincipalName = u.UserPrincipalName,
                userDisplayName = u.UserDisplayName,
                riskLevel = u.RiskLevel?.ToString(),
                riskState = u.RiskState?.ToString(),
                riskDetail = u.RiskDetail?.ToString(),
                riskLastUpdatedDateTime = u.RiskLastUpdatedDateTime,
                isDeleted = u.IsDeleted,
                isProcessing = u.IsProcessing
            }).ToList();

            // Summary by risk level
            var riskLevelSummary = result
                .GroupBy(u => u.riskLevel ?? "Unknown")
                .Select(g => new { riskLevel = g.Key, count = g.Count() })
                .ToList();

            // Summary by risk state
            var riskStateSummary = result
                .GroupBy(u => u.riskState ?? "Unknown")
                .Select(g => new { riskState = g.Key, count = g.Count() })
                .ToList();

            return Ok(new
            {
                users = result,
                summary = new
                {
                    totalRiskyUsers = result.Count,
                    highRisk = result.Count(u => u.riskLevel == "High"),
                    mediumRisk = result.Count(u => u.riskLevel == "Medium"),
                    lowRisk = result.Count(u => u.riskLevel == "Low"),
                    atRisk = result.Count(u => u.riskState == "AtRisk"),
                    confirmedCompromised = result.Count(u => u.riskState == "ConfirmedCompromised")
                },
                riskLevelSummary,
                riskStateSummary,
                lastUpdated = DateTime.UtcNow
            });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching risky users");
            return Ok(new
            {
                users = new List<object>(),
                summary = new { totalRiskyUsers = 0 },
                error = ex.Message
            });
        }
    }

    /// <summary>
    /// Get risk detections from Identity Protection
    /// </summary>
    [HttpGet("risk-detections")]
    public async Task<IActionResult> GetRiskDetections([FromQuery] int days = 30)
    {
        try
        {
            var fromDate = DateTime.UtcNow.AddDays(-days);

            var detections = await _graphClient.IdentityProtection.RiskDetections
                .GetAsync(config =>
                {
                    config.QueryParameters.Filter = $"detectedDateTime ge {fromDate:yyyy-MM-ddTHH:mm:ssZ}";
                    config.QueryParameters.Top = 200;
                    config.QueryParameters.Orderby = new[] { "detectedDateTime desc" };
                });

            var detectionList = detections?.Value ?? new List<RiskDetection>();

            var result = detectionList.Select(d => new
            {
                id = d.Id,
                userPrincipalName = d.UserPrincipalName,
                userDisplayName = d.UserDisplayName,
                riskType = d.RiskEventType,
                riskLevel = d.RiskLevel?.ToString(),
                riskState = d.RiskState?.ToString(),
                riskDetail = d.RiskDetail?.ToString(),
                detectedDateTime = d.DetectedDateTime,
                lastUpdatedDateTime = d.LastUpdatedDateTime,
                ipAddress = d.IpAddress,
                location = d.Location != null ? new
                {
                    city = d.Location.City,
                    state = d.Location.State,
                    country = d.Location.CountryOrRegion
                } : null,
                activityType = d.Activity?.ToString(),
                source = d.Source,
                detectionTimingType = d.DetectionTimingType?.ToString(),
                tokenIssuerType = d.TokenIssuerType?.ToString()
            }).ToList();

            // Summary by risk type
            var riskTypeSummary = result
                .GroupBy(d => d.riskType ?? "Unknown")
                .Select(g => new { riskType = g.Key, count = g.Count() })
                .OrderByDescending(x => x.count)
                .ToList();

            // Summary by risk level
            var riskLevelSummary = result
                .GroupBy(d => d.riskLevel ?? "Unknown")
                .Select(g => new { riskLevel = g.Key, count = g.Count() })
                .ToList();

            // Daily trend
            var dailyTrend = result
                .Where(d => d.detectedDateTime.HasValue)
                .GroupBy(d => d.detectedDateTime!.Value.Date)
                .OrderBy(g => g.Key)
                .Select(g => new
                {
                    date = g.Key,
                    count = g.Count(),
                    high = g.Count(d => d.riskLevel == "High"),
                    medium = g.Count(d => d.riskLevel == "Medium"),
                    low = g.Count(d => d.riskLevel == "Low")
                })
                .ToList();

            return Ok(new
            {
                detections = result.Take(100),
                summary = new
                {
                    totalDetections = result.Count,
                    highRisk = result.Count(d => d.riskLevel == "High"),
                    mediumRisk = result.Count(d => d.riskLevel == "Medium"),
                    lowRisk = result.Count(d => d.riskLevel == "Low"),
                    uniqueUsers = result.Select(d => d.userPrincipalName).Distinct().Count()
                },
                riskTypeSummary,
                riskLevelSummary,
                dailyTrend,
                period = new { fromDate, toDate = DateTime.UtcNow },
                lastUpdated = DateTime.UtcNow
            });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching risk detections");
            return Ok(new
            {
                detections = new List<object>(),
                summary = new { totalDetections = 0 },
                error = ex.Message
            });
        }
    }

    /// <summary>
    /// Get suspicious mailbox rules (forwarding rules)
    /// </summary>
    [HttpGet("suspicious-mailbox-rules")]
    public async Task<IActionResult> GetSuspiciousMailboxRules([FromQuery] int top = 50)
    {
        try
        {
            // Get users
            var users = await _graphClient.Users
                .GetAsync(config =>
                {
                    config.QueryParameters.Top = top;
                    config.QueryParameters.Select = new[] { "id", "userPrincipalName", "displayName", "mail" };
                    config.QueryParameters.Filter = "accountEnabled eq true and mail ne null";
                });

            var userList = users?.Value ?? new List<User>();
            var suspiciousRules = new List<object>();

            foreach (var user in userList)
            {
                try
                {
                    var rules = await _graphClient.Users[user.Id].MailFolders["inbox"].MessageRules
                        .GetAsync();

                    var ruleList = rules?.Value ?? new List<MessageRule>();

                    foreach (var rule in ruleList)
                    {
                        // Check for suspicious patterns
                        var isSuspicious = false;
                        var suspiciousReasons = new List<string>();

                        // Forward to external address
                        if (rule.Actions?.ForwardTo?.Any() == true)
                        {
                            var forwardAddresses = rule.Actions.ForwardTo
                                .Select(r => r.EmailAddress?.Address)
                                .Where(a => a != null)
                                .ToList();

                            var userDomain = user.Mail?.Split('@').LastOrDefault() ?? "";
                            if (forwardAddresses.Any(a => !a!.Contains(userDomain, StringComparison.OrdinalIgnoreCase)))
                            {
                                isSuspicious = true;
                                suspiciousReasons.Add("Forwards to external address");
                            }
                        }

                        // Forward as attachment
                        if (rule.Actions?.ForwardAsAttachmentTo?.Any() == true)
                        {
                            isSuspicious = true;
                            suspiciousReasons.Add("Forwards as attachment");
                        }

                        // Redirect
                        if (rule.Actions?.RedirectTo?.Any() == true)
                        {
                            isSuspicious = true;
                            suspiciousReasons.Add("Redirects mail");
                        }

                        // Delete without reading
                        if (rule.Actions?.Delete == true && rule.Actions?.MarkAsRead != true)
                        {
                            isSuspicious = true;
                            suspiciousReasons.Add("Auto-deletes messages");
                        }

                        // Move to deleted items
                        if (rule.Actions?.MoveToFolder?.ToLower().Contains("deleted") == true)
                        {
                            isSuspicious = true;
                            suspiciousReasons.Add("Moves to Deleted Items");
                        }

                        if (isSuspicious || rule.Actions?.ForwardTo?.Any() == true || rule.Actions?.RedirectTo?.Any() == true)
                        {
                            suspiciousRules.Add(new
                            {
                                userId = user.Id,
                                userPrincipalName = user.UserPrincipalName,
                                userDisplayName = user.DisplayName,
                                ruleId = rule.Id,
                                ruleName = rule.DisplayName,
                                isEnabled = rule.IsEnabled,
                                isSuspicious,
                                suspiciousReasons,
                                actions = new
                                {
                                    forwardTo = rule.Actions?.ForwardTo?.Select(r => r.EmailAddress?.Address).ToList(),
                                    forwardAsAttachmentTo = rule.Actions?.ForwardAsAttachmentTo?.Select(r => r.EmailAddress?.Address).ToList(),
                                    redirectTo = rule.Actions?.RedirectTo?.Select(r => r.EmailAddress?.Address).ToList(),
                                    delete = rule.Actions?.Delete,
                                    moveToFolder = rule.Actions?.MoveToFolder,
                                    markAsRead = rule.Actions?.MarkAsRead
                                },
                                conditions = new
                                {
                                    fromAddresses = rule.Conditions?.FromAddresses?.Select(a => a.EmailAddress?.Address).ToList(),
                                    subjectContains = rule.Conditions?.SubjectContains,
                                    bodyContains = rule.Conditions?.BodyContains
                                }
                            });
                        }
                    }
                }
                catch
                {
                    // Skip users where we can't access mailbox rules
                }
            }

            return Ok(new
            {
                rules = suspiciousRules.OrderByDescending(r => ((dynamic)r).isSuspicious).ToList(),
                summary = new
                {
                    totalRulesFound = suspiciousRules.Count,
                    suspiciousRules = suspiciousRules.Count(r => ((dynamic)r).isSuspicious),
                    usersScanned = userList.Count
                },
                lastUpdated = DateTime.UtcNow
            });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching mailbox rules");
            return Ok(new
            {
                rules = new List<object>(),
                summary = new { totalRulesFound = 0 },
                error = ex.Message
            });
        }
    }

    /// <summary>
    /// Get secure score
    /// </summary>
    [HttpGet("secure-score")]
    public async Task<IActionResult> GetSecureScore()
    {
        try
        {
            var secureScores = await _graphClient.Security.SecureScores
                .GetAsync(config =>
                {
                    config.QueryParameters.Top = 1;
                    config.QueryParameters.Orderby = new[] { "createdDateTime desc" };
                });

            var latestScore = secureScores?.Value?.FirstOrDefault();

            if (latestScore == null)
            {
                return Ok(new
                {
                    score = (object?)null,
                    message = "No secure score data available"
                });
            }

            var controlScores = latestScore.ControlScores?.Select(c => new
            {
                controlName = c.ControlName,
                controlCategory = c.ControlCategory,
                score = c.Score,
                description = c.Description
            }).OrderByDescending(c => c.score).ToList();

            return Ok(new
            {
                score = new
                {
                    currentScore = latestScore.CurrentScore,
                    maxScore = latestScore.MaxScore,
                    percentage = latestScore.MaxScore > 0 ? Math.Round((latestScore.CurrentScore ?? 0) / (double)latestScore.MaxScore * 100, 1) : 0,
                    createdDateTime = latestScore.CreatedDateTime,
                    licensedUserCount = latestScore.LicensedUserCount,
                    activeUserCount = latestScore.ActiveUserCount,
                    enabledServices = latestScore.EnabledServices
                },
                controlScores,
                lastUpdated = DateTime.UtcNow
            });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching secure score");
            return Ok(new
            {
                score = (object?)null,
                error = ex.Message
            });
        }
    }
}
