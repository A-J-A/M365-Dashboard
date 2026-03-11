using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Graph;
using Microsoft.Graph.Models;

namespace M365Dashboard.Api.Controllers;

[ApiController]
[Route("api/[controller]")]
[Authorize]
public class ConditionalAccessController : ControllerBase
{
    private readonly GraphServiceClient _graphClient;
    private readonly ILogger<ConditionalAccessController> _logger;

    public ConditionalAccessController(GraphServiceClient graphClient, ILogger<ConditionalAccessController> logger)
    {
        _graphClient = graphClient;
        _logger = logger;
    }

    /// <summary>
    /// Get all Conditional Access policies
    /// </summary>
    [HttpGet("policies")]
    public async Task<IActionResult> GetPolicies()
    {
        try
        {
            var policies = await _graphClient.Identity.ConditionalAccess.Policies
                .GetAsync(config =>
                {
                    config.QueryParameters.Top = 100;
                });

            var policyList = policies?.Value ?? new List<ConditionalAccessPolicy>();

            var result = policyList.Select(p => new
            {
                id = p.Id,
                displayName = p.DisplayName,
                state = p.State?.ToString(),
                createdDateTime = p.CreatedDateTime,
                modifiedDateTime = p.ModifiedDateTime,
                conditions = new
                {
                    users = new
                    {
                        includeUsers = p.Conditions?.Users?.IncludeUsers,
                        excludeUsers = p.Conditions?.Users?.ExcludeUsers,
                        includeGroups = p.Conditions?.Users?.IncludeGroups,
                        excludeGroups = p.Conditions?.Users?.ExcludeGroups,
                        includeRoles = p.Conditions?.Users?.IncludeRoles,
                        excludeRoles = p.Conditions?.Users?.ExcludeRoles
                    },
                    applications = new
                    {
                        includeApplications = p.Conditions?.Applications?.IncludeApplications,
                        excludeApplications = p.Conditions?.Applications?.ExcludeApplications,
                        includeUserActions = p.Conditions?.Applications?.IncludeUserActions
                    },
                    platforms = p.Conditions?.Platforms,
                    locations = new
                    {
                        includeLocations = p.Conditions?.Locations?.IncludeLocations,
                        excludeLocations = p.Conditions?.Locations?.ExcludeLocations
                    },
                    clientAppTypes = p.Conditions?.ClientAppTypes,
                    signInRiskLevels = p.Conditions?.SignInRiskLevels,
                    userRiskLevels = p.Conditions?.UserRiskLevels
                },
                grantControls = new
                {
                    builtInControls = p.GrantControls?.BuiltInControls?.Select(c => c.ToString()),
                    customAuthenticationFactors = p.GrantControls?.CustomAuthenticationFactors,
                    operatorControl = p.GrantControls?.Operator,
                    termsOfUse = p.GrantControls?.TermsOfUse
                },
                sessionControls = new
                {
                    applicationEnforcedRestrictions = p.SessionControls?.ApplicationEnforcedRestrictions?.IsEnabled,
                    cloudAppSecurity = p.SessionControls?.CloudAppSecurity?.IsEnabled,
                    persistentBrowser = p.SessionControls?.PersistentBrowser?.Mode?.ToString(),
                    signInFrequency = p.SessionControls?.SignInFrequency != null ? new
                    {
                        value = p.SessionControls.SignInFrequency.Value,
                        type = p.SessionControls.SignInFrequency.Type?.ToString(),
                        isEnabled = p.SessionControls.SignInFrequency.IsEnabled
                    } : null
                }
            }).ToList();

            // Summary statistics
            var summary = new
            {
                totalPolicies = policyList.Count,
                enabledPolicies = policyList.Count(p => p.State == ConditionalAccessPolicyState.Enabled),
                disabledPolicies = policyList.Count(p => p.State == ConditionalAccessPolicyState.Disabled),
                reportOnlyPolicies = policyList.Count(p => p.State == ConditionalAccessPolicyState.EnabledForReportingButNotEnforced),
                policiesRequiringMfa = policyList.Count(p => p.GrantControls?.BuiltInControls?.Contains(ConditionalAccessGrantControl.Mfa) == true),
                policiesRequiringCompliantDevice = policyList.Count(p => p.GrantControls?.BuiltInControls?.Contains(ConditionalAccessGrantControl.CompliantDevice) == true),
                policiesBlockingAccess = policyList.Count(p => p.GrantControls?.BuiltInControls?.Contains(ConditionalAccessGrantControl.Block) == true)
            };

            return Ok(new
            {
                policies = result,
                summary,
                lastUpdated = DateTime.UtcNow
            });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching Conditional Access policies");
            return Ok(new
            {
                policies = new List<object>(),
                summary = new { totalPolicies = 0 },
                error = ex.Message
            });
        }
    }

    /// <summary>
    /// Get named locations
    /// </summary>
    [HttpGet("named-locations")]
    public async Task<IActionResult> GetNamedLocations()
    {
        try
        {
            var locations = await _graphClient.Identity.ConditionalAccess.NamedLocations
                .GetAsync(config =>
                {
                    config.QueryParameters.Top = 100;
                });

            var locationList = locations?.Value ?? new List<NamedLocation>();

            var result = locationList.Select(l =>
            {
                var baseInfo = new Dictionary<string, object?>
                {
                    ["id"] = l.Id,
                    ["displayName"] = l.DisplayName,
                    ["createdDateTime"] = l.CreatedDateTime,
                    ["modifiedDateTime"] = l.ModifiedDateTime
                };

                if (l is IpNamedLocation ipLocation)
                {
                    baseInfo["type"] = "IP";
                    baseInfo["isTrusted"] = ipLocation.IsTrusted;
                    baseInfo["ipRanges"] = ipLocation.IpRanges?.Select(r =>
                    {
                        if (r is IPv4CidrRange ipv4)
                            return $"{ipv4.CidrAddress}";
                        if (r is IPv6CidrRange ipv6)
                            return $"{ipv6.CidrAddress}";
                        return "Unknown";
                    }).ToList();
                }
                else if (l is CountryNamedLocation countryLocation)
                {
                    baseInfo["type"] = "Country";
                    baseInfo["countriesAndRegions"] = countryLocation.CountriesAndRegions;
                    baseInfo["includeUnknownCountriesAndRegions"] = countryLocation.IncludeUnknownCountriesAndRegions;
                }

                return baseInfo;
            }).ToList();

            return Ok(new
            {
                locations = result,
                totalCount = result.Count,
                lastUpdated = DateTime.UtcNow
            });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching named locations");
            return Ok(new
            {
                locations = new List<object>(),
                error = ex.Message
            });
        }
    }

    /// <summary>
    /// Get sign-in logs filtered by Conditional Access results
    /// </summary>
    [HttpGet("sign-in-insights")]
    public async Task<IActionResult> GetSignInInsights([FromQuery] int days = 7)
    {
        try
        {
            var fromDate = DateTime.UtcNow.AddDays(-days);

            var signIns = await _graphClient.AuditLogs.SignIns
                .GetAsync(config =>
                {
                    config.QueryParameters.Filter = $"createdDateTime ge {fromDate:yyyy-MM-ddTHH:mm:ssZ}";
                    config.QueryParameters.Top = 999;
                    config.QueryParameters.Orderby = new[] { "createdDateTime desc" };
                });

            var signInList = signIns?.Value ?? new List<SignIn>();

            // Analyze CA results
            var caResults = new Dictionary<string, int>
            {
                ["success"] = 0,
                ["failure"] = 0,
                ["notApplied"] = 0,
                ["reportOnlySuccess"] = 0,
                ["reportOnlyFailure"] = 0
            };

            var policyHits = new Dictionary<string, PolicyHitInfo>();
            var blockReasons = new Dictionary<string, int>();

            foreach (var signIn in signInList)
            {
                if (signIn.AppliedConditionalAccessPolicies != null)
                {
                    foreach (var policy in signIn.AppliedConditionalAccessPolicies)
                    {
                        var policyName = policy.DisplayName ?? "Unknown Policy";
                        var policyId = policy.Id ?? "unknown";

                        if (!policyHits.ContainsKey(policyId))
                        {
                            policyHits[policyId] = new PolicyHitInfo
                            {
                                PolicyId = policyId,
                                PolicyName = policyName,
                                SuccessCount = 0,
                                FailureCount = 0,
                                NotAppliedCount = 0,
                                ReportOnlyCount = 0
                            };
                        }

                        switch (policy.Result)
                        {
                            case AppliedConditionalAccessPolicyResult.Success:
                                caResults["success"]++;
                                policyHits[policyId].SuccessCount++;
                                break;
                            case AppliedConditionalAccessPolicyResult.Failure:
                                caResults["failure"]++;
                                policyHits[policyId].FailureCount++;
                                // Track block reasons
                                if (policy.EnforcedGrantControls != null)
                                {
                                    foreach (var control in policy.EnforcedGrantControls)
                                    {
                                        if (!blockReasons.ContainsKey(control))
                                            blockReasons[control] = 0;
                                        blockReasons[control]++;
                                    }
                                }
                                break;
                            case AppliedConditionalAccessPolicyResult.NotApplied:
                                caResults["notApplied"]++;
                                policyHits[policyId].NotAppliedCount++;
                                break;
                            case AppliedConditionalAccessPolicyResult.ReportOnlySuccess:
                                caResults["reportOnlySuccess"]++;
                                policyHits[policyId].ReportOnlyCount++;
                                break;
                            case AppliedConditionalAccessPolicyResult.ReportOnlyFailure:
                                caResults["reportOnlyFailure"]++;
                                policyHits[policyId].ReportOnlyCount++;
                                break;
                        }
                    }
                }
            }

            // Get blocked sign-ins with details
            var blockedSignIns = signInList
                .Where(s => s.AppliedConditionalAccessPolicies?.Any(p => p.Result == AppliedConditionalAccessPolicyResult.Failure) == true)
                .Take(50)
                .Select(s => new
                {
                    id = s.Id,
                    userPrincipalName = s.UserPrincipalName,
                    userDisplayName = s.UserDisplayName,
                    appDisplayName = s.AppDisplayName,
                    ipAddress = s.IpAddress,
                    location = s.Location != null ? $"{s.Location.City}, {s.Location.CountryOrRegion}" : null,
                    createdDateTime = s.CreatedDateTime,
                    status = s.Status?.ErrorCode,
                    failureReason = s.Status?.FailureReason,
                    blockedByPolicies = s.AppliedConditionalAccessPolicies?
                        .Where(p => p.Result == AppliedConditionalAccessPolicyResult.Failure)
                        .Select(p => new { p.Id, p.DisplayName, p.EnforcedGrantControls })
                        .ToList()
                })
                .ToList();

            return Ok(new
            {
                summary = new
                {
                    totalSignIns = signInList.Count,
                    caResults,
                    blockReasons = blockReasons.OrderByDescending(kvp => kvp.Value)
                        .Select(kvp => new { reason = kvp.Key, count = kvp.Value })
                        .ToList()
                },
                policyHits = policyHits.Values
                    .OrderByDescending(p => p.SuccessCount + p.FailureCount)
                    .Take(20)
                    .ToList(),
                blockedSignIns,
                period = new { fromDate, toDate = DateTime.UtcNow },
                lastUpdated = DateTime.UtcNow
            });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching CA sign-in insights");
            return Ok(new
            {
                summary = new { totalSignIns = 0 },
                policyHits = new List<object>(),
                blockedSignIns = new List<object>(),
                error = ex.Message
            });
        }
    }

    private class PolicyHitInfo
    {
        public string PolicyId { get; set; } = "";
        public string PolicyName { get; set; } = "";
        public int SuccessCount { get; set; }
        public int FailureCount { get; set; }
        public int NotAppliedCount { get; set; }
        public int ReportOnlyCount { get; set; }
    }
}
