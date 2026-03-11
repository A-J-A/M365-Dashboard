using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Graph;
using Microsoft.Graph.Models;

namespace M365Dashboard.Api.Controllers;

[ApiController]
[Route("api/[controller]")]
[Authorize]
public class PrivilegedAccessController : ControllerBase
{
    private readonly GraphServiceClient _graphClient;
    private readonly ILogger<PrivilegedAccessController> _logger;

    public PrivilegedAccessController(GraphServiceClient graphClient, ILogger<PrivilegedAccessController> logger)
    {
        _graphClient = graphClient;
        _logger = logger;
    }

    /// <summary>
    /// Get all directory roles and their members
    /// </summary>
    [HttpGet("directory-roles")]
    public async Task<IActionResult> GetDirectoryRoles()
    {
        try
        {
            // Get all activated directory roles
            var roles = await _graphClient.DirectoryRoles
                .GetAsync(config =>
                {
                    config.QueryParameters.Expand = new[] { "members" };
                });

            var roleList = roles?.Value ?? new List<DirectoryRole>();

            var result = new List<object>();
            var privilegedRoles = new[] {
                "Global Administrator",
                "Privileged Role Administrator",
                "Privileged Authentication Administrator",
                "Security Administrator",
                "User Administrator",
                "Exchange Administrator",
                "SharePoint Administrator",
                "Teams Administrator",
                "Intune Administrator",
                "Cloud Application Administrator",
                "Application Administrator",
                "Conditional Access Administrator",
                "Authentication Administrator",
                "Password Administrator",
                "Helpdesk Administrator",
                "Billing Administrator"
            };

            foreach (var role in roleList)
            {
                var membersList = new List<object>();
                if (role.Members != null)
                {
                    foreach (var member in role.Members.OfType<User>())
                    {
                        membersList.Add(new
                        {
                            id = member.Id,
                            userPrincipalName = member.UserPrincipalName,
                            displayName = member.DisplayName,
                            mail = member.Mail,
                            accountEnabled = member.AccountEnabled
                        });
                    }
                }

                var isPrivileged = privilegedRoles.Any(pr =>
                    role.DisplayName?.Contains(pr, StringComparison.OrdinalIgnoreCase) == true);

                result.Add(new
                {
                    id = role.Id,
                    displayName = role.DisplayName,
                    description = role.Description,
                    roleTemplateId = role.RoleTemplateId,
                    memberCount = membersList.Count,
                    members = membersList,
                    isPrivileged,
                    isGlobalAdmin = role.DisplayName?.Contains("Global Administrator", StringComparison.OrdinalIgnoreCase) == true
                });
            }

            // Summary
            var summary = new
            {
                totalRoles = result.Count,
                totalAssignments = result.Sum(r => ((dynamic)r).memberCount),
                privilegedRoleCount = result.Count(r => ((dynamic)r).isPrivileged),
                globalAdminCount = result
                    .Where(r => ((dynamic)r).isGlobalAdmin)
                    .Sum(r => (int)((dynamic)r).memberCount)
            };

            return Ok(new
            {
                roles = result.OrderByDescending(r => ((dynamic)r).isPrivileged)
                    .ThenByDescending(r => ((dynamic)r).memberCount)
                    .ToList(),
                summary,
                lastUpdated = DateTime.UtcNow
            });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching directory roles");
            return Ok(new
            {
                roles = new List<object>(),
                summary = new { totalRoles = 0 },
                error = ex.Message
            });
        }
    }

    /// <summary>
    /// Get role assignment changes from audit logs
    /// </summary>
    [HttpGet("role-changes")]
    public async Task<IActionResult> GetRoleChanges([FromQuery] int days = 30)
    {
        try
        {
            var fromDate = DateTime.UtcNow.AddDays(-days);

            var auditLogs = await _graphClient.AuditLogs.DirectoryAudits
                .GetAsync(config =>
                {
                    config.QueryParameters.Filter = $"activityDateTime ge {fromDate:yyyy-MM-ddTHH:mm:ssZ} and (activityDisplayName eq 'Add member to role' or activityDisplayName eq 'Remove member from role' or activityDisplayName eq 'Add eligible member to role' or activityDisplayName eq 'Remove eligible member from role')";
                    config.QueryParameters.Top = 200;
                    config.QueryParameters.Orderby = new[] { "activityDateTime desc" };
                });

            var auditList = auditLogs?.Value ?? new List<DirectoryAudit>();

            var result = auditList.Select(a => new
            {
                id = a.Id,
                activityDateTime = a.ActivityDateTime,
                activityDisplayName = a.ActivityDisplayName,
                initiatedBy = a.InitiatedBy?.User != null ? new
                {
                    userPrincipalName = a.InitiatedBy.User.UserPrincipalName,
                    displayName = a.InitiatedBy.User.DisplayName,
                    id = a.InitiatedBy.User.Id
                } : null,
                targetResources = a.TargetResources?.Select(t => new
                {
                    displayName = t.DisplayName,
                    userPrincipalName = t.UserPrincipalName,
                    type = t.Type,
                    modifiedProperties = t.ModifiedProperties?.Select(mp => new
                    {
                        displayName = mp.DisplayName,
                        oldValue = mp.OldValue,
                        newValue = mp.NewValue
                    }).ToList()
                }).ToList(),
                result = a.Result,
                resultReason = a.ResultReason,
                category = a.Category
            }).ToList();

            // Group by activity type
            var activitySummary = result
                .GroupBy(r => r.activityDisplayName)
                .Select(g => new { activity = g.Key, count = g.Count() })
                .OrderByDescending(x => x.count)
                .ToList();

            return Ok(new
            {
                changes = result,
                activitySummary,
                totalChanges = result.Count,
                period = new { fromDate, toDate = DateTime.UtcNow },
                lastUpdated = DateTime.UtcNow
            });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching role changes");
            return Ok(new
            {
                changes = new List<object>(),
                error = ex.Message
            });
        }
    }

    /// <summary>
    /// Get privileged user sign-ins
    /// </summary>
    [HttpGet("privileged-sign-ins")]
    public async Task<IActionResult> GetPrivilegedSignIns([FromQuery] int days = 7)
    {
        try
        {
            // First get Global Admins
            var globalAdminRole = await _graphClient.DirectoryRoles
                .GetAsync(config =>
                {
                    config.QueryParameters.Filter = "displayName eq 'Global Administrator'";
                    config.QueryParameters.Expand = new[] { "members" };
                });

            var globalAdmins = globalAdminRole?.Value?.FirstOrDefault()?.Members?
                .OfType<User>()
                .Select(u => u.UserPrincipalName?.ToLower())
                .Where(upn => upn != null)
                .ToHashSet() ?? new HashSet<string?>();

            if (globalAdmins.Count == 0)
            {
                return Ok(new
                {
                    signIns = new List<object>(),
                    summary = new { totalSignIns = 0, globalAdminCount = 0 },
                    message = "No Global Administrators found"
                });
            }

            var fromDate = DateTime.UtcNow.AddDays(-days);

            var signIns = await _graphClient.AuditLogs.SignIns
                .GetAsync(config =>
                {
                    config.QueryParameters.Filter = $"createdDateTime ge {fromDate:yyyy-MM-ddTHH:mm:ssZ}";
                    config.QueryParameters.Top = 999;
                    config.QueryParameters.Orderby = new[] { "createdDateTime desc" };
                });

            var signInList = signIns?.Value ?? new List<SignIn>();

            // Filter to privileged users
            var privilegedSignIns = signInList
                .Where(s => globalAdmins.Contains(s.UserPrincipalName?.ToLower()))
                .Select(s => new
                {
                    id = s.Id,
                    userPrincipalName = s.UserPrincipalName,
                    userDisplayName = s.UserDisplayName,
                    appDisplayName = s.AppDisplayName,
                    ipAddress = s.IpAddress,
                    location = s.Location != null ? new
                    {
                        city = s.Location.City,
                        state = s.Location.State,
                        country = s.Location.CountryOrRegion,
                        geoCoordinates = s.Location.GeoCoordinates
                    } : null,
                    createdDateTime = s.CreatedDateTime,
                    status = new
                    {
                        errorCode = s.Status?.ErrorCode,
                        failureReason = s.Status?.FailureReason
                    },
                    isInteractive = s.IsInteractive,
                    riskLevel = s.RiskLevelDuringSignIn?.ToString(),
                    riskState = s.RiskState?.ToString(),
                    deviceDetail = s.DeviceDetail != null ? new
                    {
                        browser = s.DeviceDetail.Browser,
                        operatingSystem = s.DeviceDetail.OperatingSystem,
                        displayName = s.DeviceDetail.DisplayName,
                        isCompliant = s.DeviceDetail.IsCompliant,
                        isManaged = s.DeviceDetail.IsManaged
                    } : null,
                    clientAppUsed = s.ClientAppUsed,
                    conditionalAccessStatus = s.ConditionalAccessStatus?.ToString()
                })
                .ToList();

            // Analyze patterns
            var locationSummary = privilegedSignIns
                .Where(s => s.location?.country != null)
                .GroupBy(s => s.location!.country)
                .Select(g => new { country = g.Key, count = g.Count() })
                .OrderByDescending(x => x.count)
                .Take(10)
                .ToList();

            var appSummary = privilegedSignIns
                .GroupBy(s => s.appDisplayName ?? "Unknown")
                .Select(g => new { app = g.Key, count = g.Count() })
                .OrderByDescending(x => x.count)
                .Take(10)
                .ToList();

            var riskySigIns = privilegedSignIns
                .Where(s => s.riskLevel != null && s.riskLevel != "None" && s.riskLevel != "none")
                .ToList();

            return Ok(new
            {
                signIns = privilegedSignIns.Take(100),
                summary = new
                {
                    totalSignIns = privilegedSignIns.Count,
                    globalAdminCount = globalAdmins.Count,
                    uniqueAdminsSignedIn = privilegedSignIns.Select(s => s.userPrincipalName).Distinct().Count(),
                    riskySignIns = riskySigIns.Count,
                    failedSignIns = privilegedSignIns.Count(s => s.status.errorCode != 0)
                },
                locationSummary,
                appSummary,
                riskySignIns = riskySigIns.Take(20),
                period = new { fromDate, toDate = DateTime.UtcNow },
                lastUpdated = DateTime.UtcNow
            });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching privileged sign-ins");
            return Ok(new
            {
                signIns = new List<object>(),
                summary = new { totalSignIns = 0 },
                error = ex.Message
            });
        }
    }

    /// <summary>
    /// Get service principals with high-privilege permissions
    /// </summary>
    [HttpGet("privileged-apps")]
    public async Task<IActionResult> GetPrivilegedApps()
    {
        try
        {
            var servicePrincipals = await _graphClient.ServicePrincipals
                .GetAsync(config =>
                {
                    config.QueryParameters.Top = 200;
                    config.QueryParameters.Select = new[] { "id", "displayName", "appId", "servicePrincipalType", "appRoles", "oauth2PermissionScopes" };
                });

            var spList = servicePrincipals?.Value ?? new List<ServicePrincipal>();

            // Get app role assignments
            var privilegedApps = new List<object>();

            foreach (var sp in spList.Take(50)) // Limit to avoid too many API calls
            {
                try
                {
                    var appRoleAssignments = await _graphClient.ServicePrincipals[sp.Id].AppRoleAssignments
                        .GetAsync();

                    var assignments = appRoleAssignments?.Value ?? new List<AppRoleAssignment>();

                    if (assignments.Any())
                    {
                        privilegedApps.Add(new
                        {
                            id = sp.Id,
                            displayName = sp.DisplayName,
                            appId = sp.AppId,
                            type = sp.ServicePrincipalType,
                            assignmentCount = assignments.Count,
                            assignments = assignments.Select(a => new
                            {
                                resourceDisplayName = a.ResourceDisplayName,
                                appRoleId = a.AppRoleId,
                                createdDateTime = a.CreatedDateTime
                            }).ToList()
                        });
                    }
                }
                catch
                {
                    // Skip if we can't read assignments
                }
            }

            return Ok(new
            {
                apps = privilegedApps.OrderByDescending(a => ((dynamic)a).assignmentCount).ToList(),
                totalCount = privilegedApps.Count,
                lastUpdated = DateTime.UtcNow
            });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching privileged apps");
            return Ok(new
            {
                apps = new List<object>(),
                error = ex.Message
            });
        }
    }
}
