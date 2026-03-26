using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Graph;
using Microsoft.Graph.Models;

namespace M365Dashboard.Api.Controllers;

[ApiController]
[Route("api/[controller]")]
[Authorize]
public class ApplicationConsentController : ControllerBase
{
    private readonly GraphServiceClient _graphClient;
    private readonly ILogger<ApplicationConsentController> _logger;

    public ApplicationConsentController(GraphServiceClient graphClient, ILogger<ApplicationConsentController> logger)
    {
        _graphClient = graphClient;
        _logger = logger;
    }

    /// <summary>
    /// Get all OAuth2 permission grants (delegated permissions consented)
    /// </summary>
    [HttpGet("oauth2-grants")]
    public async Task<IActionResult> GetOAuth2PermissionGrants()
    {
        try
        {
            var grants = await _graphClient.Oauth2PermissionGrants
                .GetAsync(config =>
                {
                    config.QueryParameters.Top = 500;
                });

            var grantList = grants?.Value ?? new List<OAuth2PermissionGrant>();

            // Get service principal details for context
            var spIds = grantList.Select(g => g.ClientId).Distinct().ToList();
            var servicePrincipals = new Dictionary<string, ServicePrincipal>();

            foreach (var spId in spIds.Take(100))
            {
                try
                {
                    var sp = await _graphClient.ServicePrincipals[spId].GetAsync();
                    if (sp != null)
                    {
                        servicePrincipals[spId!] = sp;
                    }
                }
                catch { }
            }

            var result = grantList.Select(g =>
            {
                servicePrincipals.TryGetValue(g.ClientId ?? "", out var sp);
                return new
                {
                    id = g.Id,
                    clientId = g.ClientId,
                    clientDisplayName = sp?.DisplayName ?? "Unknown App",
                    clientAppId = sp?.AppId,
                    consentType = g.ConsentType,
                    principalId = g.PrincipalId,
                    resourceId = g.ResourceId,
                    scope = g.Scope,
                    scopes = g.Scope?.Split(' ', StringSplitOptions.RemoveEmptyEntries).ToList()
                };
            }).ToList();

            // Categorize permissions
            var highRiskPermissions = new[] {
                "Directory.ReadWrite.All", "Directory.AccessAsUser.All",
                "User.ReadWrite.All", "Group.ReadWrite.All",
                "Mail.ReadWrite", "Mail.Send", "Mail.Read",
                "Files.ReadWrite.All", "Sites.ReadWrite.All",
                "Calendars.ReadWrite", "Contacts.ReadWrite"
            };

            var grantsWithHighRisk = result
                .Where(g => g.scopes?.Any(s => highRiskPermissions.Contains(s)) == true)
                .ToList();

            // Summary by consent type
            var consentTypeSummary = result
                .GroupBy(g => g.consentType ?? "Unknown")
                .Select(g => new { consentType = g.Key, count = g.Count() })
                .ToList();

            // Top apps by permission count
            var topAppsByPermissions = result
                .GroupBy(g => g.clientDisplayName)
                .Select(g => new
                {
                    appName = g.Key,
                    grantCount = g.Count(),
                    totalScopes = g.Sum(x => x.scopes?.Count ?? 0),
                    hasHighRiskPermissions = g.Any(x => x.scopes?.Any(s => highRiskPermissions.Contains(s)) == true)
                })
                .OrderByDescending(x => x.totalScopes)
                .Take(20)
                .ToList();

            return Ok(new
            {
                grants = result,
                grantsWithHighRisk,
                summary = new
                {
                    totalGrants = result.Count,
                    adminConsentGrants = result.Count(g => g.consentType == "AllPrincipals"),
                    userConsentGrants = result.Count(g => g.consentType == "Principal"),
                    highRiskGrants = grantsWithHighRisk.Count,
                    uniqueApps = result.Select(g => g.clientId).Distinct().Count()
                },
                consentTypeSummary,
                topAppsByPermissions,
                lastUpdated = DateTime.UtcNow
            });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching OAuth2 permission grants");
            return Ok(new
            {
                grants = new List<object>(),
                summary = new { totalGrants = 0 },
                error = ex.Message
            });
        }
    }

    /// <summary>
    /// Get enterprise applications with their permissions
    /// </summary>
    [HttpGet("enterprise-apps")]
    public async Task<IActionResult> GetEnterpriseApps()
    {
        try
        {
            // Page through all service principals - Entra portal shows all Application type SPs
            var spList = new List<ServicePrincipal>();
            var spPage = await _graphClient.ServicePrincipals.GetAsync(config =>
            {
                config.QueryParameters.Top = 999;
                config.QueryParameters.Filter = "servicePrincipalType eq 'Application'";
                config.QueryParameters.Orderby = new[] { "displayName" };
            });
            while (spPage?.Value != null)
            {
                spList.AddRange(spPage.Value);
                if (spPage.OdataNextLink == null) break;
                spPage = await _graphClient.ServicePrincipals.WithUrl(spPage.OdataNextLink).GetAsync();
            }

            var result = new List<object>();

            foreach (var sp in spList)
            {
                try
                {
                    // Get app role assignments (application permissions)
                    var appRoleAssignments = await _graphClient.ServicePrincipals[sp.Id].AppRoleAssignments
                        .GetAsync();

                    var assignments = appRoleAssignments?.Value ?? new List<AppRoleAssignment>();

                    // Get OAuth2 permission grants (delegated permissions)
                    var oauth2Grants = await _graphClient.ServicePrincipals[sp.Id].Oauth2PermissionGrants
                        .GetAsync();

                    var grantsList = oauth2Grants?.Value ?? new List<OAuth2PermissionGrant>();

                    result.Add(new
                    {
                        id = sp.Id,
                        appId = sp.AppId,
                        displayName = sp.DisplayName,
                        description = sp.Description,
                        homepage = sp.Homepage,
                        loginUrl = sp.LoginUrl,
                        logoutUrl = sp.LogoutUrl,
                        publisherName = sp.AppOwnerOrganizationId?.ToString(),
                        verifiedPublisher = sp.VerifiedPublisher != null ? new
                        {
                            displayName = sp.VerifiedPublisher.DisplayName,
                            verifiedPublisherId = sp.VerifiedPublisher.VerifiedPublisherId,
                            addedDateTime = sp.VerifiedPublisher.AddedDateTime
                        } : null,
                        isVerified = sp.VerifiedPublisher != null,
                        accountEnabled = sp.AccountEnabled,
                        servicePrincipalType = sp.ServicePrincipalType,
                        signInAudience = sp.SignInAudience,
                        tags = sp.Tags,
                        appRoleAssignmentCount = assignments.Count,
                        appRoleAssignments = assignments.Select(a => new
                        {
                            resourceDisplayName = a.ResourceDisplayName,
                            appRoleId = a.AppRoleId,
                            createdDateTime = a.CreatedDateTime
                        }).ToList(),
                        oauth2GrantCount = grantsList.Count,
                        delegatedScopes = grantsList.SelectMany(g => g.Scope?.Split(' ') ?? Array.Empty<string>()).Distinct().ToList()
                    });
                }
                catch
                {
                    // Add basic info if we can't get detailed permissions
                    result.Add(new
                    {
                        id = sp.Id,
                        appId = sp.AppId,
                        displayName = sp.DisplayName,
                        description = sp.Description,
                        publisherName = sp.AppOwnerOrganizationId?.ToString(),
                        isVerified = sp.VerifiedPublisher != null,
                        accountEnabled = sp.AccountEnabled,
                        servicePrincipalType = sp.ServicePrincipalType,
                        appRoleAssignmentCount = 0,
                        oauth2GrantCount = 0
                    });
                }
            }

            // Summary
            var summary = new
            {
                totalApps = result.Count,
                verifiedApps = result.Count(r => ((dynamic)r).isVerified == true),
                unverifiedApps = result.Count(r => ((dynamic)r).isVerified != true),
                appsWithAppPermissions = result.Count(r => ((dynamic)r).appRoleAssignmentCount > 0),
                appsWithDelegatedPermissions = result.Count(r => ((dynamic)r).oauth2GrantCount > 0),
                disabledApps = result.Count(r => ((dynamic)r).accountEnabled == false)
            };

            // Apps by publisher
            var publisherSummary = result
                .GroupBy(r => ((dynamic)r).publisherName ?? "Unknown")
                .Select(g => new { publisher = (string)g.Key, count = g.Count() })
                .OrderByDescending(x => x.count)
                .Take(10)
                .ToList();

            return Ok(new
            {
                apps = result.OrderByDescending(r => ((dynamic)r).appRoleAssignmentCount + ((dynamic)r).oauth2GrantCount).ToList(),
                summary,
                publisherSummary,
                lastUpdated = DateTime.UtcNow
            });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching enterprise apps");
            return Ok(new
            {
                apps = new List<object>(),
                summary = new { totalApps = 0 },
                error = ex.Message
            });
        }
    }

    /// <summary>
    /// Get app registrations (apps registered in tenant)
    /// </summary>
    [HttpGet("app-registrations")]
    public async Task<IActionResult> GetAppRegistrations()
    {
        try
        {
            var apps = await _graphClient.Applications
                .GetAsync(config =>
                {
                    config.QueryParameters.Top = 200;
                    config.QueryParameters.Orderby = new[] { "displayName" };
                });

            var appList = apps?.Value ?? new List<Application>();

            var result = appList.Select(a => new
            {
                id = a.Id,
                appId = a.AppId,
                displayName = a.DisplayName,
                description = a.Description,
                createdDateTime = a.CreatedDateTime,
                signInAudience = a.SignInAudience,
                publisherDomain = a.PublisherDomain,
                identifierUris = a.IdentifierUris,
                web = a.Web != null ? new
                {
                    redirectUris = a.Web.RedirectUris,
                    homePageUrl = a.Web.HomePageUrl,
                    logoutUrl = a.Web.LogoutUrl
                } : null,
                spa = a.Spa != null ? new
                {
                    redirectUris = a.Spa.RedirectUris
                } : null,
                publicClient = a.PublicClient != null ? new
                {
                    redirectUris = a.PublicClient.RedirectUris
                } : null,
                requiredResourceAccess = a.RequiredResourceAccess?.Select(r => new
                {
                    resourceAppId = r.ResourceAppId,
                    resourceAccess = r.ResourceAccess?.Select(ra => new
                    {
                        id = ra.Id,
                        type = ra.Type
                    }).ToList()
                }).ToList(),
                passwordCredentials = a.PasswordCredentials?.Select(p => new
                {
                    keyId = p.KeyId,
                    displayName = p.DisplayName,
                    startDateTime = p.StartDateTime,
                    endDateTime = p.EndDateTime,
                    isExpired = p.EndDateTime < DateTime.UtcNow,
                    isExpiringSoon = p.EndDateTime < DateTime.UtcNow.AddDays(30)
                }).ToList(),
                keyCredentials = a.KeyCredentials?.Select(k => new
                {
                    keyId = k.KeyId,
                    displayName = k.DisplayName,
                    type = k.Type,
                    usage = k.Usage,
                    startDateTime = k.StartDateTime,
                    endDateTime = k.EndDateTime,
                    isExpired = k.EndDateTime < DateTime.UtcNow,
                    isExpiringSoon = k.EndDateTime < DateTime.UtcNow.AddDays(30)
                }).ToList(),
                hasExpiredCredentials = a.PasswordCredentials?.Any(p => p.EndDateTime < DateTime.UtcNow) == true ||
                                        a.KeyCredentials?.Any(k => k.EndDateTime < DateTime.UtcNow) == true,
                hasExpiringSoonCredentials = a.PasswordCredentials?.Any(p => p.EndDateTime < DateTime.UtcNow.AddDays(30)) == true ||
                                              a.KeyCredentials?.Any(k => k.EndDateTime < DateTime.UtcNow.AddDays(30)) == true
            }).ToList();

            // Summary
            var summary = new
            {
                totalApps = result.Count,
                appsWithExpiredCredentials = result.Count(r => r.hasExpiredCredentials),
                appsWithExpiringSoonCredentials = result.Count(r => r.hasExpiringSoonCredentials),
                singleTenantApps = result.Count(r => r.signInAudience == "AzureADMyOrg"),
                multiTenantApps = result.Count(r => r.signInAudience == "AzureADMultipleOrgs" || r.signInAudience == "AzureADandPersonalMicrosoftAccount")
            };

            // Apps with expiring credentials
            var expiringCredentials = result
                .Where(r => r.hasExpiringSoonCredentials)
                .Select(r => new
                {
                    r.appId,
                    r.displayName,
                    expiringPasswords = r.passwordCredentials?.Where(p => p.isExpiringSoon).ToList(),
                    expiringCertificates = r.keyCredentials?.Where(k => k.isExpiringSoon).ToList()
                })
                .ToList();

            return Ok(new
            {
                apps = result,
                summary,
                expiringCredentials,
                lastUpdated = DateTime.UtcNow
            });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching app registrations");
            return Ok(new
            {
                apps = new List<object>(),
                summary = new { totalApps = 0 },
                error = ex.Message
            });
        }
    }

    /// <summary>
    /// Enterprise app audit - covers both app registrations and enterprise apps (service principals)
    /// </summary>
    [HttpGet("enterprise-app-audit")]
    public async Task<IActionResult> GetEnterpriseAppAudit()
    {
        try
        {
            var now = DateTime.UtcNow;
            var thirtyDaysAgo = now.AddDays(-30);

            // ── App Registrations (paged) ──────────────────────────────────
            var appList = new List<Microsoft.Graph.Models.Application>();
            var appPage = await _graphClient.Applications.GetAsync(config =>
            {
                config.QueryParameters.Top = 999;
                config.QueryParameters.Select = new[]
                {
                    "id", "appId", "displayName", "createdDateTime", "description",
                    "signInAudience", "publisherDomain", "passwordCredentials",
                    "keyCredentials", "requiredResourceAccess"
                };
            });
            while (appPage?.Value != null)
            {
                appList.AddRange(appPage.Value);
                if (appPage.OdataNextLink == null) break;
                appPage = await _graphClient.Applications.WithUrl(appPage.OdataNextLink).GetAsync();
            }

            var registrations = appList.Select(a =>
            {
                var hasSecret  = a.PasswordCredentials?.Any() == true;
                var hasCert    = a.KeyCredentials?.Any() == true;
                var hasAnyCred = hasSecret || hasCert;
                var allExpired = hasAnyCred &&
                    (a.PasswordCredentials?.All(p => p.EndDateTime < now) ?? true) &&
                    (a.KeyCredentials?.All(k => k.EndDateTime < now) ?? true);

                var allExpiries = new List<DateTimeOffset?>();
                if (a.PasswordCredentials != null) allExpiries.AddRange(a.PasswordCredentials.Select(p => p.EndDateTime));
                if (a.KeyCredentials     != null) allExpiries.AddRange(a.KeyCredentials.Select(k => k.EndDateTime));
                var nextExpiry = allExpiries.Where(d => d > now).OrderBy(d => d).FirstOrDefault();
                var createdDaysAgo = a.CreatedDateTime.HasValue ? (int)(now - a.CreatedDateTime.Value).TotalDays : (int?)null;

                return new
                {
                    id = a.Id,
                    appId = a.AppId,
                    displayName = a.DisplayName ?? "(no name)",
                    description = a.Description,
                    createdDateTime = a.CreatedDateTime,
                    createdDaysAgo,
                    isNew = a.CreatedDateTime.HasValue && a.CreatedDateTime.Value >= thirtyDaysAgo,
                    signInAudience = a.SignInAudience,
                    publisherDomain = a.PublisherDomain,
                    hasCredentials = hasAnyCred,
                    noCredentials = !hasAnyCred,
                    allCredentialsExpired = allExpired,
                    secretCount = a.PasswordCredentials?.Count ?? 0,
                    certCount   = a.KeyCredentials?.Count ?? 0,
                    nextExpiry,
                    daysUntilNextExpiry = nextExpiry.HasValue ? (int)(nextExpiry.Value - now).TotalDays : (int?)null,
                    requiresResourceAccess = a.RequiredResourceAccess?.Any() == true,
                    resourceAccessCount    = a.RequiredResourceAccess?.Sum(r => r.ResourceAccess?.Count ?? 0) ?? 0,
                    appType = "Registration"
                };
            }).ToList();

            // ── Enterprise Apps (Service Principals, paged) ─────────────────
            var spList = new List<Microsoft.Graph.Models.ServicePrincipal>();
            var spPage = await _graphClient.ServicePrincipals.GetAsync(config =>
            {
                config.QueryParameters.Top = 999;
                config.QueryParameters.Filter = "servicePrincipalType eq 'Application'";
                config.QueryParameters.Select = new[]
                {
                    "id", "appId", "displayName", "description", "accountEnabled",
                    "signInAudience", "tags", "verifiedPublisher", "appOwnerOrganizationId",
                    "homepage", "servicePrincipalType"
                };
            });
            while (spPage?.Value != null)
            {
                spList.AddRange(spPage.Value);
                if (spPage.OdataNextLink == null) break;
                spPage = await _graphClient.ServicePrincipals.WithUrl(spPage.OdataNextLink).GetAsync();
            }

            // Build a set of appIds that are our own registrations (to flag them)
            var ownAppIds = new HashSet<string>(appList.Select(a => a.AppId ?? ""), StringComparer.OrdinalIgnoreCase);

            var enterpriseApps = spList.Select(sp =>
            {
                var isOwn = ownAppIds.Contains(sp.AppId ?? "");
                var ownerGuid = sp.AppOwnerOrganizationId?.ToString();
                var isMicrosoft =
                    ownerGuid == "f8cdef31-a31e-4b4a-93e4-5f571e91255a" ||
                    ownerGuid == "72f988bf-86f1-41af-91ab-2d7cd011db47" ||
                    sp.Tags?.Contains("WindowsAzureActiveDirectoryIntegratedApp") == true;

                // CreatedDateTime is not a typed property on ServicePrincipal in SDK v5 - read from AdditionalData
                DateTimeOffset? createdDt = null;
                if (sp.AdditionalData?.TryGetValue("createdDateTime", out var cdtRaw) == true
                    && cdtRaw is string cdtString
                    && DateTimeOffset.TryParse(cdtString, out var parsedDt))
                    createdDt = parsedDt;

                var createdDaysAgo = createdDt.HasValue ? (int)(now - createdDt.Value).TotalDays : (int?)null;

                return new
                {
                    id = sp.Id,
                    appId = sp.AppId,
                    displayName = sp.DisplayName ?? "(no name)",
                    description = sp.Description,
                    createdDateTime = createdDt,
                    createdDaysAgo,
                    isNew = createdDt.HasValue && createdDt.Value >= thirtyDaysAgo,
                    accountEnabled = sp.AccountEnabled,
                    signInAudience = sp.SignInAudience,
                    isOwnRegistration = isOwn,
                    isMicrosoftApp = isMicrosoft,
                    isVerified = sp.VerifiedPublisher != null,
                    publisherName = sp.VerifiedPublisher?.DisplayName,
                    homepage = sp.Homepage,
                    tags = sp.Tags,
                    appType = "EnterpriseApp"
                };
            }).ToList();

            // ── Summaries ──────────────────────────────────────────────────
            var recentRegistrations  = registrations.Where(r => r.isNew).OrderByDescending(r => r.createdDateTime).ToList();
            var noCredApps           = registrations.Where(r => r.noCredentials).OrderBy(r => r.displayName).ToList();
            var expiredCredApps      = registrations.Where(r => r.allCredentialsExpired).OrderBy(r => r.displayName).ToList();
            var recentEnterprise     = enterpriseApps.Where(e => e.isNew).OrderByDescending(e => e.createdDateTime).ToList();
            var disabledEnterprise   = enterpriseApps.Where(e => e.accountEnabled == false).OrderBy(e => e.displayName).ToList();
            var thirdPartyEnterprise = enterpriseApps.Where(e => !e.isMicrosoftApp && !e.isOwnRegistration).OrderBy(e => e.displayName).ToList();

            return Ok(new
            {
                summary = new
                {
                    totalRegistrations    = registrations.Count,
                    totalEnterpriseApps   = enterpriseApps.Count,
                    thirdPartyApps        = thirdPartyEnterprise.Count,
                    newRegistrations30d   = recentRegistrations.Count,
                    newEnterpriseApps30d  = recentEnterprise.Count,
                    noCredentials         = noCredApps.Count,
                    allCredentialsExpired = expiredCredApps.Count,
                    disabledEnterpriseApps = disabledEnterprise.Count
                },
                // App registrations
                recentRegistrations,
                noCredentialsApps    = noCredApps,
                expiredCredentialsApps = expiredCredApps,
                allRegistrationsByDate = registrations.OrderByDescending(r => r.createdDateTime).ToList(),
                // Enterprise apps
                recentEnterpriseApps   = recentEnterprise,
                allEnterpriseApps      = enterpriseApps.OrderByDescending(e => e.createdDateTime).ToList(),
                thirdPartyEnterpriseApps = thirdPartyEnterprise,
                disabledEnterpriseApps = disabledEnterprise,
                lastUpdated = now
            });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching enterprise app audit");
            return StatusCode(500, new { error = ex.Message });
        }
    }

    /// <summary>
    /// Get risky/suspicious app consents
    /// </summary>
    [HttpGet("risky-consents")]
    public async Task<IActionResult> GetRiskyConsents()
    {
        try
        {
            // Get all OAuth2 grants
            var grants = await _graphClient.Oauth2PermissionGrants
                .GetAsync(config =>
                {
                    config.QueryParameters.Top = 500;
                });

            var grantList = grants?.Value ?? new List<OAuth2PermissionGrant>();

            // Define high-risk permissions
            var highRiskPermissions = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "Directory.ReadWrite.All",
                "Directory.AccessAsUser.All",
                "User.ReadWrite.All",
                "User.ManageIdentities.All",
                "Group.ReadWrite.All",
                "RoleManagement.ReadWrite.Directory",
                "Application.ReadWrite.All",
                "AppRoleAssignment.ReadWrite.All",
                "Mail.ReadWrite",
                "Mail.Send",
                "Mail.Read",
                "MailboxSettings.ReadWrite",
                "Files.ReadWrite.All",
                "Sites.ReadWrite.All",
                "Sites.FullControl.All",
                "Calendars.ReadWrite",
                "Contacts.ReadWrite",
                "People.Read.All",
                "Notes.ReadWrite.All",
                "Chat.ReadWrite",
                "ChannelMessage.Send"
            };

            // Get service principal info
            var spIds = grantList.Select(g => g.ClientId).Distinct().ToList();
            var servicePrincipals = new Dictionary<string, ServicePrincipal>();

            foreach (var spId in spIds.Take(100))
            {
                try
                {
                    var sp = await _graphClient.ServicePrincipals[spId].GetAsync();
                    if (sp != null)
                    {
                        servicePrincipals[spId!] = sp;
                    }
                }
                catch { }
            }

            var riskyConsents = new List<object>();

            foreach (var grant in grantList)
            {
                var scopes = grant.Scope?.Split(' ', StringSplitOptions.RemoveEmptyEntries) ?? Array.Empty<string>();
                var riskyScopes = scopes.Where(s => highRiskPermissions.Contains(s)).ToList();

                if (riskyScopes.Any())
                {
                    servicePrincipals.TryGetValue(grant.ClientId ?? "", out var sp);

                    var riskScore = riskyScopes.Count;
                    var riskFactors = new List<string>();

                    if (riskyScopes.Any(s => s.Contains("Directory")))
                    {
                        riskScore += 3;
                        riskFactors.Add("Directory access");
                    }
                    if (riskyScopes.Any(s => s.Contains("Mail")))
                    {
                        riskScore += 2;
                        riskFactors.Add("Mail access");
                    }
                    if (riskyScopes.Any(s => s.Contains("ReadWrite")))
                    {
                        riskScore += 1;
                        riskFactors.Add("Write permissions");
                    }
                    if (sp?.VerifiedPublisher == null)
                    {
                        riskScore += 2;
                        riskFactors.Add("Unverified publisher");
                    }
                    if (grant.ConsentType == "AllPrincipals")
                    {
                        riskScore += 1;
                        riskFactors.Add("Admin consent (all users)");
                    }

                    riskyConsents.Add(new
                    {
                        grantId = grant.Id,
                        clientId = grant.ClientId,
                        appName = sp?.DisplayName ?? "Unknown App",
                        appId = sp?.AppId,
                        publisherName = sp?.AppOwnerOrganizationId?.ToString(),
                        isVerified = sp?.VerifiedPublisher != null,
                        consentType = grant.ConsentType,
                        allScopes = scopes,
                        riskyScopes,
                        riskScore,
                        riskLevel = riskScore >= 6 ? "High" : riskScore >= 3 ? "Medium" : "Low",
                        riskFactors,
                        principalId = grant.PrincipalId
                    });
                }
            }

            // Sort by risk score
            riskyConsents = riskyConsents.OrderByDescending(r => ((dynamic)r).riskScore).ToList();

            var summary = new
            {
                totalRiskyConsents = riskyConsents.Count,
                highRisk = riskyConsents.Count(r => ((dynamic)r).riskLevel == "High"),
                mediumRisk = riskyConsents.Count(r => ((dynamic)r).riskLevel == "Medium"),
                lowRisk = riskyConsents.Count(r => ((dynamic)r).riskLevel == "Low"),
                unverifiedApps = riskyConsents.Count(r => ((dynamic)r).isVerified == false),
                adminConsentCount = riskyConsents.Count(r => ((dynamic)r).consentType == "AllPrincipals")
            };

            return Ok(new
            {
                riskyConsents,
                summary,
                highRiskPermissionsList = highRiskPermissions.ToList(),
                lastUpdated = DateTime.UtcNow
            });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching risky consents");
            return Ok(new
            {
                riskyConsents = new List<object>(),
                summary = new { totalRiskyConsents = 0 },
                error = ex.Message
            });
        }
    }
}
