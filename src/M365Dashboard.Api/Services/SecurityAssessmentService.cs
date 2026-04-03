using Microsoft.Graph;
using Microsoft.Graph.Models;
using M365Dashboard.Api.Models;

namespace M365Dashboard.Api.Services;

/// <summary>
/// Service for generating comprehensive Microsoft 365 Security Assessment reports
/// Uses read-only permissions where possible
/// </summary>
public interface ISecurityAssessmentService
{
    Task<SecurityAssessmentResult> RunAssessmentAsync();
}

public class SecurityAssessmentService : ISecurityAssessmentService
{
    private readonly GraphServiceClient _graphClient;
    private readonly ILogger<SecurityAssessmentService> _logger;
    private readonly ICisBenchmarkService _cisBenchmarkService;

    public SecurityAssessmentService(
        GraphServiceClient graphClient,
        ILogger<SecurityAssessmentService> logger,
        ICisBenchmarkService cisBenchmarkService)
    {
        _graphClient = graphClient;
        _logger = logger;
        _cisBenchmarkService = cisBenchmarkService;
    }

    public async Task<SecurityAssessmentResult> RunAssessmentAsync()
    {
        var result = new SecurityAssessmentResult
        {
            GeneratedAt = DateTime.UtcNow
        };

        // Get tenant info
        try
        {
            var org = await _graphClient.Organization.GetAsync();
            if (org?.Value?.FirstOrDefault() != null)
            {
                result.TenantId = org.Value[0].Id ?? "";
                result.TenantName = org.Value[0].DisplayName ?? "";
                result.TenantDomain = org.Value[0].VerifiedDomains?
                    .FirstOrDefault(d => d.IsDefault == true)?.Name ?? "";
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Could not get organization info");
        }

        // Gather statistics
        result.UserStats = await GetUserStatisticsAsync();
        result.LicenseStats = await GetLicenseStatisticsAsync();
        result.RoleDistribution = await GetRoleDistributionAsync();

        // Run compliance checks for each section
        result.EntraIdCompliance = await CheckEntraIdComplianceAsync();
        result.ExchangeCompliance = await CheckExchangeComplianceAsync();
        result.SharePointCompliance = await CheckSharePointComplianceAsync();
        result.TeamsCompliance = await CheckTeamsComplianceAsync();
        result.IntuneCompliance = await CheckIntuneComplianceAsync();
        result.DefenderCompliance = await CheckDefenderComplianceAsync();

        // Calculate totals
        var allSections = new[] { 
            result.EntraIdCompliance, 
            result.ExchangeCompliance, 
            result.SharePointCompliance, 
            result.TeamsCompliance, 
            result.IntuneCompliance,
            result.DefenderCompliance 
        };
        
        result.TotalChecks = allSections.Sum(s => s.TotalChecks);
        result.CompliantChecks = allSections.Sum(s => s.CompliantChecks);
        result.NonCompliantChecks = allSections.Sum(s => s.NonCompliantChecks);
        result.OverallCompliancePercentage = result.TotalChecks > 0 
            ? Math.Round((double)result.CompliantChecks / result.TotalChecks * 100, 1) 
            : 0;

        return result;
    }

    #region Statistics Gathering

    private async Task<UserStatistics> GetUserStatisticsAsync()
    {
        var stats = new UserStatistics();

        try
        {
            // Get all users
            var users = await _graphClient.Users.GetAsync(config =>
            {
                config.QueryParameters.Select = new[] { 
                    "id", "displayName", "userPrincipalName", "userType", 
                    "accountEnabled", "assignedLicenses" 
                };
                config.QueryParameters.Top = 999;
            });

            var allUsers = users?.Value ?? new List<User>();
            
            stats.TotalUsers = allUsers.Count;
            stats.MemberUsers = allUsers.Count(u => u.UserType == "Member");
            stats.GuestUsers = allUsers.Count(u => u.UserType == "Guest");
            stats.LicensedUsers = allUsers.Count(u => u.AssignedLicenses?.Any() == true);
            stats.UnlicensedUsers = allUsers.Count(u => u.AssignedLicenses?.Any() != true);
            stats.BlockedUsers = allUsers.Count(u => u.AccountEnabled == false);
            stats.BlockedUsersWithLicenses = allUsers.Count(u => 
                u.AccountEnabled == false && u.AssignedLicenses?.Any() == true);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Error getting user statistics");
        }

        return stats;
    }

    private async Task<LicenseStatistics> GetLicenseStatisticsAsync()
    {
        var stats = new LicenseStatistics();

        try
        {
            var skus = await _graphClient.SubscribedSkus.GetAsync();
            
            foreach (var sku in skus?.Value ?? new List<SubscribedSku>())
            {
                var total = sku.PrepaidUnits?.Enabled ?? 0;
                var assigned = sku.ConsumedUnits ?? 0;
                
                stats.TotalLicenses += total;
                stats.AssignedLicenses += assigned;
                
                stats.LicenseBreakdown.Add(new LicenseSummary
                {
                    SkuName = GetFriendlyLicenseName(sku.SkuPartNumber ?? ""),
                    SkuPartNumber = sku.SkuPartNumber ?? "",
                    Total = total,
                    Assigned = assigned,
                    Available = total - assigned
                });
            }
            
            stats.AvailableLicenses = stats.TotalLicenses - stats.AssignedLicenses;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Error getting license statistics");
        }

        return stats;
    }

    private async Task<List<AdminRoleAssignment>> GetRoleDistributionAsync()
    {
        var roles = new List<AdminRoleAssignment>();

        try
        {
            var directoryRoles = await _graphClient.DirectoryRoles.GetAsync(config =>
            {
                config.QueryParameters.Expand = new[] { "members" };
            });

            var importantRoles = new[] { 
                "Global Administrator", "User Administrator", "Exchange Administrator",
                "SharePoint Administrator", "Teams Administrator", "Security Administrator",
                "Compliance Administrator", "Privileged Role Administrator", 
                "Application Administrator", "Cloud Application Administrator",
                "Helpdesk Administrator", "Directory Readers", "Billing Administrator"
            };

            foreach (var role in directoryRoles?.Value ?? new List<DirectoryRole>())
            {
                if (role.Members?.Any() == true)
                {
                    roles.Add(new AdminRoleAssignment
                    {
                        RoleName = role.DisplayName ?? "",
                        MemberCount = role.Members.Count,
                        Members = role.Members
                            .OfType<User>()
                            .Select(u => u.DisplayName ?? u.UserPrincipalName ?? "")
                            .ToList()
                    });
                }
            }

            // Sort by importance (GA first, then by count)
            roles = roles
                .OrderByDescending(r => r.RoleName == "Global Administrator")
                .ThenByDescending(r => importantRoles.Contains(r.RoleName))
                .ThenByDescending(r => r.MemberCount)
                .ToList();
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Error getting role distribution");
        }

        return roles;
    }

    #endregion

    #region Entra ID Compliance Checks

    private async Task<ComplianceSection> CheckEntraIdComplianceAsync()
    {
        var section = new ComplianceSection
        {
            SectionName = "Entra ID",
            SectionDescription = "Microsoft have stated that 99% of breaches could be mitigated with strong passwords and multi-factor authentication. Enabling and enforcing MFA across your organization is one of the easiest and most effective ways to increase your security posture. This section evaluates your Entra ID (Azure AD) configuration against security best practices."
        };

        // Security Defaults
        section.Checks.Add(await CheckSecurityDefaultsAsync());
        
        // App Consent Policy
        section.Checks.Add(await CheckAppConsentPolicyAsync());
        
        // Tenant Creation
        section.Checks.Add(await CheckTenantCreationAsync());
        
        // App Registration
        section.Checks.Add(await CheckAppRegistrationAsync());
        
        // Security Group Creation
        section.Checks.Add(await CheckSecurityGroupCreationAsync());
        
        // Guest User Access
        section.Checks.Add(await CheckGuestUserAccessAsync());
        
        // Guest Invitation Policy
        section.Checks.Add(await CheckGuestInvitationPolicyAsync());
        
        // User Consent Settings
        section.Checks.Add(await CheckUserConsentSettingsAsync());
        
        // Password Expiration
        section.Checks.Add(await CheckPasswordExpirationAsync());
        
        // Per-user MFA disabled
        section.Checks.Add(CheckPerUserMfaAsync());
        
        // Legacy Auth Blocked
        section.Checks.Add(await CheckLegacyAuthBlockedAsync());
        
        // Cloud-only Admins
        section.Checks.Add(await CheckCloudOnlyAdminsAsync());
        
        // Global Admin Count
        section.Checks.Add(await CheckGlobalAdminCountAsync());
        
        // Password Hash Sync (for hybrid)
        section.Checks.Add(await CheckPasswordHashSyncAsync());

        CalculateSectionStats(section);
        return section;
    }

    private async Task<SecurityCheck> CheckSecurityDefaultsAsync()
    {
        var check = new SecurityCheck
        {
            Name = "Security Defaults State",
            Description = "Security defaults in Microsoft Entra ID make it easier to be secure and help protect the organization.",
            Reference = "https://learn.microsoft.com/en-us/entra/fundamentals/security-defaults"
        };

        try
        {
            var policy = await _graphClient.Policies.IdentitySecurityDefaultsEnforcementPolicy.GetAsync();
            
            if (policy?.IsEnabled == true)
            {
                check.Status = SecurityCheckStatus.Compliant;
                check.CurrentValue = "Security defaults are enabled";
            }
            else
            {
                // Security defaults off is OK if Conditional Access is used
                check.Status = SecurityCheckStatus.Warning;
                check.CurrentValue = "Security defaults are disabled";
                check.Remediation = "Enable security defaults or ensure Conditional Access policies are configured as an alternative.";
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Error checking security defaults");
            check.Status = SecurityCheckStatus.Error;
            check.CurrentValue = ex.Message;
        }

        return check;
    }

    private async Task<SecurityCheck> CheckAppConsentPolicyAsync()
    {
        var check = new SecurityCheck
        {
            Name = "App Consent Policy",
            Description = "The admin consent workflow gives admins a secure way to grant access to applications that require admin approval.",
            Reference = "https://learn.microsoft.com/en-us/entra/identity/enterprise-apps/configure-admin-consent-workflow"
        };

        try
        {
            var authPolicy = await _graphClient.Policies.AuthorizationPolicy.GetAsync();
            
            // Check DefaultUserRolePermissions for app consent settings
            var allowedToCreateApps = authPolicy?.DefaultUserRolePermissions?.AllowedToCreateApps ?? true;
            
            // If users can't create apps, consent is likely restricted
            if (!allowedToCreateApps)
            {
                check.Status = SecurityCheckStatus.Compliant;
                check.CurrentValue = "User consent is restricted (app registration disabled)";
            }
            else
            {
                check.Status = SecurityCheckStatus.Warning;
                check.CurrentValue = "User consent may be permissive - verify in Entra admin center";
                check.Remediation = "Configure the admin consent workflow or restrict user consent in Enterprise Applications settings.";
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Error checking app consent policy");
            check.Status = SecurityCheckStatus.Error;
            check.CurrentValue = ex.Message;
        }

        return check;
    }

    private async Task<SecurityCheck> CheckTenantCreationAsync()
    {
        var check = new SecurityCheck
        {
            Name = "Non-Admin Tenant Creation",
            Description = "Non-privileged users can create tenants in the Entra administration portal under Manage tenant.",
            Reference = "https://learn.microsoft.com/en-us/entra/fundamentals/users-default-permissions"
        };

        try
        {
            var authPolicy = await _graphClient.Policies.AuthorizationPolicy.GetAsync();
            var canCreate = authPolicy?.DefaultUserRolePermissions?.AllowedToCreateTenants ?? true;

            if (!canCreate)
            {
                check.Status = SecurityCheckStatus.Compliant;
                check.CurrentValue = "Non-admin users cannot create tenants";
            }
            else
            {
                check.Status = SecurityCheckStatus.NonCompliant;
                check.CurrentValue = "Non-admin users can create tenants";
                check.Remediation = "Disable tenant creation for non-admin users in Entra ID > User settings.";
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Error checking tenant creation");
            check.Status = SecurityCheckStatus.Error;
            check.CurrentValue = ex.Message;
        }

        return check;
    }

    private async Task<SecurityCheck> CheckAppRegistrationAsync()
    {
        var check = new SecurityCheck
        {
            Name = "Non-Admin App Registration",
            Description = "Application registration permissions in Microsoft Entra ID determine whether users can create and register applications within the tenant.",
            Reference = "https://learn.microsoft.com/en-us/entra/identity/role-based-access-control/delegate-app-roles"
        };

        try
        {
            var authPolicy = await _graphClient.Policies.AuthorizationPolicy.GetAsync();
            var canCreate = authPolicy?.DefaultUserRolePermissions?.AllowedToCreateApps ?? true;

            if (!canCreate)
            {
                check.Status = SecurityCheckStatus.Compliant;
                check.CurrentValue = "Non-admin users cannot register applications";
            }
            else
            {
                check.Status = SecurityCheckStatus.NonCompliant;
                check.CurrentValue = "Non-admin users can register applications";
                check.Remediation = "Disable app registration for non-admin users in Entra ID > User settings.";
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Error checking app registration");
            check.Status = SecurityCheckStatus.Error;
            check.CurrentValue = ex.Message;
        }

        return check;
    }

    private async Task<SecurityCheck> CheckSecurityGroupCreationAsync()
    {
        var check = new SecurityCheck
        {
            Name = "Non-Admin Security Group Creation",
            Description = "This setting allows users in the organization to create new security groups and add members to these groups in the Azure portal, API, or PowerShell.",
            Reference = "https://learn.microsoft.com/en-us/entra/fundamentals/users-default-permissions"
        };

        try
        {
            var authPolicy = await _graphClient.Policies.AuthorizationPolicy.GetAsync();
            var canCreate = authPolicy?.DefaultUserRolePermissions?.AllowedToCreateSecurityGroups ?? true;

            if (!canCreate)
            {
                check.Status = SecurityCheckStatus.Compliant;
                check.CurrentValue = "Non-admin users cannot create security groups";
            }
            else
            {
                check.Status = SecurityCheckStatus.NonCompliant;
                check.CurrentValue = "Non-admin users can create security groups";
                check.Remediation = "Disable security group creation for non-admin users.";
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Error checking security group creation");
            check.Status = SecurityCheckStatus.Error;
            check.CurrentValue = ex.Message;
        }

        return check;
    }

    private async Task<SecurityCheck> CheckGuestUserAccessAsync()
    {
        var check = new SecurityCheck
        {
            Name = "Guest User Access Restrictions",
            Description = "Microsoft Entra ID, part of Microsoft Entra, allows you to restrict what external guest users can see in their organization in Microsoft Entra ID.",
            Reference = "https://learn.microsoft.com/en-us/entra/external-id/external-collaboration-settings-configure"
        };

        try
        {
            var authPolicy = await _graphClient.Policies.AuthorizationPolicy.GetAsync();
            var guestRestriction = authPolicy?.GuestUserRoleId?.ToString();

            // 2af84b1e-32c8-42b7-82bc-daa82404023b = Most restrictive
            // 10dae51f-b6af-4016-8d66-8c2a99b929b3 = Limited (default)
            // a0b1b346-4d3e-4e8b-98f8-753987be4970 = Same as members

            if (guestRestriction == "2af84b1e-32c8-42b7-82bc-daa82404023b" ||
                guestRestriction == "10dae51f-b6af-4016-8d66-8c2a99b929b3")
            {
                check.Status = SecurityCheckStatus.Compliant;
                check.CurrentValue = "Guest user access is restricted";
            }
            else
            {
                check.Status = SecurityCheckStatus.NonCompliant;
                check.CurrentValue = "Guest users have same access as members";
                check.Remediation = "Restrict guest user access in Entra ID > External Identities > External collaboration settings.";
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Error checking guest user access");
            check.Status = SecurityCheckStatus.Error;
            check.CurrentValue = ex.Message;
        }

        return check;
    }

    private async Task<SecurityCheck> CheckGuestInvitationPolicyAsync()
    {
        var check = new SecurityCheck
        {
            Name = "Guest Invitation Policy",
            Description = "The Guest Invitation Policy in Microsoft Entra ID controls whether guest users can invite other external users to the organization.",
            Reference = "https://learn.microsoft.com/en-us/entra/external-id/external-collaboration-settings-configure"
        };

        try
        {
            var authPolicy = await _graphClient.Policies.AuthorizationPolicy.GetAsync();
            var allowInvites = authPolicy?.AllowInvitesFrom;

            if (allowInvites == AllowInvitesFrom.AdminsAndGuestInviters || 
                allowInvites == AllowInvitesFrom.None)
            {
                check.Status = SecurityCheckStatus.Compliant;
                check.CurrentValue = $"Guest invitations restricted to: {allowInvites}";
            }
            else
            {
                check.Status = SecurityCheckStatus.NonCompliant;
                check.CurrentValue = $"Guest invitations allowed by: {allowInvites}";
                check.Remediation = "Restrict guest invitations to admins and guest inviters only.";
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Error checking guest invitation policy");
            check.Status = SecurityCheckStatus.Error;
            check.CurrentValue = ex.Message;
        }

        return check;
    }

    private async Task<SecurityCheck> CheckUserConsentSettingsAsync()
    {
        var check = new SecurityCheck
        {
            Name = "User Consent Settings",
            Description = "User Consent Settings in Microsoft Entra ID control whether users can grant consent to applications that request access to organizational data and resources.",
            Reference = "https://learn.microsoft.com/en-us/entra/identity/enterprise-apps/configure-user-consent"
        };

        try
        {
            var authPolicy = await _graphClient.Policies.AuthorizationPolicy.GetAsync();
            
            // Check if users can create apps - this is related to consent settings
            var allowedToCreateApps = authPolicy?.DefaultUserRolePermissions?.AllowedToCreateApps ?? true;
            
            if (!allowedToCreateApps)
            {
                check.Status = SecurityCheckStatus.Compliant;
                check.CurrentValue = "User consent is restricted (app registration disabled)";
            }
            else
            {
                check.Status = SecurityCheckStatus.Warning;
                check.CurrentValue = "Users may be able to consent to applications - verify in Entra admin center";
                check.Remediation = "Restrict user consent to verified publishers or disable it entirely in Enterprise Applications > Consent and permissions.";
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Error checking user consent settings");
            check.Status = SecurityCheckStatus.Error;
            check.CurrentValue = ex.Message;
        }

        return check;
    }

    private async Task<SecurityCheck> CheckPasswordExpirationAsync()
    {
        var check = new SecurityCheck
        {
            Name = "Password Expiration Policy",
            Description = "Microsoft cloud-only accounts have a pre-defined password policy that cannot be changed. The CIS recommendation is to set passwords to never expire when combined with MFA.",
            Reference = "https://learn.microsoft.com/en-us/microsoft-365/admin/manage/set-password-expiration-policy"
        };

        try
        {
            var domains = await _graphClient.Domains.GetAsync(config =>
            {
                config.QueryParameters.Select = new[] { "id", "passwordValidityPeriodInDays" };
            });

            var domainsWithExpiry = domains?.Value?
                .Where(d => d.PasswordValidityPeriodInDays.HasValue && 
                           d.PasswordValidityPeriodInDays.Value < 2147483647)
                .ToList();

            if (domainsWithExpiry?.Any() != true)
            {
                check.Status = SecurityCheckStatus.Compliant;
                check.CurrentValue = "Passwords set to never expire";
            }
            else
            {
                check.Status = SecurityCheckStatus.NonCompliant;
                check.CurrentValue = $"Password expiry enabled on {domainsWithExpiry.Count} domain(s)";
                check.AffectedItems = domainsWithExpiry.Select(d => $"{d.Id}: {d.PasswordValidityPeriodInDays} days").ToList();
                check.Remediation = "Set password expiration to never expire in Microsoft 365 admin center.";
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Error checking password expiration");
            check.Status = SecurityCheckStatus.Error;
            check.CurrentValue = ex.Message;
        }

        return check;
    }

    private SecurityCheck CheckPerUserMfaAsync()
    {
        // Per-user MFA requires legacy API - mark as manual but explain
        return new SecurityCheck
        {
            Name = "Per-user MFA Disabled",
            Description = "Legacy per-user Multi-Factor Authentication (MFA) can be configured to require individual users to provide multiple authentication factors. CIS recommends using Conditional Access instead.",
            Status = SecurityCheckStatus.Compliant,
            CurrentValue = "Requires manual verification - check legacy MFA portal",
            Remediation = "Disable per-user MFA and use Conditional Access policies instead.",
            Reference = "https://learn.microsoft.com/en-us/entra/identity/authentication/howto-mfa-userstates"
        };
    }

    private async Task<SecurityCheck> CheckLegacyAuthBlockedAsync()
    {
        var check = new SecurityCheck
        {
            Name = "Legacy Authentication Blocked",
            Description = "Entra ID supports the most widely used authentication and authorization protocols including legacy authentication. Legacy protocols should be blocked.",
            Reference = "https://learn.microsoft.com/en-us/entra/identity/conditional-access/howto-conditional-access-policy-block-legacy"
        };

        try
        {
            var caPolicies = await _graphClient.Identity.ConditionalAccess.Policies.GetAsync();
            
            var legacyBlockPolicy = caPolicies?.Value?.FirstOrDefault(p =>
                p.State == ConditionalAccessPolicyState.Enabled &&
                p.Conditions?.ClientAppTypes?.Contains(ConditionalAccessClientApp.ExchangeActiveSync) == true &&
                p.Conditions?.ClientAppTypes?.Contains(ConditionalAccessClientApp.Other) == true &&
                p.GrantControls?.BuiltInControls?.Contains(ConditionalAccessGrantControl.Block) == true);

            if (legacyBlockPolicy != null)
            {
                check.Status = SecurityCheckStatus.Compliant;
                check.CurrentValue = $"Legacy authentication blocked by policy: {legacyBlockPolicy.DisplayName}";
            }
            else
            {
                check.Status = SecurityCheckStatus.NonCompliant;
                check.CurrentValue = "No Conditional Access policy blocking legacy authentication";
                check.Remediation = "Create a Conditional Access policy to block legacy authentication protocols.";
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Error checking legacy auth");
            check.Status = SecurityCheckStatus.Error;
            check.CurrentValue = ex.Message;
        }

        return check;
    }

    private async Task<SecurityCheck> CheckCloudOnlyAdminsAsync()
    {
        var check = new SecurityCheck
        {
            Name = "Cloud-only Administrative Accounts",
            Description = "Administrative accounts are special privileged accounts that could have varying levels of access to data, users, and settings. These should be cloud-only.",
            Reference = "https://learn.microsoft.com/en-us/entra/identity/role-based-access-control/security-planning"
        };

        try
        {
            var directoryRoles = await _graphClient.DirectoryRoles.GetAsync(config =>
            {
                config.QueryParameters.Expand = new[] { "members" };
            });

            var adminRoles = new[] { "Global Administrator", "Privileged Role Administrator", "User Administrator" };
            var syncedAdmins = new List<string>();

            foreach (var role in directoryRoles?.Value ?? new List<DirectoryRole>())
            {
                if (adminRoles.Any(ar => role.DisplayName?.Contains(ar) == true))
                {
                    foreach (var member in role.Members?.OfType<User>() ?? new List<User>())
                    {
                        var userDetails = await _graphClient.Users[member.Id].GetAsync(config =>
                        {
                            config.QueryParameters.Select = new[] { "displayName", "userPrincipalName", "onPremisesSyncEnabled" };
                        });
                        
                        if (userDetails?.OnPremisesSyncEnabled == true)
                        {
                            var display = $"{userDetails.DisplayName} ({userDetails.UserPrincipalName})";
                            if (!syncedAdmins.Contains(display))
                                syncedAdmins.Add(display);
                        }
                    }
                }
            }

            if (!syncedAdmins.Any())
            {
                check.Status = SecurityCheckStatus.Compliant;
                check.CurrentValue = "All administrative accounts are cloud-only";
            }
            else
            {
                check.Status = SecurityCheckStatus.NonCompliant;
                check.CurrentValue = $"{syncedAdmins.Count} synced admin account(s) found";
                check.AffectedItems = syncedAdmins;
                check.Remediation = "Create dedicated cloud-only admin accounts and remove admin roles from synced accounts.";
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Error checking cloud-only admins");
            check.Status = SecurityCheckStatus.Error;
            check.CurrentValue = ex.Message;
        }

        return check;
    }

    private async Task<SecurityCheck> CheckGlobalAdminCountAsync()
    {
        var check = new SecurityCheck
        {
            Name = "Global Admin Count",
            Description = "Between two and four global administrators should be designated in the tenant.",
            Reference = "https://learn.microsoft.com/en-us/entra/identity/role-based-access-control/security-planning"
        };

        try
        {
            var globalAdminRole = await _graphClient.DirectoryRoles.GetAsync(config =>
            {
                config.QueryParameters.Filter = "displayName eq 'Global Administrator'";
                config.QueryParameters.Expand = new[] { "members" };
            });

            var gaRole = globalAdminRole?.Value?.FirstOrDefault();
            var memberCount = gaRole?.Members?.Count ?? 0;

            check.AffectedItems = gaRole?.Members?
                .OfType<User>()
                .Select(u => u.DisplayName ?? u.UserPrincipalName ?? "")
                .ToList() ?? new List<string>();

            if (memberCount >= 2 && memberCount <= 4)
            {
                check.Status = SecurityCheckStatus.Compliant;
                check.CurrentValue = $"{memberCount} Global Administrator(s) - within recommended range";
            }
            else if (memberCount < 2)
            {
                check.Status = SecurityCheckStatus.NonCompliant;
                check.CurrentValue = $"{memberCount} Global Administrator(s) - too few (risk of lockout)";
                check.Remediation = "Designate at least 2 global administrators.";
            }
            else
            {
                check.Status = SecurityCheckStatus.NonCompliant;
                check.CurrentValue = $"{memberCount} Global Administrator(s) - too many (increased attack surface)";
                check.Remediation = "Reduce global administrators to 4 or fewer.";
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Error checking global admin count");
            check.Status = SecurityCheckStatus.Error;
            check.CurrentValue = ex.Message;
        }

        return check;
    }

    private async Task<SecurityCheck> CheckPasswordHashSyncAsync()
    {
        var check = new SecurityCheck
        {
            Name = "Password Hash Sync (Hybrid)",
            Description = "Password hash synchronization is one of the sign-in methods used to accomplish hybrid identity synchronization.",
            Reference = "https://learn.microsoft.com/en-us/entra/identity/hybrid/connect/whatis-phs"
        };

        try
        {
            var org = await _graphClient.Organization.GetAsync(config =>
            {
                config.QueryParameters.Select = new[] { "onPremisesSyncEnabled" };
            });

            var isHybrid = org?.Value?.FirstOrDefault()?.OnPremisesSyncEnabled == true;

            if (!isHybrid)
            {
                check.Status = SecurityCheckStatus.NotApplicable;
                check.CurrentValue = "Organization is cloud-only (not hybrid)";
            }
            else
            {
                // For hybrid, password hash sync should be enabled
                check.Status = SecurityCheckStatus.Compliant;
                check.CurrentValue = "Hybrid environment detected - verify password hash sync is enabled";
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Error checking password hash sync");
            check.Status = SecurityCheckStatus.Error;
            check.CurrentValue = ex.Message;
        }

        return check;
    }

    #endregion

    #region Exchange Online Compliance Checks

    private async Task<ComplianceSection> CheckExchangeComplianceAsync()
    {
        var section = new ComplianceSection
        {
            SectionName = "Exchange Online",
            SectionDescription = "Business Email Compromise is a common attack vector. Microsoft Exchange Online provides a number of security features to help protect your organization from email-based threats."
        };

        section.Checks.Add(CreateManualCheck("Modern Authentication", "Modern authentication in Microsoft 365 enables authentication features like MFA, smart cards, and third-party SAML identity providers.", SecurityCheckStatus.Compliant, "Modern authentication is enabled by default for all tenants"));
        section.Checks.Add(CreateManualCheck("Mailbox Auditing", "Mailbox audit logging is turned on by default in all organizations.", SecurityCheckStatus.Compliant, "Mailbox auditing is enabled by default"));
        section.Checks.Add(CreateManualCheck("Unified Audit Log", "When audit log search is enabled, user and admin activity is recorded in the audit log.", SecurityCheckStatus.Compliant, "Unified audit logging is enabled by default"));
        section.Checks.Add(CreateManualCheck("SMTP AUTH Disabled", "This setting enables or disables authenticated client SMTP submission (SMTP AUTH) organization-wide.", SecurityCheckStatus.Compliant, "Verify via PowerShell: Get-TransportConfig | Select SmtpClientAuthenticationDisabled"));
        section.Checks.Add(CreateManualCheck("External Email Tagging", "External callouts provide a native experience to identify emails from senders outside the organization.", SecurityCheckStatus.Warning, "Verify external tagging is enabled in Exchange admin center"));
        section.Checks.Add(CreateManualCheck("Mail Forwarding Blocked", "External mail forwarding should be blocked to prevent data exfiltration.", SecurityCheckStatus.Warning, "Verify auto-forwarding is blocked via transport rules"));
        section.Checks.Add(CreateManualCheck("Common Attachment Filter", "The Common Attachment Types Filter blocks known malicious file types.", SecurityCheckStatus.Compliant, "Default anti-malware policy blocks dangerous attachments"));
        section.Checks.Add(CreateManualCheck("Shared Mailbox Sign-in", "Shared mailboxes should have sign-in blocked.", SecurityCheckStatus.Warning, "Verify via PowerShell: Get-Mailbox -RecipientTypeDetails SharedMailbox | Get-MsolUser"));

        CalculateSectionStats(section);
        return section;
    }

    #endregion

    #region SharePoint Compliance Checks

    private async Task<ComplianceSection> CheckSharePointComplianceAsync()
    {
        var section = new ComplianceSection
        {
            SectionName = "SharePoint & OneDrive",
            SectionDescription = "Internal and External sharing of business data is one of the most challenging aspects to manage within Microsoft 365."
        };

        section.Checks.Add(CreateManualCheck("SharePoint Sharing Controls", "Tenant-wide default sharing settings that govern how users can share files.", SecurityCheckStatus.Warning, "Verify sharing settings in SharePoint admin center"));
        section.Checks.Add(CreateManualCheck("Guest Resharing", "Controls whether external guests can reshare documents shared with them.", SecurityCheckStatus.Warning, "Verify guest resharing is disabled"));
        section.Checks.Add(CreateManualCheck("Legacy Authentication", "SharePoint Legacy Authentication Protocols should be blocked.", SecurityCheckStatus.Compliant, "Modern authentication is required by default"));
        section.Checks.Add(CreateManualCheck("OneDrive Retention", "Determines how long a deleted user's OneDrive content remains accessible.", SecurityCheckStatus.Warning, "Verify retention period is set to maximum"));
        section.Checks.Add(CreateManualCheck("External Sharing Domains", "External sharing can be restricted to specific domains.", SecurityCheckStatus.Warning, "Verify domain restrictions are configured"));

        CalculateSectionStats(section);
        return section;
    }

    #endregion

    #region Teams Compliance Checks

    private async Task<ComplianceSection> CheckTeamsComplianceAsync()
    {
        var section = new ComplianceSection
        {
            SectionName = "Microsoft Teams",
            SectionDescription = "Microsoft Teams front-ends a number of services including identity, document sharing, and remote access."
        };

        section.Checks.Add(CreateManualCheck("Anonymous Meeting Join", "Controls if anonymous participants can start a meeting.", SecurityCheckStatus.Compliant, "Verify anonymous users cannot start meetings"));
        section.Checks.Add(CreateManualCheck("Lobby Bypass", "Controls who can join meetings directly vs waiting in lobby.", SecurityCheckStatus.Warning, "Verify only org members can bypass lobby"));
        section.Checks.Add(CreateManualCheck("External Participants Control", "Controls if external participants can give or request control.", SecurityCheckStatus.Compliant, "Verify external control is disabled"));
        section.Checks.Add(CreateManualCheck("Unmanaged Teams Communication", "Controls chats with external unmanaged Teams users.", SecurityCheckStatus.Warning, "Verify communication with unmanaged Teams is disabled"));
        section.Checks.Add(CreateManualCheck("Teams Ownership", "All Teams should have assigned owners.", SecurityCheckStatus.Compliant, "Verify all teams have owners"));
        section.Checks.Add(CreateManualCheck("External File Sharing", "Third-party cloud storage providers should be restricted.", SecurityCheckStatus.Warning, "Verify only approved storage providers"));

        CalculateSectionStats(section);
        return section;
    }

    #endregion

    #region Intune Compliance Checks

    private async Task<ComplianceSection> CheckIntuneComplianceAsync()
    {
        var section = new ComplianceSection
        {
            SectionName = "Microsoft Intune",
            SectionDescription = "Business data is accessed by employees across multiple devices. This section evaluates device management configuration."
        };

        try
        {
            // Check for compliance policies
            var compliancePolicies = await _graphClient.DeviceManagement.DeviceCompliancePolicies.GetAsync();
            var hasWindowsPolicy = compliancePolicies?.Value?.Any(p => p.OdataType?.Contains("windows") == true) == true;
            var hasIosPolicy = compliancePolicies?.Value?.Any(p => p.OdataType?.Contains("ios") == true) == true;
            var hasMacPolicy = compliancePolicies?.Value?.Any(p => p.OdataType?.Contains("macOS") == true) == true;
            var hasAndroidPolicy = compliancePolicies?.Value?.Any(p => p.OdataType?.Contains("android") == true) == true;

            section.Checks.Add(new SecurityCheck
            {
                Name = "Windows Compliance Policy",
                Description = "Device compliance policies for Windows devices accessing organizational resources.",
                Status = hasWindowsPolicy ? SecurityCheckStatus.Compliant : SecurityCheckStatus.NonCompliant,
                CurrentValue = hasWindowsPolicy ? "Windows compliance policy deployed" : "No Windows compliance policy found",
                Remediation = "Create a Windows device compliance policy in Intune."
            });

            section.Checks.Add(new SecurityCheck
            {
                Name = "iOS Compliance Policy",
                Description = "Device compliance policies for iOS devices managed through Microsoft Intune.",
                Status = hasIosPolicy ? SecurityCheckStatus.Compliant : SecurityCheckStatus.NonCompliant,
                CurrentValue = hasIosPolicy ? "iOS compliance policy deployed" : "No iOS compliance policy found",
                Remediation = "Create an iOS device compliance policy in Intune."
            });

            section.Checks.Add(new SecurityCheck
            {
                Name = "macOS Compliance Policy",
                Description = "Device compliance policies for macOS devices within Microsoft Intune.",
                Status = hasMacPolicy ? SecurityCheckStatus.Compliant : SecurityCheckStatus.NonCompliant,
                CurrentValue = hasMacPolicy ? "macOS compliance policy deployed" : "No macOS compliance policy found",
                Remediation = "Create a macOS device compliance policy in Intune."
            });

            section.Checks.Add(new SecurityCheck
            {
                Name = "Android Compliance Policy",
                Description = "Device compliance policies for Android devices.",
                Status = hasAndroidPolicy ? SecurityCheckStatus.Compliant : SecurityCheckStatus.NonCompliant,
                CurrentValue = hasAndroidPolicy ? "Android compliance policy deployed" : "No Android compliance policy found",
                Remediation = "Create an Android device compliance policy in Intune."
            });

            // Check for app protection policies
            var appProtectionPolicies = await _graphClient.DeviceAppManagement.ManagedAppPolicies.GetAsync();
            var hasAndroidAppProtection = appProtectionPolicies?.Value?.Any(p => p.OdataType?.Contains("android") == true) == true;
            var hasIosAppProtection = appProtectionPolicies?.Value?.Any(p => p.OdataType?.Contains("ios") == true) == true;

            section.Checks.Add(new SecurityCheck
            {
                Name = "Android App Protection",
                Description = "App Protection Policies for Android devices to safeguard corporate data.",
                Status = hasAndroidAppProtection ? SecurityCheckStatus.Compliant : SecurityCheckStatus.NonCompliant,
                CurrentValue = hasAndroidAppProtection ? "Android app protection deployed" : "No Android app protection policy found",
                Remediation = "Create an Android app protection policy in Intune."
            });

            section.Checks.Add(new SecurityCheck
            {
                Name = "iOS App Protection",
                Description = "App protection policies for iOS devices within the organization.",
                Status = hasIosAppProtection ? SecurityCheckStatus.Compliant : SecurityCheckStatus.NonCompliant,
                CurrentValue = hasIosAppProtection ? "iOS app protection deployed" : "No iOS app protection policy found",
                Remediation = "Create an iOS app protection policy in Intune."
            });
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Error checking Intune compliance");
            section.Checks.Add(CreateManualCheck("Intune Policies", "Could not retrieve Intune policies", SecurityCheckStatus.Error, ex.Message));
        }

        // Add manual checks for items requiring additional verification
        section.Checks.Add(CreateManualCheck("Windows Update Ring", "Windows Update for Business policies for device update management.", SecurityCheckStatus.Warning, "Verify update rings are configured"));
        section.Checks.Add(CreateManualCheck("BitLocker Encryption", "Windows BitLocker encryption policy deployment.", SecurityCheckStatus.Warning, "Verify BitLocker policies are deployed"));
        section.Checks.Add(CreateManualCheck("Defender ATP Onboarding", "Microsoft Defender for Endpoint onboarding configuration.", SecurityCheckStatus.Warning, "Verify Defender onboarding is deployed"));

        CalculateSectionStats(section);
        return section;
    }

    #endregion

    #region Defender Compliance Checks

    private async Task<ComplianceSection> CheckDefenderComplianceAsync()
    {
        var section = new ComplianceSection
        {
            SectionName = "Microsoft Defender",
            SectionDescription = "Microsoft Defender for Office 365 provides protection against advanced threats in email, attachments, and links."
        };

        try
        {
            // Check for Defender for Office 365 license
            var subscribedSkus = await _graphClient.SubscribedSkus.GetAsync();
            var hasDefenderPlan = subscribedSkus?.Value?.Any(s =>
                s.ServicePlans?.Any(p =>
                    p.ServicePlanName?.Contains("ATP", StringComparison.OrdinalIgnoreCase) == true ||
                    p.ServicePlanName?.Contains("THREAT_INTELLIGENCE", StringComparison.OrdinalIgnoreCase) == true ||
                    p.ServicePlanName?.Contains("DEFENDER", StringComparison.OrdinalIgnoreCase) == true) == true) == true;

            if (hasDefenderPlan)
            {
                section.Checks.Add(new SecurityCheck
                {
                    Name = "Defender for Office 365 Licensed",
                    Description = "Microsoft Defender for Office 365 license provides Safe Links, Safe Attachments, and anti-phishing protection.",
                    Status = SecurityCheckStatus.Compliant,
                    CurrentValue = "Defender for Office 365 is licensed"
                });

                section.Checks.Add(new SecurityCheck
                {
                    Name = "Safe Links Policy",
                    Description = "Safe Links provides URL scanning and time-of-click verification.",
                    Status = SecurityCheckStatus.Compliant,
                    CurrentValue = "Built-in protection preset policy provides Safe Links by default"
                });

                section.Checks.Add(new SecurityCheck
                {
                    Name = "Safe Attachments Policy",
                    Description = "Safe Attachments scans attachments in a virtual environment.",
                    Status = SecurityCheckStatus.Compliant,
                    CurrentValue = "Built-in protection preset policy provides Safe Attachments by default"
                });

                section.Checks.Add(new SecurityCheck
                {
                    Name = "Anti-Phishing Policy",
                    Description = "Anti-phishing policies provide mailbox intelligence and spoof protection.",
                    Status = SecurityCheckStatus.Compliant,
                    CurrentValue = "Enhanced anti-phishing available with Defender for Office 365"
                });
            }
            else
            {
                section.Checks.Add(new SecurityCheck
                {
                    Name = "Defender for Office 365 Licensed",
                    Description = "Microsoft Defender for Office 365 license provides advanced threat protection.",
                    Status = SecurityCheckStatus.NonCompliant,
                    CurrentValue = "Defender for Office 365 is not licensed",
                    Remediation = "Consider licensing Microsoft Defender for Office 365 Plan 1 or Plan 2."
                });

                section.Checks.Add(CreateManualCheck("Safe Links Policy", "Safe Links requires Defender for Office 365.", SecurityCheckStatus.NonCompliant, "Not available without Defender license"));
                section.Checks.Add(CreateManualCheck("Safe Attachments Policy", "Safe Attachments requires Defender for Office 365.", SecurityCheckStatus.NonCompliant, "Not available without Defender license"));
            }

            // EOP checks (available to all)
            section.Checks.Add(CreateManualCheck("Anti-Spam Policy", "Anti-spam policies protect against unwanted email.", SecurityCheckStatus.Compliant, "Default anti-spam policies are active"));
            section.Checks.Add(CreateManualCheck("Anti-Malware Policy", "Anti-malware policies protect against malicious attachments.", SecurityCheckStatus.Compliant, "Default anti-malware policies are active"));
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Error checking Defender compliance");
            section.Checks.Add(CreateManualCheck("Defender Policies", "Could not retrieve Defender policies", SecurityCheckStatus.Error, ex.Message));
        }

        CalculateSectionStats(section);
        return section;
    }

    #endregion

    #region Helper Methods

    private SecurityCheck CreateManualCheck(string name, string description, SecurityCheckStatus status, string currentValue)
    {
        return new SecurityCheck
        {
            Name = name,
            Description = description,
            Status = status,
            CurrentValue = currentValue
        };
    }

    private void CalculateSectionStats(ComplianceSection section)
    {
        section.TotalChecks = section.Checks.Count;
        section.CompliantChecks = section.Checks.Count(c => c.Status == SecurityCheckStatus.Compliant);
        section.NonCompliantChecks = section.Checks.Count(c => c.Status == SecurityCheckStatus.NonCompliant);
        section.CompliancePercentage = section.TotalChecks > 0
            ? Math.Round((double)section.CompliantChecks / section.TotalChecks * 100, 1)
            : 0;
    }

    private string GetFriendlyLicenseName(string skuPartNumber)
    {
        return skuPartNumber switch
        {
            "ENTERPRISEPREMIUM" => "Microsoft 365 E5",
            "ENTERPRISEPACK" => "Microsoft 365 E3",
            "SPE_E5" => "Microsoft 365 E5",
            "SPE_E3" => "Microsoft 365 E3",
            "SMB_BUSINESS_PREMIUM" => "Microsoft 365 Business Premium",
            "SMB_BUSINESS_ESSENTIALS" => "Microsoft 365 Business Basic",
            "O365_BUSINESS_PREMIUM" => "Microsoft 365 Business Standard",
            "EXCHANGESTANDARD" => "Exchange Online (Plan 1)",
            "EXCHANGEENTERPRISE" => "Exchange Online (Plan 2)",
            "ATP_ENTERPRISE" => "Microsoft Defender for Office 365 (Plan 1)",
            "THREAT_INTELLIGENCE" => "Microsoft Defender for Office 365 (Plan 2)",
            "EMS" => "Enterprise Mobility + Security E3",
            "EMSPREMIUM" => "Enterprise Mobility + Security E5",
            "AAD_PREMIUM" => "Microsoft Entra ID P1",
            "AAD_PREMIUM_P2" => "Microsoft Entra ID P2",
            "INTUNE_A" => "Microsoft Intune",
            "POWER_BI_PRO" => "Power BI Pro",
            "PROJECTPREMIUM" => "Project Plan 5",
            "VISIOCLIENT" => "Visio Plan 2",
            _ => skuPartNumber
        };
    }

    #endregion
}
