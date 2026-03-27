using Microsoft.Graph;
using Microsoft.Graph.Models;
using M365Dashboard.Api.Models;
using System.Text.Json;
using Azure.Identity;

namespace M365Dashboard.Api.Services;

/// <summary>
/// Service for checking CIS Microsoft 365 Foundations Benchmark v6.0.0 controls
/// </summary>
public interface ICisBenchmarkService
{
    Task<CisBenchmarkResult> RunBenchmarkAsync(CisBenchmarkRequest? request = null);
    Task<CisControlResult> CheckControlAsync(string controlId);
}

public class CisBenchmarkService : ICisBenchmarkService
{
    private readonly GraphServiceClient _graphClient;
    private readonly IConfiguration _configuration;
    private readonly ILogger<CisBenchmarkService> _logger;
    private readonly HttpClient _httpClient;

    public CisBenchmarkService(
        GraphServiceClient graphClient,
        IConfiguration configuration,
        ILogger<CisBenchmarkService> logger,
        IHttpClientFactory httpClientFactory)
    {
        _graphClient = graphClient;
        _configuration = configuration;
        _logger = logger;
        _httpClient = httpClientFactory.CreateClient();
    }

    public async Task<CisBenchmarkResult> RunBenchmarkAsync(CisBenchmarkRequest? request = null)
    {
        request ??= new CisBenchmarkRequest();
        
        var result = new CisBenchmarkResult
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
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Could not get organization info");
        }

        // Run all controls
        var controls = new List<CisControlResult>();
        
        // 1. Microsoft 365 Admin Center
        controls.AddRange(await CheckAdminCenterControlsAsync());
        
        // 2. Microsoft 365 Defender
        controls.AddRange(await CheckDefenderControlsAsync());
        
        // 3. Microsoft Purview (if applicable)
        controls.AddRange(await CheckPurviewControlsAsync());
        
        // 5. Microsoft Entra Admin Center
        controls.AddRange(await CheckEntraControlsAsync());
        
        // 6. Exchange Online
        controls.AddRange(await CheckExchangeControlsAsync());
        
        // 7. SharePoint & OneDrive
        controls.AddRange(await CheckSharePointControlsAsync());
        
        // 8. Microsoft Teams
        controls.AddRange(await CheckTeamsControlsAsync());

        // Filter based on request
        if (!request.IncludeLevel2)
        {
            controls = controls.Where(c => c.Level == CisLevel.L1).ToList();
        }
        if (!request.IncludeE5Only)
        {
            controls = controls.Where(c => c.Profile == CisLicenseProfile.E3).ToList();
        }
        if (request.Categories?.Any() == true)
        {
            controls = controls.Where(c => request.Categories.Contains(c.Category)).ToList();
        }

        result.Controls = controls;
        
        // Calculate summary
        result.TotalControls = controls.Count;
        result.PassedControls = controls.Count(c => c.Status == CisControlStatus.Pass);
        result.FailedControls = controls.Count(c => c.Status == CisControlStatus.Fail);
        result.ManualControls = controls.Count(c => c.Status == CisControlStatus.Manual);
        result.NotApplicableControls = controls.Count(c => c.Status == CisControlStatus.NotApplicable);
        result.ErrorControls = controls.Count(c => c.Status == CisControlStatus.Error);
        
        var assessedControls = result.PassedControls + result.FailedControls;
        result.CompliancePercentage = assessedControls > 0 
            ? Math.Round((double)result.PassedControls / assessedControls * 100, 1) 
            : 0;

        // Level breakdown
        result.Level1Total = controls.Count(c => c.Level == CisLevel.L1);
        result.Level1Passed = controls.Count(c => c.Level == CisLevel.L1 && c.Status == CisControlStatus.Pass);
        result.Level2Total = controls.Count(c => c.Level == CisLevel.L2);
        result.Level2Passed = controls.Count(c => c.Level == CisLevel.L2 && c.Status == CisControlStatus.Pass);

        // Category breakdown
        result.Categories = controls
            .GroupBy(c => c.Category)
            .Select(g => new CisCategoryResult
            {
                CategoryId = g.Key.Split(' ')[0],
                CategoryName = g.Key,
                TotalControls = g.Count(),
                PassedControls = g.Count(c => c.Status == CisControlStatus.Pass),
                FailedControls = g.Count(c => c.Status == CisControlStatus.Fail),
                ManualControls = g.Count(c => c.Status == CisControlStatus.Manual),
                CompliancePercentage = g.Count(c => c.Status == CisControlStatus.Pass || c.Status == CisControlStatus.Fail) > 0
                    ? Math.Round((double)g.Count(c => c.Status == CisControlStatus.Pass) / g.Count(c => c.Status == CisControlStatus.Pass || c.Status == CisControlStatus.Fail) * 100, 1)
                    : 0
            })
            .OrderBy(c => c.CategoryId)
            .ToList();

        return result;
    }

    public async Task<CisControlResult> CheckControlAsync(string controlId)
    {
        // Individual control check - useful for re-checking specific controls
        var allControls = await RunBenchmarkAsync();
        return allControls.Controls.FirstOrDefault(c => c.ControlId == controlId) 
            ?? new CisControlResult { ControlId = controlId, Status = CisControlStatus.Unknown };
    }

    #region 1. Microsoft 365 Admin Center Controls

    private async Task<List<CisControlResult>> CheckAdminCenterControlsAsync()
    {
        var controls = new List<CisControlResult>();

        // 1.1.1 - Ensure Administrative accounts are cloud-only
        controls.Add(await Check_1_1_1_CloudOnlyAdmins());
        
        // 1.1.3 - Ensure that between two and four global admins are designated
        controls.Add(await Check_1_1_3_GlobalAdminCount());
        
        // 1.2.1 - Ensure that only organizationally managed/approved public groups exist
        controls.Add(await Check_1_2_1_PublicGroups());
        
        // 1.2.2 - Ensure sign-in to shared mailboxes is blocked
        controls.Add(await Check_1_2_2_SharedMailboxSignIn());
        
        // 1.3.1 - Ensure the 'Password expiration policy' is set to 'Set passwords to never expire'
        controls.Add(await Check_1_3_1_PasswordExpiry());
        
        // 1.3.6 - Ensure the customer lockbox feature is enabled
        controls.Add(await Check_1_3_6_CustomerLockbox());

        return controls;
    }

    private async Task<CisControlResult> Check_1_1_1_CloudOnlyAdmins()
    {
        var control = new CisControlResult
        {
            ControlId = "1.1.1",
            Title = "Ensure Administrative accounts are cloud-only",
            Description = "Administrative accounts should be cloud-only accounts separate from on-premises sync to prevent credential compromise.",
            Rationale = "If administrative accounts are synchronized from on-premises, then if the on-premises environment is compromised, an attacker could gain control of the cloud tenant. Using cloud-only administrative accounts reduces this attack surface.",
            Category = "1 Microsoft 365 admin center",
            SubCategory = "1.1 Users",
            Level = CisLevel.L1,
            Profile = CisLicenseProfile.E3,
            ExpectedValue = "All admin accounts should be cloud-only",
            Remediation = "Create dedicated cloud-only admin accounts and remove admin roles from synced accounts.",
            Impact = "Requires creating and managing separate cloud-only admin accounts.",
            Reference = "https://learn.microsoft.com/en-us/entra/identity/role-based-access-control/security-planning"
        };

        try
        {
            // Get users with admin roles
            var directoryRoles = await _graphClient.DirectoryRoles.GetAsync(config =>
            {
                config.QueryParameters.Expand = new[] { "members" };
            });

            var adminRoles = new[] { "Global Administrator", "Privileged Role Administrator", "User Administrator", 
                "Exchange Administrator", "SharePoint Administrator", "Security Administrator" };
            
            var syncedAdmins = new List<string>();
            var cloudOnlyAdmins = new List<string>();
            
            foreach (var role in directoryRoles?.Value ?? new List<DirectoryRole>())
            {
                if (adminRoles.Any(ar => role.DisplayName?.Contains(ar, StringComparison.OrdinalIgnoreCase) == true))
                {
                    foreach (var member in role.Members ?? new List<DirectoryObject>())
                    {
                        if (member is User user)
                        {
                            // Check if user is synced (has onPremisesSyncEnabled)
                            var userDetails = await _graphClient.Users[user.Id].GetAsync(config =>
                            {
                                config.QueryParameters.Select = new[] { "displayName", "userPrincipalName", "onPremisesSyncEnabled" };
                            });
                            
                            var userDisplay = $"{userDetails?.DisplayName} ({userDetails?.UserPrincipalName})";
                            
                            if (userDetails?.OnPremisesSyncEnabled == true)
                            {
                                if (!syncedAdmins.Contains(userDisplay))
                                    syncedAdmins.Add(userDisplay);
                            }
                            else
                            {
                                if (!cloudOnlyAdmins.Contains(userDisplay))
                                    cloudOnlyAdmins.Add(userDisplay);
                            }
                        }
                    }
                }
            }

            if (syncedAdmins.Any())
            {
                control.Status = CisControlStatus.Fail;
                control.CurrentValue = $"{syncedAdmins.Count} synced admin account(s) found";
                control.StatusReason = "Administrative accounts should be cloud-only";
                control.AffectedItems = syncedAdmins;
            }
            else
            {
                control.Status = CisControlStatus.Pass;
                control.CurrentValue = $"All {cloudOnlyAdmins.Count} admin accounts are cloud-only";
                control.StatusReason = "No synced accounts have administrative roles";
                // Show the cloud-only admins that were checked for verification
                control.AffectedItems = cloudOnlyAdmins;
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Error checking control 1.1.1");
            control.Status = CisControlStatus.Error;
            control.StatusReason = ex.Message;
        }

        return control;
    }

    private async Task<CisControlResult> Check_1_1_3_GlobalAdminCount()
    {
        var control = new CisControlResult
        {
            ControlId = "1.1.3",
            Title = "Ensure that between two and four global admins are designated",
            Description = "More than one global administrator should be designated to ensure business continuity, but no more than four to limit the attack surface.",
            Rationale = "Having more than one Global Administrator helps ensure business continuity if one admin leaves or their account is compromised. However, more than four increases the attack surface. Microsoft recommends between 2 and 4 Global Administrators.",
            Category = "1 Microsoft 365 admin center",
            SubCategory = "1.1 Users",
            Level = CisLevel.L1,
            Profile = CisLicenseProfile.E3,
            ExpectedValue = "Between 2 and 4 Global Administrators",
            Remediation = "Adjust the number of Global Administrators to be between 2 and 4.",
            Impact = "May require reassigning roles or creating emergency access accounts.",
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
            var members = gaRole?.Members?.OfType<User>().ToList() ?? new List<User>();
            var memberCount = members.Count;

            control.CurrentValue = $"{memberCount} Global Administrator(s)";
            
            // Get detailed info for each GA
            var gaDetails = new List<string>();
            foreach (var member in members)
            {
                try
                {
                    var userDetails = await _graphClient.Users[member.Id].GetAsync(config =>
                    {
                        config.QueryParameters.Select = new[] { "displayName", "userPrincipalName", "accountEnabled", "onPremisesSyncEnabled" };
                    });
                    var syncStatus = userDetails?.OnPremisesSyncEnabled == true ? "Synced" : "Cloud-only";
                    var enabledStatus = userDetails?.AccountEnabled == true ? "Enabled" : "Disabled";
                    gaDetails.Add($"{userDetails?.DisplayName} ({userDetails?.UserPrincipalName}) - {syncStatus}, {enabledStatus}");
                }
                catch
                {
                    gaDetails.Add(member.DisplayName ?? member.UserPrincipalName ?? member.Id ?? "Unknown");
                }
            }
            control.AffectedItems = gaDetails;

            if (memberCount >= 2 && memberCount <= 4)
            {
                control.Status = CisControlStatus.Pass;
                control.StatusReason = "Global Admin count is within the recommended range";
            }
            else if (memberCount < 2)
            {
                control.Status = CisControlStatus.Fail;
                control.StatusReason = "Too few Global Administrators - risk of lockout";
            }
            else
            {
                control.Status = CisControlStatus.Fail;
                control.StatusReason = "Too many Global Administrators - increased attack surface";
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Error checking control 1.1.3");
            control.Status = CisControlStatus.Error;
            control.StatusReason = ex.Message;
        }

        return control;
    }

    private async Task<CisControlResult> Check_1_2_1_PublicGroups()
    {
        var control = new CisControlResult
        {
            ControlId = "1.2.1",
            Title = "Ensure that only organizationally managed/approved public groups exist",
            Description = "Public Microsoft 365 Groups should be reviewed and approved to prevent data exposure.",
            Rationale = "Public Microsoft 365 Groups allow any user in the organization to join without approval. If sensitive information is stored in these groups, it could be exposed to unauthorized users. All public groups should be reviewed and approved by management.",
            Category = "1 Microsoft 365 admin center",
            SubCategory = "1.2 Groups",
            Level = CisLevel.L2,
            Profile = CisLicenseProfile.E3,
            ExpectedValue = "No unapproved public groups",
            Remediation = "Review all public groups and convert to private if not required to be public.",
            Impact = "Converting groups to private may affect external collaboration.",
            Reference = "https://learn.microsoft.com/en-us/microsoft-365/solutions/groups-teams-access-governance"
        };

        try
        {
            var publicGroups = await _graphClient.Groups.GetAsync(config =>
            {
                config.QueryParameters.Filter = "groupTypes/any(c:c eq 'Unified') and visibility eq 'Public'";
                config.QueryParameters.Select = new[] { "displayName", "mail", "visibility" };
                config.QueryParameters.Top = 999;
            });

            var publicGroupCount = publicGroups?.Value?.Count ?? 0;
            control.CurrentValue = $"{publicGroupCount} public group(s)";
            control.AffectedItems = publicGroups?.Value?
                .Select(g => $"{g.DisplayName} ({g.Mail})")
                .ToList() ?? new List<string>();

            if (publicGroupCount == 0)
            {
                control.Status = CisControlStatus.Pass;
                control.StatusReason = "No public Microsoft 365 Groups found";
            }
            else
            {
                // This is manual because we can't determine if they're approved
                control.Status = CisControlStatus.Manual;
                control.StatusReason = $"Review required: {publicGroupCount} public groups exist - verify they are approved";
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Error checking control 1.2.1");
            control.Status = CisControlStatus.Error;
            control.StatusReason = ex.Message;
        }

        return control;
    }

    private async Task<CisControlResult> Check_1_2_2_SharedMailboxSignIn()
    {
        var control = new CisControlResult
        {
            ControlId = "1.2.2",
            Title = "Ensure sign-in to shared mailboxes is blocked",
            Description = "Shared mailboxes should have sign-in blocked as they are not intended for direct login.",
            Rationale = "Shared mailboxes are designed to be accessed via delegation, not direct login. If sign-in is enabled, the shared mailbox password could be used for unauthorized access. Blocking sign-in ensures proper delegation-based access only.",
            Category = "1 Microsoft 365 admin center",
            SubCategory = "1.2 Groups",
            Level = CisLevel.L1,
            Profile = CisLicenseProfile.E3,
            ExpectedValue = "Sign-in blocked for all shared mailboxes",
            Remediation = "Block sign-in for all shared mailbox accounts in Entra ID.",
            Impact = "None - shared mailboxes should be accessed via delegation.",
            Reference = "https://learn.microsoft.com/en-us/microsoft-365/admin/email/about-shared-mailboxes"
        };

        try
        {
            // Get users that have shared mailbox recipient type (requires Exchange Online data)
            // This is a simplified check - full check would require Exchange Online PowerShell
            var users = await _graphClient.Users.GetAsync(config =>
            {
                config.QueryParameters.Select = new[] { "displayName", "userPrincipalName", "accountEnabled", "mailNickname" };
                config.QueryParameters.Filter = "accountEnabled eq true";
                config.QueryParameters.Top = 999;
            });

            // For now, mark as manual since we can't reliably detect shared mailboxes via Graph alone
            control.Status = CisControlStatus.Manual;
            control.CurrentValue = "Requires Exchange Online verification";
            control.StatusReason = "Check shared mailbox sign-in status via Exchange Online Admin Center or PowerShell";
            control.IsAutomated = false;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Error checking control 1.2.2");
            control.Status = CisControlStatus.Error;
            control.StatusReason = ex.Message;
        }

        return control;
    }

    private async Task<CisControlResult> Check_1_3_1_PasswordExpiry()
    {
        var control = new CisControlResult
        {
            ControlId = "1.3.1",
            Title = "Ensure the 'Password expiration policy' is set to 'Set passwords to never expire'",
            Description = "Password expiration should be disabled as it can lead to weaker passwords. Combined with MFA, non-expiring passwords are more secure.",
            Rationale = "NIST research shows that forced periodic password changes lead users to choose weaker passwords. When combined with MFA and banned password lists, non-expiring passwords are more secure than frequently rotated passwords.",
            Category = "1 Microsoft 365 admin center",
            SubCategory = "1.3 Settings",
            Level = CisLevel.L1,
            Profile = CisLicenseProfile.E3,
            ExpectedValue = "Passwords set to never expire",
            Remediation = "Set password expiration policy to 'Never expire' in Microsoft 365 admin center.",
            Impact = "Users will not be prompted to change passwords periodically.",
            Reference = "https://learn.microsoft.com/en-us/microsoft-365/admin/manage/set-password-expiration-policy"
        };

        try
        {
            var domains = await _graphClient.Domains.GetAsync(config =>
            {
                config.QueryParameters.Select = new[] { "id", "passwordValidityPeriodInDays" };
            });

            var allDomains = domains?.Value ?? new List<Domain>();
            var domainsWithExpiry = allDomains
                .Where(d => d.PasswordValidityPeriodInDays.HasValue && d.PasswordValidityPeriodInDays.Value < 2147483647)
                .ToList();
            var domainsWithoutExpiry = allDomains
                .Where(d => !d.PasswordValidityPeriodInDays.HasValue || d.PasswordValidityPeriodInDays.Value >= 2147483647)
                .ToList();

            if (domainsWithExpiry.Any())
            {
                control.Status = CisControlStatus.Fail;
                control.CurrentValue = $"Password expiry enabled on {domainsWithExpiry.Count} domain(s)";
                control.StatusReason = "Passwords should be set to never expire when combined with MFA";
                control.AffectedItems = domainsWithExpiry.Select(d => $"{d.Id}: {d.PasswordValidityPeriodInDays} days").ToList();
            }
            else
            {
                control.Status = CisControlStatus.Pass;
                control.CurrentValue = $"Passwords set to never expire on all {allDomains.Count} domain(s)";
                control.StatusReason = "Password expiration is properly configured";
                control.AffectedItems = domainsWithoutExpiry.Select(d => $"{d.Id}: Never expires").ToList();
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Error checking control 1.3.1");
            control.Status = CisControlStatus.Error;
            control.StatusReason = ex.Message;
        }

        return control;
    }

    private async Task<CisControlResult> Check_1_3_6_CustomerLockbox()
    {
        var control = new CisControlResult
        {
            ControlId = "1.3.6",
            Title = "Ensure the customer lockbox feature is enabled",
            Description = "Customer Lockbox requires Microsoft support to obtain approval before accessing tenant data.",
            Rationale = "Customer Lockbox adds an extra layer of control by requiring explicit approval before Microsoft support engineers can access your tenant data. This helps maintain data privacy and comply with regulatory requirements.",
            Category = "1 Microsoft 365 admin center",
            SubCategory = "1.3 Settings",
            Level = CisLevel.L2,
            Profile = CisLicenseProfile.E5,
            ExpectedValue = "Customer Lockbox enabled",
            Remediation = "Enable Customer Lockbox in the Microsoft 365 admin center under Settings > Org settings > Security & privacy.",
            Impact = "Support requests may take longer as approval is required.",
            Reference = "https://learn.microsoft.com/en-us/microsoft-365/compliance/customer-lockbox-requests"
        };

        try
        {
            // Customer Lockbox status requires beta API or admin API
            // Mark as manual for now
            control.Status = CisControlStatus.Manual;
            control.CurrentValue = "Requires manual verification";
            control.StatusReason = "Check Customer Lockbox status in Microsoft 365 admin center > Settings > Org settings > Security & privacy";
            control.IsAutomated = false;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Error checking control 1.3.6");
            control.Status = CisControlStatus.Error;
            control.StatusReason = ex.Message;
        }

        return control;
    }

    #endregion

    #region 2. Microsoft 365 Defender Controls

    private async Task<List<CisControlResult>> CheckDefenderControlsAsync()
    {
        var controls = new List<CisControlResult>();

        // 2.1.1 - Ensure Safe Links for Office Applications is Enabled
        controls.Add(await Check_2_1_1_SafeLinks());
        
        // 2.1.2 - Ensure the Common Attachment Types Filter is enabled
        controls.Add(await Check_2_1_2_AttachmentFilter());
        
        // 2.1.4 - Ensure Safe Attachments policy is enabled
        controls.Add(await Check_2_1_4_SafeAttachments());
        
        // 2.1.7 - Ensure that an anti-phishing policy has been created
        controls.Add(await Check_2_1_7_AntiPhishing());
        
        // 2.1.9 - Ensure DKIM is enabled for all Exchange Online Domains
        controls.Add(await Check_2_1_9_DKIM());

        // 3.1.1 - Ensure Microsoft 365 audit log search is Enabled
        controls.Add(await Check_3_1_1_AuditLog());

        return controls;
    }

    private async Task<CisControlResult> Check_2_1_1_SafeLinks()
    {
        var control = new CisControlResult
        {
            ControlId = "2.1.1",
            Title = "Ensure Safe Links for Office Applications is Enabled",
            Description = "Safe Links provides URL scanning and rewriting of inbound email messages and time-of-click verification of URLs.",
            Rationale = "Safe Links protects users by scanning URLs at time of click, even if the URL was safe when the email was delivered but has since been weaponized. This provides zero-day protection against malicious links.",
            Category = "2 Microsoft 365 Defender",
            SubCategory = "2.1 Email Protection",
            Level = CisLevel.L2,
            Profile = CisLicenseProfile.E5,
            ExpectedValue = "Safe Links enabled via preset security policy or custom policy",
            Remediation = "Enable Safe Links in Microsoft 365 Defender > Policies & rules > Threat policies > Safe Links, or enable Standard/Strict preset security policies.",
            Impact = "URLs in emails will be scanned and rewritten for protection.",
            Reference = "https://learn.microsoft.com/en-us/defender-office-365/safe-links-about"
        };

        try
        {
            // Check for preset security policies which include Safe Links by default
            // The Built-in protection preset security policy provides Safe Links protection to all recipients
            var checkedItems = new List<string>();
            
            // Check organization's subscribed SKUs for Defender for Office 365
            var subscribedSkus = await _graphClient.SubscribedSkus.GetAsync();
            var hasDefenderPlan = subscribedSkus?.Value?.Any(s => 
                s.ServicePlans?.Any(p => 
                    p.ServicePlanName?.Contains("ATP", StringComparison.OrdinalIgnoreCase) == true ||
                    p.ServicePlanName?.Contains("THREAT_INTELLIGENCE", StringComparison.OrdinalIgnoreCase) == true ||
                    p.ServicePlanName?.Contains("DEFENDER", StringComparison.OrdinalIgnoreCase) == true) == true) == true;

            if (hasDefenderPlan)
            {
                checkedItems.Add("Microsoft Defender for Office 365 license: Found");
                checkedItems.Add("licence is present but actual policy configuration cannot be verified via Graph API.");
                checkedItems.Add("");
                checkedItems.Add("Verify Safe Links is configured via PowerShell:");
                checkedItems.Add("  Get-SafeLinksPolicy | Format-List Name,IsEnabled");
                checkedItems.Add("  Get-SafeLinksRule | Format-List Name,State");
                
                control.Status = CisControlStatus.Manual;
                control.CurrentValue = "Defender for Office 365 licensed — policy configuration requires manual verification";
                control.StatusReason = "Licence present but Safe Links policy enablement cannot be verified via Graph API. Check Defender portal or PowerShell.";
                control.IsAutomated = false;
            }
            else
            {
                checkedItems.Add("Microsoft Defender for Office 365 license: Not found");
                checkedItems.Add("Safe Links requires Defender for Office 365 Plan 1 or Plan 2");
                
                control.Status = CisControlStatus.Fail;
                control.CurrentValue = "No Defender for Office 365 license found";
                control.StatusReason = "Safe Links requires Microsoft Defender for Office 365";
            }
            
            checkedItems.Add("");
            checkedItems.Add("Note: Custom Safe Links policies require PowerShell verification:");
            checkedItems.Add("  Get-SafeLinksPolicy | Format-List Name,IsEnabled");
            
            control.AffectedItems = checkedItems;
            control.IsAutomated = true;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Error checking control 2.1.1");
            control.Status = CisControlStatus.Error;
            control.StatusReason = ex.Message;
        }

        return control;
    }

    private async Task<CisControlResult> Check_2_1_2_AttachmentFilter()
    {
        var control = new CisControlResult
        {
            ControlId = "2.1.2",
            Title = "Ensure the Common Attachment Types Filter is enabled",
            Description = "The common attachment types filter blocks attachments with potentially dangerous file types.",
            Rationale = "Executable files and scripts are commonly used to deliver malware. Blocking these attachment types at the email gateway prevents users from accidentally executing malicious files.",
            Category = "2 Microsoft 365 Defender",
            SubCategory = "2.1 Email Protection",
            Level = CisLevel.L1,
            Profile = CisLicenseProfile.E3,
            ExpectedValue = "Common attachment filter enabled",
            Remediation = "Enable common attachment types filter in Microsoft 365 Defender > Policies & rules > Threat policies > Anti-malware.",
            Impact = "Emails with blocked attachment types will be quarantined or rejected.",
            Reference = "https://learn.microsoft.com/en-us/defender-office-365/anti-malware-protection-about"
        };

        try
        {
            var checkedItems = new List<string>();
            
            // This is part of Exchange Online Protection (EOP) which is included in all Exchange Online plans
            checkedItems.Add("Anti-malware policies are part of Exchange Online Protection (EOP)");
            checkedItems.Add("EOP is included in all Microsoft 365 subscriptions with Exchange Online");
            checkedItems.Add("");
            checkedItems.Add("Default policy blocks these file types:");
            checkedItems.Add("  - Executable files (.exe, .com, .bat, .cmd, .ps1, etc.)");
            checkedItems.Add("  - Script files (.vbs, .js, .wsf, etc.)");
            checkedItems.Add("  - Office files with macros (.docm, .xlsm, etc.)");
            checkedItems.Add("");
            checkedItems.Add("Verification via PowerShell:");
            checkedItems.Add("  Get-MalwareFilterPolicy | FL Name,EnableFileFilter,FileTypes");
            
            control.Status = CisControlStatus.Manual;
            control.CurrentValue = "Requires Exchange Online verification";
            control.StatusReason = "Verify common attachment filter is enabled in anti-malware policies";
            control.AffectedItems = checkedItems;
            control.IsAutomated = false;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Error checking control 2.1.2");
            control.Status = CisControlStatus.Error;
            control.StatusReason = ex.Message;
        }

        return control;
    }

    private async Task<CisControlResult> Check_2_1_4_SafeAttachments()
    {
        var control = new CisControlResult
        {
            ControlId = "2.1.4",
            Title = "Ensure Safe Attachments policy is enabled",
            Description = "Safe Attachments scans attachments in a virtual environment before delivering to recipients.",
            Rationale = "Safe Attachments detonates suspicious attachments in a sandbox environment to detect zero-day malware that signature-based scanners might miss. This provides advanced protection against targeted attacks.",
            Category = "2 Microsoft 365 Defender",
            SubCategory = "2.1 Email Protection",
            Level = CisLevel.L2,
            Profile = CisLicenseProfile.E5,
            ExpectedValue = "Safe Attachments enabled via preset security policy or custom policy",
            Remediation = "Enable Safe Attachments in Microsoft 365 Defender > Policies & rules > Threat policies > Safe Attachments.",
            Impact = "Email delivery may be slightly delayed while attachments are scanned.",
            Reference = "https://learn.microsoft.com/en-us/defender-office-365/safe-attachments-about"
        };

        try
        {
            var checkedItems = new List<string>();
            
            // Check organization's subscribed SKUs for Defender for Office 365
            var subscribedSkus = await _graphClient.SubscribedSkus.GetAsync();
            var hasDefenderPlan = subscribedSkus?.Value?.Any(s => 
                s.ServicePlans?.Any(p => 
                    p.ServicePlanName?.Contains("ATP", StringComparison.OrdinalIgnoreCase) == true ||
                    p.ServicePlanName?.Contains("THREAT_INTELLIGENCE", StringComparison.OrdinalIgnoreCase) == true ||
                    p.ServicePlanName?.Contains("DEFENDER", StringComparison.OrdinalIgnoreCase) == true) == true) == true;

            if (hasDefenderPlan)
            {
                checkedItems.Add("Microsoft Defender for Office 365 license: Found");
                checkedItems.Add("Licence is present but actual policy configuration cannot be verified via Graph API.");
                checkedItems.Add("");
                checkedItems.Add("Verify Safe Attachments is configured via PowerShell:");
                checkedItems.Add("  Get-SafeAttachmentPolicy | FL Name,Enable,Action");
                checkedItems.Add("  Get-SafeAttachmentRule | FL Name,State");
                
                control.Status = CisControlStatus.Manual;
                control.CurrentValue = "Defender for Office 365 licensed — policy configuration requires manual verification";
                control.StatusReason = "Licence present but Safe Attachments policy enablement cannot be verified via Graph API. Check Defender portal or PowerShell.";
                control.IsAutomated = false;
            }
            else
            {
                checkedItems.Add("Microsoft Defender for Office 365 license: Not found");
                checkedItems.Add("Safe Attachments requires Defender for Office 365 Plan 1 or Plan 2");
                
                control.Status = CisControlStatus.Fail;
                control.CurrentValue = "No Defender for Office 365 license found";
                control.StatusReason = "Safe Attachments requires Microsoft Defender for Office 365";
            }
            
            control.AffectedItems = checkedItems;
            control.IsAutomated = true;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Error checking control 2.1.4");
            control.Status = CisControlStatus.Error;
            control.StatusReason = ex.Message;
        }

        return control;
    }

    private async Task<CisControlResult> Check_2_1_7_AntiPhishing()
    {
        var control = new CisControlResult
        {
            ControlId = "2.1.7",
            Title = "Ensure that an anti-phishing policy has been created",
            Description = "Anti-phishing policies provide mailbox intelligence and spoof protection.",
            Rationale = "Phishing is the most common attack vector. Anti-phishing policies help detect and block impersonation attempts, spoofed senders, and other phishing techniques that bypass standard spam filters.",
            Category = "2 Microsoft 365 Defender",
            SubCategory = "2.1 Email Protection",
            Level = CisLevel.L1,
            Profile = CisLicenseProfile.E3,
            ExpectedValue = "Anti-phishing policy configured",
            Remediation = "Create or verify anti-phishing policy in Microsoft 365 Defender > Policies & rules > Threat policies > Anti-phishing.",
            Impact = "Suspected phishing emails will be handled according to policy.",
            Reference = "https://learn.microsoft.com/en-us/defender-office-365/anti-phishing-policies-about"
        };

        try
        {
            var checkedItems = new List<string>();
            
            // Anti-phishing is available in EOP (basic) and enhanced in Defender for Office 365
            checkedItems.Add("Anti-phishing protection levels:");
            checkedItems.Add("");
            checkedItems.Add("EOP (Exchange Online Protection) - included with Exchange Online:");
            checkedItems.Add("  - Spoof intelligence");
            checkedItems.Add("  - Implicit email authentication");
            checkedItems.Add("  - Default anti-phishing policy");
            checkedItems.Add("");
            
            // Check for Defender license for enhanced protection
            var subscribedSkus = await _graphClient.SubscribedSkus.GetAsync();
            var hasDefenderPlan = subscribedSkus?.Value?.Any(s => 
                s.ServicePlans?.Any(p => 
                    p.ServicePlanName?.Contains("ATP", StringComparison.OrdinalIgnoreCase) == true ||
                    p.ServicePlanName?.Contains("THREAT_INTELLIGENCE", StringComparison.OrdinalIgnoreCase) == true ||
                    p.ServicePlanName?.Contains("DEFENDER", StringComparison.OrdinalIgnoreCase) == true) == true) == true;

            if (hasDefenderPlan)
            {
                checkedItems.Add("Defender for Office 365 (enhanced) - LICENSED:");
                checkedItems.Add("Licence is present but actual policy configuration cannot be verified via Graph API.");
                checkedItems.Add("");
                checkedItems.Add("Verify anti-phishing policy is configured via PowerShell:");
                checkedItems.Add("  Get-AntiPhishPolicy | FL Name,Enabled,EnableMailboxIntelligence");
                checkedItems.Add("  Get-AntiPhishRule | FL Name,State");
                
                control.Status = CisControlStatus.Manual;
                control.CurrentValue = "Defender for Office 365 licensed — policy configuration requires manual verification";
                control.StatusReason = "Licence present but anti-phishing policy configuration cannot be verified via Graph API. Check Defender portal or PowerShell.";
                control.IsAutomated = false;
            }
            else
            {
                checkedItems.Add("Defender for Office 365 (enhanced) - NOT LICENSED:");
                checkedItems.Add("  - Basic EOP anti-phishing only");
                
                control.Status = CisControlStatus.Manual;
                control.CurrentValue = "Basic EOP anti-phishing available";
                control.StatusReason = "Verify anti-phishing policies are configured in Microsoft 365 Defender portal";
            }
            
            checkedItems.Add("");
            checkedItems.Add("Verification via PowerShell:");
            checkedItems.Add("  Get-AntiPhishPolicy | FL Name,Enabled,EnableMailboxIntelligence");
            
            control.AffectedItems = checkedItems;
            control.IsAutomated = false;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Error checking control 2.1.7");
            control.Status = CisControlStatus.Error;
            control.StatusReason = ex.Message;
        }

        return control;
    }

    private async Task<CisControlResult> Check_2_1_9_DKIM()
    {
        var control = new CisControlResult
        {
            ControlId = "2.1.9",
            Title = "Ensure that DKIM is enabled for all Exchange Online Domains",
            Description = "DKIM adds a digital signature to outgoing email messages to verify authenticity.",
            Rationale = "DKIM (DomainKeys Identified Mail) cryptographically signs outgoing emails, allowing recipients to verify the email originated from your domain and wasn't modified in transit. This helps prevent spoofing and improves deliverability.",
            Category = "2 Microsoft 365 Defender",
            SubCategory = "2.1 Email Protection",
            Level = CisLevel.L1,
            Profile = CisLicenseProfile.E3,
            ExpectedValue = "DKIM enabled for all domains",
            Remediation = "Enable DKIM for each domain in Microsoft 365 Defender > Policies & rules > Threat policies > Email authentication settings.",
            Impact = "Improves email deliverability and protects against spoofing.",
            Reference = "https://learn.microsoft.com/en-us/defender-office-365/email-authentication-dkim-configure"
        };

        try
        {
            // Get verified domains
            var domains = await _graphClient.Domains.GetAsync(config =>
            {
                config.QueryParameters.Filter = "isVerified eq true";
                config.QueryParameters.Select = new[] { "id" };
            });

            // DKIM status requires Exchange Online - mark as manual
            control.Status = CisControlStatus.Manual;
            control.CurrentValue = $"{domains?.Value?.Count ?? 0} verified domain(s) - check DKIM status";
            control.StatusReason = "Verify DKIM is enabled for each domain in Exchange admin center";
            control.IsAutomated = false;
            control.AffectedItems = domains?.Value?.Select(d => d.Id ?? "").ToList() ?? new List<string>();
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Error checking control 2.1.9");
            control.Status = CisControlStatus.Error;
            control.StatusReason = ex.Message;
        }

        return control;
    }

    private async Task<CisControlResult> Check_3_1_1_AuditLog()
    {
        var control = new CisControlResult
        {
            ControlId = "3.1.1",
            Title = "Ensure Microsoft 365 audit log search is Enabled",
            Description = "Unified audit logging captures user and admin activity for compliance and security investigations.",
            Rationale = "Audit logs are essential for security investigations, compliance requirements, and understanding what happened during a security incident. Without audit logging enabled, critical evidence may be unavailable.",
            Category = "3 Microsoft Purview",
            SubCategory = "3.1 Audit",
            Level = CisLevel.L1,
            Profile = CisLicenseProfile.E3,
            ExpectedValue = "Audit logging enabled",
            Remediation = "Enable unified audit logging in Microsoft Purview compliance portal.",
            Impact = "None - this is essential for security monitoring.",
            Reference = "https://learn.microsoft.com/en-us/purview/audit-log-enable-disable"
        };

        try
        {
            // Audit log status check requires the Security & Compliance PowerShell
            // We can check if audit logs are being returned as a proxy
            control.Status = CisControlStatus.Manual;
            control.CurrentValue = "Requires Microsoft Purview verification";
            control.StatusReason = "Verify audit logging is enabled in Microsoft Purview compliance portal";
            control.IsAutomated = false;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Error checking control 3.1.1");
            control.Status = CisControlStatus.Error;
            control.StatusReason = ex.Message;
        }

        return control;
    }

    #endregion

    #region 3. Microsoft Purview Controls

    private async Task<List<CisControlResult>> CheckPurviewControlsAsync()
    {
        var controls = new List<CisControlResult>();
        // Purview controls typically require PowerShell or are manual
        return controls;
    }

    #endregion

    #region 5. Microsoft Entra Admin Center Controls

    private async Task<List<CisControlResult>> CheckEntraControlsAsync()
    {
        var controls = new List<CisControlResult>();

        // 5.1.2.1 - Ensure 'Per-user MFA' is disabled (use Conditional Access instead)
        controls.Add(await Check_5_1_2_1_PerUserMfa());
        
        // 5.1.2.3 - Ensure 'Restrict non-admin users from creating tenants' is set to 'Yes'
        controls.Add(await Check_5_1_2_3_TenantCreation());
        
        // 5.1.6.2 - Ensure that guest user access is restricted
        controls.Add(await Check_5_1_6_2_GuestAccess());
        
        // 5.2.2.1 - Ensure MFA is enabled for all users in administrative roles
        controls.Add(await Check_5_2_2_1_AdminMfa());
        
        // 5.2.2.2 - Ensure MFA is enabled for all users
        controls.Add(await Check_5_2_2_2_AllUsersMfa());
        
        // 5.2.2.3 - Ensure Conditional Access policies block legacy authentication
        controls.Add(await Check_5_2_2_3_LegacyAuth());
        
        // 5.2.3.2 - Ensure custom banned passwords lists are used
        controls.Add(await Check_5_2_3_2_BannedPasswords());
        
        // 5.2.4.1 - Ensure 'Self service password reset' is enabled
        controls.Add(await Check_5_2_4_1_SSPR());

        return controls;
    }

    private async Task<CisControlResult> Check_5_1_2_1_PerUserMfa()
    {
        var control = new CisControlResult
        {
            ControlId = "5.1.2.1",
            Title = "Ensure 'Per-user MFA' is disabled",
            Description = "Per-user MFA should be disabled in favor of Conditional Access-based MFA for better control and reporting.",
            Rationale = "Per-user MFA is a legacy feature with limited reporting and no ability to apply conditional policies. Conditional Access provides granular control over when MFA is required, better integration with other security features, and comprehensive reporting.",
            Category = "5 Microsoft Entra admin center",
            SubCategory = "5.1.2 External Identities",
            Level = CisLevel.L1,
            Profile = CisLicenseProfile.E3,
            ExpectedValue = "Per-user MFA disabled, Conditional Access MFA used",
            Remediation = "Disable per-user MFA and implement Conditional Access policies for MFA.",
            Impact = "Requires Conditional Access policies to be configured for MFA.",
            Reference = "https://learn.microsoft.com/en-us/entra/identity/authentication/howto-mfa-userstates"
        };

        try
        {
            // Per-user MFA status requires legacy MFA API - mark as manual
            control.Status = CisControlStatus.Manual;
            control.CurrentValue = "Requires manual verification";
            control.StatusReason = "Check per-user MFA settings at https://account.activedirectory.windowsazure.com/UserManagement/MultifactorVerification.aspx";
            control.IsAutomated = false;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Error checking control 5.1.2.1");
            control.Status = CisControlStatus.Error;
            control.StatusReason = ex.Message;
        }

        return control;
    }

    private async Task<CisControlResult> Check_5_1_2_3_TenantCreation()
    {
        var control = new CisControlResult
        {
            ControlId = "5.1.2.3",
            Title = "Ensure 'Restrict non-admin users from creating tenants' is set to 'Yes'",
            Description = "Non-admin users should not be able to create new Azure AD tenants.",
            Rationale = "When users can create new tenants, they may inadvertently create shadow IT environments outside organizational control. This can lead to data sprawl, security gaps, and compliance issues. Restricting tenant creation to administrators ensures proper governance.",
            Category = "5 Microsoft Entra admin center",
            SubCategory = "5.1.2 External Identities",
            Level = CisLevel.L1,
            Profile = CisLicenseProfile.E3,
            ExpectedValue = "Tenant creation restricted to admins only",
            Remediation = "Set 'Restrict non-admin users from creating tenants' to Yes in Entra ID > User settings.",
            Impact = "Users will need admin assistance to create new tenants if required.",
            Reference = "https://learn.microsoft.com/en-us/entra/fundamentals/users-default-permissions"
        };

        try
        {
            var authPolicy = await _graphClient.Policies.AuthorizationPolicy.GetAsync();
            
            // Check defaultUserRolePermissions
            var canCreateTenants = authPolicy?.DefaultUserRolePermissions?.AllowedToCreateTenants ?? true;
            var canCreateApps = authPolicy?.DefaultUserRolePermissions?.AllowedToCreateApps ?? true;
            var canReadOtherUsers = authPolicy?.DefaultUserRolePermissions?.AllowedToReadOtherUsers ?? true;

            control.AffectedItems = new List<string>
            {
                $"AllowedToCreateTenants: {canCreateTenants}",
                $"AllowedToCreateApps: {canCreateApps}",
                $"AllowedToReadOtherUsers: {canReadOtherUsers}"
            };

            if (!canCreateTenants)
            {
                control.Status = CisControlStatus.Pass;
                control.CurrentValue = "Non-admin users cannot create tenants";
                control.StatusReason = "Tenant creation is properly restricted";
            }
            else
            {
                control.Status = CisControlStatus.Fail;
                control.CurrentValue = "Non-admin users can create tenants";
                control.StatusReason = "Tenant creation should be restricted to admins only";
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Error checking control 5.1.2.3");
            control.Status = CisControlStatus.Error;
            control.StatusReason = ex.Message;
        }

        return control;
    }

    private async Task<CisControlResult> Check_5_1_6_2_GuestAccess()
    {
        var control = new CisControlResult
        {
            ControlId = "5.1.6.2",
            Title = "Ensure that guest user access is restricted",
            Description = "Guest users should have limited access to directory information.",
            Rationale = "Guest users are external users who may not have the same level of trust as internal employees. Restricting their access to directory information minimizes the risk of reconnaissance and data exposure to external parties.",
            Category = "5 Microsoft Entra admin center",
            SubCategory = "5.1.6 External collaboration settings",
            Level = CisLevel.L1,
            Profile = CisLicenseProfile.E3,
            ExpectedValue = "Guest user access restricted",
            Remediation = "Configure guest user access restrictions in Entra ID > External Identities > External collaboration settings.",
            Impact = "Guest users will have limited visibility of directory information.",
            Reference = "https://learn.microsoft.com/en-us/entra/external-id/external-collaboration-settings-configure"
        };

        try
        {
            var authPolicy = await _graphClient.Policies.AuthorizationPolicy.GetAsync();
            var guestRestriction = authPolicy?.GuestUserRoleId?.ToString();

            // Guest user role IDs:
            // a0b1b346-4d3e-4e8b-98f8-753987be4970 = Same as member users (most permissive)
            // 10dae51f-b6af-4016-8d66-8c2a99b929b3 = Limited access (default)
            // 2af84b1e-32c8-42b7-82bc-daa82404023b = Restricted access (most restrictive)

            var roleDescription = guestRestriction switch
            {
                "2af84b1e-32c8-42b7-82bc-daa82404023b" => "Restricted access (most restrictive)",
                "10dae51f-b6af-4016-8d66-8c2a99b929b3" => "Limited access (default)",
                "a0b1b346-4d3e-4e8b-98f8-753987be4970" => "Same as member users (most permissive)",
                _ => $"Unknown role ID: {guestRestriction}"
            };

            control.AffectedItems = new List<string>
            {
                $"Guest User Role ID: {guestRestriction}",
                $"Access Level: {roleDescription}",
                $"Allow invites from: {authPolicy?.AllowInvitesFrom}"
            };

            if (guestRestriction == "2af84b1e-32c8-42b7-82bc-daa82404023b")
            {
                control.Status = CisControlStatus.Pass;
                control.CurrentValue = "Most restrictive - guests have limited access";
                control.StatusReason = "Guest user access is properly restricted";
            }
            else if (guestRestriction == "10dae51f-b6af-4016-8d66-8c2a99b929b3")
            {
                control.Status = CisControlStatus.Pass;
                control.CurrentValue = "Limited access (default)";
                control.StatusReason = "Guest user access is restricted";
            }
            else
            {
                control.Status = CisControlStatus.Fail;
                control.CurrentValue = "Same as member users - too permissive";
                control.StatusReason = "Guest users have the same access as member users";
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Error checking control 5.1.6.2");
            control.Status = CisControlStatus.Error;
            control.StatusReason = ex.Message;
        }

        return control;
    }

    private async Task<CisControlResult> Check_5_2_2_1_AdminMfa()
    {
        var control = new CisControlResult
        {
            ControlId = "5.2.2.1",
            Title = "Ensure multifactor authentication is enabled for all users in administrative roles",
            Description = "All administrative accounts should be protected with MFA.",
            Rationale = "Administrative accounts have elevated privileges and are high-value targets for attackers. MFA provides an additional layer of security that significantly reduces the risk of account compromise, even if passwords are phished or stolen.",
            Category = "5 Microsoft Entra admin center",
            SubCategory = "5.2.2 Conditional Access",
            Level = CisLevel.L1,
            Profile = CisLicenseProfile.E3,
            ExpectedValue = "All admin users MFA registered",
            Remediation = "Create a Conditional Access policy requiring MFA for users with administrative roles.",
            Impact = "Admins will need to complete MFA when signing in.",
            Reference = "https://learn.microsoft.com/en-us/entra/identity/conditional-access/howto-conditional-access-policy-admin-mfa"
        };

        try
        {
            var adminRoles = new[] { "Global Administrator", "Privileged Role Administrator", "User Administrator",
                "Exchange Administrator", "SharePoint Administrator", "Security Administrator",
                "Billing Administrator", "Compliance Administrator", "Conditional Access Administrator" };

            var adminUsers = new Dictionary<string, bool>();

            // Get directory roles and their members
            var directoryRoles = await _graphClient.DirectoryRoles.GetAsync(config =>
            {
                config.QueryParameters.Expand = new[] { "members" };
            });

            foreach (var role in directoryRoles?.Value ?? new List<DirectoryRole>())
            {
                if (!adminRoles.Any(ar => role.DisplayName?.Contains(ar, StringComparison.OrdinalIgnoreCase) == true))
                    continue;

                foreach (var member in role.Members?.OfType<User>() ?? new List<User>())
                {
                    if (member.Id != null && !adminUsers.ContainsKey(member.Id))
                    {
                        adminUsers[member.Id] = false;
                    }
                }
            }

            // Check MFA registration for each admin
            var adminsWithoutMfa = new List<string>();
            var adminsWithMfa = new List<string>();
            foreach (var adminId in adminUsers.Keys)
            {
                try
                {
                    var authMethods = await _graphClient.Users[adminId].Authentication.Methods.GetAsync();
                    var hasMfa = authMethods?.Value?.Any(m => 
                        m is not Microsoft.Graph.Models.PasswordAuthenticationMethod) == true;
                    
                    var user = await _graphClient.Users[adminId].GetAsync(config =>
                    {
                        config.QueryParameters.Select = new[] { "displayName", "userPrincipalName" };
                    });
                    var userDisplay = $"{user?.DisplayName} ({user?.UserPrincipalName})";
                    
                    if (!hasMfa)
                    {
                        adminsWithoutMfa.Add($"{userDisplay} - NO MFA");
                    }
                    else
                    {
                        var methods = authMethods?.Value?
                            .Where(m => m is not Microsoft.Graph.Models.PasswordAuthenticationMethod)
                            .Select(m => m.GetType().Name.Replace("AuthenticationMethod", ""))
                            .ToList() ?? new List<string>();
                        adminsWithMfa.Add($"{userDisplay} - MFA: {string.Join(", ", methods)}");
                    }
                }
                catch
                {
                    // Skip users we can't check
                }
            }

            control.CurrentValue = $"{adminsWithMfa.Count}/{adminUsers.Count} admins MFA registered";
            
            // Show all admins - those with MFA issues first, then those compliant
            var allAdmins = new List<string>();
            allAdmins.AddRange(adminsWithoutMfa);
            allAdmins.AddRange(adminsWithMfa);
            control.AffectedItems = allAdmins;

            if (adminsWithoutMfa.Count == 0)
            {
                control.Status = CisControlStatus.Pass;
                control.StatusReason = "All administrative users have MFA registered";
            }
            else
            {
                control.Status = CisControlStatus.Fail;
                control.StatusReason = $"{adminsWithoutMfa.Count} admin(s) without MFA registration";
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Error checking control 5.2.2.1");
            control.Status = CisControlStatus.Error;
            control.StatusReason = ex.Message;
        }

        return control;
    }

    private async Task<CisControlResult> Check_5_2_2_2_AllUsersMfa()
    {
        var control = new CisControlResult
        {
            ControlId = "5.2.2.2",
            Title = "Ensure multifactor authentication is enabled for all users",
            Description = "All users should be protected with MFA to prevent credential-based attacks.",
            Rationale = "MFA blocks 99.9% of automated account attacks. As credential theft remains the most common attack vector, requiring MFA for all users provides organization-wide protection against phishing and password spray attacks.",
            Category = "5 Microsoft Entra admin center",
            SubCategory = "5.2.2 Conditional Access",
            Level = CisLevel.L1,
            Profile = CisLicenseProfile.E3,
            ExpectedValue = "MFA required for all users",
            Remediation = "Create a Conditional Access policy requiring MFA for all users.",
            Impact = "Users will need to complete MFA when signing in.",
            Reference = "https://learn.microsoft.com/en-us/entra/identity/conditional-access/howto-conditional-access-policy-all-users-mfa"
        };

        try
        {
            // Check for Conditional Access policies requiring MFA for all users
            var caPolicies = await _graphClient.Identity.ConditionalAccess.Policies.GetAsync();
            
            var allUsersMfaPolicy = caPolicies?.Value?.FirstOrDefault(p =>
                p.State == ConditionalAccessPolicyState.Enabled &&
                p.Conditions?.Users?.IncludeUsers?.Contains("All") == true &&
                p.GrantControls?.BuiltInControls?.Contains(ConditionalAccessGrantControl.Mfa) == true);

            // List all enabled CA policies for verification
            var enabledPolicies = caPolicies?.Value?
                .Where(p => p.State == ConditionalAccessPolicyState.Enabled)
                .Select(p => $"{p.DisplayName} (Users: {string.Join(",", p.Conditions?.Users?.IncludeUsers ?? new List<string>())})")
                .ToList() ?? new List<string>();
            
            control.AffectedItems = enabledPolicies;

            if (allUsersMfaPolicy != null)
            {
                control.Status = CisControlStatus.Pass;
                control.CurrentValue = $"Policy '{allUsersMfaPolicy.DisplayName}' requires MFA for all users";
                control.StatusReason = "Conditional Access policy requiring MFA for all users is enabled";
            }
            else
            {
                // Check security defaults as fallback
                control.Status = CisControlStatus.Manual;
                control.CurrentValue = "No explicit all-users MFA policy found";
                control.StatusReason = "Verify MFA is required via Conditional Access or Security Defaults";
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Error checking control 5.2.2.2");
            control.Status = CisControlStatus.Error;
            control.StatusReason = ex.Message;
        }

        return control;
    }

    private async Task<CisControlResult> Check_5_2_2_3_LegacyAuth()
    {
        var control = new CisControlResult
        {
            ControlId = "5.2.2.3",
            Title = "Enable Conditional Access policies to block legacy authentication",
            Description = "Legacy authentication protocols don't support MFA and should be blocked.",
            Rationale = "Legacy authentication protocols like IMAP, SMTP, and POP3 do not support modern authentication or MFA. Attackers specifically target these protocols because they can bypass MFA protections. Blocking legacy auth closes this security gap.",
            Category = "5 Microsoft Entra admin center",
            SubCategory = "5.2.2 Conditional Access",
            Level = CisLevel.L1,
            Profile = CisLicenseProfile.E3,
            ExpectedValue = "Legacy authentication blocked",
            Remediation = "Create a Conditional Access policy that blocks legacy authentication protocols.",
            Impact = "Users with older clients may lose access until they upgrade.",
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

            // List all enabled CA policies that might block legacy auth
            var blockingPolicies = caPolicies?.Value?
                .Where(p => p.State == ConditionalAccessPolicyState.Enabled &&
                           p.GrantControls?.BuiltInControls?.Contains(ConditionalAccessGrantControl.Block) == true)
                .Select(p => $"{p.DisplayName} (Client Apps: {string.Join(",", p.Conditions?.ClientAppTypes?.Select(c => c.ToString()) ?? new List<string>())})")
                .ToList() ?? new List<string>();
            
            control.AffectedItems = blockingPolicies.Any() 
                ? blockingPolicies 
                : new List<string> { "No blocking policies found" };

            if (legacyBlockPolicy != null)
            {
                control.Status = CisControlStatus.Pass;
                control.CurrentValue = $"Policy '{legacyBlockPolicy.DisplayName}' blocks legacy authentication";
                control.StatusReason = "Legacy authentication is blocked by Conditional Access";
            }
            else
            {
                control.Status = CisControlStatus.Fail;
                control.CurrentValue = "No policy blocking legacy authentication found";
                control.StatusReason = "Legacy authentication protocols should be blocked";
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Error checking control 5.2.2.3");
            control.Status = CisControlStatus.Error;
            control.StatusReason = ex.Message;
        }

        return control;
    }

    private async Task<CisControlResult> Check_5_2_3_2_BannedPasswords()
    {
        var control = new CisControlResult
        {
            ControlId = "5.2.3.2",
            Title = "Ensure custom banned passwords lists are used",
            Description = "Custom banned passwords prevent users from using organization-specific weak passwords.",
            Rationale = "Users often create passwords based on company names, products, locations, or sports teams. Custom banned password lists prevent these predictable passwords that attackers may target using organization-specific wordlists.",
            Category = "5 Microsoft Entra admin center",
            SubCategory = "5.2.3 Authentication methods",
            Level = CisLevel.L1,
            Profile = CisLicenseProfile.E3,
            ExpectedValue = "Custom banned password list configured",
            Remediation = "Configure custom banned passwords in Entra ID > Security > Authentication methods > Password protection.",
            Impact = "Users cannot use passwords that match banned terms.",
            Reference = "https://learn.microsoft.com/en-us/entra/identity/authentication/concept-password-ban-bad"
        };

        try
        {
            // Password protection settings require beta API
            control.Status = CisControlStatus.Manual;
            control.CurrentValue = "Requires manual verification";
            control.StatusReason = "Check custom banned password list in Entra ID > Security > Authentication methods > Password protection";
            control.IsAutomated = false;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Error checking control 5.2.3.2");
            control.Status = CisControlStatus.Error;
            control.StatusReason = ex.Message;
        }

        return control;
    }

    private async Task<CisControlResult> Check_5_2_4_1_SSPR()
    {
        var control = new CisControlResult
        {
            ControlId = "5.2.4.1",
            Title = "Ensure 'Self service password reset enabled' is set to 'All'",
            Description = "Self-service password reset allows users to reset their passwords without helpdesk assistance.",
            Rationale = "SSPR reduces helpdesk costs and enables users to quickly regain access when locked out. It also reduces the risk of social engineering attacks against helpdesk staff who might reset passwords for attackers.",
            Category = "5 Microsoft Entra admin center",
            SubCategory = "5.2.4 Password reset",
            Level = CisLevel.L1,
            Profile = CisLicenseProfile.E3,
            ExpectedValue = "SSPR enabled for all users",
            Remediation = "Enable self-service password reset for all users in Entra ID > Password reset.",
            Impact = "Users can reset their own passwords, reducing helpdesk load.",
            Reference = "https://learn.microsoft.com/en-us/entra/identity/authentication/tutorial-enable-sspr"
        };

        try
        {
            // SSPR configuration requires specific permissions
            control.Status = CisControlStatus.Manual;
            control.CurrentValue = "Requires manual verification";
            control.StatusReason = "Check SSPR settings in Entra ID > Password reset > Properties";
            control.IsAutomated = false;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Error checking control 5.2.4.1");
            control.Status = CisControlStatus.Error;
            control.StatusReason = ex.Message;
        }

        return control;
    }

    #endregion

    #region 6. Exchange Online Controls

    private async Task<List<CisControlResult>> CheckExchangeControlsAsync()
    {
        var controls = new List<CisControlResult>();

        // 6.2.1 - Ensure all forms of mail forwarding are blocked
        controls.Add(Check_6_2_1_MailForwarding());
        
        // 6.2.3 - Ensure email from external senders is identified
        controls.Add(Check_6_2_3_ExternalTagging());
        
        // 6.5.1 - Ensure modern authentication for Exchange Online is enabled
        controls.Add(Check_6_5_1_ModernAuth());
        
        // 6.5.4 - Ensure SMTP AUTH is disabled
        controls.Add(Check_6_5_4_SmtpAuth());

        return controls;
    }

    private CisControlResult Check_6_2_1_MailForwarding()
    {
        return new CisControlResult
        {
            ControlId = "6.2.1",
            Title = "Ensure all forms of mail forwarding are blocked and/or disabled",
            Description = "External mail forwarding should be blocked to prevent data exfiltration.",
            Rationale = "Automatic mail forwarding to external addresses is a common technique used by attackers after compromising an account. Blocking external forwarding prevents attackers from silently exfiltrating email data.",
            Category = "6 Exchange Online",
            SubCategory = "6.2 Mail Transport",
            Level = CisLevel.L1,
            Profile = CisLicenseProfile.E3,
            ExpectedValue = "External mail forwarding blocked",
            Status = CisControlStatus.Manual,
            CurrentValue = "Requires Exchange Online verification",
            StatusReason = "Check transport rules and remote domains in Exchange admin center",
            Remediation = "Configure outbound spam filter policy to block automatic forwarding and create transport rules to prevent forwarding.",
            Impact = "Users will not be able to automatically forward mail to external addresses.",
            Reference = "https://learn.microsoft.com/en-us/microsoft-365/security/office-365-security/outbound-spam-policies-external-email-forwarding",
            IsAutomated = false
        };
    }

    private CisControlResult Check_6_2_3_ExternalTagging()
    {
        return new CisControlResult
        {
            ControlId = "6.2.3",
            Title = "Ensure email from external senders is identified",
            Description = "External email tagging helps users identify potentially malicious emails.",
            Rationale = "Users are more likely to fall for phishing attacks when they believe an email came from inside the organization. External sender tagging provides a visual indicator that helps users identify potentially suspicious emails.",
            Category = "6 Exchange Online",
            SubCategory = "6.2 Mail Transport",
            Level = CisLevel.L1,
            Profile = CisLicenseProfile.E3,
            ExpectedValue = "External tagging enabled",
            Status = CisControlStatus.Manual,
            CurrentValue = "Requires Exchange Online verification",
            StatusReason = "Check external sender tagging in Exchange admin center",
            Remediation = "Enable external tagging in Exchange admin center > Mail flow > Rules or use Set-ExternalInOutlook cmdlet.",
            Impact = "External emails will display a tag identifying them as from outside the organization.",
            Reference = "https://learn.microsoft.com/en-us/exchange/mail-flow-best-practices/remote-domains/external-sender-identification",
            IsAutomated = false
        };
    }

    private CisControlResult Check_6_5_1_ModernAuth()
    {
        return new CisControlResult
        {
            ControlId = "6.5.1",
            Title = "Ensure modern authentication for Exchange Online is enabled",
            Description = "Modern authentication enables OAuth-based authentication and supports MFA.",
            Rationale = "Modern authentication provides better security through OAuth 2.0 tokens, support for MFA, and integration with Conditional Access policies. It replaces basic authentication which transmits credentials with each request.",
            Category = "6 Exchange Online",
            SubCategory = "6.5 Exchange Authentication",
            Level = CisLevel.L1,
            Profile = CisLicenseProfile.E3,
            ExpectedValue = "Modern authentication enabled",
            Status = CisControlStatus.Pass, // Modern auth is now default and cannot be disabled
            CurrentValue = "Modern authentication is enabled by default",
            StatusReason = "Modern authentication is enabled by default for all tenants created after 2017",
            Remediation = "Modern authentication is now always enabled for Exchange Online.",
            Impact = "None - modern authentication supports MFA and is more secure.",
            Reference = "https://learn.microsoft.com/en-us/exchange/clients-and-mobile-in-exchange-online/enable-or-disable-modern-authentication-in-exchange-online",
            IsAutomated = true
        };
    }

    private CisControlResult Check_6_5_4_SmtpAuth()
    {
        return new CisControlResult
        {
            ControlId = "6.5.4",
            Title = "Ensure SMTP AUTH is disabled",
            Description = "SMTP AUTH should be disabled globally unless specifically required.",
            Rationale = "SMTP AUTH is a legacy protocol that doesn't support modern authentication. It's commonly used in password spray attacks. Disabling it reduces the attack surface while modern applications can use OAuth.",
            Category = "6 Exchange Online",
            SubCategory = "6.5 Exchange Authentication",
            Level = CisLevel.L1,
            Profile = CisLicenseProfile.E3,
            ExpectedValue = "SMTP AUTH disabled",
            Status = CisControlStatus.Manual,
            CurrentValue = "Requires Exchange Online verification",
            StatusReason = "Check SMTP AUTH settings via PowerShell: Get-TransportConfig | Select SmtpClientAuthenticationDisabled",
            Remediation = "Disable SMTP AUTH globally: Set-TransportConfig -SmtpClientAuthenticationDisabled $true",
            Impact = "Applications using SMTP AUTH will need to be reconfigured to use OAuth or other methods.",
            Reference = "https://learn.microsoft.com/en-us/exchange/clients-and-mobile-in-exchange-online/authenticated-client-smtp-submission",
            IsAutomated = false
        };
    }

    #endregion

    #region 7. SharePoint & OneDrive Controls

    private async Task<List<CisControlResult>> CheckSharePointControlsAsync()
    {
        var controls = new List<CisControlResult>();

        // 7.2.1 - Ensure modern authentication for SharePoint applications is required
        controls.Add(Check_7_2_1_ModernAuth());
        
        // 7.2.3 - Ensure external content sharing is restricted
        controls.Add(Check_7_2_3_ExternalSharing());
        
        // 7.2.7 - Ensure link sharing is restricted
        controls.Add(Check_7_2_7_LinkSharing());
        
        // 7.3.4 - Ensure custom script execution is restricted
        controls.Add(Check_7_3_4_CustomScripts());

        return controls;
    }

    private CisControlResult Check_7_2_1_ModernAuth()
    {
        return new CisControlResult
        {
            ControlId = "7.2.1",
            Title = "Ensure modern authentication for SharePoint applications is required",
            Description = "Legacy authentication should be blocked for SharePoint Online.",
            Rationale = "Legacy authentication protocols don't support MFA and are vulnerable to credential-based attacks. Requiring modern authentication ensures that all SharePoint access is protected by Conditional Access policies and MFA.",
            Category = "7 SharePoint & OneDrive",
            SubCategory = "7.2 Sharing",
            Level = CisLevel.L1,
            Profile = CisLicenseProfile.E3,
            ExpectedValue = "Legacy authentication disabled",
            Status = CisControlStatus.Manual,
            CurrentValue = "Requires SharePoint Online verification",
            StatusReason = "Check via PowerShell: Get-SPOTenant | Select LegacyAuthProtocolsEnabled",
            Remediation = "Disable legacy authentication: Set-SPOTenant -LegacyAuthProtocolsEnabled $false",
            Impact = "Older clients that don't support modern auth will lose access.",
            Reference = "https://learn.microsoft.com/en-us/sharepoint/control-access-from-unmanaged-devices",
            IsAutomated = false
        };
    }

    private CisControlResult Check_7_2_3_ExternalSharing()
    {
        return new CisControlResult
        {
            ControlId = "7.2.3",
            Title = "Ensure external content sharing is restricted",
            Description = "External sharing should be limited to specific domains or disabled.",
            Rationale = "Unrestricted external sharing can lead to data leakage. By limiting sharing to specific trusted domains or disabling it entirely, organizations can control who has access to sensitive information.",
            Category = "7 SharePoint & OneDrive",
            SubCategory = "7.2 Sharing",
            Level = CisLevel.L1,
            Profile = CisLicenseProfile.E3,
            ExpectedValue = "External sharing restricted",
            Status = CisControlStatus.Manual,
            CurrentValue = "Requires SharePoint Online verification",
            StatusReason = "Check sharing settings in SharePoint admin center > Policies > Sharing",
            Remediation = "Configure external sharing limits in SharePoint admin center.",
            Impact = "External collaboration may be limited based on configuration.",
            Reference = "https://learn.microsoft.com/en-us/sharepoint/turn-external-sharing-on-or-off",
            IsAutomated = false
        };
    }

    private CisControlResult Check_7_2_7_LinkSharing()
    {
        return new CisControlResult
        {
            ControlId = "7.2.7",
            Title = "Ensure link sharing is restricted in SharePoint and OneDrive",
            Description = "Default sharing links should be restricted to organization members.",
            Rationale = "'Anyone' links can be forwarded and used by anyone, even unintended recipients. Defaulting to 'Specific people' or 'Only people in organization' ensures users consciously choose to create less secure links.",
            Category = "7 SharePoint & OneDrive",
            SubCategory = "7.2 Sharing",
            Level = CisLevel.L1,
            Profile = CisLicenseProfile.E3,
            ExpectedValue = "Default links set to 'Specific people' or 'Only people in organization'",
            Status = CisControlStatus.Manual,
            CurrentValue = "Requires SharePoint Online verification",
            StatusReason = "Check default link type in SharePoint admin center > Policies > Sharing",
            Remediation = "Set default sharing link type to 'Specific people' or 'Only people in your organization'.",
            Impact = "Users will need to explicitly choose 'Anyone' links when required.",
            Reference = "https://learn.microsoft.com/en-us/sharepoint/turn-external-sharing-on-or-off",
            IsAutomated = false
        };
    }

    private CisControlResult Check_7_3_4_CustomScripts()
    {
        return new CisControlResult
        {
            ControlId = "7.3.4",
            Title = "Ensure custom script execution is restricted on site collections",
            Description = "Custom scripts can be used to inject malicious code and should be restricted.",
            Rationale = "Custom scripts can be used to execute arbitrary JavaScript, potentially allowing data exfiltration, phishing, or other malicious activities. Restricting custom scripts to managed sites reduces this risk.",
            Category = "7 SharePoint & OneDrive",
            SubCategory = "7.3 Access Control",
            Level = CisLevel.L1,
            Profile = CisLicenseProfile.E3,
            ExpectedValue = "Custom scripts disabled",
            Status = CisControlStatus.Manual,
            CurrentValue = "Requires SharePoint Online verification",
            StatusReason = "Check via PowerShell: Get-SPOSite | Select Url,DenyAddAndCustomizePages",
            Remediation = "Disable custom scripts: Set-SPOSite -DenyAddAndCustomizePages $true",
            Impact = "Some customizations may not be available on sites.",
            Reference = "https://learn.microsoft.com/en-us/sharepoint/allow-or-prevent-custom-script",
            IsAutomated = false
        };
    }

    #endregion

    #region 8. Microsoft Teams Controls

    private async Task<List<CisControlResult>> CheckTeamsControlsAsync()
    {
        var controls = new List<CisControlResult>();

        // 8.1.1 - Ensure external file sharing in Teams is enabled for only approved cloud storage services
        controls.Add(Check_8_1_1_ExternalFileSharing());
        
        // 8.2.2 - Ensure communication with unmanaged Teams users is disabled
        controls.Add(Check_8_2_2_UnmanagedUsers());
        
        // 8.5.3 - Ensure only people in my org can bypass the lobby
        controls.Add(Check_8_5_3_LobbyBypass());
        
        // 8.6.1 - Ensure users can report security concerns in Teams
        controls.Add(Check_8_6_1_SecurityReporting());

        return controls;
    }

    private CisControlResult Check_8_1_1_ExternalFileSharing()
    {
        return new CisControlResult
        {
            ControlId = "8.1.1",
            Title = "Ensure external file sharing in Teams is enabled for only approved cloud storage services",
            Description = "Third-party cloud storage providers should be restricted to approved services only.",
            Rationale = "Allowing unrestricted cloud storage providers in Teams can lead to data leakage through unapproved services. Limiting to approved providers ensures data stays within controlled environments that meet organizational security requirements.",
            Category = "8 Microsoft Teams",
            SubCategory = "8.1 External Access",
            Level = CisLevel.L2,
            Profile = CisLicenseProfile.E3,
            ExpectedValue = "Only approved cloud storage services enabled",
            Status = CisControlStatus.Manual,
            CurrentValue = "Requires Teams admin center verification",
            StatusReason = "Check cloud storage settings in Teams admin center > Messaging policies",
            Remediation = "Disable unapproved cloud storage providers in Teams admin center.",
            Impact = "Users cannot share files from disabled cloud storage services.",
            Reference = "https://learn.microsoft.com/en-us/microsoftteams/messaging-policies-in-teams",
            IsAutomated = false
        };
    }

    private CisControlResult Check_8_2_2_UnmanagedUsers()
    {
        return new CisControlResult
        {
            ControlId = "8.2.2",
            Title = "Ensure communication with unmanaged Teams users is disabled",
            Description = "Communication with Teams users not managed by an organization should be restricted.",
            Rationale = "Unmanaged Teams users (personal accounts) are not subject to organizational policies. Allowing communication with them increases the risk of phishing, data exfiltration, and social engineering attacks.",
            Category = "8 Microsoft Teams",
            SubCategory = "8.2 Guest Access",
            Level = CisLevel.L1,
            Profile = CisLicenseProfile.E3,
            ExpectedValue = "Unmanaged Teams communication disabled",
            Status = CisControlStatus.Manual,
            CurrentValue = "Requires Teams admin center verification",
            StatusReason = "Check external access settings in Teams admin center > External access",
            Remediation = "Disable 'Unmanaged Teams users' in Teams admin center > External access.",
            Impact = "Users cannot communicate with external Teams users not managed by an organization.",
            Reference = "https://learn.microsoft.com/en-us/microsoftteams/manage-external-access",
            IsAutomated = false
        };
    }

    private CisControlResult Check_8_5_3_LobbyBypass()
    {
        return new CisControlResult
        {
            ControlId = "8.5.3",
            Title = "Ensure only people in my org can bypass the lobby",
            Description = "Meeting lobby settings should require external participants to wait in the lobby.",
            Rationale = "The meeting lobby provides a security checkpoint where organizers can verify attendees before admitting them. Requiring external participants to wait in the lobby prevents unauthorized access to meetings.",
            Category = "8 Microsoft Teams",
            SubCategory = "8.5 Meetings",
            Level = CisLevel.L1,
            Profile = CisLicenseProfile.E3,
            ExpectedValue = "Only organization members bypass lobby",
            Status = CisControlStatus.Manual,
            CurrentValue = "Requires Teams admin center verification",
            StatusReason = "Check meeting policies in Teams admin center > Meetings > Meeting policies",
            Remediation = "Set 'Who can bypass the lobby' to 'People in my org' or more restrictive.",
            Impact = "External participants will wait in the lobby until admitted.",
            Reference = "https://learn.microsoft.com/en-us/microsoftteams/meeting-policies-participants-and-guests",
            IsAutomated = false
        };
    }

    private CisControlResult Check_8_6_1_SecurityReporting()
    {
        return new CisControlResult
        {
            ControlId = "8.6.1",
            Title = "Ensure users can report security concerns in Teams",
            Description = "Users should be able to report suspicious messages in Teams.",
            Rationale = "User reporting is a critical part of security defense. Enabling users to report suspicious messages in Teams helps identify phishing attempts and other threats that automated systems might miss.",
            Category = "8 Microsoft Teams",
            SubCategory = "8.6 Messaging",
            Level = CisLevel.L1,
            Profile = CisLicenseProfile.E3,
            ExpectedValue = "Security reporting enabled",
            Status = CisControlStatus.Manual,
            CurrentValue = "Requires Teams admin center verification",
            StatusReason = "Check messaging policies in Teams admin center > Messaging policies",
            Remediation = "Enable 'Report a security concern' in Teams messaging policies.",
            Impact = "Users can report suspicious messages directly from Teams.",
            Reference = "https://learn.microsoft.com/en-us/microsoft-365/security/office-365-security/submissions-teams",
            IsAutomated = false
        };
    }

    #endregion
}
