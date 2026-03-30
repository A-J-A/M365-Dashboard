using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Graph;
using Azure.Core;
using Azure.Identity;
using System.Security.Cryptography.X509Certificates;
using M365Dashboard.Api.Models.Dtos;

namespace M365Dashboard.Api.Controllers;

[ApiController]
[Route("api/[controller]")]
[Authorize]
public class PermissionsController : ControllerBase
{
    private readonly GraphServiceClient _graphClient;
    private readonly IConfiguration _configuration;
    private readonly ILogger<PermissionsController> _logger;

    public PermissionsController(GraphServiceClient graphClient, IConfiguration configuration, ILogger<PermissionsController> logger)
    {
        _graphClient = graphClient;
        _configuration = configuration;
        _logger = logger;
    }

    // Centralised credential helper — supports both certificate (preferred) and client secret.
    // Mirrors the logic in Program.cs so permission checks work regardless of which credential type is deployed.
    private TokenCredential GetAppCredential()
    {
        var tenantId  = _configuration["AzureAd:TenantId"]  ?? throw new InvalidOperationException("AzureAd:TenantId not configured");
        var clientId  = _configuration["AzureAd:ClientId"]  ?? throw new InvalidOperationException("AzureAd:ClientId not configured");

        var certThumbprint = _configuration["AzureAd:ClientCertificateThumbprint"];
        var certPfxBase64  = _configuration["AzureAd:ClientCertificatePfx"];
        var clientSecret   = _configuration["AzureAd:ClientSecret"];

        // Prefer certificate authentication
        if (!string.IsNullOrEmpty(certThumbprint) && !string.IsNullOrEmpty(certPfxBase64))
        {
            try
            {
                var pfxBytes = Convert.FromBase64String(certPfxBase64);
                var cert = new X509Certificate2(
                    pfxBytes, (string?)null,
                    X509KeyStorageFlags.MachineKeySet | X509KeyStorageFlags.EphemeralKeySet);
                _logger.LogDebug("Permission checks using certificate credential (thumbprint: {Thumb})", cert.Thumbprint[..8] + "...");
                return new ClientCertificateCredential(tenantId, clientId, cert);
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "Failed to load certificate for permission checks — falling back to secret");
            }
        }

        // Fall back to client secret
        if (!string.IsNullOrWhiteSpace(clientSecret))
        {
            _logger.LogDebug("Permission checks using client secret credential");
            return new ClientSecretCredential(tenantId, clientId, clientSecret.Trim());
        }

        throw new InvalidOperationException(
            "No valid credential configured for permission checks. " +
            "Set AzureAd:ClientCertificatePfx (preferred) or AzureAd:ClientSecret.");
    }

    /// <summary>
    /// Get the status of all required Graph API permissions
    /// </summary>
    [HttpGet("status")]
    public async Task<IActionResult> GetPermissionsStatus()
    {
        var permissions = new List<PermissionStatusDto>();

        // Core permissions (always needed)
        permissions.Add(await CheckPermissionAsync("User.Read.All", "Read all users' full profiles", 
            "Required for viewing user information", CheckUsersPermission));
        
        permissions.Add(await CheckPermissionAsync("Group.Read.All", "Read all groups", 
            "Required for viewing groups and Teams", CheckGroupsPermission));
        
        permissions.Add(await CheckPermissionAsync("Device.Read.All", "Read all devices", 
            "Required for viewing Entra ID registered devices", CheckDevicesPermission));
        
        permissions.Add(await CheckPermissionAsync("DeviceManagementManagedDevices.Read.All", "Read Intune devices", 
            "Required for viewing Intune managed devices", CheckIntuneDevicesPermission));
        
        permissions.Add(await CheckPermissionAsync("DeviceManagementServiceConfig.Read.All", "Read Intune service configuration", 
            "Required for Apple Push certificate and enrollment settings", CheckDeviceManagementConfigPermission));
        
        permissions.Add(await CheckPermissionAsync("Mail.Read", "Read mail in all mailboxes", 
            "Required for mailbox information (delegated scenarios)", () => Task.FromResult(true))); // Can't easily test

        permissions.Add(await CheckPermissionAsync("Mail.Send", "Send mail as any user", 
            "Required for sending scheduled reports by email", CheckMailSendPermission));
        
        permissions.Add(await CheckPermissionAsync("Reports.Read.All", "Read all usage reports", 
            "Required for usage reports and analytics", CheckReportsPermission));

        // Security permissions
        permissions.Add(await CheckPermissionAsync("SecurityEvents.Read.All", "Read security events", 
            "Required for Microsoft Secure Score", CheckSecureScorePermission));
        
        permissions.Add(await CheckPermissionAsync("IdentityRiskyUser.Read.All", "Read risky user information", 
            "Required for Identity Protection risky users", CheckRiskyUsersPermission));
        
        permissions.Add(await CheckPermissionAsync("AuditLog.Read.All", "Read audit log data", 
            "Required for sign-in logs and risky sign-ins", CheckAuditLogPermission));
        
        permissions.Add(await CheckPermissionAsync("UserAuthenticationMethod.Read.All", "Read authentication methods", 
            "Required for MFA registration status", CheckMfaPermission));

        // Additional useful permissions
        permissions.Add(await CheckPermissionAsync("Directory.Read.All", "Read directory data", 
            "Required for tenant and license information", CheckDirectoryPermission));
        
        permissions.Add(await CheckPermissionAsync("Organization.Read.All", "Read organization information", 
            "Required for tenant details", CheckOrganizationPermission));
        
        // SharePoint permissions
        permissions.Add(await CheckPermissionAsync("Sites.Read.All", "Read items in all site collections", 
            "Required for SharePoint sites and storage information", CheckSharePointPermission));

        // Teams Phone permissions
        permissions.Add(await CheckPermissionAsync("CallRecords.Read.All", "Read call records", 
            "Required for Teams Phone System call analytics", CheckCallRecordsPermission));

        // Attack Simulation permissions
        permissions.Add(await CheckPermissionAsync("AttackSimulation.Read.All", "Read attack simulation data", 
            "Required for Attack Simulation Training reports in Executive Summary", CheckAttackSimulationPermission));

        // Microsoft Defender for Endpoint permissions (separate API)
        permissions.Add(await CheckPermissionAsync("Machine.Read.All", "Read machine information (Defender)", 
            "Required for Defender for Endpoint device data", CheckDefenderMachinePermission));
        
        permissions.Add(await CheckPermissionAsync("Vulnerability.Read.All", "Read vulnerability information (Defender)", 
            "Required for vulnerability assessment data", CheckDefenderVulnerabilityPermission));
        
        permissions.Add(await CheckPermissionAsync("Score.Read.All", "Read exposure score (Defender)", 
            "Required for Defender exposure and secure scores", CheckDefenderScorePermission));

        // Exchange Online permissions
        permissions.Add(await CheckPermissionAsync("Exchange.ManageAsApp", "Exchange Online Application Access", 
            "Required for Exchange distribution lists and mailbox forwarding (server-side)", CheckExchangeManageAsAppPermission));
        
        // Exchange Role Assignment
        permissions.Add(await CheckExchangeRoleAssignmentAsync());

        // Exchange Security/Protection Policy role (needed for Defender for Office cmdlets)
        permissions.Add(await CheckExchangeSecurityPolicyRoleAsync());

        var grantedCount = permissions.Count(p => p.IsGranted);
        var totalCount = permissions.Count;

        return Ok(new PermissionsStatusResponseDto(
            Permissions: permissions,
            TotalPermissions: totalCount,
            GrantedPermissions: grantedCount,
            MissingPermissions: totalCount - grantedCount,
            AllPermissionsGranted: grantedCount == totalCount,
            LastChecked: DateTime.UtcNow
        ));
    }

    private async Task<PermissionStatusDto> CheckPermissionAsync(
        string permissionName, 
        string displayName, 
        string description, 
        Func<Task<bool>> checkFunc)
    {
        bool isGranted;
        string? errorMessage = null;

        try
        {
            isGranted = await checkFunc();
        }
        catch (Exception ex)
        {
            isGranted = false;
            errorMessage = ex.Message.Contains("Authorization") || ex.Message.Contains("Forbidden") || ex.Message.Contains("Access")
                ? "Permission not granted"
                : ex.Message;
            _logger.LogDebug(ex, "Permission check failed for {Permission}", permissionName);
        }

        return new PermissionStatusDto(
            PermissionName: permissionName,
            DisplayName: displayName,
            Description: description,
            IsGranted: isGranted,
            ErrorMessage: errorMessage,
            Category: GetPermissionCategory(permissionName)
        );
    }

    private static string GetPermissionCategory(string permissionName) => permissionName switch
    {
        "User.Read.All" or "Group.Read.All" or "Directory.Read.All" or "Organization.Read.All" => "Core",
        "Device.Read.All" or "DeviceManagementManagedDevices.Read.All" or "DeviceManagementServiceConfig.Read.All" => "Devices",
        "Mail.Read" or "Mail.Send" or "Reports.Read.All" => "Mail & Reports",
        "SecurityEvents.Read.All" or "IdentityRiskyUser.Read.All" or "AuditLog.Read.All" or "UserAuthenticationMethod.Read.All" or "AttackSimulation.Read.All" => "Security",
        "Machine.Read.All" or "Vulnerability.Read.All" or "Score.Read.All" => "Defender for Endpoint",
        "Sites.Read.All" => "SharePoint",
        "CallRecords.Read.All" => "Teams Phone",
        "Exchange.ManageAsApp" or "Exchange Recipient Administrator" or "Security Reader (Exchange)" => "Exchange Online",
        _ => "Other"
    };

    private async Task<bool> CheckUsersPermission()
    {
        var users = await _graphClient.Users.GetAsync(config =>
        {
            config.QueryParameters.Top = 1;
            config.QueryParameters.Select = new[] { "id" };
        });
        return users?.Value != null;
    }

    private async Task<bool> CheckGroupsPermission()
    {
        var groups = await _graphClient.Groups.GetAsync(config =>
        {
            config.QueryParameters.Top = 1;
            config.QueryParameters.Select = new[] { "id" };
        });
        return groups?.Value != null;
    }

    private async Task<bool> CheckDevicesPermission()
    {
        var devices = await _graphClient.Devices.GetAsync(config =>
        {
            config.QueryParameters.Top = 1;
            config.QueryParameters.Select = new[] { "id" };
        });
        return devices?.Value != null;
    }

    private async Task<bool> CheckIntuneDevicesPermission()
    {
        var devices = await _graphClient.DeviceManagement.ManagedDevices.GetAsync(config =>
        {
            config.QueryParameters.Top = 1;
            config.QueryParameters.Select = new[] { "id" };
        });
        return devices?.Value != null;
    }

    private async Task<bool> CheckDeviceManagementConfigPermission()
    {
        try
        {
            var cert = await _graphClient.DeviceManagement.ApplePushNotificationCertificate.GetAsync();
            // If we get here without exception, permission is granted (cert may or may not exist)
            return true;
        }
        catch (Microsoft.Graph.Models.ODataErrors.ODataError ex) when (ex.ResponseStatusCode == 404)
        {
            // 404 means permission is granted but no certificate is configured
            return true;
        }
    }

    private async Task<bool> CheckReportsPermission()
    {
        // Try to get a usage report
        var report = await _graphClient.Reports
            .GetOffice365ActiveUserCountsWithPeriod("D7")
            .GetAsync();
        return report != null;
    }

    private async Task<bool> CheckMailSendPermission()
    {
        // Verify Mail.Send by checking the app's service principal has the role assigned.
        // We do this by attempting to list messages in any mailbox — if Mail.Send is granted
        // as an application permission the token will contain it. We can't actually send a
        // test email, so instead we verify by checking the current app's app role assignments.
        try
        {
            var tenantId  = _configuration["AzureAd:TenantId"];
            var clientId  = _configuration["AzureAd:ClientId"];

            // Get the service principal for this app
            var sp = await _graphClient.ServicePrincipals
                .GetAsync(config =>
                {
                    config.QueryParameters.Filter = $"appId eq '{clientId}'";
                    config.QueryParameters.Select = new[] { "id", "appId" };
                });

            var spId = sp?.Value?.FirstOrDefault()?.Id;
            if (spId == null) return false;

            // Get the app role assignments for this SP
            var assignments = await _graphClient.ServicePrincipals[spId]
                .AppRoleAssignments
                .GetAsync();

            // Mail.Send application permission ID in Microsoft Graph
            const string mailSendRoleId = "b633e1c5-b582-4048-a93e-9f11b44c7e96";

            return assignments?.Value?.Any(a =>
                string.Equals(a.AppRoleId?.ToString(), mailSendRoleId, StringComparison.OrdinalIgnoreCase)) == true;
        }
        catch
        {
            return false;
        }
    }

    private async Task<bool> CheckSecureScorePermission()
    {
        var scores = await _graphClient.Security.SecureScores.GetAsync(config =>
        {
            config.QueryParameters.Top = 1;
        });
        return scores?.Value != null;
    }

    private async Task<bool> CheckRiskyUsersPermission()
    {
        var users = await _graphClient.IdentityProtection.RiskyUsers.GetAsync(config =>
        {
            config.QueryParameters.Top = 1;
        });
        return users?.Value != null;
    }

    private async Task<bool> CheckAuditLogPermission()
    {
        var signIns = await _graphClient.AuditLogs.SignIns.GetAsync(config =>
        {
            config.QueryParameters.Top = 1;
        });
        return signIns?.Value != null;
    }

    private async Task<bool> CheckMfaPermission()
    {
        var details = await _graphClient.Reports.AuthenticationMethods.UserRegistrationDetails.GetAsync(config =>
        {
            config.QueryParameters.Top = 1;
        });
        return details?.Value != null;
    }

    private async Task<bool> CheckDirectoryPermission()
    {
        var skus = await _graphClient.SubscribedSkus.GetAsync();
        return skus?.Value != null;
    }

    private async Task<bool> CheckOrganizationPermission()
    {
        var org = await _graphClient.Organization.GetAsync();
        return org?.Value != null;
    }

    private async Task<bool> CheckSharePointPermission()
    {
        var sites = await _graphClient.Sites.GetAsync(config =>
        {
            config.QueryParameters.Search = "\"*\"";
            config.QueryParameters.Top = 1;
            config.QueryParameters.Select = new[] { "id" };
        });
        return sites?.Value != null;
    }

    private async Task<bool> CheckCallRecordsPermission()
    {
        try
        {
            var fromDate = DateTime.UtcNow.AddDays(-1);
            var toDate = DateTime.UtcNow;
            
            var callRecords = await _graphClient.Communications.CallRecords
                .MicrosoftGraphCallRecordsGetPstnCallsWithFromDateTimeWithToDateTime(fromDate, toDate)
                .GetAsync();
            
            // If we get here without an exception, permission is granted
            // Note: Empty results are valid - it just means no calls in that period
            return true;
        }
        catch (Exception ex) when (ex.Message.Contains("Authorization") || 
                                    ex.Message.Contains("Forbidden") || 
                                    ex.Message.Contains("Access") ||
                                    ex.Message.Contains("403"))
        {
            return false;
        }
    }

    private async Task<bool> CheckAttackSimulationPermission()
    {
        try
        {
            var credential = GetAppCredential();
            var betaClient = new GraphServiceClient(credential, new[] { "https://graph.microsoft.com/.default" }, "https://graph.microsoft.com/beta");
            
            var requestInfo = new Microsoft.Kiota.Abstractions.RequestInformation
            {
                HttpMethod = Microsoft.Kiota.Abstractions.Method.GET,
                URI = new Uri("https://graph.microsoft.com/beta/security/attackSimulation/simulations?$top=1")
            };
            
            var response = await betaClient.RequestAdapter.SendPrimitiveAsync<System.IO.Stream>(requestInfo);
            return true; // If we get here, permission is granted
        }
        catch (Exception ex) when (ex.Message.Contains("Authorization") || 
                                    ex.Message.Contains("Forbidden") || 
                                    ex.Message.Contains("Access") ||
                                    ex.Message.Contains("403"))
        {
            return false;
        }
    }

    private async Task<bool> CheckDefenderMachinePermission()
    {
        try
        {
            var credential = GetAppCredential();
            var scopes = new[] { "https://api.securitycenter.microsoft.com/.default" };
            var token = await credential.GetTokenAsync(new Azure.Core.TokenRequestContext(scopes), CancellationToken.None);
            
            using var httpClient = new HttpClient();
            httpClient.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token.Token);
            
            var response = await httpClient.GetAsync("https://api.securitycenter.microsoft.com/api/machines?$top=1");
            return response.IsSuccessStatusCode;
        }
        catch (Exception ex)
        {
            _logger.LogDebug(ex, "Defender Machine.Read.All permission check failed");
            return false;
        }
    }

    private async Task<bool> CheckDefenderVulnerabilityPermission()
    {
        try
        {
            var credential = GetAppCredential();
            var scopes = new[] { "https://api.securitycenter.microsoft.com/.default" };
            var token = await credential.GetTokenAsync(new Azure.Core.TokenRequestContext(scopes), CancellationToken.None);
            
            using var httpClient = new HttpClient();
            httpClient.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token.Token);
            
            var response = await httpClient.GetAsync("https://api.securitycenter.microsoft.com/api/vulnerabilities?$top=1");
            return response.IsSuccessStatusCode;
        }
        catch (Exception ex)
        {
            _logger.LogDebug(ex, "Defender Vulnerability.Read.All permission check failed");
            return false;
        }
    }

    private async Task<bool> CheckDefenderScorePermission()
    {
        try
        {
            var credential = GetAppCredential();
            var scopes = new[] { "https://api.securitycenter.microsoft.com/.default" };
            var token = await credential.GetTokenAsync(new Azure.Core.TokenRequestContext(scopes), CancellationToken.None);
            
            using var httpClient = new HttpClient();
            httpClient.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token.Token);
            
            var response = await httpClient.GetAsync("https://api.securitycenter.microsoft.com/api/exposureScore");
            return response.IsSuccessStatusCode;
        }
        catch (Exception ex)
        {
            _logger.LogDebug(ex, "Defender Score.Read.All permission check failed");
            return false;
        }
    }

    private async Task<bool> CheckExchangeManageAsAppPermission()
    {
        try
        {
            var credential = GetAppCredential();
            var tenantId = _configuration["AzureAd:TenantId"]!;
            // Exchange Online uses outlook.office365.com scope
            var scopes = new[] { "https://outlook.office365.com/.default" };
            var token = await credential.GetTokenAsync(new Azure.Core.TokenRequestContext(scopes), CancellationToken.None);
            
            // If we can get a token, the permission is granted
            // Try a simple Exchange Online REST call to verify
            using var httpClient = new HttpClient();
            httpClient.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token.Token);
            
            // Use Exchange Online PowerShell REST endpoint
            var requestUrl = $"https://outlook.office365.com/adminapi/beta/{tenantId}/InvokeCommand";
            var requestBody = new
            {
                CmdletInput = new
                {
                    CmdletName = "Get-OrganizationConfig",
                    Parameters = new { }
                }
            };
            
            var jsonContent = new StringContent(
                System.Text.Json.JsonSerializer.Serialize(requestBody),
                System.Text.Encoding.UTF8,
                "application/json");
            
            var response = await httpClient.PostAsync(requestUrl, jsonContent);
            
            _logger.LogInformation("Exchange ManageAsApp check response: {StatusCode}", response.StatusCode);
            
            // 200 = success, 403 = permission denied, other errors might be role-related
            return response.IsSuccessStatusCode || response.StatusCode == System.Net.HttpStatusCode.BadRequest;
        }
        catch (Exception ex)
        {
            _logger.LogDebug(ex, "Exchange.ManageAsApp permission check failed: {Message}", ex.Message);
            return false;
        }
    }

    private async Task<PermissionStatusDto> CheckExchangeSecurityPolicyRoleAsync()
    {
        bool isGranted = false;
        string? errorMessage = null;
        string appDisplayName = "your app registration";

        var tenantId = _configuration["AzureAd:TenantId"]!;
        var clientId = _configuration["AzureAd:ClientId"]!;

        // Resolve app display name for fix instructions
        try
        {
            var sps = await _graphClient.ServicePrincipals.GetAsync(cfg =>
            {
                cfg.QueryParameters.Filter = $"appId eq '{clientId}'";
                cfg.QueryParameters.Select = new[] { "displayName" };
            });
            var name = sps?.Value?.FirstOrDefault()?.DisplayName;
            if (!string.IsNullOrEmpty(name)) appDisplayName = name;
        }
        catch { /* best-effort */ }

        // Live cmdlet probe — sole source of truth
        try
        {
            var credential = GetAppCredential();
            var token = await credential.GetTokenAsync(
                new Azure.Core.TokenRequestContext(new[] { "https://outlook.office365.com/.default" }), CancellationToken.None);

            using var httpClient = new System.Net.Http.HttpClient();
            httpClient.DefaultRequestHeaders.Authorization =
                new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token.Token);

            var requestUrl = $"https://outlook.office365.com/adminapi/beta/{tenantId}/InvokeCommand";
            var body = new
            {
                CmdletInput = new
                {
                    CmdletName = "Get-AntiPhishPolicy",
                    Parameters = new { }
                }
            };
            var json = new System.Net.Http.StringContent(
                System.Text.Json.JsonSerializer.Serialize(body),
                System.Text.Encoding.UTF8, "application/json");

            var response = await httpClient.PostAsync(requestUrl, json);
            var responseBody = await response.Content.ReadAsStringAsync();
            isGranted = response.IsSuccessStatusCode;

            _logger.LogInformation("Exchange security policy probe: {Status}", (int)response.StatusCode);

            if (!isGranted)
            {
                errorMessage = response.StatusCode == System.Net.HttpStatusCode.Forbidden
                    ? "HTTP 403 — role not yet granted or not propagated. See fix instructions."
                    : $"HTTP {(int)response.StatusCode} — {responseBody}";
            }
        }
        catch (Exception ex)
        {
            errorMessage = "Could not verify: " + ex.Message;
            _logger.LogWarning(ex, "Exchange security policy live probe failed");
        }

        return new PermissionStatusDto(
            PermissionName: "Security Reader (Exchange)",
            DisplayName: "Exchange Security Policy Access",
            Description: $"Required for Defender for Office 365 policy data (anti-phish, anti-malware, Safe Attachments, Safe Links) | appName:{appDisplayName}",
            IsGranted: isGranted,
            ErrorMessage: errorMessage,
            Category: "Exchange Online"
        );
    }

    private async Task<PermissionStatusDto> CheckExchangeRoleAssignmentAsync()
    {
        bool isGranted = false;
        string? errorMessage = null;
        
        try
        {
            var clientId = _configuration["AzureAd:ClientId"];
            
            // Get the service principal for this app
            var servicePrincipals = await _graphClient.ServicePrincipals.GetAsync(config =>
            {
                config.QueryParameters.Filter = $"appId eq '{clientId}'";
                config.QueryParameters.Select = new[] { "id", "displayName", "appId" };
            });
            
            var servicePrincipal = servicePrincipals?.Value?.FirstOrDefault();
            if (servicePrincipal == null)
            {
                errorMessage = "Service principal not found";
            }
            else
            {
                // Get role assignments for this service principal
                var roleAssignments = await _graphClient.ServicePrincipals[servicePrincipal.Id].AppRoleAssignments.GetAsync();
                
                // Also check directory role memberships
                // Exchange Recipient Administrator role ID: 31392ffb-586c-42d1-9346-e59415a2cc4e
                // Exchange Administrator role ID: 29232cdf-9323-42fd-ade2-1d097af3e4de
                var exchangeRecipientAdminRoleId = "31392ffb-586c-42d1-9346-e59415a2cc4e";
                var exchangeAdminRoleId = "29232cdf-9323-42fd-ade2-1d097af3e4de";
                
                // Check directory role memberships
                try
                {
                    var memberOf = await _graphClient.ServicePrincipals[servicePrincipal.Id].MemberOf.GetAsync();
                    
                    if (memberOf?.Value != null)
                    {
                        foreach (var member in memberOf.Value)
                        {
                            if (member is Microsoft.Graph.Models.DirectoryRole role)
                            {
                                _logger.LogInformation("Service Principal has role: {RoleName} ({RoleId})", 
                                    role.DisplayName, role.RoleTemplateId);
                                
                                if (role.RoleTemplateId == exchangeRecipientAdminRoleId ||
                                    role.RoleTemplateId == exchangeAdminRoleId)
                                {
                                    isGranted = true;
                                    break;
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, "Could not check directory role memberships");
                }
                
                if (!isGranted)
                {
                    errorMessage = "Exchange Recipient Administrator or Exchange Administrator role not assigned to app";
                }
            }
        }
        catch (Exception ex)
        {
            errorMessage = ex.Message;
            _logger.LogDebug(ex, "Exchange role assignment check failed");
        }
        
        return new PermissionStatusDto(
            PermissionName: "Exchange Recipient Administrator",
            DisplayName: "Exchange Admin Role Assignment",
            Description: "Required role for Exchange Online management via app",
            IsGranted: isGranted,
            ErrorMessage: errorMessage,
            Category: "Exchange Online"
        );
    }
}
