using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using Azure.Identity;

namespace M365Dashboard.Api.Services;

// ── DTOs ─────────────────────────────────────────────────────────────────────

public record DefenderOverviewDto(
    AntiPhishResultDto AntiPhish,
    AntiMalwareResultDto AntiMalware,
    AntiSpamResultDto AntiSpam,
    OutboundSpamResultDto OutboundSpam,
    SafeAttachmentsResultDto SafeAttachments,
    SafeLinksResultDto SafeLinks,
    DateTime LastUpdated
);

// Anti-phish
public record AntiPhishPolicyDto(
    string Identity,
    string Name,
    bool Enabled,
    bool IsDefault,
    bool EnableSpoofIntelligence,
    bool EnableUnauthenticatedSender,
    bool EnableViaTag,
    string PhishThresholdLevel,
    string SpoofQuarantineTag,
    bool EnableFirstContactSafetyTips,
    bool EnableSimilarUsersSafetyTips,
    bool EnableSimilarDomainsSafetyTips,
    bool EnableUnusualCharactersSafetyTips,
    bool EnableMailboxIntelligence,
    bool EnableMailboxIntelligenceProtection,
    string MailboxIntelligenceProtectionAction,
    bool EnableOrganizationDomainsProtection,
    bool EnableTargetedDomainsProtection,
    bool EnableTargetedUserProtection,
    string TargetedUserProtectionAction,
    string TargetedDomainProtectionAction,
    List<string> TargetedUsersToProtect,
    List<string> TargetedDomainsToProtect,
    DateTime? WhenCreated,
    DateTime? WhenChanged
);

public record AntiPhishResultDto(
    List<AntiPhishPolicyDto> Policies,
    int TotalCount,
    string? Error
);

// Anti-malware
public record AntiMalwarePolicyDto(
    string Identity,
    string Name,
    bool Enabled,
    bool IsDefault,
    string Action,
    bool EnableFileFilter,
    List<string> FileTypes,
    bool ZapEnabled,
    bool EnableInternalSenderAdminNotifications,
    string? InternalSenderAdminAddress,
    bool EnableExternalSenderAdminNotifications,
    string? ExternalSenderAdminAddress,
    DateTime? WhenCreated,
    DateTime? WhenChanged
);

public record AntiMalwareResultDto(
    List<AntiMalwarePolicyDto> Policies,
    int TotalCount,
    string? Error
);

// Anti-spam (inbound)
public record AntiSpamPolicyDto(
    string Identity,
    string Name,
    bool Enabled,
    bool IsDefault,
    string SpamAction,
    string HighConfidenceSpamAction,
    string PhishSpamAction,
    string HighConfidencePhishAction,
    string BulkSpamAction,
    int BulkThreshold,
    bool ZapEnabled,
    bool SpamZapEnabled,
    bool PhishZapEnabled,
    bool EnableEndUserSpamNotifications,
    int EndUserSpamNotificationFrequency,
    bool AllowedSenderDomainsPresent,
    bool BlockedSenderDomainsPresent,
    string QuarantineTag,
    DateTime? WhenCreated,
    DateTime? WhenChanged
);

public record AntiSpamResultDto(
    List<AntiSpamPolicyDto> Policies,
    int TotalCount,
    string? Error
);

// Outbound spam
public record OutboundSpamPolicyDto(
    string Identity,
    string Name,
    bool Enabled,
    bool IsDefault,
    string ActionWhenThresholdReached,
    int RecipientLimitExternalPerHour,
    int RecipientLimitInternalPerHour,
    int RecipientLimitPerDay,
    bool AutoForwardingMode,
    string AutoForwardingModeValue,
    DateTime? WhenCreated,
    DateTime? WhenChanged
);

public record OutboundSpamResultDto(
    List<OutboundSpamPolicyDto> Policies,
    int TotalCount,
    string? Error
);

// Safe Attachments
public record SafeAttachmentsPolicyDto(
    string Identity,
    string Name,
    bool Enabled,
    bool IsDefault,
    string Action,
    bool ActionOnError,
    bool Redirect,
    string? RedirectAddress,
    bool QuarantineTag,
    DateTime? WhenCreated,
    DateTime? WhenChanged
);

public record SafeAttachmentsResultDto(
    List<SafeAttachmentsPolicyDto> Policies,
    int TotalCount,
    bool IsLicensed,
    string? Error
);

// Safe Links
public record SafeLinksPolicyDto(
    string Identity,
    string Name,
    bool Enabled,
    bool IsDefault,
    bool AllowClickThrough,
    bool DisableUrlRewrite,
    bool EnableForInternalSenders,
    bool EnableOrganizationBranding,
    bool EnableSafeLinksForEmail,
    bool EnableSafeLinksForTeams,
    bool EnableSafeLinksForOffice,
    bool TrackClicks,
    bool ScanUrls,
    List<string> DoNotRewriteUrls,
    DateTime? WhenCreated,
    DateTime? WhenChanged
);

public record SafeLinksResultDto(
    List<SafeLinksPolicyDto> Policies,
    int TotalCount,
    bool IsLicensed,
    string? Error
);

// ── Interface ─────────────────────────────────────────────────────────────────

public interface IDefenderForOfficeService
{
    Task<DefenderOverviewDto> GetOverviewAsync();
    Task<AntiPhishResultDto> GetAntiPhishPoliciesAsync();
    Task<AntiMalwareResultDto> GetAntiMalwarePoliciesAsync();
    Task<AntiSpamResultDto> GetAntiSpamPoliciesAsync();
    Task<OutboundSpamResultDto> GetOutboundSpamPoliciesAsync();
    Task<SafeAttachmentsResultDto> GetSafeAttachmentsPoliciesAsync();
    Task<SafeLinksResultDto> GetSafeLinksPoliciesAsync();
}

// ── Implementation ────────────────────────────────────────────────────────────

public class DefenderForOfficeService : IDefenderForOfficeService
{
    private readonly IConfiguration _configuration;
    private readonly ILogger<DefenderForOfficeService> _logger;
    private readonly IHttpClientFactory _httpClientFactory;

    public DefenderForOfficeService(
        IConfiguration configuration,
        ILogger<DefenderForOfficeService> logger,
        IHttpClientFactory httpClientFactory)
    {
        _configuration = configuration;
        _logger = logger;
        _httpClientFactory = httpClientFactory;
    }

    private async Task<string> GetExchangeTokenAsync()
    {
        var tenantId = _configuration["AzureAd:TenantId"];
        var clientId = _configuration["AzureAd:ClientId"];
        var clientSecret = _configuration["AzureAd:ClientSecret"];

        var credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
        var token = await credential.GetTokenAsync(
            new Azure.Core.TokenRequestContext(new[] { "https://outlook.office365.com/.default" }));
        return token.Token;
    }

    private async Task<JsonDocument?> InvokeExchangeCommandAsync(string cmdletName, Dictionary<string, object>? parameters = null)
    {
        var tenantId = _configuration["AzureAd:TenantId"];
        var token = await GetExchangeTokenAsync();

        using var httpClient = _httpClientFactory.CreateClient();
        httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);

        var requestUrl = $"https://outlook.office365.com/adminapi/beta/{tenantId}/InvokeCommand";
        var requestBody = new
        {
            CmdletInput = new
            {
                CmdletName = cmdletName,
                Parameters = parameters ?? new Dictionary<string, object>()
            }
        };

        var jsonContent = new StringContent(
            JsonSerializer.Serialize(requestBody),
            Encoding.UTF8,
            "application/json");

        _logger.LogInformation("Defender: Invoking {Cmdlet}", cmdletName);

        var response = await httpClient.PostAsync(requestUrl, jsonContent);
        var responseContent = await response.Content.ReadAsStringAsync();

        if (!response.IsSuccessStatusCode)
        {
            _logger.LogError("Defender Exchange API error: {StatusCode} - {Content}", response.StatusCode, responseContent);
            throw new Exception($"Exchange API returned {response.StatusCode}: {responseContent}");
        }

        if (string.IsNullOrWhiteSpace(responseContent)) return null;
        return JsonDocument.Parse(responseContent);
    }

    // ── Public methods ────────────────────────────────────────────────────────

    public async Task<DefenderOverviewDto> GetOverviewAsync()
    {
        var tasks = new[]
        {
            GetAntiPhishPoliciesAsync().ContinueWith(t => (object)t.Result),
            GetAntiMalwarePoliciesAsync().ContinueWith(t => (object)t.Result),
            GetAntiSpamPoliciesAsync().ContinueWith(t => (object)t.Result),
            GetOutboundSpamPoliciesAsync().ContinueWith(t => (object)t.Result),
            GetSafeAttachmentsPoliciesAsync().ContinueWith(t => (object)t.Result),
            GetSafeLinksPoliciesAsync().ContinueWith(t => (object)t.Result),
        };

        await Task.WhenAll(tasks);

        return new DefenderOverviewDto(
            AntiPhish: (AntiPhishResultDto)tasks[0].Result,
            AntiMalware: (AntiMalwareResultDto)tasks[1].Result,
            AntiSpam: (AntiSpamResultDto)tasks[2].Result,
            OutboundSpam: (OutboundSpamResultDto)tasks[3].Result,
            SafeAttachments: (SafeAttachmentsResultDto)tasks[4].Result,
            SafeLinks: (SafeLinksResultDto)tasks[5].Result,
            LastUpdated: DateTime.UtcNow
        );
    }

    public async Task<AntiPhishResultDto> GetAntiPhishPoliciesAsync()
    {
        try
        {
            var result = await InvokeExchangeCommandAsync("Get-AntiPhishPolicy");
            var policies = new List<AntiPhishPolicyDto>();

            if (result != null && result.RootElement.TryGetProperty("value", out var arr) && arr.ValueKind == JsonValueKind.Array)
            {
                foreach (var item in arr.EnumerateArray())
                {
                    try
                    {
                        var name = GetString(item, "Name") ?? GetString(item, "Identity") ?? "Unknown";
                        policies.Add(new AntiPhishPolicyDto(
                            Identity: GetString(item, "Identity") ?? name,
                            Name: name,
                            Enabled: GetBool(item, "Enabled"),
                            IsDefault: GetBool(item, "IsDefault") || name.Equals("Office365 AntiPhish Default", StringComparison.OrdinalIgnoreCase),
                            EnableSpoofIntelligence: GetBool(item, "EnableSpoofIntelligence"),
                            EnableUnauthenticatedSender: GetBool(item, "EnableUnauthenticatedSender"),
                            EnableViaTag: GetBool(item, "EnableViaTag"),
                            PhishThresholdLevel: GetString(item, "PhishThresholdLevel") ?? "1",
                            SpoofQuarantineTag: GetString(item, "SpoofQuarantineTag") ?? "",
                            EnableFirstContactSafetyTips: GetBool(item, "EnableFirstContactSafetyTips"),
                            EnableSimilarUsersSafetyTips: GetBool(item, "EnableSimilarUsersSafetyTips"),
                            EnableSimilarDomainsSafetyTips: GetBool(item, "EnableSimilarDomainsSafetyTips"),
                            EnableUnusualCharactersSafetyTips: GetBool(item, "EnableUnusualCharactersSafetyTips"),
                            EnableMailboxIntelligence: GetBool(item, "EnableMailboxIntelligence"),
                            EnableMailboxIntelligenceProtection: GetBool(item, "EnableMailboxIntelligenceProtection"),
                            MailboxIntelligenceProtectionAction: GetString(item, "MailboxIntelligenceProtectionAction") ?? "NoAction",
                            EnableOrganizationDomainsProtection: GetBool(item, "EnableOrganizationDomainsProtection"),
                            EnableTargetedDomainsProtection: GetBool(item, "EnableTargetedDomainsProtection"),
                            EnableTargetedUserProtection: GetBool(item, "EnableTargetedUserProtection"),
                            TargetedUserProtectionAction: GetString(item, "TargetedUserProtectionAction") ?? "NoAction",
                            TargetedDomainProtectionAction: GetString(item, "TargetedDomainProtectionAction") ?? "NoAction",
                            TargetedUsersToProtect: GetStringArray(item, "TargetedUsersToProtect"),
                            TargetedDomainsToProtect: GetStringArray(item, "TargetedDomainsToProtect"),
                            WhenCreated: GetDateTime(item, "WhenCreated"),
                            WhenChanged: GetDateTime(item, "WhenChanged")
                        ));
                    }
                    catch (Exception ex)
                    {
                        _logger.LogWarning(ex, "Failed to parse anti-phish policy item");
                    }
                }
            }

            return new AntiPhishResultDto(
                Policies: policies.OrderByDescending(p => p.IsDefault).ThenBy(p => p.Name).ToList(),
                TotalCount: policies.Count,
                Error: null
            );
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching anti-phish policies");
            return new AntiPhishResultDto(new List<AntiPhishPolicyDto>(), 0, ex.Message);
        }
    }

    public async Task<AntiMalwareResultDto> GetAntiMalwarePoliciesAsync()
    {
        try
        {
            var result = await InvokeExchangeCommandAsync("Get-MalwareFilterPolicy");
            var policies = new List<AntiMalwarePolicyDto>();

            if (result != null && result.RootElement.TryGetProperty("value", out var arr) && arr.ValueKind == JsonValueKind.Array)
            {
                foreach (var item in arr.EnumerateArray())
                {
                    try
                    {
                        var name = GetString(item, "Name") ?? GetString(item, "Identity") ?? "Unknown";
                        policies.Add(new AntiMalwarePolicyDto(
                            Identity: GetString(item, "Identity") ?? name,
                            Name: name,
                            Enabled: true, // malware filter is always on; no explicit Enabled flag
                            IsDefault: name.Equals("Default", StringComparison.OrdinalIgnoreCase),
                            Action: GetString(item, "Action") ?? "DeleteMessage",
                            EnableFileFilter: GetBool(item, "EnableFileFilter"),
                            FileTypes: GetStringArray(item, "FileTypes"),
                            ZapEnabled: GetBool(item, "ZapEnabled"),
                            EnableInternalSenderAdminNotifications: GetBool(item, "EnableInternalSenderAdminNotifications"),
                            InternalSenderAdminAddress: GetString(item, "InternalSenderAdminAddress"),
                            EnableExternalSenderAdminNotifications: GetBool(item, "EnableExternalSenderAdminNotifications"),
                            ExternalSenderAdminAddress: GetString(item, "ExternalSenderAdminAddress"),
                            WhenCreated: GetDateTime(item, "WhenCreated"),
                            WhenChanged: GetDateTime(item, "WhenChanged")
                        ));
                    }
                    catch (Exception ex)
                    {
                        _logger.LogWarning(ex, "Failed to parse anti-malware policy item");
                    }
                }
            }

            return new AntiMalwareResultDto(
                Policies: policies.OrderByDescending(p => p.IsDefault).ThenBy(p => p.Name).ToList(),
                TotalCount: policies.Count,
                Error: null
            );
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching anti-malware policies");
            return new AntiMalwareResultDto(new List<AntiMalwarePolicyDto>(), 0, ex.Message);
        }
    }

    public async Task<AntiSpamResultDto> GetAntiSpamPoliciesAsync()
    {
        try
        {
            var result = await InvokeExchangeCommandAsync("Get-HostedContentFilterPolicy");
            var policies = new List<AntiSpamPolicyDto>();

            if (result != null && result.RootElement.TryGetProperty("value", out var arr) && arr.ValueKind == JsonValueKind.Array)
            {
                foreach (var item in arr.EnumerateArray())
                {
                    try
                    {
                        var name = GetString(item, "Name") ?? GetString(item, "Identity") ?? "Unknown";
                        var allowedSenderDomains = GetStringArray(item, "AllowedSenderDomains");
                        var blockedSenderDomains = GetStringArray(item, "BlockedSenderDomains");

                        policies.Add(new AntiSpamPolicyDto(
                            Identity: GetString(item, "Identity") ?? name,
                            Name: name,
                            Enabled: GetBool(item, "Enabled"),
                            IsDefault: name.Equals("Default", StringComparison.OrdinalIgnoreCase),
                            SpamAction: GetString(item, "SpamAction") ?? "MoveToJmf",
                            HighConfidenceSpamAction: GetString(item, "HighConfidenceSpamAction") ?? "MoveToJmf",
                            PhishSpamAction: GetString(item, "PhishSpamAction") ?? "Quarantine",
                            HighConfidencePhishAction: GetString(item, "HighConfidencePhishAction") ?? "Quarantine",
                            BulkSpamAction: GetString(item, "BulkSpamAction") ?? "MoveToJmf",
                            BulkThreshold: GetInt(item, "BulkThreshold"),
                            ZapEnabled: GetBool(item, "ZapEnabled"),
                            SpamZapEnabled: GetBool(item, "SpamZapEnabled"),
                            PhishZapEnabled: GetBool(item, "PhishZapEnabled"),
                            EnableEndUserSpamNotifications: GetBool(item, "EnableEndUserSpamNotifications"),
                            EndUserSpamNotificationFrequency: GetInt(item, "EndUserSpamNotificationFrequency"),
                            AllowedSenderDomainsPresent: allowedSenderDomains.Count > 0,
                            BlockedSenderDomainsPresent: blockedSenderDomains.Count > 0,
                            QuarantineTag: GetString(item, "SpamQuarantineTag") ?? "",
                            WhenCreated: GetDateTime(item, "WhenCreated"),
                            WhenChanged: GetDateTime(item, "WhenChanged")
                        ));
                    }
                    catch (Exception ex)
                    {
                        _logger.LogWarning(ex, "Failed to parse anti-spam policy item");
                    }
                }
            }

            return new AntiSpamResultDto(
                Policies: policies.OrderByDescending(p => p.IsDefault).ThenBy(p => p.Name).ToList(),
                TotalCount: policies.Count,
                Error: null
            );
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching anti-spam policies");
            return new AntiSpamResultDto(new List<AntiSpamPolicyDto>(), 0, ex.Message);
        }
    }

    public async Task<OutboundSpamResultDto> GetOutboundSpamPoliciesAsync()
    {
        try
        {
            var result = await InvokeExchangeCommandAsync("Get-HostedOutboundSpamFilterPolicy");
            var policies = new List<OutboundSpamPolicyDto>();

            if (result != null && result.RootElement.TryGetProperty("value", out var arr) && arr.ValueKind == JsonValueKind.Array)
            {
                foreach (var item in arr.EnumerateArray())
                {
                    try
                    {
                        var name = GetString(item, "Name") ?? GetString(item, "Identity") ?? "Unknown";
                        var autoForwardRaw = GetString(item, "AutoForwardingMode") ?? "Automatic";

                        policies.Add(new OutboundSpamPolicyDto(
                            Identity: GetString(item, "Identity") ?? name,
                            Name: name,
                            Enabled: GetBool(item, "Enabled"),
                            IsDefault: name.Equals("Default", StringComparison.OrdinalIgnoreCase),
                            ActionWhenThresholdReached: GetString(item, "ActionWhenThresholdReached") ?? "BlockUserForToday",
                            RecipientLimitExternalPerHour: GetInt(item, "RecipientLimitExternalPerHour"),
                            RecipientLimitInternalPerHour: GetInt(item, "RecipientLimitInternalPerHour"),
                            RecipientLimitPerDay: GetInt(item, "RecipientLimitPerDay"),
                            AutoForwardingMode: !autoForwardRaw.Equals("Off", StringComparison.OrdinalIgnoreCase),
                            AutoForwardingModeValue: autoForwardRaw,
                            WhenCreated: GetDateTime(item, "WhenCreated"),
                            WhenChanged: GetDateTime(item, "WhenChanged")
                        ));
                    }
                    catch (Exception ex)
                    {
                        _logger.LogWarning(ex, "Failed to parse outbound spam policy item");
                    }
                }
            }

            return new OutboundSpamResultDto(
                Policies: policies.OrderByDescending(p => p.IsDefault).ThenBy(p => p.Name).ToList(),
                TotalCount: policies.Count,
                Error: null
            );
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching outbound spam policies");
            return new OutboundSpamResultDto(new List<OutboundSpamPolicyDto>(), 0, ex.Message);
        }
    }

    public async Task<SafeAttachmentsResultDto> GetSafeAttachmentsPoliciesAsync()
    {
        try
        {
            var result = await InvokeExchangeCommandAsync("Get-SafeAttachmentPolicy");
            var policies = new List<SafeAttachmentsPolicyDto>();

            if (result != null && result.RootElement.TryGetProperty("value", out var arr) && arr.ValueKind == JsonValueKind.Array)
            {
                foreach (var item in arr.EnumerateArray())
                {
                    try
                    {
                        var name = GetString(item, "Name") ?? GetString(item, "Identity") ?? "Unknown";
                        policies.Add(new SafeAttachmentsPolicyDto(
                            Identity: GetString(item, "Identity") ?? name,
                            Name: name,
                            Enabled: GetBool(item, "Enable"),
                            IsDefault: GetBool(item, "IsDefault"),
                            Action: GetString(item, "Action") ?? "Block",
                            ActionOnError: GetBool(item, "ActionOnError"),
                            Redirect: GetBool(item, "Redirect"),
                            RedirectAddress: GetString(item, "RedirectAddress"),
                            QuarantineTag: GetBool(item, "QuarantineTag"),
                            WhenCreated: GetDateTime(item, "WhenCreated"),
                            WhenChanged: GetDateTime(item, "WhenChanged")
                        ));
                    }
                    catch (Exception ex)
                    {
                        _logger.LogWarning(ex, "Failed to parse Safe Attachments policy item");
                    }
                }
            }

            return new SafeAttachmentsResultDto(
                Policies: policies.OrderByDescending(p => p.IsDefault).ThenBy(p => p.Name).ToList(),
                TotalCount: policies.Count,
                IsLicensed: true,
                Error: null
            );
        }
        catch (Exception ex)
        {
            var isLicensingError = ex.Message.Contains("license", StringComparison.OrdinalIgnoreCase)
                                || ex.Message.Contains("not available", StringComparison.OrdinalIgnoreCase)
                                || ex.Message.Contains("subscription", StringComparison.OrdinalIgnoreCase)
                                || ex.Message.Contains("cmdlet", StringComparison.OrdinalIgnoreCase);

            _logger.LogWarning(ex, "Error fetching Safe Attachments policies (may not be licensed)");
            return new SafeAttachmentsResultDto(new List<SafeAttachmentsPolicyDto>(), 0, !isLicensingError, ex.Message);
        }
    }

    public async Task<SafeLinksResultDto> GetSafeLinksPoliciesAsync()
    {
        try
        {
            var result = await InvokeExchangeCommandAsync("Get-SafeLinksPolicy");
            var policies = new List<SafeLinksPolicyDto>();

            if (result != null && result.RootElement.TryGetProperty("value", out var arr) && arr.ValueKind == JsonValueKind.Array)
            {
                foreach (var item in arr.EnumerateArray())
                {
                    try
                    {
                        var name = GetString(item, "Name") ?? GetString(item, "Identity") ?? "Unknown";
                        policies.Add(new SafeLinksPolicyDto(
                            Identity: GetString(item, "Identity") ?? name,
                            Name: name,
                            Enabled: GetBool(item, "IsEnabled"),
                            IsDefault: GetBool(item, "IsDefault"),
                            AllowClickThrough: GetBool(item, "AllowClickThrough"),
                            DisableUrlRewrite: GetBool(item, "DisableUrlRewrite"),
                            EnableForInternalSenders: GetBool(item, "EnableForInternalSenders"),
                            EnableOrganizationBranding: GetBool(item, "EnableOrganizationBranding"),
                            EnableSafeLinksForEmail: GetBool(item, "EnableSafeLinksForEmail"),
                            EnableSafeLinksForTeams: GetBool(item, "EnableSafeLinksForTeams"),
                            EnableSafeLinksForOffice: GetBool(item, "EnableSafeLinksForOffice"),
                            TrackClicks: GetBool(item, "TrackClicks"),
                            ScanUrls: GetBool(item, "ScanUrls"),
                            DoNotRewriteUrls: GetStringArray(item, "DoNotRewriteUrls"),
                            WhenCreated: GetDateTime(item, "WhenCreated"),
                            WhenChanged: GetDateTime(item, "WhenChanged")
                        ));
                    }
                    catch (Exception ex)
                    {
                        _logger.LogWarning(ex, "Failed to parse Safe Links policy item");
                    }
                }
            }

            return new SafeLinksResultDto(
                Policies: policies.OrderByDescending(p => p.IsDefault).ThenBy(p => p.Name).ToList(),
                TotalCount: policies.Count,
                IsLicensed: true,
                Error: null
            );
        }
        catch (Exception ex)
        {
            var isLicensingError = ex.Message.Contains("license", StringComparison.OrdinalIgnoreCase)
                                || ex.Message.Contains("not available", StringComparison.OrdinalIgnoreCase)
                                || ex.Message.Contains("subscription", StringComparison.OrdinalIgnoreCase)
                                || ex.Message.Contains("cmdlet", StringComparison.OrdinalIgnoreCase);

            _logger.LogWarning(ex, "Error fetching Safe Links policies (may not be licensed)");
            return new SafeLinksResultDto(new List<SafeLinksPolicyDto>(), 0, !isLicensingError, ex.Message);
        }
    }

    // ── Helpers ───────────────────────────────────────────────────────────────

    private static string? GetString(JsonElement el, string prop)
    {
        if (el.TryGetProperty(prop, out var v) && v.ValueKind == JsonValueKind.String) return v.GetString();
        return null;
    }

    private static bool GetBool(JsonElement el, string prop)
    {
        if (el.TryGetProperty(prop, out var v))
        {
            if (v.ValueKind == JsonValueKind.True) return true;
            if (v.ValueKind == JsonValueKind.False) return false;
        }
        return false;
    }

    private static int GetInt(JsonElement el, string prop)
    {
        if (el.TryGetProperty(prop, out var v) && v.ValueKind == JsonValueKind.Number)
            return v.GetInt32();
        return 0;
    }

    private static DateTime? GetDateTime(JsonElement el, string prop)
    {
        var s = GetString(el, prop);
        if (!string.IsNullOrEmpty(s) && DateTime.TryParse(s, out var dt)) return dt;
        return null;
    }

    private static List<string> GetStringArray(JsonElement el, string prop)
    {
        var result = new List<string>();
        if (!el.TryGetProperty(prop, out var v)) return result;
        if (v.ValueKind == JsonValueKind.Array)
        {
            foreach (var item in v.EnumerateArray())
                if (item.ValueKind == JsonValueKind.String && !string.IsNullOrEmpty(item.GetString()))
                    result.Add(item.GetString()!);
        }
        else if (v.ValueKind == JsonValueKind.String && !string.IsNullOrEmpty(v.GetString()))
        {
            result.Add(v.GetString()!);
        }
        return result;
    }
}
