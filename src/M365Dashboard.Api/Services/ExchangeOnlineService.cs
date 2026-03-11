using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using Azure.Identity;
using M365Dashboard.Api.Models.Dtos;

namespace M365Dashboard.Api.Services;

public interface IExchangeOnlineService
{
    Task<ExchangeDistributionListResultDto> GetDistributionListsAsync(int take = 100);
    Task<ExchangeDistributionListDetailDto?> GetDistributionListAsync(string identity);
    Task<List<ExchangeDistributionListMemberDto>> GetDistributionListMembersAsync(string identity);
    Task<object> DebugGetRecipientsAsync();
    Task<MailboxForwardingResultDto> GetMailboxesWithForwardingAsync(int take = 500);
    Task<InboxRuleForwardingResultDto> GetInboxRulesWithForwardingAsync(int take = 100);
}

public class ExchangeOnlineService : IExchangeOnlineService
{
    private readonly IConfiguration _configuration;
    private readonly ILogger<ExchangeOnlineService> _logger;
    private readonly IHttpClientFactory _httpClientFactory;

    public ExchangeOnlineService(
        IConfiguration configuration,
        ILogger<ExchangeOnlineService> logger,
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
        var scopes = new[] { "https://outlook.office365.com/.default" };
        var token = await credential.GetTokenAsync(new Azure.Core.TokenRequestContext(scopes));
        
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

        _logger.LogInformation("Invoking Exchange command: {Cmdlet}", cmdletName);

        var response = await httpClient.PostAsync(requestUrl, jsonContent);
        var responseContent = await response.Content.ReadAsStringAsync();

        if (!response.IsSuccessStatusCode)
        {
            _logger.LogError("Exchange API error: {StatusCode} - {Content}", response.StatusCode, responseContent);
            throw new Exception($"Exchange API returned {response.StatusCode}: {responseContent}");
        }

        // Always log the response for debugging
        _logger.LogInformation("Exchange response for {Cmdlet}: {Content}", cmdletName, responseContent);

        if (string.IsNullOrWhiteSpace(responseContent))
        {
            _logger.LogWarning("Exchange returned empty response for {Cmdlet}", cmdletName);
            return null;
        }

        return JsonDocument.Parse(responseContent);
    }

    public async Task<ExchangeDistributionListResultDto> GetDistributionListsAsync(int take = 100)
    {
        try
        {
            _logger.LogInformation("Fetching Exchange distribution lists");

            var distributionLists = new List<ExchangeDistributionListDto>();

            // Fetch regular Distribution Groups
            var dgParams = new Dictionary<string, object>
            {
                { "ResultSize", take }
            };
            var dgResult = await InvokeExchangeCommandAsync("Get-DistributionGroup", dgParams);

            if (dgResult != null)
            {
                if (dgResult.RootElement.TryGetProperty("value", out var valueArray) && valueArray.ValueKind == JsonValueKind.Array)
                {
                    foreach (var item in valueArray.EnumerateArray())
                    {
                        try
                        {
                            var dl = ParseDistributionList(item, false);
                            if (dl != null)
                            {
                                distributionLists.Add(dl);
                            }
                        }
                        catch (Exception ex)
                        {
                            _logger.LogWarning(ex, "Failed to parse distribution list item");
                        }
                    }
                }
            }

            _logger.LogInformation("Found {Count} regular distribution groups", distributionLists.Count);

            // Fetch Dynamic Distribution Groups
            var ddgParams = new Dictionary<string, object>
            {
                { "ResultSize", take }
            };
            var ddgResult = await InvokeExchangeCommandAsync("Get-DynamicDistributionGroup", ddgParams);

            if (ddgResult != null)
            {
                if (ddgResult.RootElement.TryGetProperty("value", out var valueArray) && valueArray.ValueKind == JsonValueKind.Array)
                {
                    foreach (var item in valueArray.EnumerateArray())
                    {
                        try
                        {
                            var dl = ParseDistributionList(item, true);
                            if (dl != null)
                            {
                                distributionLists.Add(dl);
                            }
                        }
                        catch (Exception ex)
                        {
                            _logger.LogWarning(ex, "Failed to parse dynamic distribution list item");
                        }
                    }
                }
            }

            _logger.LogInformation("Found {Count} total distribution lists (including dynamic)", distributionLists.Count);

            return new ExchangeDistributionListResultDto(
                DistributionLists: distributionLists.OrderBy(d => d.DisplayName).ToList(),
                TotalCount: distributionLists.Count
            );
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching Exchange distribution lists");
            throw;
        }
    }

    public async Task<ExchangeDistributionListDetailDto?> GetDistributionListAsync(string identity)
    {
        try
        {
            _logger.LogInformation("Fetching Exchange distribution list: {Identity}", identity);

            var parameters = new Dictionary<string, object>
            {
                { "Identity", identity }
            };

            // Try regular distribution group first
            JsonDocument? result = null;
            var isDynamic = false;
            
            try
            {
                result = await InvokeExchangeCommandAsync("Get-DistributionGroup", parameters);
            }
            catch (Exception ex)
            {
                _logger.LogDebug(ex, "Get-DistributionGroup failed for {Identity}, trying dynamic", identity);
            }

            // Check if we got results
            JsonElement valueArray;
            if (result == null || !result.RootElement.TryGetProperty("value", out valueArray) || 
                valueArray.ValueKind != JsonValueKind.Array || valueArray.GetArrayLength() == 0)
            {
                // Not found as regular DG, try dynamic distribution group
                try
                {
                    result = await InvokeExchangeCommandAsync("Get-DynamicDistributionGroup", parameters);
                    isDynamic = true;
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "Get-DynamicDistributionGroup also failed for {Identity}", identity);
                    throw;
                }
            }

            if (result != null)
            {
                if (result.RootElement.TryGetProperty("value", out valueArray) && 
                    valueArray.ValueKind == JsonValueKind.Array &&
                    valueArray.GetArrayLength() > 0)
                {
                    var item = valueArray[0];
                    var dl = ParseDistributionList(item, isDynamic);
                    
                    if (dl != null)
                    {
                        // Get members (only for regular DLs, dynamic DLs use filters)
                        var members = new List<ExchangeDistributionListMemberDto>();
                        var recipientFilter = isDynamic ? GetStringProperty(item, "RecipientFilter") : null;
                        
                        if (!isDynamic)
                        {
                            members = await GetDistributionListMembersAsync(identity);
                        }
                        else if (!string.IsNullOrEmpty(recipientFilter))
                        {
                            // For dynamic DDLs, preview the calculated members
                            members = await GetDynamicDistributionListMembersAsync(recipientFilter);
                        }

                        return new ExchangeDistributionListDetailDto(
                            Id: dl.Id,
                            DisplayName: dl.DisplayName,
                            PrimarySmtpAddress: dl.PrimarySmtpAddress,
                            Alias: dl.Alias,
                            Description: GetStringProperty(item, "Description"),
                            ManagedBy: GetStringArrayProperty(item, "ManagedBy"),
                            GroupType: dl.GroupType,
                            RecipientType: dl.RecipientType,
                            MemberCount: isDynamic ? 0 : members.Count,
                            Members: members,
                            WhenCreated: dl.WhenCreated,
                            WhenChanged: GetDateTimeProperty(item, "WhenChanged"),
                            HiddenFromAddressListsEnabled: dl.HiddenFromAddressListsEnabled,
                            RequireSenderAuthenticationEnabled: GetBoolProperty(item, "RequireSenderAuthenticationEnabled"),
                            AcceptMessagesOnlyFromSendersOrMembers: GetStringArrayProperty(item, "AcceptMessagesOnlyFromSendersOrMembers"),
                            EmailAddresses: GetStringArrayProperty(item, "EmailAddresses"),
                            IsDynamic: isDynamic,
                            RecipientFilter: isDynamic ? GetStringProperty(item, "RecipientFilter") : null
                        );
                    }
                }
            }

            return null;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching Exchange distribution list: {Identity}", identity);
            throw;
        }
    }

    public async Task<List<ExchangeDistributionListMemberDto>> GetDistributionListMembersAsync(string identity)
    {
        try
        {
            _logger.LogInformation("Fetching members for distribution list: {Identity}", identity);

            var parameters = new Dictionary<string, object>
            {
                { "Identity", identity },
                { "ResultSize", 1000 }
            };

            var result = await InvokeExchangeCommandAsync("Get-DistributionGroupMember", parameters);

            var members = new List<ExchangeDistributionListMemberDto>();

            if (result != null)
            {
                if (result.RootElement.TryGetProperty("value", out var valueArray) && valueArray.ValueKind == JsonValueKind.Array)
                {
                    foreach (var item in valueArray.EnumerateArray())
                    {
                        try
                        {
                            var member = new ExchangeDistributionListMemberDto(
                                Id: GetStringProperty(item, "Guid") ?? GetStringProperty(item, "Identity") ?? "",
                                DisplayName: GetStringProperty(item, "DisplayName") ?? "Unknown",
                                PrimarySmtpAddress: GetStringProperty(item, "PrimarySmtpAddress"),
                                RecipientType: GetStringProperty(item, "RecipientType") ?? "Unknown"
                            );
                            members.Add(member);
                        }
                        catch (Exception ex)
                        {
                            _logger.LogWarning(ex, "Failed to parse distribution list member");
                        }
                    }
                }
            }

            _logger.LogInformation("Found {Count} members in distribution list", members.Count);

            return members;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching distribution list members: {Identity}", identity);
            throw;
        }
    }

    private async Task<List<ExchangeDistributionListMemberDto>> GetDynamicDistributionListMembersAsync(string recipientFilter)
    {
        try
        {
            _logger.LogInformation("Fetching preview members for dynamic DDL with filter: {Filter}", recipientFilter);

            var parameters = new Dictionary<string, object>
            {
                { "RecipientPreviewFilter", recipientFilter },
                { "ResultSize", 500 }
            };

            var result = await InvokeExchangeCommandAsync("Get-Recipient", parameters);

            var members = new List<ExchangeDistributionListMemberDto>();

            if (result != null)
            {
                if (result.RootElement.TryGetProperty("value", out var valueArray) && valueArray.ValueKind == JsonValueKind.Array)
                {
                    foreach (var item in valueArray.EnumerateArray())
                    {
                        try
                        {
                            var member = new ExchangeDistributionListMemberDto(
                                Id: GetStringProperty(item, "Guid") ?? GetStringProperty(item, "Identity") ?? "",
                                DisplayName: GetStringProperty(item, "DisplayName") ?? "Unknown",
                                PrimarySmtpAddress: GetStringProperty(item, "PrimarySmtpAddress"),
                                RecipientType: GetStringProperty(item, "RecipientType") ?? "Unknown"
                            );
                            members.Add(member);
                        }
                        catch (Exception ex)
                        {
                            _logger.LogWarning(ex, "Failed to parse dynamic DDL member");
                        }
                    }
                }
            }

            _logger.LogInformation("Found {Count} preview members for dynamic DDL", members.Count);

            return members;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching dynamic DDL members");
            return new List<ExchangeDistributionListMemberDto>();
        }
    }

    private ExchangeDistributionListDto? ParseDistributionList(JsonElement item, bool isDynamic)
    {
        var id = GetStringProperty(item, "Guid") ?? GetStringProperty(item, "ExternalDirectoryObjectId");
        var displayName = GetStringProperty(item, "DisplayName");

        if (string.IsNullOrEmpty(id) || string.IsNullOrEmpty(displayName))
        {
            return null;
        }

        var groupType = isDynamic ? "Dynamic Distribution" : (GetStringProperty(item, "GroupType") ?? "Distribution");
        var recipientType = GetStringProperty(item, "RecipientType") ?? (isDynamic ? "DynamicDistributionGroup" : "MailUniversalDistributionGroup");

        return new ExchangeDistributionListDto(
            Id: id,
            DisplayName: displayName,
            PrimarySmtpAddress: GetStringProperty(item, "PrimarySmtpAddress"),
            Alias: GetStringProperty(item, "Alias"),
            GroupType: groupType,
            RecipientType: recipientType,
            MemberCount: GetIntProperty(item, "GroupMemberCount"),
            WhenCreated: GetDateTimeProperty(item, "WhenCreated"),
            HiddenFromAddressListsEnabled: GetBoolProperty(item, "HiddenFromAddressListsEnabled")
        );
    }

    private static string? GetStringProperty(JsonElement element, string propertyName)
    {
        if (element.TryGetProperty(propertyName, out var prop) && prop.ValueKind == JsonValueKind.String)
        {
            return prop.GetString();
        }
        return null;
    }

    private static int GetIntProperty(JsonElement element, string propertyName)
    {
        if (element.TryGetProperty(propertyName, out var prop))
        {
            if (prop.ValueKind == JsonValueKind.Number)
            {
                return prop.GetInt32();
            }
        }
        return 0;
    }

    private static bool GetBoolProperty(JsonElement element, string propertyName)
    {
        if (element.TryGetProperty(propertyName, out var prop))
        {
            if (prop.ValueKind == JsonValueKind.True) return true;
            if (prop.ValueKind == JsonValueKind.False) return false;
        }
        return false;
    }

    private static DateTime? GetDateTimeProperty(JsonElement element, string propertyName)
    {
        if (element.TryGetProperty(propertyName, out var prop) && prop.ValueKind == JsonValueKind.String)
        {
            var str = prop.GetString();
            if (!string.IsNullOrEmpty(str) && DateTime.TryParse(str, out var dt))
            {
                return dt;
            }
        }
        return null;
    }

    private static List<string> GetStringArrayProperty(JsonElement element, string propertyName)
    {
        var result = new List<string>();
        if (element.TryGetProperty(propertyName, out var prop))
        {
            if (prop.ValueKind == JsonValueKind.Array)
            {
                foreach (var item in prop.EnumerateArray())
                {
                    if (item.ValueKind == JsonValueKind.String)
                    {
                        var str = item.GetString();
                        if (!string.IsNullOrEmpty(str))
                        {
                            result.Add(str);
                        }
                    }
                }
            }
            else if (prop.ValueKind == JsonValueKind.String)
            {
                var str = prop.GetString();
                if (!string.IsNullOrEmpty(str))
                {
                    result.Add(str);
                }
            }
        }
        return result;
    }

    public async Task<object> DebugGetRecipientsAsync()
    {
        try
        {
            _logger.LogInformation("Debug: Fetching all recipients from Exchange");

            // Try Get-Recipient to see all mail-enabled objects
            var recipientParams = new Dictionary<string, object>
            {
                { "ResultSize", 50 }
            };
            var recipientResult = await InvokeExchangeCommandAsync("Get-Recipient", recipientParams);
            
            var recipients = new List<object>();
            if (recipientResult != null && recipientResult.RootElement.TryGetProperty("value", out var recipientArray))
            {
                foreach (var item in recipientArray.EnumerateArray())
                {
                    recipients.Add(new
                    {
                        DisplayName = GetStringProperty(item, "DisplayName"),
                        PrimarySmtpAddress = GetStringProperty(item, "PrimarySmtpAddress"),
                        RecipientType = GetStringProperty(item, "RecipientType"),
                        RecipientTypeDetails = GetStringProperty(item, "RecipientTypeDetails")
                    });
                }
            }

            // Also try Get-Group to see all groups
            var groupParams = new Dictionary<string, object>
            {
                { "ResultSize", 50 }
            };
            var groupResult = await InvokeExchangeCommandAsync("Get-Group", groupParams);
            
            var groups = new List<object>();
            if (groupResult != null && groupResult.RootElement.TryGetProperty("value", out var groupArray))
            {
                foreach (var item in groupArray.EnumerateArray())
                {
                    groups.Add(new
                    {
                        DisplayName = GetStringProperty(item, "DisplayName"),
                        PrimarySmtpAddress = GetStringProperty(item, "PrimarySmtpAddress"),
                        GroupType = GetStringProperty(item, "GroupType"),
                        RecipientType = GetStringProperty(item, "RecipientType"),
                        RecipientTypeDetails = GetStringProperty(item, "RecipientTypeDetails")
                    });
                }
            }

            return new
            {
                success = true,
                recipientCount = recipients.Count,
                recipients = recipients,
                groupCount = groups.Count,
                groups = groups
            };
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Debug recipients failed");
            return new { success = false, error = ex.Message };
        }
    }

    public async Task<MailboxForwardingResultDto> GetMailboxesWithForwardingAsync(int take = 500)
    {
        try
        {
            _logger.LogInformation("Fetching mailboxes with forwarding enabled");

            var parameters = new Dictionary<string, object>
            {
                { "ResultSize", take }
            };

            var result = await InvokeExchangeCommandAsync("Get-Mailbox", parameters);

            var mailboxesWithForwarding = new List<MailboxForwardingDto>();

            if (result != null)
            {
                if (result.RootElement.TryGetProperty("value", out var valueArray) && valueArray.ValueKind == JsonValueKind.Array)
                {
                    foreach (var item in valueArray.EnumerateArray())
                    {
                        try
                        {
                            var forwardingAddress = GetStringProperty(item, "ForwardingAddress");
                            var forwardingSmtpAddress = GetStringProperty(item, "ForwardingSmtpAddress");
                            var deliverToMailboxAndForward = GetBoolProperty(item, "DeliverToMailboxAndForward");

                            // Only include if forwarding is configured
                            if (!string.IsNullOrEmpty(forwardingAddress) || !string.IsNullOrEmpty(forwardingSmtpAddress))
                            {
                                var primarySmtp = GetStringProperty(item, "PrimarySmtpAddress") ?? "";
                                var forwardTarget = forwardingSmtpAddress ?? forwardingAddress ?? "";
                                
                                // Determine if external (simple check - not matching primary domain)
                                var primaryDomain = primarySmtp.Contains("@") ? primarySmtp.Split('@')[1].ToLower() : "";
                                var targetDomain = forwardTarget.Contains("@") ? forwardTarget.Split('@')[1].ToLower() : "";
                                var isExternal = !string.IsNullOrEmpty(targetDomain) && 
                                                 !string.IsNullOrEmpty(primaryDomain) && 
                                                 !targetDomain.Equals(primaryDomain, StringComparison.OrdinalIgnoreCase) &&
                                                 !targetDomain.EndsWith(".onmicrosoft.com", StringComparison.OrdinalIgnoreCase);

                                mailboxesWithForwarding.Add(new MailboxForwardingDto(
                                    Id: GetStringProperty(item, "Guid") ?? GetStringProperty(item, "ExternalDirectoryObjectId") ?? "",
                                    DisplayName: GetStringProperty(item, "DisplayName") ?? "Unknown",
                                    UserPrincipalName: GetStringProperty(item, "UserPrincipalName") ?? "",
                                    PrimarySmtpAddress: primarySmtp,
                                    ForwardingAddress: forwardingAddress,
                                    ForwardingSmtpAddress: forwardingSmtpAddress,
                                    ForwardingTarget: forwardTarget,
                                    ForwardingTargetDomain: targetDomain,
                                    DeliverToMailboxAndForward: deliverToMailboxAndForward,
                                    IsExternal: isExternal,
                                    RecipientTypeDetails: GetStringProperty(item, "RecipientTypeDetails") ?? "UserMailbox"
                                ));
                            }
                        }
                        catch (Exception ex)
                        {
                            _logger.LogWarning(ex, "Failed to parse mailbox for forwarding check");
                        }
                    }
                }
            }

            var externalCount = mailboxesWithForwarding.Count(m => m.IsExternal);
            var internalCount = mailboxesWithForwarding.Count(m => !m.IsExternal);

            _logger.LogInformation("Found {Count} mailboxes with forwarding ({External} external, {Internal} internal)", 
                mailboxesWithForwarding.Count, externalCount, internalCount);

            return new MailboxForwardingResultDto(
                Mailboxes: mailboxesWithForwarding.OrderByDescending(m => m.IsExternal).ThenBy(m => m.DisplayName).ToList(),
                TotalCount: mailboxesWithForwarding.Count,
                ExternalCount: externalCount,
                InternalCount: internalCount,
                LastUpdated: DateTime.UtcNow
            );
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching mailboxes with forwarding");
            throw;
        }
    }

    public async Task<InboxRuleForwardingResultDto> GetInboxRulesWithForwardingAsync(int take = 100)
    {
        try
        {
            _logger.LogInformation("Fetching inbox rules with forwarding for up to {Take} mailboxes", take);

            // First get mailboxes
            var mailboxParams = new Dictionary<string, object>
            {
                { "ResultSize", take },
                { "RecipientTypeDetails", "UserMailbox" }
            };

            var mailboxResult = await InvokeExchangeCommandAsync("Get-Mailbox", mailboxParams);

            var forwardingRules = new List<InboxRuleForwardingDto>();
            var tenantDomains = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            int mailboxesScanned = 0;
            int mailboxesWithForwarding = 0;

            // Collect tenant domains from mailbox addresses
            var mailboxes = new List<(string Identity, string DisplayName, string Email)>();
            
            if (mailboxResult != null && mailboxResult.RootElement.TryGetProperty("value", out var mailboxArray) && mailboxArray.ValueKind == JsonValueKind.Array)
            {
                foreach (var item in mailboxArray.EnumerateArray())
                {
                    var email = GetStringProperty(item, "PrimarySmtpAddress") ?? "";
                    var identity = GetStringProperty(item, "UserPrincipalName") ?? email;
                    var displayName = GetStringProperty(item, "DisplayName") ?? identity;
                    
                    if (!string.IsNullOrEmpty(identity))
                    {
                        mailboxes.Add((identity, displayName, email));
                        
                        // Extract domain for tenant domain list
                        if (email.Contains('@'))
                        {
                            var domain = email.Split('@')[1];
                            if (!domain.EndsWith(".onmicrosoft.com", StringComparison.OrdinalIgnoreCase))
                            {
                                tenantDomains.Add(domain);
                            }
                        }
                    }
                }
            }

            _logger.LogInformation("Found {Count} mailboxes to scan for inbox rules", mailboxes.Count);

            // Now get inbox rules for each mailbox
            foreach (var (identity, displayName, email) in mailboxes)
            {
                mailboxesScanned++;
                
                try
                {
                    var ruleParams = new Dictionary<string, object>
                    {
                        { "Mailbox", identity }
                    };

                    var ruleResult = await InvokeExchangeCommandAsync("Get-InboxRule", ruleParams);

                    if (ruleResult != null && ruleResult.RootElement.TryGetProperty("value", out var ruleArray) && ruleArray.ValueKind == JsonValueKind.Array)
                    {
                        var hasForwardingRule = false;
                        
                        foreach (var rule in ruleArray.EnumerateArray())
                        {
                            // Check for forwarding actions
                            var forwardTo = GetStringArrayProperty(rule, "ForwardTo");
                            var forwardAsAttachmentTo = GetStringArrayProperty(rule, "ForwardAsAttachmentTo");
                            var redirectTo = GetStringArrayProperty(rule, "RedirectTo");

                            var allForwardTargets = new List<(string Address, string Type)>();
                            
                            foreach (var addr in forwardTo)
                            {
                                var cleanAddr = ExtractEmailFromRecipient(addr);
                                if (!string.IsNullOrEmpty(cleanAddr))
                                    allForwardTargets.Add((cleanAddr, "Forward"));
                            }
                            foreach (var addr in forwardAsAttachmentTo)
                            {
                                var cleanAddr = ExtractEmailFromRecipient(addr);
                                if (!string.IsNullOrEmpty(cleanAddr))
                                    allForwardTargets.Add((cleanAddr, "ForwardAsAttachment"));
                            }
                            foreach (var addr in redirectTo)
                            {
                                var cleanAddr = ExtractEmailFromRecipient(addr);
                                if (!string.IsNullOrEmpty(cleanAddr))
                                    allForwardTargets.Add((cleanAddr, "Redirect"));
                            }

                            if (allForwardTargets.Any())
                            {
                                hasForwardingRule = true;
                                var ruleName = GetStringProperty(rule, "Name") ?? "Unnamed Rule";
                                var ruleEnabled = GetBoolProperty(rule, "Enabled");
                                var ruleId = GetStringProperty(rule, "RuleIdentity") ?? GetStringProperty(rule, "Identity") ?? "";

                                foreach (var (targetAddress, forwardType) in allForwardTargets)
                                {
                                    var targetDomain = targetAddress.Contains('@') ? targetAddress.Split('@')[1].ToLower() : "";
                                    var isExternal = !string.IsNullOrEmpty(targetDomain) && !tenantDomains.Contains(targetDomain);

                                    forwardingRules.Add(new InboxRuleForwardingDto(
                                        UserId: identity,
                                        UserPrincipalName: identity,
                                        DisplayName: displayName,
                                        Mail: email,
                                        RuleId: ruleId,
                                        RuleName: ruleName,
                                        RuleEnabled: ruleEnabled,
                                        ForwardingType: forwardType,
                                        ForwardingTarget: targetAddress,
                                        ForwardingTargetDomain: targetDomain,
                                        IsExternal: isExternal,
                                        DeliverToMailbox: forwardType != "Redirect",
                                        RiskLevel: isExternal ? "High" : "Low"
                                    ));
                                }
                            }
                        }
                        
                        if (hasForwardingRule)
                            mailboxesWithForwarding++;
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogWarning("Could not get inbox rules for {Identity}: {Error}", identity, ex.Message);
                }
            }

            var externalCount = forwardingRules.Count(r => r.IsExternal);
            var internalCount = forwardingRules.Count(r => !r.IsExternal);

            _logger.LogInformation("Inbox rule scan complete. Scanned: {Scanned}, With forwarding: {WithForwarding}, Rules found: {Rules}",
                mailboxesScanned, mailboxesWithForwarding, forwardingRules.Count);

            return new InboxRuleForwardingResultDto(
                ForwardingRules: forwardingRules.OrderByDescending(r => r.IsExternal).ThenBy(r => r.DisplayName).ToList(),
                TotalMailboxesScanned: mailboxesScanned,
                MailboxesWithForwarding: mailboxesWithForwarding,
                TotalForwardingRules: forwardingRules.Count,
                ExternalForwardingCount: externalCount,
                InternalForwardingCount: internalCount,
                TenantDomains: tenantDomains.ToList(),
                LastUpdated: DateTime.UtcNow
            );
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching inbox rules with forwarding");
            throw;
        }
    }

    private static string ExtractEmailFromRecipient(string recipient)
    {
        // Exchange returns recipients in format like "User Name [SMTP:user@domain.com]" or just "user@domain.com"
        if (string.IsNullOrEmpty(recipient))
            return "";

        // Check for SMTP: format
        var smtpIndex = recipient.IndexOf("SMTP:", StringComparison.OrdinalIgnoreCase);
        if (smtpIndex >= 0)
        {
            var start = smtpIndex + 5;
            var end = recipient.IndexOf(']', start);
            if (end > start)
                return recipient.Substring(start, end - start).Trim();
            return recipient.Substring(start).Trim().TrimEnd(']');
        }

        // Check if it looks like an email already
        if (recipient.Contains('@'))
        {
            // Remove any surrounding brackets or quotes
            return recipient.Trim().Trim('[', ']', '"', '\'');
        }

        return recipient;
    }
}

// DTOs for Exchange Distribution Lists
public record ExchangeDistributionListDto(
    string Id,
    string DisplayName,
    string? PrimarySmtpAddress,
    string? Alias,
    string GroupType,
    string RecipientType,
    int MemberCount,
    DateTime? WhenCreated,
    bool HiddenFromAddressListsEnabled
);

public record ExchangeDistributionListResultDto(
    List<ExchangeDistributionListDto> DistributionLists,
    int TotalCount
);

public record ExchangeDistributionListMemberDto(
    string Id,
    string DisplayName,
    string? PrimarySmtpAddress,
    string RecipientType
);

public record ExchangeDistributionListDetailDto(
    string Id,
    string DisplayName,
    string? PrimarySmtpAddress,
    string? Alias,
    string? Description,
    List<string> ManagedBy,
    string GroupType,
    string RecipientType,
    int MemberCount,
    List<ExchangeDistributionListMemberDto> Members,
    DateTime? WhenCreated,
    DateTime? WhenChanged,
    bool HiddenFromAddressListsEnabled,
    bool RequireSenderAuthenticationEnabled,
    List<string> AcceptMessagesOnlyFromSendersOrMembers,
    List<string> EmailAddresses,
    bool IsDynamic = false,
    string? RecipientFilter = null
);

// DTOs for Mailbox Forwarding
public record MailboxForwardingDto(
    string Id,
    string DisplayName,
    string UserPrincipalName,
    string PrimarySmtpAddress,
    string? ForwardingAddress,
    string? ForwardingSmtpAddress,
    string ForwardingTarget,
    string ForwardingTargetDomain,
    bool DeliverToMailboxAndForward,
    bool IsExternal,
    string RecipientTypeDetails
);

public record MailboxForwardingResultDto(
    List<MailboxForwardingDto> Mailboxes,
    int TotalCount,
    int ExternalCount,
    int InternalCount,
    DateTime LastUpdated
);

// DTOs for Inbox Rule Forwarding
public record InboxRuleForwardingDto(
    string UserId,
    string UserPrincipalName,
    string DisplayName,
    string Mail,
    string RuleId,
    string RuleName,
    bool RuleEnabled,
    string ForwardingType,
    string ForwardingTarget,
    string ForwardingTargetDomain,
    bool IsExternal,
    bool DeliverToMailbox,
    string RiskLevel
);

public record InboxRuleForwardingResultDto(
    List<InboxRuleForwardingDto> ForwardingRules,
    int TotalMailboxesScanned,
    int MailboxesWithForwarding,
    int TotalForwardingRules,
    int ExternalForwardingCount,
    int InternalForwardingCount,
    List<string> TenantDomains,
    DateTime LastUpdated
);
