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
    Task<MailboxAccessResultDto> GetMailboxAccessForUserAsync(string userEmail);
    Task<MailboxAccessResultDto> GetMailboxDelegatesAsync(string mailboxEmail);
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

    /// <summary>
    /// Returns all mailboxes that the given user has been granted access to.
    /// Checks Full Access (Get-MailboxPermission) and Send As (Get-RecipientPermission)
    /// across all mailboxes in the tenant.
    /// </summary>
    public async Task<MailboxAccessResultDto> GetMailboxAccessForUserAsync(string userEmail)
    {
        var tenantId = _configuration["AzureAd:TenantId"]!;
        var token = await GetExchangeTokenAsync();

        using var httpClient = _httpClientFactory.CreateClient();
        httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);

        var requestUrl = $"https://outlook.office365.com/adminapi/beta/{tenantId}/InvokeCommand";

        // ── 1. Full Access permissions ─────────────────────────────────────────
        // Get-MailboxPermission -Identity <each mailbox> is too slow at scale.
        // Instead enumerate all mailboxes and filter client-side, OR use:
        // Get-MailboxPermission on each mailbox is O(n) — instead search all
        // mailbox permissions for the trustee = userEmail.
        // Exchange REST does not support a direct "who does this user have access to"
        // query, so we page Get-Mailbox and then batch-check.
        // For MSP use the simpler approach: Get all mailboxes, then for each check
        // permissions — but that's too slow. Use the supported approach:
        // Get-MailboxPermission -Identity ALL and filter server-side via OData-ish where.
        // The most practical approach via REST API: run Get-Mailbox (all), then
        // Get-EXOMailboxPermission which supports -User parameter to find all grants.

        var fullAccessMailboxes = new List<MailboxAccessEntryDto>();
        var sendAsMailboxes     = new List<MailboxAccessEntryDto>();
        var sendOnBehalfMailboxes = new List<MailboxAccessEntryDto>();

        // Get all user mailboxes first (needed for both queries)
        var allMailboxes = await GetAllMailboxesAsync(httpClient, requestUrl);
        _logger.LogInformation("Checking {Count} mailboxes for access by {User}", allMailboxes.Count, userEmail);

        foreach (var mailbox in allMailboxes)
        {
            var mbxEmail = mailbox.PrimarySmtpAddress;
            if (string.IsNullOrEmpty(mbxEmail)) continue;

            // Skip the user's own mailbox
            if (mbxEmail.Equals(userEmail, StringComparison.OrdinalIgnoreCase)) continue;

            try
            {
                // Full Access
                var permBody = new
                {
                    CmdletInput = new
                    {
                        CmdletName = "Get-MailboxPermission",
                        Parameters = new Dictionary<string, object>
                        {
                            ["Identity"] = mbxEmail,
                            ["User"]     = userEmail
                        }
                    }
                };

                var permResponse = await httpClient.PostAsync(requestUrl,
                    new StringContent(JsonSerializer.Serialize(permBody), Encoding.UTF8, "application/json"));

                if (permResponse.IsSuccessStatusCode)
                {
                    var permJson = await permResponse.Content.ReadAsStringAsync();
                    var permDoc  = JsonDocument.Parse(permJson);
                    var perms    = permDoc.RootElement.GetProperty("value");

                    foreach (var perm in perms.EnumerateArray())
                    {
                        var rights = perm.TryGetProperty("AccessRights", out var ar)
                            ? string.Join(", ", ar.EnumerateArray().Select(x => x.GetString() ?? ""))
                            : "";

                        var isDenied = perm.TryGetProperty("Deny", out var deny) && deny.GetBoolean();
                        if (isDenied) continue;

                        if (rights.Contains("FullAccess", StringComparison.OrdinalIgnoreCase))
                        {
                            fullAccessMailboxes.Add(new MailboxAccessEntryDto(
                                MailboxEmail:       mbxEmail,
                                MailboxDisplayName: mailbox.DisplayName,
                                MailboxType:        mailbox.RecipientTypeDetails ?? "UserMailbox",
                                Permission:         "Full Access",
                                GrantedTo:          userEmail,
                                IsInherited:        perm.TryGetProperty("IsInherited", out var inh) && inh.GetBoolean()
                            ));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogDebug(ex, "Full access check failed for {Mailbox}", mbxEmail);
            }

            try
            {
                // Send As
                var saBody = new
                {
                    CmdletInput = new
                    {
                        CmdletName = "Get-RecipientPermission",
                        Parameters = new Dictionary<string, object>
                        {
                            ["Identity"] = mbxEmail,
                            ["Trustee"]  = userEmail
                        }
                    }
                };

                var saResponse = await httpClient.PostAsync(requestUrl,
                    new StringContent(JsonSerializer.Serialize(saBody), Encoding.UTF8, "application/json"));

                if (saResponse.IsSuccessStatusCode)
                {
                    var saJson = await saResponse.Content.ReadAsStringAsync();
                    var saDoc  = JsonDocument.Parse(saJson);
                    var grants = saDoc.RootElement.GetProperty("value");

                    foreach (var grant in grants.EnumerateArray())
                    {
                        var rights = grant.TryGetProperty("AccessRights", out var ar)
                            ? string.Join(", ", ar.EnumerateArray().Select(x => x.GetString() ?? ""))
                            : "";

                        if (rights.Contains("SendAs", StringComparison.OrdinalIgnoreCase))
                        {
                            sendAsMailboxes.Add(new MailboxAccessEntryDto(
                                MailboxEmail:       mbxEmail,
                                MailboxDisplayName: mailbox.DisplayName,
                                MailboxType:        mailbox.RecipientTypeDetails ?? "UserMailbox",
                                Permission:         "Send As",
                                GrantedTo:          userEmail,
                                IsInherited:        false
                            ));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogDebug(ex, "Send As check failed for {Mailbox}", mbxEmail);
            }

            // Send on Behalf — stored on the mailbox itself
            try
            {
                if (mailbox.GrantSendOnBehalfTo != null &&
                    mailbox.GrantSendOnBehalfTo.Any(g => g.Contains(userEmail.Split('@')[0], StringComparison.OrdinalIgnoreCase)))
                {
                    sendOnBehalfMailboxes.Add(new MailboxAccessEntryDto(
                        MailboxEmail:       mbxEmail,
                        MailboxDisplayName: mailbox.DisplayName,
                        MailboxType:        mailbox.RecipientTypeDetails ?? "UserMailbox",
                        Permission:         "Send on Behalf",
                        GrantedTo:          userEmail,
                        IsInherited:        false
                    ));
                }
            }
            catch { /* best-effort */ }
        }

        var all = fullAccessMailboxes
            .Concat(sendAsMailboxes)
            .Concat(sendOnBehalfMailboxes)
            .ToList();

        return new MailboxAccessResultDto(
            SubjectEmail:         userEmail,
            QueryType:            "AccessByUser",
            FullAccessMailboxes:  fullAccessMailboxes,
            SendAsMailboxes:      sendAsMailboxes,
            SendOnBehalfMailboxes: sendOnBehalfMailboxes,
            TotalCount:           all.Count,
            MailboxesChecked:     allMailboxes.Count,
            LastUpdated:          DateTime.UtcNow
        );
    }

    /// <summary>
    /// Returns all users/groups that have been granted access to the given mailbox.
    /// </summary>
    public async Task<MailboxAccessResultDto> GetMailboxDelegatesAsync(string mailboxEmail)
    {
        var tenantId = _configuration["AzureAd:TenantId"]!;
        var token = await GetExchangeTokenAsync();

        using var httpClient = _httpClientFactory.CreateClient();
        httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);

        var requestUrl = $"https://outlook.office365.com/adminapi/beta/{tenantId}/InvokeCommand";

        var fullAccessEntries  = new List<MailboxAccessEntryDto>();
        var sendAsEntries      = new List<MailboxAccessEntryDto>();
        var sendOnBehalfEntries = new List<MailboxAccessEntryDto>();

        // ── Full Access ────────────────────────────────────────────────────────
        try
        {
            var permDoc = await InvokeExchangeCommandAsync("Get-MailboxPermission",
                new Dictionary<string, object> { ["Identity"] = mailboxEmail });

            if (permDoc != null)
            {
                foreach (var perm in permDoc.RootElement.GetProperty("value").EnumerateArray())
                {
                    var user     = perm.TryGetProperty("User", out var u) ? u.GetString() ?? "" : "";
                    var isDenied = perm.TryGetProperty("Deny", out var deny) && deny.GetBoolean();
                    var isInherited = perm.TryGetProperty("IsInherited", out var inh) && inh.GetBoolean();
                    var rights   = perm.TryGetProperty("AccessRights", out var ar)
                        ? string.Join(", ", ar.EnumerateArray().Select(x => x.GetString() ?? ""))
                        : "";

                    // Skip system/inherited NT AUTHORITY entries and denies
                    if (isDenied) continue;
                    if (user.StartsWith("NT AUTHORITY", StringComparison.OrdinalIgnoreCase)) continue;
                    if (!rights.Contains("FullAccess", StringComparison.OrdinalIgnoreCase)) continue;

                    fullAccessEntries.Add(new MailboxAccessEntryDto(
                        MailboxEmail:       mailboxEmail,
                        MailboxDisplayName: mailboxEmail,
                        MailboxType:        "UserMailbox",
                        Permission:         "Full Access",
                        GrantedTo:          user,
                        IsInherited:        isInherited
                    ));
                }
            }
        }
        catch (Exception ex) { _logger.LogWarning(ex, "Get-MailboxPermission failed for {Mailbox}", mailboxEmail); }

        // ── Send As ────────────────────────────────────────────────────────────
        try
        {
            var saDoc = await InvokeExchangeCommandAsync("Get-RecipientPermission",
                new Dictionary<string, object> { ["Identity"] = mailboxEmail });

            if (saDoc != null)
            {
                foreach (var grant in saDoc.RootElement.GetProperty("value").EnumerateArray())
                {
                    var trustee = grant.TryGetProperty("Trustee", out var t) ? t.GetString() ?? "" : "";
                    var rights  = grant.TryGetProperty("AccessRights", out var ar)
                        ? string.Join(", ", ar.EnumerateArray().Select(x => x.GetString() ?? ""))
                        : "";

                    if (trustee.StartsWith("NT AUTHORITY", StringComparison.OrdinalIgnoreCase)) continue;
                    if (!rights.Contains("SendAs", StringComparison.OrdinalIgnoreCase)) continue;

                    sendAsEntries.Add(new MailboxAccessEntryDto(
                        MailboxEmail:       mailboxEmail,
                        MailboxDisplayName: mailboxEmail,
                        MailboxType:        "UserMailbox",
                        Permission:         "Send As",
                        GrantedTo:          trustee,
                        IsInherited:        false
                    ));
                }
            }
        }
        catch (Exception ex) { _logger.LogWarning(ex, "Get-RecipientPermission failed for {Mailbox}", mailboxEmail); }

        // ── Send on Behalf ─────────────────────────────────────────────────────
        try
        {
            var mbxDoc = await InvokeExchangeCommandAsync("Get-Mailbox",
                new Dictionary<string, object> { ["Identity"] = mailboxEmail });

            if (mbxDoc != null)
            {
                var mbxArr = mbxDoc.RootElement.GetProperty("value");
                if (mbxArr.GetArrayLength() > 0)
                {
                    var mbx = mbxArr[0];
                    if (mbx.TryGetProperty("GrantSendOnBehalfTo", out var sob))
                    {
                        foreach (var entry in sob.EnumerateArray())
                        {
                            var delegate_ = entry.GetString() ?? "";
                            if (string.IsNullOrEmpty(delegate_)) continue;

                            sendOnBehalfEntries.Add(new MailboxAccessEntryDto(
                                MailboxEmail:       mailboxEmail,
                                MailboxDisplayName: mailboxEmail,
                                MailboxType:        "UserMailbox",
                                Permission:         "Send on Behalf",
                                GrantedTo:          delegate_,
                                IsInherited:        false
                            ));
                        }
                    }
                }
            }
        }
        catch (Exception ex) { _logger.LogWarning(ex, "Get-Mailbox GrantSendOnBehalfTo failed for {Mailbox}", mailboxEmail); }

        var all = fullAccessEntries.Concat(sendAsEntries).Concat(sendOnBehalfEntries).ToList();

        return new MailboxAccessResultDto(
            SubjectEmail:          mailboxEmail,
            QueryType:             "DelegatesOnMailbox",
            FullAccessMailboxes:   fullAccessEntries,
            SendAsMailboxes:       sendAsEntries,
            SendOnBehalfMailboxes: sendOnBehalfEntries,
            TotalCount:            all.Count,
            MailboxesChecked:      1,
            LastUpdated:           DateTime.UtcNow
        );
    }

    // Helper: fetch all mailboxes (used by GetMailboxAccessForUserAsync)
    private record MailboxSummary(
        string? DisplayName,
        string? PrimarySmtpAddress,
        string? RecipientTypeDetails,
        List<string>? GrantSendOnBehalfTo);

    private async Task<List<MailboxSummary>> GetAllMailboxesAsync(HttpClient httpClient, string requestUrl)
    {
        var mailboxes = new List<MailboxSummary>();
        var page = 0;
        const int pageSize = 100;
        string? skipToken = null;

        while (true)
        {
            var parameters = new Dictionary<string, object>
            {
                ["ResultSize"] = pageSize.ToString()
            };
            if (skipToken != null)
                parameters["SkipToken"] = skipToken;

            var body = new
            {
                CmdletInput = new
                {
                    CmdletName = "Get-Mailbox",
                    Parameters = parameters
                }
            };

            var response = await httpClient.PostAsync(requestUrl,
                new StringContent(JsonSerializer.Serialize(body), Encoding.UTF8, "application/json"));

            if (!response.IsSuccessStatusCode) break;

            var json = await response.Content.ReadAsStringAsync();
            var doc  = JsonDocument.Parse(json);
            var arr  = doc.RootElement.GetProperty("value");

            foreach (var mbx in arr.EnumerateArray())
            {
                var sob = new List<string>();
                if (mbx.TryGetProperty("GrantSendOnBehalfTo", out var sobArr))
                    sob = sobArr.EnumerateArray().Select(x => x.GetString() ?? "").ToList();

                mailboxes.Add(new MailboxSummary(
                    DisplayName:          mbx.TryGetProperty("DisplayName",         out var dn) ? dn.GetString() : null,
                    PrimarySmtpAddress:   mbx.TryGetProperty("PrimarySmtpAddress",  out var ps) ? ps.GetString() : null,
                    RecipientTypeDetails: mbx.TryGetProperty("RecipientTypeDetails", out var rt) ? rt.GetString() : null,
                    GrantSendOnBehalfTo:  sob
                ));
            }

            // Check for OdataNextLink / SkipToken for pagination
            if (doc.RootElement.TryGetProperty("@odata.nextLink", out var next))
            {
                var nextVal = next.GetString();
                if (!string.IsNullOrEmpty(nextVal))
                {
                    // Extract skiptoken parameter
                    var uri = new Uri(nextVal);
                    var qs  = System.Web.HttpUtility.ParseQueryString(uri.Query);
                    skipToken = qs["$skiptoken"] ?? qs["skipToken"];
                    if (skipToken == null) break;
                    page++;
                    if (page > 20) break; // cap at 2000 mailboxes
                    continue;
                }
            }

            break; // No more pages
        }

        return mailboxes;
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

// DTOs for Mailbox Access / Delegation
public record MailboxAccessEntryDto(
    string MailboxEmail,
    string? MailboxDisplayName,
    string MailboxType,
    string Permission,       // "Full Access" | "Send As" | "Send on Behalf"
    string GrantedTo,        // UPN or display name of the delegate
    bool IsInherited
);

public record MailboxAccessResultDto(
    string SubjectEmail,
    string QueryType,        // "AccessByUser" | "DelegatesOnMailbox"
    List<MailboxAccessEntryDto> FullAccessMailboxes,
    List<MailboxAccessEntryDto> SendAsMailboxes,
    List<MailboxAccessEntryDto> SendOnBehalfMailboxes,
    int TotalCount,
    int MailboxesChecked,
    DateTime LastUpdated
);
