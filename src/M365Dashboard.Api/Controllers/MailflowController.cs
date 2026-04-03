using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Identity.Web;
using System.Net.Http.Headers;
using System.Text.Json;

namespace M365Dashboard.Api.Controllers;

[ApiController]
[Route("api/[controller]")]
[Authorize]
public class MailFlowController : ControllerBase
{
    private readonly GraphServiceClient _graphClient;
    private readonly ITokenAcquisition _tokenAcquisition;
    private readonly IHttpClientFactory _httpClientFactory;
    private readonly ILogger<MailFlowController> _logger;

    public MailFlowController(
        GraphServiceClient graphClient, 
        ITokenAcquisition tokenAcquisition,
        IHttpClientFactory httpClientFactory,
        ILogger<MailFlowController> logger)
    {
        _graphClient = graphClient;
        _tokenAcquisition = tokenAcquisition;
        _httpClientFactory = httpClientFactory;
        _logger = logger;
    }

    /// <summary>
    /// Get combined mail flow summary with traffic data and top senders/recipients
    /// </summary>
    [HttpGet("summary")]
    public async Task<IActionResult> GetMailflowSummary([FromQuery] int days = 30)
    {
        try
        {
            // Validate days parameter - Graph API only supports D7, D30, D90, D180
            var period = days <= 7 ? "D7" : days <= 30 ? "D30" : days <= 90 ? "D90" : "D180";

            // Get email activity counts - this provides actual message counts per day
            var activityResponse = await _graphClient.Reports
                .GetEmailActivityCountsWithPeriod(period)
                .GetAsync();

            var dailyTraffic = new List<object>();
            int totalSent = 0, totalReceived = 0, totalRead = 0;

            if (activityResponse != null)
            {
                using var reader = new StreamReader(activityResponse);
                var csvContent = await reader.ReadToEndAsync();
                _logger.LogInformation("Email Activity Counts CSV (first 1500 chars): {Content}", 
                    csvContent.Length > 1500 ? csvContent.Substring(0, 1500) : csvContent);
                
                var lines = csvContent.Split('\n', StringSplitOptions.RemoveEmptyEntries);

                if (lines.Length > 1)
                {
                    var headers = lines[0].Split(',').Select(h => h.Trim('"').Trim()).ToArray();
                    _logger.LogInformation("CSV Headers: {Headers}", string.Join(", ", headers));
                    
                    // Find column indices dynamically
                    var dateIndex = Array.FindIndex(headers, h => h.Equals("Report Date", StringComparison.OrdinalIgnoreCase));
                    var sendIndex = Array.FindIndex(headers, h => h.Equals("Send", StringComparison.OrdinalIgnoreCase));
                    var receiveIndex = Array.FindIndex(headers, h => h.Equals("Receive", StringComparison.OrdinalIgnoreCase));
                    var readIndex = Array.FindIndex(headers, h => h.Equals("Read", StringComparison.OrdinalIgnoreCase));
                    var reportRefreshIndex = Array.FindIndex(headers, h => h.Equals("Report Refresh Date", StringComparison.OrdinalIgnoreCase));

                    _logger.LogInformation("Column indices - Date: {DateIdx}, Send: {SendIdx}, Receive: {ReceiveIdx}, Read: {ReadIdx}", 
                        dateIndex, sendIndex, receiveIndex, readIndex);

                    // Create a dictionary to aggregate by date (in case of duplicates)
                    var trafficByDate = new Dictionary<string, (int sent, int received, int read)>();

                    for (int i = 1; i < lines.Length; i++)
                    {
                        var values = ParseCsvLine(lines[i]);
                        
                        var dateStr = dateIndex >= 0 && dateIndex < values.Length ? values[dateIndex].Trim('"') : "";
                        var sent = sendIndex >= 0 && sendIndex < values.Length && int.TryParse(values[sendIndex].Trim('"'), out var s) ? s : 0;
                        var received = receiveIndex >= 0 && receiveIndex < values.Length && int.TryParse(values[receiveIndex].Trim('"'), out var r) ? r : 0;
                        var read = readIndex >= 0 && readIndex < values.Length && int.TryParse(values[readIndex].Trim('"'), out var rd) ? rd : 0;

                        if (!string.IsNullOrEmpty(dateStr))
                        {
                            if (trafficByDate.ContainsKey(dateStr))
                            {
                                var existing = trafficByDate[dateStr];
                                trafficByDate[dateStr] = (existing.sent + sent, existing.received + received, existing.read + read);
                            }
                            else
                            {
                                trafficByDate[dateStr] = (sent, received, read);
                            }
                        }
                    }

                    // Convert to list and calculate totals
                    foreach (var kvp in trafficByDate)
                    {
                        totalSent += kvp.Value.sent;
                        totalReceived += kvp.Value.received;
                        totalRead += kvp.Value.read;

                        dailyTraffic.Add(new
                        {
                            date = kvp.Key,
                            messagesSent = kvp.Value.sent,
                            messagesReceived = kvp.Value.received,
                            messagesRead = kvp.Value.read,
                            spamReceived = 0,
                            malwareReceived = 0,
                            goodMail = kvp.Value.received
                        });
                    }
                }
            }

            _logger.LogInformation("Parsed {Count} daily traffic records, Total Sent: {Sent}, Total Received: {Received}, Total Read: {Read}", 
                dailyTraffic.Count, totalSent, totalReceived, totalRead);

            // Get user activity for top senders/recipients
            var userActivityResponse = await _graphClient.Reports
                .GetEmailActivityUserDetailWithPeriod(period)
                .GetAsync();

            var topSenders = new List<object>();
            var topRecipients = new List<object>();

            if (userActivityResponse != null)
            {
                using var reader = new StreamReader(userActivityResponse);
                var csvContent = await reader.ReadToEndAsync();
                var lines = csvContent.Split('\n', StringSplitOptions.RemoveEmptyEntries);

                if (lines.Length > 1)
                {
                    var headers = lines[0].Split(',').Select(h => h.Trim('"').Trim()).ToArray();
                    var upnIndex = Array.FindIndex(headers, h => h.Contains("User Principal Name", StringComparison.OrdinalIgnoreCase));
                    var displayNameIndex = Array.FindIndex(headers, h => h.Contains("Display Name", StringComparison.OrdinalIgnoreCase));
                    var sendIndex = Array.FindIndex(headers, h => h.Contains("Send Count", StringComparison.OrdinalIgnoreCase));
                    var receiveIndex = Array.FindIndex(headers, h => h.Contains("Receive Count", StringComparison.OrdinalIgnoreCase));

                    var users = new List<(string upn, string displayName, int sent, int received)>();

                    for (int i = 1; i < lines.Length; i++)
                    {
                        var values = ParseCsvLine(lines[i]);
                        if (values.Length > Math.Max(sendIndex, receiveIndex) && sendIndex >= 0 && receiveIndex >= 0)
                        {
                            var upn = upnIndex >= 0 && upnIndex < values.Length ? values[upnIndex].Trim('"') : "";
                            var displayName = displayNameIndex >= 0 && displayNameIndex < values.Length ? values[displayNameIndex].Trim('"') : "";
                            var sent = int.TryParse(values[sendIndex].Trim('"'), out var sc) ? sc : 0;
                            var received = int.TryParse(values[receiveIndex].Trim('"'), out var rc) ? rc : 0;

                            if (!string.IsNullOrEmpty(upn) && (sent > 0 || received > 0))
                            {
                                users.Add((upn, displayName, sent, received));
                            }
                        }
                    }

                    topSenders = users
                        .Where(u => u.sent > 0)
                        .OrderByDescending(u => u.sent)
                        .Take(10)
                        .Select(u => new { userPrincipalName = u.upn, displayName = u.displayName, messageCount = u.sent })
                        .Cast<object>()
                        .ToList();

                    topRecipients = users
                        .Where(u => u.received > 0)
                        .OrderByDescending(u => u.received)
                        .Take(10)
                        .Select(u => new { userPrincipalName = u.upn, displayName = u.displayName, messageCount = u.received })
                        .Cast<object>()
                        .ToList();
                }
            }

            // Sort daily traffic by date ascending for proper chart display
            var sortedTraffic = dailyTraffic
                .OrderBy(d => ((dynamic)d).date)
                .ToList();

            return Ok(new
            {
                totalMessagesSent = totalSent,
                totalMessagesReceived = totalReceived,
                totalMessagesRead = totalRead,
                totalSpamBlocked = 0,
                totalMalwareBlocked = 0,
                averageMessagesPerDay = dailyTraffic.Count > 0 ? (int)Math.Round((double)(totalSent + totalReceived) / dailyTraffic.Count) : 0,
                dailyTraffic = sortedTraffic,
                topSenders,
                topRecipients,
                lastUpdated = DateTime.UtcNow
            });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching mailflow summary");
            return Ok(new
            {
                totalMessagesSent = 0,
                totalMessagesReceived = 0,
                totalMessagesRead = 0,
                totalSpamBlocked = 0,
                totalMalwareBlocked = 0,
                averageMessagesPerDay = 0,
                dailyTraffic = new List<object>(),
                topSenders = new List<object>(),
                topRecipients = new List<object>(),
                lastUpdated = DateTime.UtcNow,
                error = ex.Message
            });
        }
    }

    /// <summary>
    /// Get mail traffic summary using reporting API
    /// </summary>
    [HttpGet("traffic-summary")]
    public async Task<IActionResult> GetMailTrafficSummary([FromQuery] int days = 7)
    {
        try
        {
            var fromDate = DateTime.UtcNow.AddDays(-days).Date;
            var toDate = DateTime.UtcNow.Date;

            // Get email activity counts
            var activityResponse = await _graphClient.Reports
                .GetEmailActivityCountsWithPeriod($"D{days}")
                .GetAsync();

            // Parse the CSV response
            var trafficData = new List<object>();
            if (activityResponse != null)
            {
                using var reader = new StreamReader(activityResponse);
                var csvContent = await reader.ReadToEndAsync();
                var lines = csvContent.Split('\n', StringSplitOptions.RemoveEmptyEntries);
                
                if (lines.Length > 1)
                {
                    var headers = lines[0].Split(',');
                    for (int i = 1; i < lines.Length; i++)
                    {
                        var values = lines[i].Split(',');
                        if (values.Length >= 5)
                        {
                            trafficData.Add(new
                            {
                                reportDate = values[1].Trim('"'),
                                send = int.TryParse(values[2], out var s) ? s : 0,
                                receive = int.TryParse(values[3], out var r) ? r : 0,
                                read = int.TryParse(values[4], out var rd) ? rd : 0
                            });
                        }
                    }
                }
            }

            // Calculate totals
            var totalSent = trafficData.Sum(d => ((dynamic)d).send);
            var totalReceived = trafficData.Sum(d => ((dynamic)d).receive);
            var totalRead = trafficData.Sum(d => ((dynamic)d).read);

            return Ok(new
            {
                dailyData = trafficData.OrderBy(d => ((dynamic)d).reportDate).ToList(),
                summary = new
                {
                    totalSent,
                    totalReceived,
                    totalRead,
                    avgDailySent = days > 0 ? Math.Round((double)totalSent / days, 0) : 0,
                    avgDailyReceived = days > 0 ? Math.Round((double)totalReceived / days, 0) : 0
                },
                period = new { fromDate, toDate },
                lastUpdated = DateTime.UtcNow
            });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching mail traffic summary");
            return Ok(new
            {
                dailyData = new List<object>(),
                summary = new { totalSent = 0, totalReceived = 0 },
                error = ex.Message
            });
        }
    }

    /// <summary>
    /// Get top mail senders
    /// </summary>
    [HttpGet("top-senders")]
    public async Task<IActionResult> GetTopSenders([FromQuery] int days = 7)
    {
        try
        {
            var activityResponse = await _graphClient.Reports
                .GetEmailActivityUserDetailWithPeriod($"D{days}")
                .GetAsync();

            var senders = new List<dynamic>();
            if (activityResponse != null)
            {
                using var reader = new StreamReader(activityResponse);
                var csvContent = await reader.ReadToEndAsync();
                var lines = csvContent.Split('\n', StringSplitOptions.RemoveEmptyEntries);

                if (lines.Length > 1)
                {
                    var headers = lines[0].Split(',');
                    var upnIndex = Array.IndexOf(headers, "\"User Principal Name\"");
                    var displayNameIndex = Array.IndexOf(headers, "\"Display Name\"");
                    var sendIndex = Array.IndexOf(headers, "\"Send Count\"");
                    var receiveIndex = Array.IndexOf(headers, "\"Receive Count\"");
                    var readIndex = Array.IndexOf(headers, "\"Read Count\"");

                    for (int i = 1; i < lines.Length; i++)
                    {
                        var values = lines[i].Split(',');
                        if (values.Length > Math.Max(Math.Max(sendIndex, receiveIndex), readIndex))
                        {
                            senders.Add(new
                            {
                                userPrincipalName = upnIndex >= 0 ? values[upnIndex].Trim('"') : "",
                                displayName = displayNameIndex >= 0 ? values[displayNameIndex].Trim('"') : "",
                                sendCount = sendIndex >= 0 && int.TryParse(values[sendIndex], out var s) ? s : 0,
                                receiveCount = receiveIndex >= 0 && int.TryParse(values[receiveIndex], out var r) ? r : 0,
                                readCount = readIndex >= 0 && int.TryParse(values[readIndex], out var rd) ? rd : 0
                            });
                        }
                    }
                }
            }

            return Ok(new
            {
                topSenders = senders.OrderByDescending(s => s.sendCount).Take(20).ToList(),
                topReceivers = senders.OrderByDescending(s => s.receiveCount).Take(20).ToList(),
                totalUsers = senders.Count,
                lastUpdated = DateTime.UtcNow
            });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching top senders");
            return Ok(new
            {
                topSenders = new List<object>(),
                error = ex.Message
            });
        }
    }

    /// <summary>
    /// Get transport rules (mail flow rules)
    /// </summary>
    [HttpGet("transport-rules")]
    public IActionResult GetTransportRules()
    {
        try
        {
            // Note: Transport rules require Exchange Online PowerShell or EWS
            // This is a placeholder that would need Exchange Online Management module
            // For now, we'll return info about what would be available

            return Ok(new
            {
                rules = new List<object>(),
                message = "Transport rules require Exchange Online PowerShell connection. Consider using the Exchange Admin Center for detailed rule management.",
                alternatives = new[]
                {
                    "Use Exchange Admin Center (EAC) for transport rules",
                    "Use Exchange Online PowerShell: Get-TransportRule",
                    "Consider using Microsoft Graph beta endpoints when available"
                },
                lastUpdated = DateTime.UtcNow
            });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching transport rules");
            return Ok(new
            {
                rules = new List<object>(),
                error = ex.Message
            });
        }
    }

    /// <summary>
    /// Get mailbox forwarding configurations (inbox rules)
    /// </summary>
    [HttpGet("forwarding-configs")]
    public async Task<IActionResult> GetForwardingConfigs([FromQuery] int top = 100)
    {
        try
        {
            var users = await _graphClient.Users
                .GetAsync(config =>
                {
                    config.QueryParameters.Top = top;
                    config.QueryParameters.Select = new[] { "id", "userPrincipalName", "displayName", "mail" };
                    config.QueryParameters.Filter = "accountEnabled eq true";
                });

            // Filter to users with mail addresses
            var userList = (users?.Value ?? new List<User>())
                .Where(u => !string.IsNullOrEmpty(u.Mail))
                .ToList();
            var forwardingUsers = new List<object>();

            foreach (var user in userList)
            {
                try
                {
                    // Check mailbox settings for forwarding
                    var mailboxSettings = await _graphClient.Users[user.Id].MailboxSettings
                        .GetAsync();

                    // Check automatic replies (which can contain forwarding)
                    var automaticReplies = mailboxSettings?.AutomaticRepliesSetting;
                    
                    // Check for any inbox rules with forwarding
                    var rules = await _graphClient.Users[user.Id].MailFolders["inbox"].MessageRules
                        .GetAsync();

                    var forwardingRules = rules?.Value?
                        .Where(r => r.Actions?.ForwardTo?.Any() == true || 
                                    r.Actions?.RedirectTo?.Any() == true ||
                                    r.Actions?.ForwardAsAttachmentTo?.Any() == true)
                        .Select(r => new
                        {
                            ruleId = r.Id,
                            ruleName = r.DisplayName,
                            isEnabled = r.IsEnabled,
                            forwardTo = r.Actions?.ForwardTo?.Select(a => a.EmailAddress?.Address).ToList(),
                            redirectTo = r.Actions?.RedirectTo?.Select(a => a.EmailAddress?.Address).ToList(),
                            forwardAsAttachmentTo = r.Actions?.ForwardAsAttachmentTo?.Select(a => a.EmailAddress?.Address).ToList()
                        })
                        .ToList();

                    if (forwardingRules?.Any() == true || automaticReplies?.Status != AutomaticRepliesStatus.Disabled)
                    {
                        forwardingUsers.Add(new
                        {
                            userId = user.Id,
                            userPrincipalName = user.UserPrincipalName,
                            displayName = user.DisplayName,
                            mail = user.Mail,
                            autoReplyStatus = automaticReplies?.Status?.ToString(),
                            forwardingRules
                        });
                    }
                }
                catch
                {
                    // Skip users where we can't access settings
                }
            }

            return Ok(new
            {
                users = forwardingUsers,
                summary = new
                {
                    usersScanned = userList.Count,
                    usersWithForwarding = forwardingUsers.Count
                },
                lastUpdated = DateTime.UtcNow
            });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching forwarding configs");
            return Ok(new
            {
                users = new List<object>(),
                error = ex.Message
            });
        }
    }

    /// <summary>
    /// Get comprehensive mailbox forwarding report including inbox rules
    /// </summary>
    [HttpGet("forwarding-report")]
    public async Task<IActionResult> GetForwardingReport([FromQuery] int top = 500)
    {
        try
        {
            _logger.LogInformation("Starting forwarding report scan for up to {Top} users", top);
            
            // Get all users with mailboxes
            var users = await _graphClient.Users
                .GetAsync(config =>
                {
                    config.QueryParameters.Top = top;
                    config.QueryParameters.Select = new[] { "id", "userPrincipalName", "displayName", "mail", "department", "jobTitle" };
                    config.QueryParameters.Filter = "accountEnabled eq true";
                });

            // Filter to users with mail addresses
            var userList = (users?.Value ?? new List<User>())
                .Where(u => !string.IsNullOrEmpty(u.Mail))
                .ToList();
            var forwardingMailboxes = new List<object>();
            var tenantDomains = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            int usersScanned = 0;
            int usersWithInboxRules = 0;
            int internalForwardingCount = 0;
            int externalForwardingCount = 0;

            // Get verified tenant domains from the organization
            try
            {
                var domains = await _graphClient.Domains.GetAsync();
                if (domains?.Value != null)
                {
                    foreach (var domain in domains.Value)
                    {
                        if (domain.Id != null && domain.IsVerified == true)
                        {
                            tenantDomains.Add(domain.Id);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning("Could not fetch tenant domains: {Error}", ex.Message);
            }

            _logger.LogInformation("Found {DomainCount} tenant domains: {Domains}", 
                tenantDomains.Count, string.Join(", ", tenantDomains));

            foreach (var user in userList)
            {
                usersScanned++;
                
                try
                {
                    // Check for inbox rules with forwarding actions
                    var rules = await _graphClient.Users[user.Id].MailFolders["inbox"].MessageRules
                        .GetAsync();

                    var forwardingRules = rules?.Value?
                        .Where(r => r.Actions?.ForwardTo?.Any() == true || 
                                    r.Actions?.RedirectTo?.Any() == true ||
                                    r.Actions?.ForwardAsAttachmentTo?.Any() == true)
                        .ToList();

                    if (forwardingRules?.Any() == true)
                    {
                        usersWithInboxRules++;
                        
                        foreach (var rule in forwardingRules)
                        {
                            // Collect all forwarding addresses from this rule
                            var allForwardAddresses = new List<string>();
                            
                            if (rule.Actions?.ForwardTo != null)
                                allForwardAddresses.AddRange(rule.Actions.ForwardTo.Select(a => a.EmailAddress?.Address ?? "").Where(a => !string.IsNullOrEmpty(a)));
                            
                            if (rule.Actions?.RedirectTo != null)
                                allForwardAddresses.AddRange(rule.Actions.RedirectTo.Select(a => a.EmailAddress?.Address ?? "").Where(a => !string.IsNullOrEmpty(a)));
                            
                            if (rule.Actions?.ForwardAsAttachmentTo != null)
                                allForwardAddresses.AddRange(rule.Actions.ForwardAsAttachmentTo.Select(a => a.EmailAddress?.Address ?? "").Where(a => !string.IsNullOrEmpty(a)));

                            foreach (var forwardAddress in allForwardAddresses.Distinct())
                            {
                                var isExternal = true;
                                var forwardDomain = forwardAddress.Contains('@') ? forwardAddress.Split('@')[1] : "";
                                
                                if (!string.IsNullOrEmpty(forwardDomain) && tenantDomains.Contains(forwardDomain))
                                {
                                    isExternal = false;
                                    internalForwardingCount++;
                                }
                                else
                                {
                                    externalForwardingCount++;
                                }

                                var forwardType = "Forward";
                                if (rule.Actions?.RedirectTo?.Any(a => a.EmailAddress?.Address == forwardAddress) == true)
                                    forwardType = "Redirect";
                                else if (rule.Actions?.ForwardAsAttachmentTo?.Any(a => a.EmailAddress?.Address == forwardAddress) == true)
                                    forwardType = "ForwardAsAttachment";

                                forwardingMailboxes.Add(new
                                {
                                    userId = user.Id,
                                    userPrincipalName = user.UserPrincipalName,
                                    displayName = user.DisplayName ?? user.UserPrincipalName,
                                    mail = user.Mail,
                                    department = user.Department,
                                    jobTitle = user.JobTitle,
                                    forwardingType = forwardType,
                                    forwardingTarget = forwardAddress,
                                    forwardingTargetDomain = forwardDomain,
                                    isExternal = isExternal,
                                    source = "InboxRule",
                                    ruleName = rule.DisplayName ?? "Unnamed Rule",
                                    ruleEnabled = rule.IsEnabled ?? false,
                                    deliverToMailbox = forwardType != "Redirect", // Redirect doesn't keep a copy
                                    riskLevel = isExternal ? "High" : "Low"
                                });
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogWarning("Could not check forwarding for user {UPN}: {Error}", 
                        user.UserPrincipalName, ex.Message);
                }
            }

            // Group by user for summary
            var uniqueUsersWithForwarding = forwardingMailboxes
                .GroupBy(m => ((dynamic)m).userPrincipalName)
                .Count();

            var uniqueExternalTargets = forwardingMailboxes
                .Where(m => ((dynamic)m).isExternal == true)
                .Select(m => ((dynamic)m).forwardingTarget)
                .Distinct()
                .Count();

            _logger.LogInformation("Forwarding report complete. Scanned: {Scanned}, With rules: {WithRules}, External: {External}, Internal: {Internal}",
                usersScanned, usersWithInboxRules, externalForwardingCount, internalForwardingCount);

            return Ok(new
            {
                forwardingMailboxes = forwardingMailboxes
                    .OrderByDescending(m => ((dynamic)m).isExternal)
                    .ThenBy(m => ((dynamic)m).displayName)
                    .ToList(),
                summary = new
                {
                    totalMailboxesScanned = usersScanned,
                    mailboxesWithForwarding = uniqueUsersWithForwarding,
                    totalForwardingRules = forwardingMailboxes.Count,
                    externalForwardingCount,
                    internalForwardingCount,
                    uniqueExternalTargets,
                    percentWithForwarding = usersScanned > 0 
                        ? Math.Round((double)uniqueUsersWithForwarding / usersScanned * 100, 1) 
                        : 0
                },
                tenantDomains = tenantDomains.ToList(),
                note = "This report scans inbox rules via Graph API (requires Mail.Read application permission for all users). For Exchange-level forwarding, see the 'Mailbox Fwd' tab.",
                exchangeAdminUrl = "https://admin.exchange.microsoft.com/#/mailboxes",
                lastUpdated = DateTime.UtcNow
            });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error generating forwarding report");
            return Ok(new
            {
                forwardingMailboxes = new List<object>(),
                summary = new
                {
                    totalMailboxesScanned = 0,
                    mailboxesWithForwarding = 0,
                    totalForwardingRules = 0,
                    externalForwardingCount = 0,
                    internalForwardingCount = 0
                },
                error = ex.Message,
                lastUpdated = DateTime.UtcNow
            });
        }
    }

    /// <summary>
    /// Get message trace summary (requires Exchange Online permissions)
    /// </summary>
    [HttpGet("message-trace-summary")]
    public IActionResult GetMessageTraceSummary([FromQuery] int hours = 24)
    {
        try
        {
            // Message trace requires Exchange Online Management or EWS
            // Graph API has limited message trace capabilities
            // We'll provide guidance and use available Graph data

            return Ok(new
            {
                message = "Detailed message trace requires Exchange Online Management Shell or the Security & Compliance Center.",
                graphAlternatives = new[]
                {
                    "Use mail activity reports for aggregate data",
                    "Check user mailbox rules for forwarding patterns",
                    "Use Security & Compliance Center for detailed traces"
                },
                securityCenterUrl = "https://security.microsoft.com/messagetrace",
                exchangeAdminUrl = "https://admin.exchange.microsoft.com",
                lastUpdated = DateTime.UtcNow
            });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error in message trace summary");
            return Ok(new
            {
                error = ex.Message
            });
        }
    }

    /// <summary>
    /// Get spam/phishing report summary
    /// </summary>
    [HttpGet("threat-summary")]
    public async Task<IActionResult> GetThreatSummary([FromQuery] int days = 7)
    {
        try
        {
            // This would ideally use the Office 365 Management Activity API
            // or the Security & Compliance Center API
            // Graph has limited direct access to threat protection data

            // Try to get what we can from security alerts
            var fromDate = DateTime.UtcNow.AddDays(-days);

            var alerts = await _graphClient.Security.Alerts_v2
                .GetAsync(config =>
                {
                    config.QueryParameters.Filter = $"createdDateTime ge {fromDate:yyyy-MM-ddTHH:mm:ssZ}";
                    config.QueryParameters.Top = 100;
                });

            var alertList = alerts?.Value ?? new List<Microsoft.Graph.Models.Security.Alert>();

            var mailAlerts = alertList
                .Where(a => a.ServiceSource?.ToString()?.Contains("Office", StringComparison.OrdinalIgnoreCase) == true ||
                            a.Category?.Contains("Email", StringComparison.OrdinalIgnoreCase) == true ||
                            a.Category?.Contains("Phish", StringComparison.OrdinalIgnoreCase) == true)
                .Select(a => new
                {
                    id = a.Id,
                    title = a.Title,
                    severity = a.Severity?.ToString(),
                    status = a.Status?.ToString(),
                    category = a.Category,
                    createdDateTime = a.CreatedDateTime,
                    description = a.Description
                })
                .ToList();

            return Ok(new
            {
                mailRelatedAlerts = mailAlerts,
                summary = new
                {
                    totalMailAlerts = mailAlerts.Count,
                    highSeverity = mailAlerts.Count(a => a.severity == "High"),
                    mediumSeverity = mailAlerts.Count(a => a.severity == "Medium")
                },
                additionalResources = new
                {
                    threatExplorer = "https://security.microsoft.com/threatexplorer",
                    quarantine = "https://security.microsoft.com/quarantine",
                    attackSimulator = "https://security.microsoft.com/attacksimulator"
                },
                period = new { fromDate, toDate = DateTime.UtcNow },
                lastUpdated = DateTime.UtcNow
            });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching threat summary");
            return Ok(new
            {
                mailRelatedAlerts = new List<object>(),
                error = ex.Message
            });
        }
    }

    /// <summary>
    /// Get mailbox storage usage report - shows mailboxes approaching quota
    /// </summary>
    [HttpGet("storage-usage")]
    public async Task<IActionResult> GetMailboxStorageUsage([FromQuery] int thresholdPercent = 80)
    {
        try
        {
            // Get mailbox usage detail report
            var usageResponse = await _graphClient.Reports
                .GetMailboxUsageDetailWithPeriod("D7")
                .GetAsync();

            var mailboxes = new List<object>();
            
            if (usageResponse != null)
            {
                using var reader = new StreamReader(usageResponse);
                var csvContent = await reader.ReadToEndAsync();
                var lines = csvContent.Split('\n', StringSplitOptions.RemoveEmptyEntries);

                if (lines.Length > 1)
                {
                    var headers = lines[0].Split(',').Select(h => h.Trim('"')).ToArray();
                    
                    // Find column indices
                    var upnIndex = Array.FindIndex(headers, h => h.Contains("User Principal Name", StringComparison.OrdinalIgnoreCase));
                    var displayNameIndex = Array.FindIndex(headers, h => h.Contains("Display Name", StringComparison.OrdinalIgnoreCase));
                    var storageUsedIndex = Array.FindIndex(headers, h => h.Contains("Storage Used", StringComparison.OrdinalIgnoreCase) && !h.Contains("Quota"));
                    var issueWarningQuotaIndex = Array.FindIndex(headers, h => h.Contains("Issue Warning Quota", StringComparison.OrdinalIgnoreCase));
                    var prohibitSendQuotaIndex = Array.FindIndex(headers, h => h.Contains("Prohibit Send Quota", StringComparison.OrdinalIgnoreCase));
                    var prohibitSendReceiveQuotaIndex = Array.FindIndex(headers, h => h.Contains("Prohibit Send/Receive Quota", StringComparison.OrdinalIgnoreCase));
                    var deletedItemSizeIndex = Array.FindIndex(headers, h => h.Contains("Deleted Item Size", StringComparison.OrdinalIgnoreCase));
                    var itemCountIndex = Array.FindIndex(headers, h => h.Contains("Item Count", StringComparison.OrdinalIgnoreCase));
                    var lastActivityIndex = Array.FindIndex(headers, h => h.Contains("Last Activity Date", StringComparison.OrdinalIgnoreCase));
                    var hasArchiveIndex = Array.FindIndex(headers, h => h.Contains("Has Archive", StringComparison.OrdinalIgnoreCase));
                    var isDeletedIndex = Array.FindIndex(headers, h => h.Contains("Is Deleted", StringComparison.OrdinalIgnoreCase));

                    for (int i = 1; i < lines.Length; i++)
                    {
                        var values = ParseCsvLine(lines[i]);
                        if (values.Length <= Math.Max(upnIndex, storageUsedIndex)) continue;

                        var storageUsedBytes = ParseBytes(storageUsedIndex >= 0 ? values[storageUsedIndex] : "0");
                        var prohibitSendQuotaBytes = ParseBytes(prohibitSendQuotaIndex >= 0 ? values[prohibitSendQuotaIndex] : "0");
                        var issueWarningQuotaBytes = ParseBytes(issueWarningQuotaIndex >= 0 ? values[issueWarningQuotaIndex] : "0");
                        
                        // Use prohibit send quota as the main quota, fallback to warning quota
                        var quotaBytes = prohibitSendQuotaBytes > 0 ? prohibitSendQuotaBytes : issueWarningQuotaBytes;
                        var percentUsed = quotaBytes > 0 ? Math.Round((double)storageUsedBytes / quotaBytes * 100, 1) : 0;

                        mailboxes.Add(new
                        {
                            userPrincipalName = upnIndex >= 0 ? values[upnIndex].Trim('"') : "",
                            displayName = displayNameIndex >= 0 ? values[displayNameIndex].Trim('"') : "",
                            storageUsedBytes,
                            storageUsedGB = Math.Round((double)storageUsedBytes / (1024 * 1024 * 1024), 2),
                            issueWarningQuotaBytes,
                            issueWarningQuotaGB = Math.Round((double)issueWarningQuotaBytes / (1024 * 1024 * 1024), 2),
                            prohibitSendQuotaBytes,
                            prohibitSendQuotaGB = Math.Round((double)prohibitSendQuotaBytes / (1024 * 1024 * 1024), 2),
                            percentUsed,
                            deletedItemSizeBytes = ParseBytes(deletedItemSizeIndex >= 0 ? values[deletedItemSizeIndex] : "0"),
                            itemCount = itemCountIndex >= 0 && long.TryParse(values[itemCountIndex].Trim('"'), out var ic) ? ic : 0,
                            lastActivityDate = lastActivityIndex >= 0 ? values[lastActivityIndex].Trim('"') : null,
                            hasArchive = hasArchiveIndex >= 0 && values[hasArchiveIndex].Trim('"').Equals("True", StringComparison.OrdinalIgnoreCase),
                            isDeleted = isDeletedIndex >= 0 && values[isDeletedIndex].Trim('"').Equals("True", StringComparison.OrdinalIgnoreCase),
                            isNearQuota = percentUsed >= thresholdPercent,
                            isOverWarning = issueWarningQuotaBytes > 0 && storageUsedBytes >= issueWarningQuotaBytes,
                            status = percentUsed >= 100 ? "Full" : percentUsed >= 90 ? "Critical" : percentUsed >= thresholdPercent ? "Warning" : "OK"
                        });
                    }
                }
            }

            // Sort by percent used descending and filter to those near quota
            var nearQuotaMailboxes = mailboxes
                .Where(m => ((dynamic)m).percentUsed >= thresholdPercent)
                .OrderByDescending(m => ((dynamic)m).percentUsed)
                .ToList();

            var allMailboxesSorted = mailboxes
                .OrderByDescending(m => ((dynamic)m).percentUsed)
                .Take(100)
                .ToList();

            // Summary stats
            var totalMailboxes = mailboxes.Count;
            var criticalCount = mailboxes.Count(m => ((dynamic)m).percentUsed >= 90);
            var warningCount = mailboxes.Count(m => ((dynamic)m).percentUsed >= thresholdPercent && ((dynamic)m).percentUsed < 90);
            var totalStorageUsedBytes = mailboxes.Sum(m => (long)((dynamic)m).storageUsedBytes);
            var archiveEnabledCount = mailboxes.Count(m => ((dynamic)m).hasArchive == true);

            return Ok(new
            {
                nearQuotaMailboxes,
                allMailboxes = allMailboxesSorted,
                summary = new
                {
                    totalMailboxes,
                    criticalCount,
                    warningCount,
                    okCount = totalMailboxes - criticalCount - warningCount,
                    totalStorageUsedGB = Math.Round((double)totalStorageUsedBytes / (1024 * 1024 * 1024), 2),
                    averagePercentUsed = totalMailboxes > 0 ? Math.Round(mailboxes.Average(m => (double)((dynamic)m).percentUsed), 1) : 0,
                    archiveEnabledCount
                },
                thresholdPercent,
                lastUpdated = DateTime.UtcNow
            });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching mailbox storage usage");
            return Ok(new
            {
                nearQuotaMailboxes = new List<object>(),
                allMailboxes = new List<object>(),
                summary = new { totalMailboxes = 0, criticalCount = 0, warningCount = 0 },
                error = ex.Message
            });
        }
    }

    /// <summary>
    /// Get list of mailboxes using the mailbox usage report (includes proper mailbox types)
    /// </summary>
    [HttpGet("mailboxes")]
    public async Task<IActionResult> GetMailboxes([FromQuery] int top = 500)
    {
        try
        {
            // Use mailbox usage detail report - this includes actual mailbox data
            var usageResponse = await _graphClient.Reports
                .GetMailboxUsageDetailWithPeriod("D7")
                .GetAsync();

            var mailboxList = new List<object>();
            
            if (usageResponse != null)
            {
                using var reader = new StreamReader(usageResponse);
                var csvContent = await reader.ReadToEndAsync();
                var lines = csvContent.Split('\n', StringSplitOptions.RemoveEmptyEntries);

                if (lines.Length > 1)
                {
                    var headers = lines[0].Split(',').Select(h => h.Trim('"').Trim()).ToArray();
                    
                    // Find column indices
                    var upnIndex = Array.FindIndex(headers, h => h.Equals("User Principal Name", StringComparison.OrdinalIgnoreCase));
                    var displayNameIndex = Array.FindIndex(headers, h => h.Equals("Display Name", StringComparison.OrdinalIgnoreCase));
                    var isDeletedIndex = Array.FindIndex(headers, h => h.Equals("Is Deleted", StringComparison.OrdinalIgnoreCase));
                    var deletedDateIndex = Array.FindIndex(headers, h => h.Equals("Deleted Date", StringComparison.OrdinalIgnoreCase));
                    var createdDateIndex = Array.FindIndex(headers, h => h.Equals("Created Date", StringComparison.OrdinalIgnoreCase));
                    var lastActivityIndex = Array.FindIndex(headers, h => h.Equals("Last Activity Date", StringComparison.OrdinalIgnoreCase));
                    var itemCountIndex = Array.FindIndex(headers, h => h.Equals("Item Count", StringComparison.OrdinalIgnoreCase));
                    var storageUsedIndex = Array.FindIndex(headers, h => h.Contains("Storage Used", StringComparison.OrdinalIgnoreCase) && !h.Contains("Quota"));
                    var issueWarningQuotaIndex = Array.FindIndex(headers, h => h.Contains("Issue Warning Quota", StringComparison.OrdinalIgnoreCase));
                    var prohibitSendQuotaIndex = Array.FindIndex(headers, h => h.Contains("Prohibit Send Quota", StringComparison.OrdinalIgnoreCase));
                    var hasArchiveIndex = Array.FindIndex(headers, h => h.Equals("Has Archive", StringComparison.OrdinalIgnoreCase));
                    var recipientTypeIndex = Array.FindIndex(headers, h => h.Equals("Recipient Type", StringComparison.OrdinalIgnoreCase));

                    _logger.LogInformation("Mailbox CSV Headers: {Headers}", string.Join(", ", headers));
                    _logger.LogInformation("Column indices - UPN: {UPN}, DisplayName: {DN}, RecipientType: {RT}", upnIndex, displayNameIndex, recipientTypeIndex);

                    for (int i = 1; i < lines.Length && mailboxList.Count < top; i++)
                    {
                        var values = ParseCsvLine(lines[i]);
                        if (values.Length <= upnIndex) continue;

                        var upn = upnIndex >= 0 && upnIndex < values.Length ? values[upnIndex].Trim('"') : "";
                        var displayName = displayNameIndex >= 0 && displayNameIndex < values.Length ? values[displayNameIndex].Trim('"') : "";
                        var isDeleted = isDeletedIndex >= 0 && isDeletedIndex < values.Length && values[isDeletedIndex].Trim('"').Equals("True", StringComparison.OrdinalIgnoreCase);
                        var createdDate = createdDateIndex >= 0 && createdDateIndex < values.Length ? values[createdDateIndex].Trim('"') : null;
                        var lastActivity = lastActivityIndex >= 0 && lastActivityIndex < values.Length ? values[lastActivityIndex].Trim('"') : null;
                        var hasArchive = hasArchiveIndex >= 0 && hasArchiveIndex < values.Length && values[hasArchiveIndex].Trim('"').Equals("True", StringComparison.OrdinalIgnoreCase);
                        var storageUsedBytes = ParseBytes(storageUsedIndex >= 0 && storageUsedIndex < values.Length ? values[storageUsedIndex] : "0");
                        var recipientType = recipientTypeIndex >= 0 && recipientTypeIndex < values.Length ? values[recipientTypeIndex].Trim('"') : "";

                        // Skip deleted mailboxes
                        if (isDeleted) continue;

                        // Skip if no UPN
                        if (string.IsNullOrEmpty(upn)) continue;

                        // Determine mailbox type from recipient type or infer from other data
                        var mailboxType = "UserMailbox";
                        var mailboxTypeDetails = "UserMailbox";
                        
                        if (!string.IsNullOrEmpty(recipientType))
                        {
                            mailboxType = recipientType;
                            mailboxTypeDetails = recipientType;
                        }

                        // Determine if mailbox is active (had activity in last 30 days)
                        var isActive = !string.IsNullOrEmpty(lastActivity) && 
                            DateTime.TryParse(lastActivity, out var lastActivityDate) && 
                            lastActivityDate > DateTime.UtcNow.AddDays(-30);

                        mailboxList.Add(new
                        {
                            id = upn,
                            displayName = displayName,
                            userPrincipalName = upn,
                            mail = upn,
                            recipientType = mailboxType,
                            recipientTypeDetails = mailboxTypeDetails,
                            whenCreated = createdDate,
                            isMailboxEnabled = isActive,
                            primarySmtpAddress = upn,
                            hasArchive = hasArchive,
                            storageUsedBytes = storageUsedBytes,
                            storageUsedGB = Math.Round((double)storageUsedBytes / (1024 * 1024 * 1024), 2),
                            lastActivityDate = lastActivity,
                            emailAddresses = new[] { upn }
                        });
                    }
                }
            }

            _logger.LogInformation("Returning {Count} mailboxes", mailboxList.Count);

            return Ok(new
            {
                mailboxes = mailboxList.OrderBy(m => ((dynamic)m).displayName).ToList(),
                totalCount = mailboxList.Count,
                filteredCount = mailboxList.Count
            });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching mailboxes");
            return Ok(new
            {
                mailboxes = new List<object>(),
                totalCount = 0,
                filteredCount = 0,
                error = ex.Message
            });
        }
    }

    /// <summary>
    /// Get mailbox statistics summary
    /// </summary>
    [HttpGet("stats")]
    public async Task<IActionResult> GetMailboxStats()
    {
        try
        {
            // Get mailbox usage report for statistics
            var usageResponse = await _graphClient.Reports
                .GetMailboxUsageMailboxCountsWithPeriod("D7")
                .GetAsync();

            int totalMailboxes = 0, activeMailboxes = 0, inactiveMailboxes = 0;

            if (usageResponse != null)
            {
                using var reader = new StreamReader(usageResponse);
                var csvContent = await reader.ReadToEndAsync();
                var lines = csvContent.Split('\n', StringSplitOptions.RemoveEmptyEntries);

                if (lines.Length > 1)
                {
                    var headers = lines[0].Split(',').Select(h => h.Trim('"').Trim()).ToArray();
                    var totalIndex = Array.FindIndex(headers, h => h.Equals("Total", StringComparison.OrdinalIgnoreCase));
                    var activeIndex = Array.FindIndex(headers, h => h.Equals("Active", StringComparison.OrdinalIgnoreCase));
                    var inactiveIndex = Array.FindIndex(headers, h => h.Equals("Inactive", StringComparison.OrdinalIgnoreCase));

                    // Get the most recent row (last data row)
                    var lastLine = lines[lines.Length - 1];
                    var values = ParseCsvLine(lastLine);

                    totalMailboxes = totalIndex >= 0 && totalIndex < values.Length && int.TryParse(values[totalIndex].Trim('"'), out var t) ? t : 0;
                    activeMailboxes = activeIndex >= 0 && activeIndex < values.Length && int.TryParse(values[activeIndex].Trim('"'), out var a) ? a : 0;
                    inactiveMailboxes = inactiveIndex >= 0 && inactiveIndex < values.Length && int.TryParse(values[inactiveIndex].Trim('"'), out var ia) ? ia : 0;
                }
            }

            // Get user mailbox count - count users with mail attribute set
            var users = await _graphClient.Users
                .GetAsync(config =>
                {
                    config.QueryParameters.Select = new[] { "id", "mail" };
                    config.QueryParameters.Top = 999;
                });

            var userMailboxes = (users?.Value ?? new List<User>())
                .Count(u => !string.IsNullOrEmpty(u.Mail));

            return Ok(new
            {
                totalMailboxes = totalMailboxes > 0 ? totalMailboxes : userMailboxes,
                userMailboxes,
                sharedMailboxes = 0, // Would need Exchange Online PowerShell to get this accurately
                roomMailboxes = 0,
                equipmentMailboxes = 0,
                activeMailboxes,
                inactiveMailboxes,
                totalStorageUsedBytes = 0L,
                mailboxesNearQuota = 0,
                mailboxesWithForwarding = 0,
                mailboxesOnHold = 0,
                mailboxesWithArchive = 0,
                lastUpdated = DateTime.UtcNow
            });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching mailbox stats");
            return Ok(new
            {
                totalMailboxes = 0,
                userMailboxes = 0,
                sharedMailboxes = 0,
                activeMailboxes = 0,
                inactiveMailboxes = 0,
                lastUpdated = DateTime.UtcNow,
                error = ex.Message
            });
        }
    }

    // Helper method to parse CSV lines that may contain quoted values with commas
    private static string[] ParseCsvLine(string line)
    {
        var values = new List<string>();
        var current = new System.Text.StringBuilder();
        var inQuotes = false;

        foreach (var c in line)
        {
            if (c == '"')
            {
                inQuotes = !inQuotes;
            }
            else if (c == ',' && !inQuotes)
            {
                values.Add(current.ToString());
                current.Clear();
            }
            else
            {
                current.Append(c);
            }
        }
        values.Add(current.ToString());

        return values.ToArray();
    }

    // Helper method to parse byte values from the report (handles formats like "50 GB", "1024 MB", etc.)
    private static long ParseBytes(string value)
    {
        if (string.IsNullOrWhiteSpace(value)) return 0;
        
        value = value.Trim().Trim('"');
        
        // Try parsing as plain number first
        if (long.TryParse(value, out var plainBytes))
            return plainBytes;

        // Parse values like "50 GB", "1024 MB", etc.
        var parts = value.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length >= 2 && double.TryParse(parts[0], out var num))
        {
            var unit = parts[1].ToUpperInvariant();
            return unit switch
            {
                "B" or "BYTE" or "BYTES" => (long)num,
                "KB" => (long)(num * 1024),
                "MB" => (long)(num * 1024 * 1024),
                "GB" => (long)(num * 1024 * 1024 * 1024),
                "TB" => (long)(num * 1024 * 1024 * 1024 * 1024),
                _ => (long)num
            };
        }

        return 0;
    }
}
