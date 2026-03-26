using Microsoft.Graph;
using Microsoft.Graph.Models;
using M365Dashboard.Api.Models.Dtos;

namespace M365Dashboard.Api.Services;

/// <summary>
/// Service for interacting with Microsoft Graph API using Application Permissions.
/// This allows the app to read tenant-wide data regardless of the signed-in user's permissions.
/// </summary>
public interface IGraphService
{
    Task<UserProfileDto> GetUserProfileAsync(string userId);
    Task<ActiveUsersDataDto> GetActiveUsersAsync(DateTime startDate, DateTime endDate);
    Task<SignInAnalyticsDto> GetSignInAnalyticsAsync(DateTime startDate, DateTime endDate);
    Task<LicenseUsageDto> GetLicenseUsageAsync();
    Task<DeviceComplianceDto> GetDeviceComplianceAsync();
    Task<MailActivityDto> GetMailActivityAsync(DateTime startDate, DateTime endDate);
    Task<TeamsActivityDto> GetTeamsActivityAsync(DateTime startDate, DateTime endDate);
    
    // User management methods
    Task<UserListResultDto> GetUsersAsync(string? filter, string? orderBy, bool ascending, int take);
    Task<UserDetailDto> GetUserDetailsAsync(string userId);
    Task<UserStatsDto> GetUserStatsAsync();
    
    // Group management methods
    Task<GroupListResultDto> GetGroupsAsync(string? filter, string? orderBy, bool ascending, int take);
    Task<GroupDetailDto> GetGroupDetailsAsync(string groupId);
    Task<GroupStatsDto> GetGroupStatsAsync();
    
    // Device management methods (Intune)
    Task<DeviceListResultDto> GetDevicesAsync(string? filter, string? orderBy, bool ascending, int take);
    Task<DeviceDetailDto> GetDeviceDetailsAsync(string deviceId);
    Task<DeviceStatsDto> GetDeviceStatsAsync();
    
    // Mailflow methods
    Task<MailboxListResultDto> GetMailboxesAsync(string? filter, string? orderBy, bool ascending, int take);
    Task<MailboxStatsDto> GetMailboxStatsAsync();
    Task<MailflowSummaryDto> GetMailflowSummaryAsync(int days);
    
    // Security methods
    Task<SecurityOverviewDto> GetSecurityOverviewAsync();
    Task<SecurityScoreDto?> GetSecureScoreAsync();
    Task<List<RiskyUserDto>> GetRiskyUsersAsync();
    Task<List<RiskySignInDto>> GetRiskySignInsAsync(int hours = 24);
    Task<SecurityStatsDto> GetSecurityStatsAsync();
    Task<MfaRegistrationListDto> GetMfaRegistrationDetailsAsync();
    
    // App Registration Credentials
    Task<AppCredentialStatusDto> GetAppCredentialStatusAsync(int thresholdDays = 45);
    
    // Public Groups Report
    Task<PublicGroupsReportDto> GetPublicGroupsAsync();
    
    // Distribution Lists
    Task<GroupListResultDto> GetDistributionListsAsync(int take = 200);
    
    // Stale Privileged Accounts Report
    Task<StalePrivilegedAccountsReportDto> GetStalePrivilegedAccountsAsync(int inactiveDaysThreshold = 30);
    
    // Conditional Access Break Glass Report
    Task<CABreakGlassReportDto> GetCABreakGlassReportAsync(List<string> breakGlassUpns);
    Task<BreakGlassAccountDto?> ResolveUserAsync(string userPrincipalName);
}

public class GraphService : IGraphService
{
    private readonly GraphServiceClient _graphClient;
    private readonly ILogger<GraphService> _logger;

    public GraphService(GraphServiceClient graphClient, ILogger<GraphService> logger)
    {
        _graphClient = graphClient;
        _logger = logger;
    }

    /// <summary>
    /// Get a user's profile by their ID (from the token's oid claim)
    /// </summary>
    public async Task<UserProfileDto> GetUserProfileAsync(string userId)
    {
        try
        {
            var user = await _graphClient.Users[userId].GetAsync(config =>
            {
                config.QueryParameters.Select = new[] { "id", "displayName", "mail", "userPrincipalName", "jobTitle", "department" };
            });

            if (user == null)
            {
                throw new InvalidOperationException("Failed to retrieve user profile");
            }

            // Get user's profile photo
            string? photoBase64 = null;
            try
            {
                var photoStream = await _graphClient.Users[userId].Photo.Content.GetAsync();
                if (photoStream != null)
                {
                    using var memoryStream = new MemoryStream();
                    await photoStream.CopyToAsync(memoryStream);
                    photoBase64 = $"data:image/jpeg;base64,{Convert.ToBase64String(memoryStream.ToArray())}";
                }
            }
            catch (Exception ex)
            {
                _logger.LogDebug(ex, "User {UserId} does not have a profile photo", userId);
            }

            return new UserProfileDto(
                Id: user.Id ?? string.Empty,
                DisplayName: user.DisplayName ?? "Unknown",
                Email: user.Mail ?? user.UserPrincipalName ?? string.Empty,
                JobTitle: user.JobTitle,
                Department: user.Department,
                ProfilePhoto: photoBase64,
                Roles: new List<string>() // Roles come from the token, not Graph
            );
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error retrieving user profile for {UserId}", userId);
            throw;
        }
    }

    public async Task<ActiveUsersDataDto> GetActiveUsersAsync(DateTime startDate, DateTime endDate)
    {
        try
        {
            var period = GetReportPeriod(startDate, endDate);
            _logger.LogInformation("Fetching active users report for period {Period}", period);
            
            // Using application permissions - no user context needed
            var response = await _graphClient.Reports
                .GetOffice365ActiveUserCountsWithPeriod(period)
                .GetAsync();

            var data = await ParseActiveUsersReportAsync(response);
            _logger.LogInformation("Retrieved {Count} active user records", data.Count);

            return new ActiveUsersDataDto(
                DailyActiveUsers: data.Count > 0 ? data.Last().Count : 0,
                WeeklyActiveUsers: data.TakeLast(7).Any() ? (int)data.TakeLast(7).Average(d => d.Count) : 0,
                MonthlyActiveUsers: data.TakeLast(30).Any() ? (int)data.TakeLast(30).Average(d => d.Count) : 0,
                Trend: data.Select(d => new DailyActiveUsersTrendDto(d.Date, d.Count)).ToList(),
                LastUpdated: DateTime.UtcNow
            );
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error retrieving active users data - returning mock data");
            return GetMockActiveUsersData();
        }
    }

    public async Task<SignInAnalyticsDto> GetSignInAnalyticsAsync(DateTime startDate, DateTime endDate)
    {
        try
        {
            // Query sign-in logs using application permissions
            var signIns = await _graphClient.AuditLogs.SignIns.GetAsync(config =>
            {
                config.QueryParameters.Filter = $"createdDateTime ge {startDate:yyyy-MM-ddTHH:mm:ssZ} and createdDateTime le {endDate:yyyy-MM-ddTHH:mm:ssZ}";
                config.QueryParameters.Top = 999;
                config.QueryParameters.Select = new[] { "createdDateTime", "status", "location", "riskState", "riskLevelDuringSignIn", "userPrincipalName" };
                config.QueryParameters.Orderby = new[] { "createdDateTime desc" };
            });

            var signInList = signIns?.Value ?? new List<SignIn>();

            var successful = signInList.Count(s => s.Status?.ErrorCode == 0);
            var failed = signInList.Count(s => s.Status?.ErrorCode != 0);
            var risky = signInList.Count(s => 
                s.RiskState == RiskState.AtRisk || 
                s.RiskLevelDuringSignIn == RiskLevel.High ||
                s.RiskLevelDuringSignIn == RiskLevel.Medium);

            var trend = signInList
                .Where(s => s.CreatedDateTime.HasValue)
                .GroupBy(s => s.CreatedDateTime!.Value.Date)
                .OrderBy(g => g.Key)
                .Select(g => new SignInTrendDto(
                    g.Key,
                    g.Count(s => s.Status?.ErrorCode == 0),
                    g.Count(s => s.Status?.ErrorCode != 0)
                ))
                .ToList();

            var topLocations = signInList
                .Where(s => s.Location?.City != null)
                .GroupBy(s => $"{s.Location?.City}, {s.Location?.CountryOrRegion}")
                .OrderByDescending(g => g.Count())
                .Take(5)
                .Select(g => new TopSignInLocationDto(g.Key, g.Count()))
                .ToList();

            return new SignInAnalyticsDto(
                TotalSignIns: signInList.Count,
                SuccessfulSignIns: successful,
                FailedSignIns: failed,
                RiskySignIns: risky,
                SuccessRate: signInList.Count > 0 ? Math.Round((double)successful / signInList.Count * 100, 1) : 0,
                Trend: trend,
                TopLocations: topLocations,
                LastUpdated: DateTime.UtcNow
            );
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error retrieving sign-in analytics");
            return GetMockSignInAnalyticsData();
        }
    }

    public async Task<LicenseUsageDto> GetLicenseUsageAsync()
    {
        try
        {
            var subscribedSkus = await _graphClient.SubscribedSkus.GetAsync();
            
            var licenses = subscribedSkus?.Value?
                .Where(sku => sku.CapabilityStatus == "Enabled")
                .Select(sku => new LicenseSkuDto(
                    SkuId: sku.SkuId?.ToString() ?? string.Empty,
                    SkuName: sku.SkuPartNumber ?? "Unknown",
                    ConsumedUnits: sku.ConsumedUnits ?? 0,
                    PrepaidUnits: sku.PrepaidUnits?.Enabled ?? 0,
                    UtilizationPercent: sku.PrepaidUnits?.Enabled > 0 
                        ? Math.Round((double)(sku.ConsumedUnits ?? 0) / sku.PrepaidUnits.Enabled.Value * 100, 1) 
                        : 0
                ))
                .OrderByDescending(l => l.ConsumedUnits)
                .ToList() ?? new List<LicenseSkuDto>();

            var totalConsumed = licenses.Sum(l => l.ConsumedUnits);
            var totalAvailable = licenses.Sum(l => l.PrepaidUnits);

            return new LicenseUsageDto(
                Licenses: licenses,
                TotalConsumed: totalConsumed,
                TotalAvailable: totalAvailable,
                OverallUtilization: totalAvailable > 0 ? Math.Round((double)totalConsumed / totalAvailable * 100, 1) : 0,
                LastUpdated: DateTime.UtcNow
            );
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error retrieving license usage");
            return GetMockLicenseUsageData();
        }
    }

    public async Task<DeviceComplianceDto> GetDeviceComplianceAsync()
    {
        try
        {
            var devices = await _graphClient.DeviceManagement.ManagedDevices.GetAsync(config =>
            {
                config.QueryParameters.Select = new[] { "complianceState", "operatingSystem", "deviceName", "lastSyncDateTime" };
                config.QueryParameters.Top = 999;
            });

            var deviceList = devices?.Value ?? new List<ManagedDevice>();

            var compliant = deviceList.Count(d => d.ComplianceState == ComplianceState.Compliant);
            var nonCompliant = deviceList.Count(d => d.ComplianceState == ComplianceState.Noncompliant);
            var unknown = deviceList.Count(d => d.ComplianceState == ComplianceState.Unknown || 
                                                 d.ComplianceState == null);

            var byPlatform = deviceList
                .GroupBy(d => d.OperatingSystem ?? "Unknown")
                .Select(g => new DeviceByPlatformDto(
                    Platform: g.Key,
                    Total: g.Count(),
                    Compliant: g.Count(d => d.ComplianceState == ComplianceState.Compliant),
                    NonCompliant: g.Count(d => d.ComplianceState == ComplianceState.Noncompliant)
                ))
                .OrderByDescending(p => p.Total)
                .ToList();

            return new DeviceComplianceDto(
                TotalDevices: deviceList.Count,
                CompliantDevices: compliant,
                NonCompliantDevices: nonCompliant,
                UnknownDevices: unknown,
                ComplianceRate: deviceList.Count > 0 ? Math.Round((double)compliant / deviceList.Count * 100, 1) : 0,
                ByPlatform: byPlatform,
                LastUpdated: DateTime.UtcNow
            );
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error retrieving device compliance data");
            return GetMockDeviceComplianceData();
        }
    }

    public async Task<MailActivityDto> GetMailActivityAsync(DateTime startDate, DateTime endDate)
    {
        try
        {
            var period = GetReportPeriod(startDate, endDate);
            var response = await _graphClient.Reports
                .GetEmailActivityCountsWithPeriod(period)
                .GetAsync();

            var data = await ParseMailActivityReportAsync(response);

            return new MailActivityDto(
                TotalEmailsSent: data.Sum(d => d.Sent),
                TotalEmailsReceived: data.Sum(d => d.Received),
                TotalEmailsRead: data.Sum(d => d.Sent + d.Received) / 2, // Approximation
                Trend: data.Select(d => new MailActivityTrendDto(d.Date, d.Sent, d.Received)).ToList(),
                LastUpdated: DateTime.UtcNow
            );
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error retrieving mail activity data");
            return GetMockMailActivityData();
        }
    }

    public async Task<TeamsActivityDto> GetTeamsActivityAsync(DateTime startDate, DateTime endDate)
    {
        try
        {
            var period = GetReportPeriod(startDate, endDate);
            _logger.LogInformation("Fetching Teams activity report for period {Period} (days: {Days})", period, (endDate - startDate).Days);
            
            // Get activity counts (messages, calls, meetings)
            var activityResponse = await _graphClient.Reports
                .GetTeamsUserActivityCountsWithPeriod(period)
                .GetAsync();

            var data = await ParseTeamsActivityReportAsync(activityResponse);

            return new TeamsActivityDto(
                TotalMessages: data.Sum(d => d.Messages),
                TotalCalls: data.Sum(d => d.Calls),
                TotalMeetings: data.Sum(d => d.Meetings),
                ActiveUsers: 0, // Not available from this report
                Trend: data.Select(d => new TeamsActivityTrendDto(d.Date, d.Messages, d.Calls, d.Meetings)).ToList(),
                LastUpdated: DateTime.UtcNow
            );
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error retrieving Teams activity data");
            return GetMockTeamsActivityData();
        }
    }

    #region Helper Methods

    private static string GetReportPeriod(DateTime startDate, DateTime endDate)
    {
        var days = (endDate - startDate).Days;
        return days switch
        {
            <= 7 => "D7",
            <= 30 => "D30",
            <= 90 => "D90",
            _ => "D180"
        };
    }

    private async Task<List<(DateTime Date, int Count)>> ParseActiveUsersReportAsync(Stream? response)
    {
        if (response == null) return new List<(DateTime, int)>();
        
        using var reader = new StreamReader(response);
        var csv = await reader.ReadToEndAsync();
        var lines = csv.Split('\n').Skip(1); // Skip header

        return lines
            .Where(line => !string.IsNullOrWhiteSpace(line))
            .Select(line =>
            {
                var parts = line.Split(',');
                return (
                    Date: DateTime.TryParse(parts.ElementAtOrDefault(0), out var date) ? date : DateTime.MinValue,
                    Count: int.TryParse(parts.LastOrDefault()?.Trim(), out var count) ? count : 0
                );
            })
            .Where(x => x.Date != DateTime.MinValue)
            .OrderBy(x => x.Date)
            .ToList();
    }

    private async Task<List<(DateTime Date, long Sent, long Received)>> ParseMailActivityReportAsync(Stream? response)
    {
        if (response == null) return new List<(DateTime, long, long)>();
        
        using var reader = new StreamReader(response);
        var csv = await reader.ReadToEndAsync();
        var lines = csv.Split('\n').Skip(1);

        return lines
            .Where(line => !string.IsNullOrWhiteSpace(line))
            .Select(line =>
            {
                var parts = line.Split(',');
                return (
                    Date: DateTime.TryParse(parts.ElementAtOrDefault(0), out var date) ? date : DateTime.MinValue,
                    Sent: long.TryParse(parts.ElementAtOrDefault(1)?.Trim(), out var sent) ? sent : 0,
                    Received: long.TryParse(parts.ElementAtOrDefault(2)?.Trim(), out var received) ? received : 0
                );
            })
            .Where(x => x.Date != DateTime.MinValue)
            .OrderBy(x => x.Date)
            .ToList();
    }

    private async Task<List<(DateTime Date, long Messages, long Calls, long Meetings, int ActiveUsers)>> ParseTeamsActivityReportAsync(Stream? response)
    {
        if (response == null) 
        {
            _logger.LogWarning("Teams Activity: Response stream is null");
            return new List<(DateTime, long, long, long, int)>();
        }
        
        using var reader = new StreamReader(response);
        var csv = await reader.ReadToEndAsync();
        
        _logger.LogInformation("Teams Activity CSV raw content (first 2000 chars): {Content}", 
            csv.Length > 2000 ? csv.Substring(0, 2000) : csv);
        
        var lines = csv.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
        if (lines.Length == 0) 
        {
            _logger.LogWarning("Teams Activity: No lines in CSV");
            return new List<(DateTime, long, long, long, int)>();
        }
        
        // Parse header to find column indices
        var header = lines[0].ToLowerInvariant();
        var headerParts = header.Split(',').Select(h => h.Trim().Trim('"').Replace(" ", "")).ToArray();
        
        _logger.LogInformation("Teams Activity CSV headers ({Count}): {Headers}", headerParts.Length, string.Join(" | ", headerParts));
        
        // Find column indices by name - Microsoft Graph Teams User Activity Counts report columns:
        // Report Refresh Date, Report Date, Team Chat Messages, Private Chat Messages, Calls, Meetings, 
        // Meetings Organized Count, Meetings Attended Count, Audio Duration, Video Duration, Screen Share Duration,
        // Post Messages, Reply Messages, Report Period
        var dateIndex = Array.FindIndex(headerParts, h => h == "reportdate" || h.Contains("reportdate"));
        var teamChatIndex = Array.FindIndex(headerParts, h => h.Contains("teamchat"));
        var privateChatIndex = Array.FindIndex(headerParts, h => h.Contains("privatechat"));
        var callsIndex = Array.FindIndex(headerParts, h => h == "calls" || (h.Contains("call") && !h.Contains("chat")));
        var meetingsIndex = Array.FindIndex(headerParts, h => h == "meetings");
        var meetingsOrganizedIndex = Array.FindIndex(headerParts, h => h.Contains("meetingsorganized"));
        var meetingsAttendedIndex = Array.FindIndex(headerParts, h => h.Contains("meetingsattended"));
        var postMessagesIndex = Array.FindIndex(headerParts, h => h == "postmessages");
        var replyMessagesIndex = Array.FindIndex(headerParts, h => h == "replymessages");
        
        _logger.LogInformation("Teams Activity column indices - ReportDate:{Date}, TeamChat:{TeamChat}, PrivateChat:{PrivateChat}, Calls:{Calls}, Meetings:{Meetings}, MeetingsOrganized:{MeetingsOrg}, MeetingsAttended:{MeetingsAtt}, Post:{Post}, Reply:{Reply}",
            dateIndex, teamChatIndex, privateChatIndex, callsIndex, meetingsIndex, meetingsOrganizedIndex, meetingsAttendedIndex, postMessagesIndex, replyMessagesIndex);
        
        // If reportdate not found, try just "date"
        if (dateIndex < 0)
        {
            dateIndex = Array.FindIndex(headerParts, h => h.Contains("date") && !h.Contains("refresh"));
            _logger.LogInformation("Teams Activity: Fallback date index: {Index}", dateIndex);
        }

        var result = new List<(DateTime Date, long Messages, long Calls, long Meetings, int ActiveUsers)>();
        
        for (int i = 1; i < lines.Length; i++)
        {
            var line = lines[i];
            if (string.IsNullOrWhiteSpace(line)) continue;
            
            var parts = line.Split(',').Select(p => p.Trim().Trim('"')).ToArray();
            
            if (i <= 3) // Log first few data rows for debugging
            {
                _logger.LogInformation("Teams Activity row {Row}: {Parts}", i, string.Join(" | ", parts));
            }
            
            // Try to get the date
            DateTime date = DateTime.MinValue;
            if (dateIndex >= 0 && dateIndex < parts.Length)
            {
                var dateStr = parts[dateIndex];
                // Try various date formats
                if (!DateTime.TryParse(dateStr, out date))
                {
                    // Try yyyy-MM-dd format explicitly
                    DateTime.TryParseExact(dateStr, new[] { "yyyy-MM-dd", "MM/dd/yyyy", "dd/MM/yyyy" }, 
                        System.Globalization.CultureInfo.InvariantCulture, 
                        System.Globalization.DateTimeStyles.None, out date);
                }
                if (date == DateTime.MinValue && i <= 3)
                {
                    _logger.LogWarning("Teams Activity: Could not parse date from '{DateStr}' at index {Index}", dateStr, dateIndex);
                }
            }
            
            if (date == DateTime.MinValue) continue;
            
            // Get messages (team chat + private chat + post messages + reply messages)
            long teamChat = 0, privateChat = 0, postMessages = 0, replyMessages = 0;
            if (teamChatIndex >= 0 && teamChatIndex < parts.Length)
                long.TryParse(parts[teamChatIndex], out teamChat);
            if (privateChatIndex >= 0 && privateChatIndex < parts.Length)
                long.TryParse(parts[privateChatIndex], out privateChat);
            if (postMessagesIndex >= 0 && postMessagesIndex < parts.Length)
                long.TryParse(parts[postMessagesIndex], out postMessages);
            if (replyMessagesIndex >= 0 && replyMessagesIndex < parts.Length)
                long.TryParse(parts[replyMessagesIndex], out replyMessages);
            var messages = teamChat + privateChat + postMessages + replyMessages;
            
            // Get calls
            long calls = 0;
            if (callsIndex >= 0 && callsIndex < parts.Length)
                long.TryParse(parts[callsIndex], out calls);
            
            // Get meetings (meetings + organized + attended - but avoid double counting)
            // The "Meetings" column seems to always be 0, so use organized + attended
            long meetingsBasic = 0, meetingsOrganized = 0, meetingsAttended = 0;
            if (meetingsIndex >= 0 && meetingsIndex < parts.Length)
                long.TryParse(parts[meetingsIndex], out meetingsBasic);
            if (meetingsOrganizedIndex >= 0 && meetingsOrganizedIndex < parts.Length)
                long.TryParse(parts[meetingsOrganizedIndex], out meetingsOrganized);
            if (meetingsAttendedIndex >= 0 && meetingsAttendedIndex < parts.Length)
                long.TryParse(parts[meetingsAttendedIndex], out meetingsAttended);
            // Use the basic meetings count if available, otherwise use organized + attended
            var meetings = meetingsBasic > 0 ? meetingsBasic : (meetingsOrganized + meetingsAttended);
            
            result.Add((date, messages, calls, meetings, 0)); // ActiveUsers not in this report
        }
        
        _logger.LogInformation("Teams Activity: Parsed {Count} records with valid dates", result.Count);
        
        return result.OrderBy(x => x.Date).ToList();
    }

    #endregion

    #region Mock Data (for development/testing)

    private static ActiveUsersDataDto GetMockActiveUsersData()
    {
        var trend = Enumerable.Range(0, 30)
            .Select(i => new DailyActiveUsersTrendDto(
                DateTime.UtcNow.AddDays(-30 + i),
                Random.Shared.Next(150, 200)))
            .ToList();

        return new ActiveUsersDataDto(180, 175, 190, trend, DateTime.UtcNow);
    }

    private static SignInAnalyticsDto GetMockSignInAnalyticsData()
    {
        var trend = Enumerable.Range(0, 7)
            .Select(i => new SignInTrendDto(
                DateTime.UtcNow.AddDays(-7 + i),
                Random.Shared.Next(450, 500),
                Random.Shared.Next(5, 15)))
            .ToList();

        return new SignInAnalyticsDto(3500, 3420, 80, 5, 97.7, trend, 
            new List<TopSignInLocationDto> { new("London, GB", 2100), new("Manchester, GB", 800) }, 
            DateTime.UtcNow);
    }

    private static LicenseUsageDto GetMockLicenseUsageData()
    {
        return new LicenseUsageDto(
            new List<LicenseSkuDto>
            {
                new("guid1", "ENTERPRISEPACK", 180, 200, 90.0),
                new("guid2", "EMS", 150, 200, 75.0),
                new("guid3", "POWER_BI_PRO", 45, 50, 90.0)
            },
            375, 450, 83.3, DateTime.UtcNow);
    }

    private static DeviceComplianceDto GetMockDeviceComplianceData()
    {
        return new DeviceComplianceDto(250, 220, 20, 10, 88.0,
            new List<DeviceByPlatformDto>
            {
                new("Windows", 150, 140, 10),
                new("iOS", 60, 55, 5),
                new("Android", 40, 25, 5)
            },
            DateTime.UtcNow);
    }

    private static MailActivityDto GetMockMailActivityData()
    {
        var trend = Enumerable.Range(0, 30)
            .Select(i => new MailActivityTrendDto(
                DateTime.UtcNow.AddDays(-30 + i),
                Random.Shared.Next(8000, 12000),
                Random.Shared.Next(15000, 20000)))
            .ToList();

        return new MailActivityDto(300000, 550000, 480000, trend, DateTime.UtcNow);
    }

    private static TeamsActivityDto GetMockTeamsActivityData()
    {
        var trend = Enumerable.Range(0, 30)
            .Select(i => new TeamsActivityTrendDto(
                DateTime.UtcNow.AddDays(-30 + i),
                Random.Shared.Next(5000, 8000),
                Random.Shared.Next(200, 400),
                Random.Shared.Next(50, 100)))
            .ToList();

        return new TeamsActivityDto(180000, 9000, 2200, 175, trend, DateTime.UtcNow);
    }

    #endregion

    #region User Management Methods

    public async Task<UserListResultDto> GetUsersAsync(string? filter, string? orderBy, bool ascending, int take)
    {
        try
        {
            _logger.LogInformation("Fetching users from tenant");

            // Build the select properties
            // Note: signInActivity requires AuditLog.Read.All and Azure AD Premium
            var selectProperties = new[]
            {
                "id", "displayName", "userPrincipalName", "mail", "userType",
                "accountEnabled", "createdDateTime", "jobTitle", "department",
                "officeLocation", "city", "country", "mobilePhone", "businessPhones",
                "assignedLicenses", "usageLocation"
            };

            // Try to include signInActivity if available
            List<User> userList;
            long? totalCount = null;
            
            try
            {
                // First try with signInActivity
                var usersWithSignIn = await _graphClient.Users.GetAsync(config =>
                {
                    config.QueryParameters.Select = selectProperties.Concat(new[] { "signInActivity" }).ToArray();
                    config.QueryParameters.Top = take;
                    config.QueryParameters.Count = true;
                    config.Headers.Add("ConsistencyLevel", "eventual");

                    if (!string.IsNullOrEmpty(filter))
                    {
                        config.QueryParameters.Filter = $"startsWith(displayName, '{filter}') or startsWith(userPrincipalName, '{filter}')";
                    }

                    if (!string.IsNullOrEmpty(orderBy))
                    {
                        config.QueryParameters.Orderby = new[] { ascending ? orderBy : $"{orderBy} desc" };
                    }
                });
                
                userList = usersWithSignIn?.Value ?? new List<User>();
                totalCount = usersWithSignIn?.OdataCount;
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "Could not fetch users with signInActivity, trying without it");
                
                // Fallback without signInActivity
                var usersWithoutSignIn = await _graphClient.Users.GetAsync(config =>
                {
                    config.QueryParameters.Select = selectProperties;
                    config.QueryParameters.Top = take;
                    config.QueryParameters.Count = true;
                    config.Headers.Add("ConsistencyLevel", "eventual");

                    if (!string.IsNullOrEmpty(filter))
                    {
                        config.QueryParameters.Filter = $"startsWith(displayName, '{filter}') or startsWith(userPrincipalName, '{filter}')";
                    }

                    if (!string.IsNullOrEmpty(orderBy))
                    {
                        config.QueryParameters.Orderby = new[] { ascending ? orderBy : $"{orderBy} desc" };
                    }
                });
                
                userList = usersWithoutSignIn?.Value ?? new List<User>();
                totalCount = usersWithoutSignIn?.OdataCount;
            }

            // Get license SKU names for mapping
            var skuNames = await GetLicenseSkuNamesAsync();
            
            // Get MFA registration details for all users
            var mfaDetails = await GetMfaRegistrationMapAsync();

            var tenantUsers = userList.Select(u => {
                var mfa = mfaDetails.GetValueOrDefault(u.Id ?? "");
                return new TenantUserDto(
                    Id: u.Id ?? string.Empty,
                    DisplayName: u.DisplayName ?? "Unknown",
                    UserPrincipalName: u.UserPrincipalName ?? string.Empty,
                    Mail: u.Mail,
                    UserType: u.UserType ?? "Member",
                    AccountEnabled: u.AccountEnabled ?? false,
                    CreatedDateTime: u.CreatedDateTime?.DateTime,
                    LastSignInDateTime: u.SignInActivity?.LastSignInDateTime?.DateTime,
                    LastNonInteractiveSignInDateTime: u.SignInActivity?.LastNonInteractiveSignInDateTime?.DateTime,
                    JobTitle: u.JobTitle,
                    Department: u.Department,
                    OfficeLocation: u.OfficeLocation,
                    City: u.City,
                    Country: u.Country,
                    MobilePhone: u.MobilePhone,
                    BusinessPhones: u.BusinessPhones?.FirstOrDefault(),
                    AssignedLicenses: u.AssignedLicenses?.Select(l => skuNames.GetValueOrDefault(l.SkuId?.ToString() ?? "", l.SkuId?.ToString() ?? "Unknown")).ToList(),
                    HasMailbox: !string.IsNullOrEmpty(u.Mail),
                    ManagerDisplayName: null,
                    ProfilePhoto: null,
                    IsMfaRegistered: mfa?.IsMfaRegistered ?? false,
                    IsMfaCapable: mfa?.IsMfaCapable ?? false,
                    DefaultMfaMethod: mfa?.DefaultMfaMethod,
                    MfaMethods: mfa?.MfaMethods,
                    UsageLocation: u.UsageLocation
                );
            }).ToList();

            return new UserListResultDto(
                Users: tenantUsers,
                TotalCount: (int)(totalCount ?? userList.Count),
                FilteredCount: tenantUsers.Count,
                NextLink: null
            );
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error retrieving users");
            throw;
        }
    }

    private record MfaInfo(bool IsMfaRegistered, bool IsMfaCapable, string? DefaultMfaMethod, List<string>? MfaMethods);

    private async Task<Dictionary<string, MfaInfo>> GetMfaRegistrationMapAsync()
    {
        var result = new Dictionary<string, MfaInfo>();
        
        try
        {
            _logger.LogInformation("Fetching MFA registration details");
            
            // Use the authentication methods registration details report
            var registrationDetails = await _graphClient.Reports.AuthenticationMethods.UserRegistrationDetails.GetAsync(config =>
            {
                config.QueryParameters.Top = 999;
            });

            if (registrationDetails?.Value != null)
            {
                foreach (var detail in registrationDetails.Value)
                {
                    if (detail.Id == null) continue;
                    
                    var methods = detail.MethodsRegistered?.ToList() ?? new List<string>();
                    var isMfaRegistered = detail.IsMfaRegistered ?? false;
                    var isMfaCapable = detail.IsMfaCapable ?? false;
                    
                    // Get default method from AdditionalData or use first registered method
                    string? defaultMethod = null;
                    if (detail.AdditionalData?.TryGetValue("defaultMfaMethod", out var defaultMfaValue) == true)
                    {
                        defaultMethod = defaultMfaValue?.ToString();
                    }
                    else if (methods.Count > 0)
                    {
                        // Use first strong method as default if not specified
                        defaultMethod = methods.FirstOrDefault(m => 
                            m.Contains("authenticator", StringComparison.OrdinalIgnoreCase) ||
                            m.Contains("fido", StringComparison.OrdinalIgnoreCase) ||
                            m.Contains("phone", StringComparison.OrdinalIgnoreCase));
                    }
                    
                    // Convert method names to friendly names
                    var friendlyMethods = methods.Select(m => GetFriendlyMfaMethodName(m) ?? m).ToList();
                    var friendlyDefaultMethod = GetFriendlyMfaMethodName(defaultMethod);
                    
                    result[detail.Id] = new MfaInfo(isMfaRegistered, isMfaCapable, friendlyDefaultMethod, friendlyMethods);
                }
            }
            
            // Page through if needed
            while (registrationDetails?.OdataNextLink != null)
            {
                registrationDetails = await _graphClient.Reports.AuthenticationMethods.UserRegistrationDetails
                    .WithUrl(registrationDetails.OdataNextLink)
                    .GetAsync();
                    
                if (registrationDetails?.Value != null)
                {
                    foreach (var detail in registrationDetails.Value)
                    {
                        if (detail.Id == null) continue;
                        
                        var methods = detail.MethodsRegistered?.ToList() ?? new List<string>();
                        var isMfaRegistered = detail.IsMfaRegistered ?? false;
                        var isMfaCapable = detail.IsMfaCapable ?? false;
                        
                        string? defaultMethod = null;
                        if (detail.AdditionalData?.TryGetValue("defaultMfaMethod", out var defaultMfaValue) == true)
                        {
                            defaultMethod = defaultMfaValue?.ToString();
                        }
                        else if (methods.Count > 0)
                        {
                            defaultMethod = methods.FirstOrDefault(m => 
                                m.Contains("authenticator", StringComparison.OrdinalIgnoreCase) ||
                                m.Contains("fido", StringComparison.OrdinalIgnoreCase) ||
                                m.Contains("phone", StringComparison.OrdinalIgnoreCase));
                        }
                        
                        var friendlyMethods = methods.Select(m => GetFriendlyMfaMethodName(m) ?? m).ToList();
                        var friendlyDefaultMethod = GetFriendlyMfaMethodName(defaultMethod);
                        
                        result[detail.Id] = new MfaInfo(isMfaRegistered, isMfaCapable, friendlyDefaultMethod, friendlyMethods);
                    }
                }
            }
            
            _logger.LogInformation("Retrieved MFA details for {Count} users", result.Count);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Could not fetch MFA registration details - MFA status will be unavailable");
        }
        
        return result;
    }

    private static string? GetFriendlyMfaMethodName(string? method)
    {
        if (string.IsNullOrEmpty(method)) return null;
        
        return method.ToLowerInvariant() switch
        {
            "microsoftauthenticatorpush" => "Authenticator App",
            "microsoftauthenticatorpasswordless" => "Authenticator Passwordless",
            "softwareoath" => "Authenticator TOTP",
            "passkeyfido2" => "Passkey (FIDO2)",
            "fido2" => "FIDO2 Key",
            "windowshelloforbusiness" => "Windows Hello",
            "sms" => "SMS",
            "voicemobile" => "Voice Call (Mobile)",
            "voicealternatephonenumber" => "Voice Call (Alternate)",
            "voiceoffice" => "Voice Call (Office)",
            "email" => "Email",
            "temporaryaccesspass" => "Temporary Access Pass",
            "mobilephone" => "Mobile Phone",
            "officephone" => "Office Phone",
            "alternatephonenumber" => "Alternate Phone",
            "none" => "None",
            _ => method
        };
    }

    public async Task<UserDetailDto> GetUserDetailsAsync(string userId)
    {
        try
        {
            _logger.LogInformation("Fetching detailed user info for {UserId}", userId);

            var user = await _graphClient.Users[userId].GetAsync(config =>
            {
                config.QueryParameters.Select = new[]
                {
                    "id", "displayName", "userPrincipalName", "mail", "userType",
                    "accountEnabled", "createdDateTime", "jobTitle", "department",
                    "companyName", "officeLocation", "streetAddress", "city", "state",
                    "postalCode", "country", "mobilePhone", "businessPhones",
                    "assignedLicenses", "signInActivity", "lastPasswordChangeDateTime",
                    "onPremisesSamAccountName", "onPremisesSyncEnabled", "onPremisesLastSyncDateTime"
                };
            });

            if (user == null)
            {
                throw new InvalidOperationException($"User {userId} not found");
            }

            // Get manager
            string? managerId = null;
            string? managerDisplayName = null;
            try
            {
                var manager = await _graphClient.Users[userId].Manager.GetAsync();
                if (manager is User managerUser)
                {
                    managerId = managerUser.Id;
                    managerDisplayName = managerUser.DisplayName;
                }
            }
            catch { /* No manager assigned */ }

            // Get group memberships
            var groups = new List<GroupMembershipDto>();
            try
            {
                var memberOf = await _graphClient.Users[userId].MemberOf.GetAsync(config =>
                {
                    config.QueryParameters.Select = new[] { "id", "displayName", "description", "groupTypes" };
                    config.QueryParameters.Top = 50;
                });

                groups = memberOf?.Value?
                    .OfType<Group>()
                    .Select(g => new GroupMembershipDto(
                        Id: g.Id ?? string.Empty,
                        DisplayName: g.DisplayName ?? "Unknown",
                        Description: g.Description,
                        GroupType: g.GroupTypes?.Contains("Unified") == true ? "Microsoft 365" :
                                   g.SecurityEnabled == true ? "Security" : "Distribution"
                    )).ToList() ?? new List<GroupMembershipDto>();
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "Could not fetch group memberships for user {UserId}", userId);
            }

            // Get license details
            var skuNames = await GetLicenseSkuNamesAsync();
            var licenses = user.AssignedLicenses?
                .Select(l => new LicenseDetailDto(
                    SkuId: l.SkuId?.ToString() ?? string.Empty,
                    SkuPartNumber: skuNames.GetValueOrDefault(l.SkuId?.ToString() ?? "", "Unknown"),
                    DisplayName: GetFriendlyLicenseName(skuNames.GetValueOrDefault(l.SkuId?.ToString() ?? "", "Unknown"))
                )).ToList() ?? new List<LicenseDetailDto>();

            // Get direct reports
            var directReports = new List<string>();
            try
            {
                var reports = await _graphClient.Users[userId].DirectReports.GetAsync(config =>
                {
                    config.QueryParameters.Select = new[] { "displayName" };
                });
                directReports = reports?.Value?
                    .OfType<User>()
                    .Select(u => u.DisplayName ?? "Unknown")
                    .ToList() ?? new List<string>();
            }
            catch { /* No direct reports */ }

            // Try to get profile photo
            string? photoBase64 = null;
            try
            {
                var photoStream = await _graphClient.Users[userId].Photo.Content.GetAsync();
                if (photoStream != null)
                {
                    using var memoryStream = new MemoryStream();
                    await photoStream.CopyToAsync(memoryStream);
                    photoBase64 = $"data:image/jpeg;base64,{Convert.ToBase64String(memoryStream.ToArray())}";
                }
            }
            catch { /* No photo */ }

            return new UserDetailDto(
                Id: user.Id ?? string.Empty,
                DisplayName: user.DisplayName ?? "Unknown",
                UserPrincipalName: user.UserPrincipalName ?? string.Empty,
                Mail: user.Mail,
                UserType: user.UserType ?? "Member",
                AccountEnabled: user.AccountEnabled ?? false,
                CreatedDateTime: user.CreatedDateTime?.DateTime,
                LastSignInDateTime: user.SignInActivity?.LastSignInDateTime?.DateTime,
                LastNonInteractiveSignInDateTime: user.SignInActivity?.LastNonInteractiveSignInDateTime?.DateTime,
                LastPasswordChangeDateTime: user.LastPasswordChangeDateTime?.DateTime,
                JobTitle: user.JobTitle,
                Department: user.Department,
                CompanyName: user.CompanyName,
                OfficeLocation: user.OfficeLocation,
                StreetAddress: user.StreetAddress,
                City: user.City,
                State: user.State,
                PostalCode: user.PostalCode,
                Country: user.Country,
                MobilePhone: user.MobilePhone,
                BusinessPhones: user.BusinessPhones?.ToList(),
                Licenses: licenses,
                GroupMemberships: groups,
                ManagerId: managerId,
                ManagerDisplayName: managerDisplayName,
                DirectReports: directReports,
                ProfilePhoto: photoBase64,
                OnPremisesSamAccountName: user.OnPremisesSamAccountName,
                OnPremisesSyncEnabled: user.OnPremisesSyncEnabled,
                OnPremisesLastSyncDateTime: user.OnPremisesLastSyncDateTime?.DateTime
            );
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error retrieving user details for {UserId}", userId);
            throw;
        }
    }

    public async Task<UserStatsDto> GetUserStatsAsync()
    {
        try
        {
            _logger.LogInformation("Calculating user statistics");

            // Get all users with minimal properties for stats
            var allUsers = new List<User>();
            
            // Try with signInActivity first, fall back without it
            string[] selectProps;
            bool hasSignInActivity = false;
            
            try
            {
                var testResponse = await _graphClient.Users.GetAsync(config =>
                {
                    config.QueryParameters.Select = new[] { "id", "userType", "accountEnabled", "assignedLicenses", "signInActivity" };
                    config.QueryParameters.Top = 1;
                    config.Headers.Add("ConsistencyLevel", "eventual");
                });
                hasSignInActivity = true;
                selectProps = new[] { "id", "userType", "accountEnabled", "assignedLicenses", "signInActivity" };
            }
            catch
            {
                _logger.LogWarning("signInActivity not available, fetching users without it");
                selectProps = new[] { "id", "userType", "accountEnabled", "assignedLicenses" };
            }

            var response = await _graphClient.Users.GetAsync(config =>
            {
                config.QueryParameters.Select = selectProps;
                config.QueryParameters.Top = 999;
                config.Headers.Add("ConsistencyLevel", "eventual");
                config.QueryParameters.Count = true;
            });

            if (response?.Value != null)
            {
                allUsers.AddRange(response.Value);
            }

            // Page through results if needed
            while (response?.OdataNextLink != null)
            {
                response = await _graphClient.Users
                    .WithUrl(response.OdataNextLink)
                    .GetAsync();
                if (response?.Value != null)
                {
                    allUsers.AddRange(response.Value);
                }
            }

            var thirtyDaysAgo = DateTime.UtcNow.AddDays(-30);

            // Get deleted users count
            int deletedUsersCount = 0;
            try
            {
                var deletedUsers = await _graphClient.Directory.DeletedItems.GraphUser.GetAsync(config =>
                {
                    config.QueryParameters.Count = true;
                    config.QueryParameters.Top = 1;
                    config.Headers.Add("ConsistencyLevel", "eventual");
                });
                deletedUsersCount = (int)(deletedUsers?.OdataCount ?? 0);
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "Could not fetch deleted users count");
            }

            // Get MFA stats
            int mfaRegistered = 0;
            int mfaNotRegistered = 0;
            try
            {
                var mfaDetails = await _graphClient.Reports.AuthenticationMethods.UserRegistrationDetails.GetAsync(config =>
                {
                    config.QueryParameters.Top = 999;
                });

                var allMfaDetails = mfaDetails?.Value?.ToList() ?? new List<Microsoft.Graph.Models.UserRegistrationDetails>();
                
                while (mfaDetails?.OdataNextLink != null)
                {
                    mfaDetails = await _graphClient.Reports.AuthenticationMethods.UserRegistrationDetails
                        .WithUrl(mfaDetails.OdataNextLink)
                        .GetAsync();
                    if (mfaDetails?.Value != null)
                    {
                        allMfaDetails.AddRange(mfaDetails.Value);
                    }
                }

                mfaRegistered = allMfaDetails.Count(m => m.IsMfaRegistered == true);
                mfaNotRegistered = allMfaDetails.Count(m => m.IsMfaRegistered != true);
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "Could not fetch MFA registration stats");
            }

            var stats = new UserStatsDto(
                TotalUsers: allUsers.Count,
                EnabledUsers: allUsers.Count(u => u.AccountEnabled == true),
                DisabledUsers: allUsers.Count(u => u.AccountEnabled == false),
                MemberUsers: allUsers.Count(u => u.UserType == "Member"),
                GuestUsers: allUsers.Count(u => u.UserType == "Guest"),
                LicensedUsers: allUsers.Count(u => u.AssignedLicenses?.Any() == true),
                UnlicensedUsers: allUsers.Count(u => u.AssignedLicenses?.Any() != true),
                UsersSignedInLast30Days: hasSignInActivity ? allUsers.Count(u => 
                    u.SignInActivity?.LastSignInDateTime?.DateTime > thirtyDaysAgo ||
                    u.SignInActivity?.LastNonInteractiveSignInDateTime?.DateTime > thirtyDaysAgo) : 0,
                UsersNeverSignedIn: hasSignInActivity ? allUsers.Count(u => 
                    u.SignInActivity?.LastSignInDateTime == null && 
                    u.SignInActivity?.LastNonInteractiveSignInDateTime == null) : 0,
                DeletedUsers: deletedUsersCount,
                MfaRegistered: mfaRegistered,
                MfaNotRegistered: mfaNotRegistered,
                LastUpdated: DateTime.UtcNow
            );

            return stats;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error calculating user statistics");
            throw;
        }
    }

    private async Task<Dictionary<string, string>> GetLicenseSkuNamesAsync()
    {
        try
        {
            var skus = await _graphClient.SubscribedSkus.GetAsync();
            return skus?.Value?
                .Where(s => s.SkuId != null && s.SkuPartNumber != null)
                .ToDictionary(s => s.SkuId!.Value.ToString(), s => s.SkuPartNumber!)
                ?? new Dictionary<string, string>();
        }
        catch
        {
            return new Dictionary<string, string>();
        }
    }

    private static string? GetFriendlyLicenseName(string skuPartNumber)
    {
        // Map common SKU part numbers to friendly names
        return skuPartNumber switch
        {
            "ENTERPRISEPACK" => "Office 365 E3",
            "ENTERPRISEPREMIUM" => "Office 365 E5",
            "SPE_E3" => "Microsoft 365 E3",
            "SPE_E5" => "Microsoft 365 E5",
            "SPB" => "Microsoft 365 Business Premium",
            "O365_BUSINESS_PREMIUM" => "Microsoft 365 Business Standard",
            "O365_BUSINESS_ESSENTIALS" => "Microsoft 365 Business Basic",
            "EXCHANGESTANDARD" => "Exchange Online (Plan 1)",
            "EXCHANGEENTERPRISE" => "Exchange Online (Plan 2)",
            "EMS" => "Enterprise Mobility + Security E3",
            "EMSPREMIUM" => "Enterprise Mobility + Security E5",
            "POWER_BI_PRO" => "Power BI Pro",
            "POWER_BI_STANDARD" => "Power BI Free",
            "PROJECTPREMIUM" => "Project Plan 5",
            "VISIOCLIENT" => "Visio Plan 2",
            "TEAMS_EXPLORATORY" => "Microsoft Teams Exploratory",
            "FLOW_FREE" => "Power Automate Free",
            "POWERAPPS_VIRAL" => "Power Apps Plan 2 Trial",
            "AAD_PREMIUM" => "Azure AD Premium P1",
            "AAD_PREMIUM_P2" => "Azure AD Premium P2",
            "INTUNE_A" => "Microsoft Intune",
            "ATP_ENTERPRISE" => "Microsoft Defender for Office 365 (Plan 1)",
            "THREAT_INTELLIGENCE" => "Microsoft Defender for Office 365 (Plan 2)",
            "WIN_DEF_ATP" => "Microsoft Defender for Endpoint",
            "IDENTITY_THREAT_PROTECTION" => "Microsoft 365 E5 Security",
            "M365_F1" => "Microsoft 365 F1",
            "SPE_F1" => "Microsoft 365 F3",
            "DESKLESSPACK" => "Office 365 F3",
            "MCOSTANDARD" => "Skype for Business Online (Plan 2)",
            "STREAM" => "Microsoft Stream",
            "FORMS_PRO" => "Dynamics 365 Customer Voice",
            "WINDOWS_STORE" => "Windows Store for Business",
            _ => skuPartNumber
        };
    }

    #endregion

    #region Group Management Methods

    public async Task<GroupListResultDto> GetGroupsAsync(string? filter, string? orderBy, bool ascending, int take)
    {
        try
        {
            _logger.LogInformation("Fetching groups from tenant");

            var selectProperties = new[]
            {
                "id", "displayName", "description", "mail", "mailEnabled", "securityEnabled",
                "groupTypes", "visibility", "createdDateTime", "renewedDateTime"
            };

            var response = await _graphClient.Groups.GetAsync(config =>
            {
                config.QueryParameters.Select = selectProperties;
                config.QueryParameters.Top = take;
                config.QueryParameters.Count = true;
                config.Headers.Add("ConsistencyLevel", "eventual");

                if (!string.IsNullOrEmpty(filter))
                {
                    config.QueryParameters.Filter = $"startsWith(displayName, '{filter}')";
                }

                if (!string.IsNullOrEmpty(orderBy))
                {
                    config.QueryParameters.Orderby = new[] { ascending ? orderBy : $"{orderBy} desc" };
                }
            });

            var groupList = response?.Value ?? new List<Group>();
            var totalCount = response?.OdataCount;

            // Build group DTOs with member/owner counts
            var tenantGroups = new List<TenantGroupDto>();
            
            foreach (var group in groupList)
            {
                var groupType = GetGroupType(group);
                var isM365Group = group.GroupTypes?.Contains("Unified") == true;
                
                // Get member count
                var memberCount = 0;
                var ownerCount = 0;
                var isTeam = false;
                string? teamWebUrl = null;

                try
                {
                    // Get member count - fetch minimal data and use count
                    var membersResponse = await _graphClient.Groups[group.Id].Members.GetAsync(config =>
                    {
                        config.QueryParameters.Select = new[] { "id" };
                        config.QueryParameters.Top = 999;
                    });
                    memberCount = membersResponse?.Value?.Count ?? 0;
                }
                catch (Exception ex)
                {
                    _logger.LogDebug(ex, "Could not get member count for group {GroupId}", group.Id);
                }

                try
                {
                    // Get owner count - fetch minimal data and use count
                    var ownersResponse = await _graphClient.Groups[group.Id].Owners.GetAsync(config =>
                    {
                        config.QueryParameters.Select = new[] { "id" };
                        config.QueryParameters.Top = 999;
                    });
                    ownerCount = ownersResponse?.Value?.Count ?? 0;
                }
                catch (Exception ex)
                {
                    _logger.LogDebug(ex, "Could not get owner count for group {GroupId}", group.Id);
                }

                // Check if M365 group is a Team
                if (isM365Group)
                {
                    try
                    {
                        var team = await _graphClient.Teams[group.Id].GetAsync();
                        if (team != null)
                        {
                            isTeam = true;
                            teamWebUrl = team.WebUrl;
                        }
                    }
                    catch { /* Not a team */ }
                }

                tenantGroups.Add(new TenantGroupDto(
                    Id: group.Id ?? string.Empty,
                    DisplayName: group.DisplayName ?? "Unknown",
                    Description: group.Description,
                    Mail: group.Mail,
                    GroupType: groupType,
                    MailEnabled: group.MailEnabled,
                    SecurityEnabled: group.SecurityEnabled,
                    Visibility: group.Visibility,
                    CreatedDateTime: group.CreatedDateTime?.DateTime,
                    RenewedDateTime: group.RenewedDateTime?.DateTime,
                    MemberCount: memberCount,
                    OwnerCount: ownerCount,
                    IsTeam: isTeam,
                    TeamWebUrl: teamWebUrl,
                    ResourceProvisioningOptions: null
                ));
            }

            return new GroupListResultDto(
                Groups: tenantGroups,
                TotalCount: (int)(totalCount ?? groupList.Count),
                FilteredCount: tenantGroups.Count,
                NextLink: null
            );
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error retrieving groups");
            throw;
        }
    }

    public async Task<GroupDetailDto> GetGroupDetailsAsync(string groupId)
    {
        try
        {
            _logger.LogInformation("Fetching detailed group info for {GroupId}", groupId);

            var group = await _graphClient.Groups[groupId].GetAsync(config =>
            {
                config.QueryParameters.Select = new[]
                {
                    "id", "displayName", "description", "mail", "mailEnabled", "securityEnabled",
                    "groupTypes", "visibility", "createdDateTime", "renewedDateTime", "expirationDateTime"
                };
            });

            if (group == null)
            {
                throw new InvalidOperationException($"Group {groupId} not found");
            }

            // Get members
            var members = new List<GroupMemberDto>();
            try
            {
                var membersResponse = await _graphClient.Groups[groupId].Members.GetAsync(config =>
                {
                    config.QueryParameters.Top = 100;
                    config.QueryParameters.Select = new[] { "id", "displayName", "userPrincipalName", "mail" };
                });

                if (membersResponse?.Value != null)
                {
                    foreach (var member in membersResponse.Value)
                    {
                        var rawType = member.OdataType?.Replace("#microsoft.graph.", "") ?? "Unknown";
                        var memberType = rawType switch
                        {
                            "user" => "User",
                            "group" => "Group",
                            "device" => "Device",
                            "servicePrincipal" => "Service Principal",
                            "orgContact" => "Contact",
                            _ => rawType
                        };

                        if (member is User user)
                        {
                            members.Add(new GroupMemberDto(
                                Id: user.Id ?? string.Empty,
                                DisplayName: user.DisplayName ?? "Unknown",
                                UserPrincipalName: user.UserPrincipalName,
                                Mail: user.Mail,
                                MemberType: "User"
                            ));
                        }
                        else if (member is Group nestedGroup)
                        {
                            members.Add(new GroupMemberDto(
                                Id: nestedGroup.Id ?? string.Empty,
                                DisplayName: nestedGroup.DisplayName ?? "Unknown",
                                UserPrincipalName: null,
                                Mail: nestedGroup.Mail,
                                MemberType: "Group"
                            ));
                        }
                        else
                        {
                            // Device, ServicePrincipal, Contact, or other directory object.
                            // The members endpoint with $select doesn't populate AdditionalData for
                            // non-user types — store a placeholder and resolve devices in a second pass.
                            var data = member.AdditionalData ?? new Dictionary<string, object>();
                            var displayName = data.TryGetValue("displayName", out var dn) && dn?.ToString() is { Length: > 0 } dnStr ? dnStr : null;
                            var mail = data.TryGetValue("mail", out var m) ? m?.ToString() : null;
                            var upn = data.TryGetValue("userPrincipalName", out var u) ? u?.ToString() : null;

                            members.Add(new GroupMemberDto(
                                Id: member.Id ?? string.Empty,
                                DisplayName: displayName ?? $"{memberType} ({member.Id?[..8]}…)",
                                UserPrincipalName: upn,
                                Mail: mail,
                                MemberType: memberType
                            ));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "Could not fetch members for group {GroupId}", groupId);
            }

            // Second pass: resolve device members by fetching their display name from /devices/{id}
            var deviceMembers = members.Where(m => m.MemberType == "Device").ToList();
            foreach (var dm in deviceMembers)
            {
                try
                {
                    var device = await _graphClient.Devices[dm.Id].GetAsync(config =>
                    {
                        config.QueryParameters.Select = new[] { "id", "displayName", "deviceId", "operatingSystem", "operatingSystemVersion" };
                    });
                    if (device != null)
                    {
                        var idx = members.FindIndex(m => m.Id == dm.Id);
                        if (idx >= 0)
                        {
                            var os = device.OperatingSystem != null ? $" ({device.OperatingSystem})" : string.Empty;
                            members[idx] = dm with
                            {
                                DisplayName = device.DisplayName ?? dm.DisplayName,
                                UserPrincipalName = device.DeviceId  // Entra device GUID — useful as a subtitle
                            };
                        }
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogDebug(ex, "Could not resolve device {DeviceId}", dm.Id);
                }
            }

            // Get owners
            var owners = new List<GroupMemberDto>();
            try
            {
                var ownersResponse = await _graphClient.Groups[groupId].Owners.GetAsync(config =>
                {
                    config.QueryParameters.Top = 100;
                    config.QueryParameters.Select = new[] { "id", "displayName", "userPrincipalName", "mail" };
                });

                if (ownersResponse?.Value != null)
                {
                    foreach (var owner in ownersResponse.Value)
                    {
                        if (owner is User user)
                        {
                            owners.Add(new GroupMemberDto(
                                Id: user.Id ?? string.Empty,
                                DisplayName: user.DisplayName ?? "Unknown",
                                UserPrincipalName: user.UserPrincipalName,
                                Mail: user.Mail,
                                MemberType: "User"
                            ));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "Could not fetch owners for group {GroupId}", groupId);
            }

            // Check if it's a Team
            var isTeam = false;
            string? teamWebUrl = null;
            bool? isArchived = null;

            if (group.GroupTypes?.Contains("Unified") == true)
            {
                try
                {
                    var team = await _graphClient.Teams[groupId].GetAsync();
                    if (team != null)
                    {
                        isTeam = true;
                        teamWebUrl = team.WebUrl;
                        isArchived = team.IsArchived;
                    }
                }
                catch { /* Not a team or no access */ }
            }

            return new GroupDetailDto(
                Id: group.Id ?? string.Empty,
                DisplayName: group.DisplayName ?? "Unknown",
                Description: group.Description,
                Mail: group.Mail,
                GroupType: GetGroupType(group),
                MailEnabled: group.MailEnabled,
                SecurityEnabled: group.SecurityEnabled,
                Visibility: group.Visibility,
                CreatedDateTime: group.CreatedDateTime?.DateTime,
                RenewedDateTime: group.RenewedDateTime?.DateTime,
                ExpirationDateTime: group.ExpirationDateTime?.DateTime,
                Members: members,
                Owners: owners,
                IsTeam: isTeam,
                TeamWebUrl: teamWebUrl,
                IsArchived: isArchived,
                ResourceProvisioningOptions: null
            );
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error retrieving group details for {GroupId}", groupId);
            throw;
        }
    }

    public async Task<GroupStatsDto> GetGroupStatsAsync()
    {
        try
        {
            _logger.LogInformation("Calculating group statistics");

            var allGroups = new List<Group>();

            var response = await _graphClient.Groups.GetAsync(config =>
            {
                config.QueryParameters.Select = new[] { "id", "groupTypes", "mailEnabled", "securityEnabled", "visibility" };
                config.QueryParameters.Top = 999;
                config.Headers.Add("ConsistencyLevel", "eventual");
                config.QueryParameters.Count = true;
            });

            if (response?.Value != null)
            {
                allGroups.AddRange(response.Value);
            }

            // Page through results
            while (response?.OdataNextLink != null)
            {
                response = await _graphClient.Groups
                    .WithUrl(response.OdataNextLink)
                    .GetAsync();
                if (response?.Value != null)
                {
                    allGroups.AddRange(response.Value);
                }
            }

            var m365Groups = allGroups.Count(g => g.GroupTypes?.Contains("Unified") == true);
            var securityGroups = allGroups.Count(g => g.SecurityEnabled == true && g.MailEnabled != true && g.GroupTypes?.Contains("Unified") != true);
            var distributionGroups = allGroups.Count(g => g.MailEnabled == true && g.SecurityEnabled != true && g.GroupTypes?.Contains("Unified") != true);
            var publicGroups = allGroups.Count(g => g.Visibility?.Equals("Public", StringComparison.OrdinalIgnoreCase) == true);
            var privateGroups = allGroups.Count(g => g.Visibility?.Equals("Private", StringComparison.OrdinalIgnoreCase) == true);

            // Count Teams-enabled groups (M365 groups that are Teams)
            var teamsEnabled = 0;
            var m365GroupIds = allGroups
                .Where(g => g.GroupTypes?.Contains("Unified") == true)
                .Take(50) // Limit for performance
                .ToList();

            foreach (var group in m365GroupIds)
            {
                try
                {
                    var team = await _graphClient.Teams[group.Id].GetAsync();
                    if (team != null)
                    {
                        teamsEnabled++;
                    }
                }
                catch { /* Not a team */ }
            }

            // Extrapolate Teams count based on sample
            if (m365GroupIds.Count > 0 && m365Groups > 50)
            {
                teamsEnabled = (int)((double)teamsEnabled / m365GroupIds.Count * m365Groups);
            }

            return new GroupStatsDto(
                TotalGroups: allGroups.Count,
                Microsoft365Groups: m365Groups,
                SecurityGroups: securityGroups,
                DistributionGroups: distributionGroups,
                TeamsEnabled: teamsEnabled,
                PublicGroups: publicGroups,
                PrivateGroups: privateGroups,
                GroupsWithNoOwner: 0, // Skip for performance
                GroupsWithNoMembers: 0, // Skip for performance
                LastUpdated: DateTime.UtcNow
            );
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error calculating group statistics");
            throw;
        }
    }

    private static string GetGroupType(Group group)
    {
        if (group.GroupTypes?.Contains("Unified") == true)
        {
            return "Microsoft 365";
        }
        if (group.SecurityEnabled == true && group.MailEnabled == true)
        {
            return "Mail-enabled Security";
        }
        if (group.SecurityEnabled == true)
        {
            return "Security";
        }
        if (group.MailEnabled == true)
        {
            return "Distribution";
        }
        return "Other";
    }

    public async Task<GroupListResultDto> GetDistributionListsAsync(int take = 200)
    {
        try
        {
            _logger.LogInformation("Fetching distribution lists from tenant");

            var selectProperties = new[]
            {
                "id", "displayName", "description", "mail", "mailEnabled", "securityEnabled",
                "groupTypes", "visibility", "createdDateTime", "renewedDateTime", "proxyAddresses"
            };

            // Use filter to get mail-enabled groups that are NOT security groups and NOT M365 groups
            // Distribution lists: mailEnabled eq true AND securityEnabled eq false
            var response = await _graphClient.Groups.GetAsync(config =>
            {
                config.QueryParameters.Select = selectProperties;
                config.QueryParameters.Top = take;
                config.QueryParameters.Count = true;
                config.QueryParameters.Filter = "mailEnabled eq true and securityEnabled eq false";
                config.Headers.Add("ConsistencyLevel", "eventual");
            });

            var groupList = response?.Value ?? new List<Group>();
            var totalCount = response?.OdataCount;

            _logger.LogInformation("Found {Count} mail-enabled non-security groups", groupList.Count);

            // Filter out M365 groups (Unified) on the client side since Graph doesn't support NOT contains
            var distributionLists = groupList
                .Where(g => g.GroupTypes == null || !g.GroupTypes.Contains("Unified"))
                .ToList();

            _logger.LogInformation("After filtering out M365 groups: {Count} distribution lists", distributionLists.Count);

            // Log details for debugging
            foreach (var group in distributionLists.Take(10))
            {
                _logger.LogInformation(
                    "DDL: {Name} | Mail: {Mail} | MailEnabled: {MailEnabled} | SecurityEnabled: {SecurityEnabled} | GroupTypes: {GroupTypes}",
                    group.DisplayName,
                    group.Mail,
                    group.MailEnabled,
                    group.SecurityEnabled,
                    group.GroupTypes != null ? string.Join(", ", group.GroupTypes) : "none");
            }

            var tenantGroups = new List<TenantGroupDto>();

            foreach (var group in distributionLists)
            {
                var memberCount = 0;
                var ownerCount = 0;

                try
                {
                    var membersResponse = await _graphClient.Groups[group.Id].Members.GetAsync(config =>
                    {
                        config.QueryParameters.Select = new[] { "id" };
                        config.QueryParameters.Top = 999;
                    });
                    memberCount = membersResponse?.Value?.Count ?? 0;
                }
                catch (Exception ex)
                {
                    _logger.LogDebug(ex, "Could not get member count for group {GroupId}", group.Id);
                }

                try
                {
                    var ownersResponse = await _graphClient.Groups[group.Id].Owners.GetAsync(config =>
                    {
                        config.QueryParameters.Select = new[] { "id" };
                        config.QueryParameters.Top = 999;
                    });
                    ownerCount = ownersResponse?.Value?.Count ?? 0;
                }
                catch (Exception ex)
                {
                    _logger.LogDebug(ex, "Could not get owner count for group {GroupId}", group.Id);
                }

                tenantGroups.Add(new TenantGroupDto(
                    Id: group.Id ?? string.Empty,
                    DisplayName: group.DisplayName ?? "Unknown",
                    Description: group.Description,
                    Mail: group.Mail,
                    GroupType: "Distribution",
                    MailEnabled: group.MailEnabled,
                    SecurityEnabled: group.SecurityEnabled,
                    Visibility: group.Visibility,
                    CreatedDateTime: group.CreatedDateTime?.DateTime,
                    RenewedDateTime: group.RenewedDateTime?.DateTime,
                    MemberCount: memberCount,
                    OwnerCount: ownerCount,
                    IsTeam: false,
                    TeamWebUrl: null,
                    ResourceProvisioningOptions: null
                ));
            }

            return new GroupListResultDto(
                Groups: tenantGroups.OrderBy(g => g.DisplayName).ToList(),
                TotalCount: (int)(totalCount ?? distributionLists.Count),
                FilteredCount: tenantGroups.Count,
                NextLink: null
            );
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error retrieving distribution lists");
            throw;
        }
    }

    #endregion

    #region Device Management Methods (Intune)

    /// <summary>
    /// Helper method to convert Windows SKU number to a friendly edition name
    /// </summary>
    private static string? GetWindowsEditionFromSkuNumber(int? skuNumber)
    {
        if (skuNumber == null) return null;
        
        // Windows SKU numbers: https://docs.microsoft.com/en-us/dotnet/api/microsoft.powershell.commands.operatingsystemsku
        return skuNumber.Value switch
        {
            0 => "Undefined",
            1 => "Ultimate",
            2 => "Home Basic",
            3 => "Home Premium",
            4 => "Enterprise",
            5 => "Home Basic N",
            6 => "Business",
            7 => "Server Standard",
            8 => "Server Datacenter",
            9 => "Small Business Server",
            10 => "Server Enterprise",
            11 => "Starter",
            12 => "Server Datacenter Core",
            13 => "Server Standard Core",
            14 => "Server Enterprise Core",
            27 => "Enterprise N",
            28 => "Ultimate N",
            48 => "Pro",
            49 => "Pro N",
            72 => "Enterprise S",
            98 => "Home",
            99 => "Home N",
            100 => "Home Single Language",
            101 => "Home China",
            121 => "Education",
            122 => "Education N",
            125 => "Enterprise S N",
            126 => "Pro for Workstations",
            161 => "Pro Education",
            162 => "Pro Education N",
            175 => "Enterprise (Multi-session)",
            188 => "Server Datacenter Azure Edition",
            191 => "IoT Enterprise",
            _ => $"SKU {skuNumber.Value}"
        };
    }

    /// <summary>
    /// Helper to extract a string from AdditionalData dictionary
    /// </summary>
    private static string? GetAdditionalDataString(IDictionary<string, object>? data, string key)
    {
        if (data == null) return null;
        if (data.TryGetValue(key, out var value))
        {
            return value?.ToString();
        }
        return null;
    }

    /// <summary>
    /// Helper to extract an integer from AdditionalData dictionary
    /// </summary>
    private static int? GetAdditionalDataInt(IDictionary<string, object>? data, string key)
    {
        if (data == null) return null;
        if (data.TryGetValue(key, out var value))
        {
            if (value is int intValue) return intValue;
            if (value is long longValue) return (int)longValue;
            if (int.TryParse(value?.ToString(), out var parsed)) return parsed;
        }
        return null;
    }

    public async Task<DeviceListResultDto> GetDevicesAsync(string? filter, string? orderBy, bool ascending, int take)
    {
        try
        {
            _logger.LogInformation("Fetching Intune managed devices");

            var response = await _graphClient.DeviceManagement.ManagedDevices.GetAsync(config =>
            {
                config.QueryParameters.Top = take;
                config.QueryParameters.Select = new[]
                {
                    "id", "deviceName", "userDisplayName", "userPrincipalName",
                    "managedDeviceOwnerType", "operatingSystem", "osVersion",
                    "complianceState", "deviceEnrollmentType",
                    "lastSyncDateTime", "enrolledDateTime", "model", "manufacturer",
                    "serialNumber", "jailBroken", "isEncrypted", "isSupervised",
                    "deviceRegistrationState", "managementAgent", "totalStorageSpaceInBytes",
                    "freeStorageSpaceInBytes", "wiFiMacAddress", "ethernetMacAddress",
                    "imei", "phoneNumber", "azureADDeviceId", "azureADRegistered",
                    "deviceCategoryDisplayName"
                    // Note: skuFamily and skuNumber are only available in beta API
                };

                if (!string.IsNullOrEmpty(filter))
                {
                    config.QueryParameters.Filter = $"contains(deviceName, '{filter}') or contains(userDisplayName, '{filter}')";
                }

                if (!string.IsNullOrEmpty(orderBy))
                {
                    config.QueryParameters.Orderby = new[] { ascending ? orderBy : $"{orderBy} desc" };
                }
            });

            var deviceList = response?.Value ?? new List<ManagedDevice>();

            var devices = deviceList.Select(d => {
                // Note: skuFamily and skuNumber are only available in beta API
                // For v1.0 API, we use Model as a fallback display value
                // To get Windows edition (Pro/Enterprise/Home/Education), would need beta API
                
                return new IntuneDeviceDto(
                    Id: d.Id ?? string.Empty,
                    DeviceName: d.DeviceName ?? "Unknown",
                    UserDisplayName: d.UserDisplayName,
                    UserPrincipalName: d.UserPrincipalName,
                    ManagedDeviceOwnerType: d.ManagedDeviceOwnerType?.ToString(),
                    OperatingSystem: d.OperatingSystem,
                    OsVersion: d.OsVersion,
                    ComplianceState: d.ComplianceState?.ToString(),
                    ManagementState: null,
                    DeviceEnrollmentType: d.DeviceEnrollmentType?.ToString(),
                    LastSyncDateTime: d.LastSyncDateTime?.DateTime,
                    EnrolledDateTime: d.EnrolledDateTime?.DateTime,
                    Model: d.Model,
                    Manufacturer: d.Manufacturer,
                    SerialNumber: d.SerialNumber,
                    JailBroken: d.JailBroken,
                    IsEncrypted: d.IsEncrypted,
                    IsSupervised: d.IsSupervised,
                    DeviceRegistrationState: d.DeviceRegistrationState?.ToString(),
                    ManagementAgent: d.ManagementAgent?.ToString(),
                    TotalStorageSpaceInBytes: d.TotalStorageSpaceInBytes,
                    FreeStorageSpaceInBytes: d.FreeStorageSpaceInBytes,
                    WiFiMacAddress: d.WiFiMacAddress,
                    EthernetMacAddress: d.EthernetMacAddress,
                    Imei: d.Imei,
                    PhoneNumber: d.PhoneNumber,
                    AzureAdDeviceId: d.AzureADDeviceId,
                    AzureAdRegistered: d.AzureADRegistered,
                    DeviceCategoryDisplayName: d.DeviceCategoryDisplayName,
                    SkuFamily: null, // Only available via beta API
                    WindowsEdition: null // Only available via beta API (derived from skuNumber)
                );
            }).ToList();

            return new DeviceListResultDto(
                Devices: devices,
                TotalCount: devices.Count,
                FilteredCount: devices.Count,
                NextLink: null
            );
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error retrieving Intune devices");
            throw;
        }
    }

    public async Task<DeviceDetailDto> GetDeviceDetailsAsync(string deviceId)
    {
        try
        {
            _logger.LogInformation("Fetching device details for {DeviceId}", deviceId);

            var device = await _graphClient.DeviceManagement.ManagedDevices[deviceId].GetAsync();

            if (device == null)
            {
                throw new InvalidOperationException($"Device {deviceId} not found");
            }

            return new DeviceDetailDto(
                Id: device.Id ?? string.Empty,
                DeviceName: device.DeviceName ?? "Unknown",
                UserDisplayName: device.UserDisplayName,
                UserPrincipalName: device.UserPrincipalName,
                UserId: device.UserId,
                EmailAddress: device.EmailAddress,
                ManagedDeviceOwnerType: device.ManagedDeviceOwnerType?.ToString(),
                OperatingSystem: device.OperatingSystem,
                OsVersion: device.OsVersion,
                ComplianceState: device.ComplianceState?.ToString(),
                ManagementState: null,
                DeviceEnrollmentType: device.DeviceEnrollmentType?.ToString(),
                LastSyncDateTime: device.LastSyncDateTime?.DateTime,
                EnrolledDateTime: device.EnrolledDateTime?.DateTime,
                ComplianceGracePeriodExpirationDateTime: device.ComplianceGracePeriodExpirationDateTime?.DateTime,
                Model: device.Model,
                Manufacturer: device.Manufacturer,
                SerialNumber: device.SerialNumber,
                JailBroken: device.JailBroken,
                IsEncrypted: device.IsEncrypted,
                IsSupervised: device.IsSupervised,
                DeviceRegistrationState: device.DeviceRegistrationState?.ToString(),
                ManagementAgent: device.ManagementAgent?.ToString(),
                TotalStorageSpaceInBytes: device.TotalStorageSpaceInBytes,
                FreeStorageSpaceInBytes: device.FreeStorageSpaceInBytes,
                WiFiMacAddress: device.WiFiMacAddress,
                EthernetMacAddress: device.EthernetMacAddress,
                Imei: device.Imei,
                Meid: device.Meid,
                PhoneNumber: device.PhoneNumber,
                SubscriberCarrier: device.SubscriberCarrier,
                AzureAdDeviceId: device.AzureADDeviceId,
                AzureAdRegistered: device.AzureADRegistered,
                DeviceCategoryDisplayName: device.DeviceCategoryDisplayName,
                ConfigurationManagerClientEnabledFeatures: null,
                Notes: device.Notes
            );
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error retrieving device details for {DeviceId}", deviceId);
            throw;
        }
    }

    public async Task<DeviceStatsDto> GetDeviceStatsAsync()
    {
        try
        {
            _logger.LogInformation("Calculating device statistics");

            var allDevices = new List<ManagedDevice>();

            var response = await _graphClient.DeviceManagement.ManagedDevices.GetAsync(config =>
            {
                config.QueryParameters.Select = new[] 
                { 
                    "id", "complianceState", "operatingSystem", "managedDeviceOwnerType", 
                    "isEncrypted", "managementAgent" 
                };
                config.QueryParameters.Top = 999;
            });

            if (response?.Value != null)
            {
                allDevices.AddRange(response.Value);
            }

            // Page through results
            while (response?.OdataNextLink != null)
            {
                response = await _graphClient.DeviceManagement.ManagedDevices
                    .WithUrl(response.OdataNextLink)
                    .GetAsync();
                if (response?.Value != null)
                {
                    allDevices.AddRange(response.Value);
                }
            }

            return new DeviceStatsDto(
                TotalDevices: allDevices.Count,
                CompliantDevices: allDevices.Count(d => d.ComplianceState == ComplianceState.Compliant),
                NonCompliantDevices: allDevices.Count(d => d.ComplianceState == ComplianceState.Noncompliant),
                InGracePeriod: allDevices.Count(d => d.ComplianceState == ComplianceState.InGracePeriod),
                ConfigurationManagerDevices: allDevices.Count(d => d.ComplianceState == ComplianceState.ConfigManager),
                WindowsDevices: allDevices.Count(d => d.OperatingSystem?.Contains("Windows", StringComparison.OrdinalIgnoreCase) == true),
                MacOsDevices: allDevices.Count(d => d.OperatingSystem?.Contains("macOS", StringComparison.OrdinalIgnoreCase) == true || 
                                                      d.OperatingSystem?.Contains("Mac OS", StringComparison.OrdinalIgnoreCase) == true),
                IosDevices: allDevices.Count(d => d.OperatingSystem?.Contains("iOS", StringComparison.OrdinalIgnoreCase) == true ||
                                                   d.OperatingSystem?.Contains("iPadOS", StringComparison.OrdinalIgnoreCase) == true),
                AndroidDevices: allDevices.Count(d => d.OperatingSystem?.Contains("Android", StringComparison.OrdinalIgnoreCase) == true),
                LinuxDevices: allDevices.Count(d => d.OperatingSystem?.Contains("Linux", StringComparison.OrdinalIgnoreCase) == true),
                CorporateDevices: allDevices.Count(d => d.ManagedDeviceOwnerType == ManagedDeviceOwnerType.Company),
                PersonalDevices: allDevices.Count(d => d.ManagedDeviceOwnerType == ManagedDeviceOwnerType.Personal),
                ManagedDevices: allDevices.Count, // All devices in Intune are managed
                EncryptedDevices: allDevices.Count(d => d.IsEncrypted == true),
                LastUpdated: DateTime.UtcNow
            );
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error calculating device statistics");
            throw;
        }
    }

    #endregion

    #region Mailflow Methods

    public async Task<MailboxListResultDto> GetMailboxesAsync(string? filter, string? orderBy, bool ascending, int take)
    {
        try
        {
            _logger.LogInformation("Fetching mailboxes from tenant");

            // Get users with mail property to identify mailboxes
            // Filter out guest users - they don't have real mailboxes
            var response = await _graphClient.Users.GetAsync(config =>
            {
                config.QueryParameters.Select = new[]
                {
                    "id", "displayName", "userPrincipalName", "mail", "userType",
                    "accountEnabled", "createdDateTime", "assignedLicenses", "proxyAddresses"
                };
                config.QueryParameters.Top = take;
                config.QueryParameters.Count = true;
                config.Headers.Add("ConsistencyLevel", "eventual");
                config.QueryParameters.Filter = "mail ne null and userType eq 'Member'";

                if (!string.IsNullOrEmpty(orderBy))
                {
                    config.QueryParameters.Orderby = new[] { ascending ? orderBy : $"{orderBy} desc" };
                }
            });

            var userList = response?.Value ?? new List<User>();

            // Apply additional search filter
            if (!string.IsNullOrEmpty(filter))
            {
                var lowerFilter = filter.ToLower();
                userList = userList.Where(u =>
                    (u.DisplayName?.ToLower().Contains(lowerFilter) ?? false) ||
                    (u.UserPrincipalName?.ToLower().Contains(lowerFilter) ?? false) ||
                    (u.Mail?.ToLower().Contains(lowerFilter) ?? false)
                ).ToList();
            }

            var mailboxes = userList.Select(u => new MailboxDto(
                Id: u.Id ?? string.Empty,
                DisplayName: u.DisplayName ?? "Unknown",
                UserPrincipalName: u.UserPrincipalName ?? string.Empty,
                Mail: u.Mail,
                RecipientType: "UserMailbox",
                RecipientTypeDetails: "UserMailbox",
                WhenCreated: u.CreatedDateTime?.DateTime,
                WhenMailboxCreated: u.CreatedDateTime?.DateTime,
                HiddenFromAddressListsEnabled: false,
                IsMailboxEnabled: u.AccountEnabled,
                ProhibitSendQuota: null,
                ProhibitSendReceiveQuota: null,
                IssueWarningQuota: null,
                TotalItemSize: null,
                ItemCount: null,
                LastLogonTime: null,
                PrimarySmtpAddress: u.Mail,
                EmailAddresses: u.ProxyAddresses?.Where(p => p.StartsWith("smtp:", StringComparison.OrdinalIgnoreCase) || p.StartsWith("SMTP:")).ToList(),
                ForwardingAddress: null,
                ForwardingSmtpAddress: null,
                DeliverToMailboxAndForward: null,
                ArchiveStatus: null,
                ArchiveQuota: null,
                LitigationHoldEnabled: null,
                RetentionPolicy: null
            )).ToList();

            return new MailboxListResultDto(
                Mailboxes: mailboxes,
                TotalCount: (int)(response?.OdataCount ?? mailboxes.Count),
                FilteredCount: mailboxes.Count,
                NextLink: null
            );
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error retrieving mailboxes");
            throw;
        }
    }

    public async Task<MailboxStatsDto> GetMailboxStatsAsync()
    {
        try
        {
            _logger.LogInformation("Calculating mailbox statistics");

            // Get mailbox usage details from reports API
            long totalStorageUsed = 0;
            int totalMailboxes = 0;
            int activeMailboxes = 0;
            int inactiveMailboxes = 0;
            int sharedMailboxes = 0;
            int roomMailboxes = 0;
            int equipmentMailboxes = 0;
            int mailboxesNearQuota = 0;

            try
            {
                var mailboxUsageResponse = await _graphClient.Reports
                    .GetMailboxUsageDetailWithPeriod("D30")
                    .GetAsync();

                if (mailboxUsageResponse != null)
                {
                    using var reader = new StreamReader(mailboxUsageResponse);
                    var csv = await reader.ReadToEndAsync();
                    var lines = csv.Split('\n');

                    // Parse header to find column indices
                    var header = lines.FirstOrDefault()?.Split(',') ?? Array.Empty<string>();
                    
                    _logger.LogInformation("Mailbox Usage CSV Headers: {Headers}", 
                        string.Join(" | ", header.Select((h, i) => $"[{i}]={h.Trim()}")));

                    var storageUsedIdx = Array.FindIndex(header, h => 
                        h.Trim().Equals("Storage Used (Byte)", StringComparison.OrdinalIgnoreCase));
                    var recipientTypeIdx = Array.FindIndex(header, h => 
                        h.Trim().Equals("Recipient Type", StringComparison.OrdinalIgnoreCase));
                    var lastActivityIdx = Array.FindIndex(header, h => 
                        h.Trim().Equals("Last Activity Date", StringComparison.OrdinalIgnoreCase));
                    var issueWarningQuotaIdx = Array.FindIndex(header, h => 
                        h.Trim().Equals("Issue Warning Quota (Byte)", StringComparison.OrdinalIgnoreCase));
                    var prohibitSendQuotaIdx = Array.FindIndex(header, h => 
                        h.Trim().Equals("Prohibit Send Quota (Byte)", StringComparison.OrdinalIgnoreCase));
                    var isDeletedIdx = Array.FindIndex(header, h => 
                        h.Trim().Equals("Is Deleted", StringComparison.OrdinalIgnoreCase));

                    _logger.LogInformation("Column indices - StorageUsed: {StorageIdx}, RecipientType: {TypeIdx}, LastActivity: {ActivityIdx}",
                        storageUsedIdx, recipientTypeIdx, lastActivityIdx);

                    var thirtyDaysAgo = DateTime.UtcNow.AddDays(-30);

                    foreach (var line in lines.Skip(1).Where(l => !string.IsNullOrWhiteSpace(l)))
                    {
                        var parts = line.Split(',');
                        
                        // Check if deleted
                        var isDeleted = isDeletedIdx >= 0 && isDeletedIdx < parts.Length &&
                            parts[isDeletedIdx]?.Trim().Equals("TRUE", StringComparison.OrdinalIgnoreCase) == true;
                        
                        if (isDeleted) continue;

                        totalMailboxes++;

                        // Get storage used
                        if (storageUsedIdx >= 0 && storageUsedIdx < parts.Length)
                        {
                            if (long.TryParse(parts[storageUsedIdx]?.Trim(), out var storage))
                            {
                                totalStorageUsed += storage;

                                // Check if near quota (>80% of prohibit send quota)
                                if (prohibitSendQuotaIdx >= 0 && prohibitSendQuotaIdx < parts.Length)
                                {
                                    if (long.TryParse(parts[prohibitSendQuotaIdx]?.Trim(), out var quota) && quota > 0)
                                    {
                                        if ((double)storage / quota > 0.8)
                                        {
                                            mailboxesNearQuota++;
                                        }
                                    }
                                }
                            }
                        }

                        // Get recipient type
                        if (recipientTypeIdx >= 0 && recipientTypeIdx < parts.Length)
                        {
                            var recipientType = parts[recipientTypeIdx]?.Trim() ?? "";
                            if (recipientType.Contains("Shared", StringComparison.OrdinalIgnoreCase))
                                sharedMailboxes++;
                            else if (recipientType.Contains("Room", StringComparison.OrdinalIgnoreCase))
                                roomMailboxes++;
                            else if (recipientType.Contains("Equipment", StringComparison.OrdinalIgnoreCase))
                                equipmentMailboxes++;
                        }

                        // Check if active (had activity in last 30 days)
                        if (lastActivityIdx >= 0 && lastActivityIdx < parts.Length)
                        {
                            var lastActivityStr = parts[lastActivityIdx]?.Trim();
                            if (DateTime.TryParse(lastActivityStr, out var lastActivity))
                            {
                                if (lastActivity >= thirtyDaysAgo)
                                    activeMailboxes++;
                                else
                                    inactiveMailboxes++;
                            }
                            else
                            {
                                // No activity date means inactive
                                inactiveMailboxes++;
                            }
                        }
                    }

                    _logger.LogInformation("Parsed mailbox usage: Total={Total}, Active={Active}, Storage={StorageGB} GB",
                        totalMailboxes, activeMailboxes, Math.Round(totalStorageUsed / 1073741824.0, 2));
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "Could not fetch mailbox usage report, falling back to user count");
                
                // Fallback to user count
                var allUsers = new List<User>();
                var response = await _graphClient.Users.GetAsync(config =>
                {
                    config.QueryParameters.Select = new[] { "id", "mail", "userType", "accountEnabled" };
                    config.QueryParameters.Top = 999;
                    config.QueryParameters.Filter = "mail ne null and userType eq 'Member'";
                    config.Headers.Add("ConsistencyLevel", "eventual");
                    config.QueryParameters.Count = true;
                });

                if (response?.Value != null)
                {
                    allUsers.AddRange(response.Value);
                }

                while (response?.OdataNextLink != null)
                {
                    response = await _graphClient.Users.WithUrl(response.OdataNextLink).GetAsync();
                    if (response?.Value != null)
                    {
                        allUsers.AddRange(response.Value);
                    }
                }

                totalMailboxes = allUsers.Count;
                activeMailboxes = allUsers.Count(u => u.AccountEnabled == true);
                inactiveMailboxes = allUsers.Count(u => u.AccountEnabled == false);
            }

            // Calculate user mailboxes (total - shared - room - equipment)
            var userMailboxes = totalMailboxes - sharedMailboxes - roomMailboxes - equipmentMailboxes;

            return new MailboxStatsDto(
                TotalMailboxes: totalMailboxes,
                UserMailboxes: userMailboxes > 0 ? userMailboxes : totalMailboxes,
                SharedMailboxes: sharedMailboxes,
                RoomMailboxes: roomMailboxes,
                EquipmentMailboxes: equipmentMailboxes,
                ActiveMailboxes: activeMailboxes,
                InactiveMailboxes: inactiveMailboxes,
                TotalStorageUsedBytes: totalStorageUsed,
                MailboxesNearQuota: mailboxesNearQuota,
                MailboxesWithForwarding: 0,
                MailboxesOnHold: 0,
                MailboxesWithArchive: 0,
                LastUpdated: DateTime.UtcNow
            );
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error calculating mailbox statistics");
            throw;
        }
    }

    public async Task<MailflowSummaryDto> GetMailflowSummaryAsync(int days)
    {
        try
        {
            _logger.LogInformation("Fetching mailflow summary for {Days} days", days);

            var period = days switch
            {
                <= 7 => "D7",
                <= 30 => "D30",
                <= 90 => "D90",
                _ => "D180"
            };

            // Get email activity counts
            var activityResponse = await _graphClient.Reports
                .GetEmailActivityCountsWithPeriod(period)
                .GetAsync();

            var dailyTraffic = new List<MailTrafficReportDto>();
            long totalSent = 0;
            long totalReceived = 0;

            if (activityResponse != null)
            {
                using var reader = new StreamReader(activityResponse);
                var csv = await reader.ReadToEndAsync();
                
                _logger.LogInformation("Email Activity CSV (first 500 chars): {Csv}", csv.Length > 500 ? csv.Substring(0, 500) : csv);
                
                var lines = csv.Split('\n');
                
                // Parse header to find column indices
                var header = lines.FirstOrDefault()?.Split(',') ?? Array.Empty<string>();
                
                _logger.LogInformation("CSV Headers: {Headers}", string.Join(" | ", header.Select((h, i) => $"[{i}]={h.Trim()}")));
                
                var reportDateIdx = Array.FindIndex(header, h => h.Trim().Equals("Report Date", StringComparison.OrdinalIgnoreCase));
                var sendIdx = Array.FindIndex(header, h => h.Trim().Equals("Send", StringComparison.OrdinalIgnoreCase));
                var receiveIdx = Array.FindIndex(header, h => h.Trim().Equals("Receive", StringComparison.OrdinalIgnoreCase));
                var readIdx = Array.FindIndex(header, h => h.Trim().Equals("Read", StringComparison.OrdinalIgnoreCase));

                _logger.LogInformation("Column indices - ReportDate: {DateIdx}, Send: {SendIdx}, Receive: {ReceiveIdx}, Read: {ReadIdx}", 
                    reportDateIdx, sendIdx, receiveIdx, readIdx);

                foreach (var line in lines.Skip(1).Where(l => !string.IsNullOrWhiteSpace(l)))
                {
                    var parts = line.Split(',');
                    
                    // Get date from the correct column
                    var dateStr = reportDateIdx >= 0 && reportDateIdx < parts.Length 
                        ? parts[reportDateIdx]?.Trim() 
                        : parts.ElementAtOrDefault(0)?.Trim();
                    
                    if (DateTime.TryParse(dateStr, out var date))
                    {
                        var sent = sendIdx >= 0 && sendIdx < parts.Length 
                            ? (long.TryParse(parts[sendIdx]?.Trim(), out var s) ? s : 0) 
                            : 0;
                        var received = receiveIdx >= 0 && receiveIdx < parts.Length 
                            ? (long.TryParse(parts[receiveIdx]?.Trim(), out var r) ? r : 0) 
                            : 0;

                        dailyTraffic.Add(new MailTrafficReportDto(
                            Date: date,
                            MessagesSent: sent,
                            MessagesReceived: received,
                            SpamReceived: 0,
                            MalwareReceived: 0,
                            GoodMail: sent + received
                        ));

                        totalSent += sent;
                        totalReceived += received;
                    }
                    else
                    {
                        _logger.LogWarning("Could not parse date from: {DateStr} in line: {Line}", dateStr, line.Length > 100 ? line.Substring(0, 100) : line);
                    }
                }
                
                _logger.LogInformation("Parsed {Count} daily traffic records. Date range: {MinDate} to {MaxDate}", 
                    dailyTraffic.Count,
                    dailyTraffic.MinBy(d => d.Date)?.Date.ToString("yyyy-MM-dd") ?? "N/A",
                    dailyTraffic.MaxBy(d => d.Date)?.Date.ToString("yyyy-MM-dd") ?? "N/A");
            }

            // Get top senders from user activity report
            var topSenders = new List<TopSenderDto>();
            var topRecipients = new List<TopRecipientDto>();

            try
            {
                var userActivityResponse = await _graphClient.Reports
                    .GetEmailActivityUserDetailWithPeriod(period)
                    .GetAsync();

                if (userActivityResponse != null)
                {
                    using var reader = new StreamReader(userActivityResponse);
                    var csv = await reader.ReadToEndAsync();
                    var lines = csv.Split('\n');
                    
                    // Parse header
                    var header = lines.FirstOrDefault()?.Split(',') ?? Array.Empty<string>();
                    var upnIdx = Array.FindIndex(header, h => h.Trim().Equals("User Principal Name", StringComparison.OrdinalIgnoreCase));
                    var displayNameIdx = Array.FindIndex(header, h => h.Trim().Equals("Display Name", StringComparison.OrdinalIgnoreCase));
                    var sendCountIdx = Array.FindIndex(header, h => h.Trim().Equals("Send Count", StringComparison.OrdinalIgnoreCase));
                    var receiveCountIdx = Array.FindIndex(header, h => h.Trim().Equals("Receive Count", StringComparison.OrdinalIgnoreCase));

                    var userActivities = new List<(string Upn, string DisplayName, long Sent, long Received)>();

                    foreach (var line in lines.Skip(1).Where(l => !string.IsNullOrWhiteSpace(l)))
                    {
                        var parts = line.Split(',');
                        
                        var upn = upnIdx >= 0 && upnIdx < parts.Length ? parts[upnIdx]?.Trim('"', ' ') ?? "" : "";
                        var displayName = displayNameIdx >= 0 && displayNameIdx < parts.Length ? parts[displayNameIdx]?.Trim('"', ' ') ?? "" : "";
                        var sent = sendCountIdx >= 0 && sendCountIdx < parts.Length 
                            ? (long.TryParse(parts[sendCountIdx]?.Trim(), out var s) ? s : 0) 
                            : 0;
                        var received = receiveCountIdx >= 0 && receiveCountIdx < parts.Length 
                            ? (long.TryParse(parts[receiveCountIdx]?.Trim(), out var r) ? r : 0) 
                            : 0;

                        if (!string.IsNullOrEmpty(upn) && (sent > 0 || received > 0))
                        {
                            userActivities.Add((upn, displayName, sent, received));
                        }
                    }

                    topSenders = userActivities
                        .OrderByDescending(u => u.Sent)
                        .Take(10)
                        .Select(u => new TopSenderDto(u.Upn, u.DisplayName, u.Sent))
                        .ToList();

                    topRecipients = userActivities
                        .OrderByDescending(u => u.Received)
                        .Take(10)
                        .Select(u => new TopRecipientDto(u.Upn, u.DisplayName, u.Received))
                        .ToList();
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "Could not fetch user activity details");
            }

            var avgPerDay = dailyTraffic.Count > 0
                ? (totalSent + totalReceived) / (double)dailyTraffic.Count
                : 0;

            return new MailflowSummaryDto(
                TotalMessagesSent: totalSent,
                TotalMessagesReceived: totalReceived,
                TotalSpamBlocked: 0, // Would need security reports
                TotalMalwareBlocked: 0,
                AverageMessagesPerDay: Math.Round(avgPerDay, 0),
                DailyTraffic: dailyTraffic.OrderBy(d => d.Date).ToList(),
                TopSenders: topSenders,
                TopRecipients: topRecipients,
                LastUpdated: DateTime.UtcNow
            );
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching mailflow summary");
            throw;
        }
    }

    #endregion

    #region Security Methods

    public async Task<SecurityOverviewDto> GetSecurityOverviewAsync()
    {
        try
        {
            _logger.LogInformation("Fetching security overview");

            // Fetch all data in parallel
            var secureScoreTask = GetSecureScoreAsync();
            var riskyUsersTask = GetRiskyUsersAsync();
            var riskySignInsTask = GetRiskySignInsAsync(24);
            var statsTask = GetSecurityStatsAsync();

            await Task.WhenAll(secureScoreTask, riskyUsersTask, riskySignInsTask, statsTask);

            return new SecurityOverviewDto(
                SecureScore: await secureScoreTask,
                Stats: await statsTask,
                RiskyUsers: await riskyUsersTask,
                RiskySignIns: await riskySignInsTask,
                RecentAlerts: new List<SecurityAlertDto>(), // Alerts require additional permissions
                LastUpdated: DateTime.UtcNow
            );
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching security overview");
            throw;
        }
    }

    public async Task<SecurityScoreDto?> GetSecureScoreAsync()
    {
        try
        {
            _logger.LogInformation("Fetching secure score");

            var secureScores = await _graphClient.Security.SecureScores.GetAsync(config =>
            {
                config.QueryParameters.Top = 1;
                config.QueryParameters.Orderby = new[] { "createdDateTime desc" };
            });

            var latestScore = secureScores?.Value?.FirstOrDefault();
            if (latestScore == null)
            {
                return null;
            }

            // Log raw control score data to diagnose the issue
            if (latestScore.ControlScores?.Any() == true)
            {
                var firstControl = latestScore.ControlScores.First();
                _logger.LogInformation("First control - Name: {Name}, Score: {Score}, AdditionalData keys: {Keys}",
                    firstControl.ControlName,
                    firstControl.Score,
                    firstControl.AdditionalData != null ? string.Join(", ", firstControl.AdditionalData.Keys) : "null");
                
                if (firstControl.AdditionalData != null)
                {
                    foreach (var kvp in firstControl.AdditionalData)
                    {
                        _logger.LogInformation("  AdditionalData[{Key}] = {Value} (Type: {Type})", 
                            kvp.Key, kvp.Value, kvp.Value?.GetType().Name ?? "null");
                    }
                }
            }

            var controlScores = latestScore.ControlScores?.Select(cs => {
                // Try to extract score from multiple possible locations
                double controlScore = cs.Score ?? 0;
                double controlMaxScore = 0;
                
                // Check AdditionalData for score and maxScore
                if (cs.AdditionalData != null)
                {
                    // Try to get score if it's 0 from the main property
                    if (controlScore == 0 && cs.AdditionalData.TryGetValue("score", out var scoreObj))
                    {
                        controlScore = ConvertToDouble(scoreObj);
                    }
                    
                    // Get maxScore directly if available
                    if (cs.AdditionalData.TryGetValue("maxScore", out var maxScoreObj))
                    {
                        controlMaxScore = ConvertToDouble(maxScoreObj);
                    }
                    
                    // If maxScore is still 0, calculate it from scoreInPercentage
                    // Formula: if score is X and that represents Y%, then maxScore = X / (Y/100)
                    if (controlMaxScore == 0 && cs.AdditionalData.TryGetValue("scoreInPercentage", out var percentObj))
                    {
                        var percent = ConvertToDouble(percentObj);
                        if (percent > 0 && controlScore > 0)
                        {
                            // score / (percent/100) = maxScore
                            // e.g., score=5, percent=100 -> maxScore = 5 / 1.0 = 5
                            // e.g., score=3, percent=60 -> maxScore = 3 / 0.6 = 5
                            controlMaxScore = Math.Round(controlScore / (percent / 100.0), 2);
                        }
                        else if (percent == 0 && controlScore == 0)
                        {
                            // Control not implemented - try to estimate maxScore from control name patterns
                            // Most controls have maxScore between 1-10, default to reasonable estimate
                            controlMaxScore = 5; // Default estimate
                        }
                    }
                }
                
                return new SecurityControlScoreDto(
                    ControlName: cs.ControlName ?? "Unknown",
                    ControlCategory: cs.ControlCategory ?? "Unknown",
                    Description: cs.Description,
                    Score: controlScore,
                    MaxScore: controlMaxScore,
                    Implementation: null,
                    UserImpact: null,
                    Threats: null
                );
            }).ToList() ?? new List<SecurityControlScoreDto>();

            var currentScore = latestScore.CurrentScore ?? 0;
            var maxScoreTotal = latestScore.MaxScore ?? 100;
            
            // Calculate percentage - don't round intermediate values
            var percentage = maxScoreTotal > 0 ? (currentScore / maxScoreTotal) * 100.0 : 0;

            _logger.LogInformation("Secure Score: {Current} / {Max} = {Percentage}% with {ControlCount} controls", 
                currentScore, maxScoreTotal, percentage, controlScores.Count);
            
            // Log a sample of control scores for debugging
            foreach (var ctrl in controlScores.Take(3))
            {
                _logger.LogInformation("  Control: {Name}, Score: {Score} / {Max}", 
                    ctrl.ControlName, ctrl.Score, ctrl.MaxScore);
            }

            // Fetch control profiles to get accurate maxScore per control
            // This avoids the unreliable scoreInPercentage estimation
            Dictionary<string, double> profileMaxScores = new(StringComparer.OrdinalIgnoreCase);
            try
            {
                var profiles = await _graphClient.Security.SecureScoreControlProfiles.GetAsync(config =>
                {
                    config.QueryParameters.Top = 200;
                    config.QueryParameters.Select = new[] { "controlName", "maxScore", "controlCategory" };
                });
                if (profiles?.Value != null)
                {
                    foreach (var p in profiles.Value)
                    {
                        if (!string.IsNullOrEmpty(p.ControlName))
                            profileMaxScores[p.ControlName] = p.MaxScore ?? 0;
                    }
                    _logger.LogInformation("Fetched {Count} control profiles for maxScore lookup", profiles.Value.Count);
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "Could not fetch control profiles - category percentages may be inaccurate");
            }

            // Calculate category scores using actual maxScores from profiles
            double CalcScore(string category) =>
                controlScores.Where(c => string.Equals(c.ControlCategory, category, StringComparison.OrdinalIgnoreCase)).Sum(c => c.Score);
            double CalcMax(string category) =>
                controlScores
                    .Where(c => string.Equals(c.ControlCategory, category, StringComparison.OrdinalIgnoreCase))
                    .Sum(c => profileMaxScores.TryGetValue(c.ControlName, out var m) && m > 0 ? m : c.MaxScore);
            double CalcPct(double score, double max) => max > 0 ? Math.Round(score / max * 100, 1) : 0;

            var idScore  = CalcScore("Identity"); var idMax  = CalcMax("Identity");
            var devScore = CalcScore("Device");   var devMax = CalcMax("Device");
            var appScore = CalcScore("Apps");     var appMax = CalcMax("Apps");
            var datScore = CalcScore("Data");     var datMax = CalcMax("Data");

            _logger.LogInformation("Category scores - Identity: {IS}/{IM} ({IP}%), Device: {DS}/{DM}, Apps: {AS}/{AM}",
                idScore, idMax, CalcPct(idScore, idMax), devScore, devMax, appScore, appMax);

            return new SecurityScoreDto(
                CurrentScore: currentScore,
                MaxScore: maxScoreTotal,
                PercentageScore: percentage,
                ControlScores: controlScores,
                LastUpdated: latestScore.CreatedDateTime?.DateTime ?? DateTime.UtcNow,
                IdentityScore: idScore,  IdentityMaxScore: idMax,  IdentityPercentage: CalcPct(idScore, idMax),
                DeviceScore: devScore,   DeviceMaxScore: devMax,   DevicePercentage: CalcPct(devScore, devMax),
                AppsScore: appScore,     AppsMaxScore: appMax,     AppsPercentage: CalcPct(appScore, appMax),
                DataScore: datScore,     DataMaxScore: datMax,     DataPercentage: CalcPct(datScore, datMax)
            );
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Could not fetch secure score - may require SecurityEvents.Read.All permission");
            return null;
        }
    }
    
    /// <summary>
    /// Helper to convert various numeric types to double
    /// </summary>
    private static double ConvertToDouble(object? value)
    {
        if (value == null) return 0;
        
        return value switch
        {
            double d => d,
            float f => f,
            int i => i,
            long l => l,
            decimal dec => (double)dec,
            System.Text.Json.JsonElement je => je.ValueKind == System.Text.Json.JsonValueKind.Number 
                ? je.GetDouble() 
                : 0,
            _ => double.TryParse(value.ToString(), out var parsed) ? parsed : 0
        };
    }

    public async Task<List<RiskyUserDto>> GetRiskyUsersAsync()
    {
        try
        {
            _logger.LogInformation("Fetching risky users");

            var riskyUsers = await _graphClient.IdentityProtection.RiskyUsers.GetAsync(config =>
            {
                config.QueryParameters.Top = 100;
                config.QueryParameters.Filter = "riskState ne 'remediated' and riskState ne 'dismissed'";
                config.QueryParameters.Orderby = new[] { "riskLastUpdatedDateTime desc" };
            });

            return riskyUsers?.Value?.Select(ru => new RiskyUserDto(
                Id: ru.Id ?? string.Empty,
                UserPrincipalName: ru.UserPrincipalName ?? "Unknown",
                DisplayName: ru.UserDisplayName,
                RiskLevel: ru.RiskLevel?.ToString() ?? "Unknown",
                RiskState: ru.RiskState?.ToString() ?? "Unknown",
                RiskDetail: ru.RiskDetail?.ToString(),
                RiskLastUpdatedDateTime: ru.RiskLastUpdatedDateTime?.DateTime,
                IsDeleted: ru.IsDeleted ?? false,
                IsProcessing: ru.IsProcessing ?? false
            )).ToList() ?? new List<RiskyUserDto>();
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Could not fetch risky users - may require IdentityRiskyUser.Read.All permission");
            return new List<RiskyUserDto>();
        }
    }

    public async Task<List<RiskySignInDto>> GetRiskySignInsAsync(int hours = 24)
    {
        try
        {
            _logger.LogInformation("Fetching risky sign-ins from last {Hours} hours", hours);

            var cutoffTime = DateTime.UtcNow.AddHours(-hours);

            var signIns = await _graphClient.AuditLogs.SignIns.GetAsync(config =>
            {
                config.QueryParameters.Top = 100;
                config.QueryParameters.Filter = $"createdDateTime ge {cutoffTime:yyyy-MM-ddTHH:mm:ssZ} and (riskLevelDuringSignIn eq 'high' or riskLevelDuringSignIn eq 'medium' or riskLevelDuringSignIn eq 'low')";
                config.QueryParameters.Orderby = new[] { "createdDateTime desc" };
            });

            return signIns?.Value?.Select(si => new RiskySignInDto(
                Id: si.Id ?? string.Empty,
                UserPrincipalName: si.UserPrincipalName ?? "Unknown",
                DisplayName: si.UserDisplayName,
                CreatedDateTime: si.CreatedDateTime?.DateTime,
                IpAddress: si.IpAddress,
                Location: si.Location != null ? $"{si.Location.City}, {si.Location.CountryOrRegion}" : null,
                RiskLevel: si.RiskLevelDuringSignIn?.ToString() ?? "None",
                RiskState: si.RiskState?.ToString() ?? "None",
                RiskDetail: si.RiskDetail?.ToString(),
                ClientAppUsed: si.ClientAppUsed,
                DeviceDetail: si.DeviceDetail?.Browser ?? si.DeviceDetail?.OperatingSystem
            )).ToList() ?? new List<RiskySignInDto>();
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Could not fetch risky sign-ins - may require AuditLog.Read.All permission");
            return new List<RiskySignInDto>();
        }
    }

    public async Task<SecurityStatsDto> GetSecurityStatsAsync()
    {
        try
        {
            _logger.LogInformation("Calculating security statistics");

            // Get risky users counts
            var riskyUsers = await GetRiskyUsersAsync();
            var highRisk = riskyUsers.Count(r => r.RiskLevel.Equals("High", StringComparison.OrdinalIgnoreCase));
            var mediumRisk = riskyUsers.Count(r => r.RiskLevel.Equals("Medium", StringComparison.OrdinalIgnoreCase));
            var lowRisk = riskyUsers.Count(r => r.RiskLevel.Equals("Low", StringComparison.OrdinalIgnoreCase));

            // Get risky sign-ins in last 24 hours
            var riskySignIns = await GetRiskySignInsAsync(24);

            // Get MFA stats from authentication methods report
            int mfaRegistered = 0;
            int mfaNotRegistered = 0;
            try
            {
                var mfaDetails = await _graphClient.Reports.AuthenticationMethods.UserRegistrationDetails.GetAsync(config =>
                {
                    config.QueryParameters.Top = 999;
                });

                var allMfaDetails = mfaDetails?.Value?.ToList() ?? new List<Microsoft.Graph.Models.UserRegistrationDetails>();
                
                // Page through if needed
                while (mfaDetails?.OdataNextLink != null)
                {
                    mfaDetails = await _graphClient.Reports.AuthenticationMethods.UserRegistrationDetails
                        .WithUrl(mfaDetails.OdataNextLink)
                        .GetAsync();
                    if (mfaDetails?.Value != null)
                    {
                        allMfaDetails.AddRange(mfaDetails.Value);
                    }
                }

                mfaRegistered = allMfaDetails.Count(m => m.IsMfaRegistered == true);
                mfaNotRegistered = allMfaDetails.Count(m => m.IsMfaRegistered != true);
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "Could not fetch MFA registration details");
            }

            var totalUsers = mfaRegistered + mfaNotRegistered;
            var mfaPercentage = totalUsers > 0 ? (double)mfaRegistered / totalUsers * 100 : 0;

            return new SecurityStatsDto(
                TotalRiskyUsers: riskyUsers.Count,
                HighRiskUsers: highRisk,
                MediumRiskUsers: mediumRisk,
                LowRiskUsers: lowRisk,
                UsersAtRisk: riskyUsers.Count(r => r.RiskState.Equals("AtRisk", StringComparison.OrdinalIgnoreCase)),
                RiskySignInsLast24Hours: riskySignIns.Count,
                ActiveAlerts: 0, // Would need Security.Alert.Read.All
                HighSeverityAlerts: 0,
                MediumSeverityAlerts: 0,
                LowSeverityAlerts: 0,
                MfaRegisteredUsers: mfaRegistered,
                MfaNotRegisteredUsers: mfaNotRegistered,
                MfaRegistrationPercentage: Math.Round(mfaPercentage, 1),
                LastUpdated: DateTime.UtcNow
            );
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error calculating security statistics");
            return new SecurityStatsDto(
                TotalRiskyUsers: 0,
                HighRiskUsers: 0,
                MediumRiskUsers: 0,
                LowRiskUsers: 0,
                UsersAtRisk: 0,
                RiskySignInsLast24Hours: 0,
                ActiveAlerts: 0,
                HighSeverityAlerts: 0,
                MediumSeverityAlerts: 0,
                LowSeverityAlerts: 0,
                MfaRegisteredUsers: 0,
                MfaNotRegisteredUsers: 0,
                MfaRegistrationPercentage: 0,
                LastUpdated: DateTime.UtcNow
            );
        }
    }

    public async Task<MfaRegistrationListDto> GetMfaRegistrationDetailsAsync()
    {
        try
        {
            _logger.LogInformation("Fetching MFA registration details for all users");

            var allMfaDetails = new List<Microsoft.Graph.Models.UserRegistrationDetails>();
            
            var mfaDetails = await _graphClient.Reports.AuthenticationMethods.UserRegistrationDetails.GetAsync(config =>
            {
                config.QueryParameters.Top = 999;
            });

            if (mfaDetails?.Value != null)
            {
                allMfaDetails.AddRange(mfaDetails.Value);
            }

            // Page through if needed
            while (mfaDetails?.OdataNextLink != null)
            {
                mfaDetails = await _graphClient.Reports.AuthenticationMethods.UserRegistrationDetails
                    .WithUrl(mfaDetails.OdataNextLink)
                    .GetAsync();
                if (mfaDetails?.Value != null)
                {
                    allMfaDetails.AddRange(mfaDetails.Value);
                }
            }

            // Get user display names
            var userNames = new Dictionary<string, string>();
            try
            {
                var users = await _graphClient.Users.GetAsync(config =>
                {
                    config.QueryParameters.Select = new[] { "id", "displayName", "userPrincipalName" };
                    config.QueryParameters.Top = 999;
                });

                if (users?.Value != null)
                {
                    foreach (var user in users.Value)
                    {
                        if (user.Id != null)
                        {
                            userNames[user.Id] = user.DisplayName ?? user.UserPrincipalName ?? "Unknown";
                        }
                    }
                }

                // Page through users
                while (users?.OdataNextLink != null)
                {
                    users = await _graphClient.Users.WithUrl(users.OdataNextLink).GetAsync();
                    if (users?.Value != null)
                    {
                        foreach (var user in users.Value)
                        {
                            if (user.Id != null)
                            {
                                userNames[user.Id] = user.DisplayName ?? user.UserPrincipalName ?? "Unknown";
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "Could not fetch user display names");
            }

            var mfaUsers = allMfaDetails.Select(m => {
                // Get default method from AdditionalData or use first registered method
                string? defaultMethod = null;
                if (m.AdditionalData?.TryGetValue("defaultMfaMethod", out var defaultMfaValue) == true)
                {
                    defaultMethod = defaultMfaValue?.ToString();
                }
                else if (m.MethodsRegistered?.Any() == true)
                {
                    defaultMethod = m.MethodsRegistered.FirstOrDefault(method => 
                        method.Contains("authenticator", StringComparison.OrdinalIgnoreCase) ||
                        method.Contains("fido", StringComparison.OrdinalIgnoreCase) ||
                        method.Contains("phone", StringComparison.OrdinalIgnoreCase));
                }
                
                return new MfaUserDetailDto(
                    Id: m.Id ?? string.Empty,
                    UserPrincipalName: m.UserPrincipalName ?? "Unknown",
                    DisplayName: userNames.GetValueOrDefault(m.Id ?? "", m.UserPrincipalName ?? "Unknown"),
                    IsMfaRegistered: m.IsMfaRegistered ?? false,
                    IsMfaCapable: m.IsMfaCapable ?? false,
                    DefaultMfaMethod: GetFriendlyMfaMethodName(defaultMethod),
                    MethodsRegistered: m.MethodsRegistered?.Select(method => GetFriendlyMfaMethodName(method) ?? method).ToList(),
                    IsAdmin: m.IsAdmin ?? false,
                    LastUpdated: null
                );
            }).OrderBy(u => u.IsMfaRegistered).ThenBy(u => u.DisplayName).ToList();

            var registeredCount = mfaUsers.Count(m => m.IsMfaRegistered);
            var notRegisteredCount = mfaUsers.Count(m => !m.IsMfaRegistered);
            var totalCount = mfaUsers.Count;
            var percentage = totalCount > 0 ? (double)registeredCount / totalCount * 100 : 0;

            return new MfaRegistrationListDto(
                Users: mfaUsers,
                TotalCount: mfaUsers.Count,
                MfaRegisteredCount: registeredCount,
                MfaNotRegisteredCount: notRegisteredCount,
                MfaRegistrationPercentage: Math.Round(percentage, 1),
                LastUpdated: DateTime.UtcNow
            );
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching MFA registration details");
            throw;
        }
    }

    #endregion

    #region App Registration Credentials

    public async Task<AppCredentialStatusDto> GetAppCredentialStatusAsync(int thresholdDays = 45)
    {
        try
        {
            _logger.LogInformation("Fetching app registration credentials with {ThresholdDays} day threshold", thresholdDays);

            var today = DateTime.UtcNow;
            var expirationThreshold = today.AddDays(thresholdDays);

            var expiringSecrets = new List<AppCredentialDetailDto>();
            var expiredSecrets = new List<AppCredentialDetailDto>();
            var expiringCertificates = new List<AppCredentialDetailDto>();
            var expiredCertificates = new List<AppCredentialDetailDto>();

            var appsWithExpiringSecrets = new HashSet<string>();
            var appsWithExpiredSecrets = new HashSet<string>();
            var appsWithExpiringCertificates = new HashSet<string>();
            var appsWithExpiredCertificates = new HashSet<string>();

            int totalApps = 0;

            // Fetch all applications
            var applications = await _graphClient.Applications.GetAsync(config =>
            {
                config.QueryParameters.Select = new[] { "id", "appId", "displayName", "passwordCredentials", "keyCredentials" };
                config.QueryParameters.Top = 999;
            });

            var allApps = applications?.Value ?? new List<Application>();

            // Page through if needed
            while (applications?.OdataNextLink != null)
            {
                applications = await _graphClient.Applications
                    .WithUrl(applications.OdataNextLink)
                    .GetAsync();
                if (applications?.Value != null)
                {
                    allApps.AddRange(applications.Value);
                }
            }

            totalApps = allApps.Count;
            _logger.LogInformation("Found {Count} app registrations to analyze", totalApps);

            foreach (var app in allApps)
            {
                var appId = app.AppId ?? app.Id ?? "Unknown";
                var appDisplayName = app.DisplayName ?? "Unknown";

                // Check password credentials (secrets)
                if (app.PasswordCredentials != null)
                {
                    foreach (var secret in app.PasswordCredentials)
                    {
                        if (secret.EndDateTime == null) continue;

                        var endDate = secret.EndDateTime.Value.DateTime;
                        var daysUntilExpiry = (int)(endDate - today).TotalDays;

                        if (endDate < today)
                        {
                            // Expired
                            expiredSecrets.Add(new AppCredentialDetailDto(
                                AppId: appId,
                                AppDisplayName: appDisplayName,
                                CredentialType: "Secret",
                                KeyId: secret.KeyId?.ToString(),
                                DisplayName: secret.DisplayName,
                                StartDateTime: secret.StartDateTime?.DateTime,
                                EndDateTime: endDate,
                                DaysUntilExpiry: daysUntilExpiry,
                                Status: "Expired"
                            ));
                            appsWithExpiredSecrets.Add(appId);
                        }
                        else if (endDate < expirationThreshold)
                        {
                            // Expiring within threshold
                            expiringSecrets.Add(new AppCredentialDetailDto(
                                AppId: appId,
                                AppDisplayName: appDisplayName,
                                CredentialType: "Secret",
                                KeyId: secret.KeyId?.ToString(),
                                DisplayName: secret.DisplayName,
                                StartDateTime: secret.StartDateTime?.DateTime,
                                EndDateTime: endDate,
                                DaysUntilExpiry: daysUntilExpiry,
                                Status: daysUntilExpiry <= 7 ? "Critical" : daysUntilExpiry <= 14 ? "Warning" : "Expiring"
                            ));
                            appsWithExpiringSecrets.Add(appId);
                        }
                    }
                }

                // Check key credentials (certificates)
                if (app.KeyCredentials != null)
                {
                    foreach (var cert in app.KeyCredentials)
                    {
                        if (cert.EndDateTime == null) continue;

                        var endDate = cert.EndDateTime.Value.DateTime;
                        var daysUntilExpiry = (int)(endDate - today).TotalDays;

                        if (endDate < today)
                        {
                            // Expired
                            expiredCertificates.Add(new AppCredentialDetailDto(
                                AppId: appId,
                                AppDisplayName: appDisplayName,
                                CredentialType: "Certificate",
                                KeyId: cert.KeyId?.ToString(),
                                DisplayName: cert.DisplayName,
                                StartDateTime: cert.StartDateTime?.DateTime,
                                EndDateTime: endDate,
                                DaysUntilExpiry: daysUntilExpiry,
                                Status: "Expired"
                            ));
                            appsWithExpiredCertificates.Add(appId);
                        }
                        else if (endDate < expirationThreshold)
                        {
                            // Expiring within threshold
                            expiringCertificates.Add(new AppCredentialDetailDto(
                                AppId: appId,
                                AppDisplayName: appDisplayName,
                                CredentialType: "Certificate",
                                KeyId: cert.KeyId?.ToString(),
                                DisplayName: cert.DisplayName,
                                StartDateTime: cert.StartDateTime?.DateTime,
                                EndDateTime: endDate,
                                DaysUntilExpiry: daysUntilExpiry,
                                Status: daysUntilExpiry <= 7 ? "Critical" : daysUntilExpiry <= 14 ? "Warning" : "Expiring"
                            ));
                            appsWithExpiringCertificates.Add(appId);
                        }
                    }
                }
            }

            // Sort results
            expiringSecrets = expiringSecrets.OrderBy(s => s.EndDateTime).ToList();
            expiredSecrets = expiredSecrets.OrderByDescending(s => s.EndDateTime).ToList();
            expiringCertificates = expiringCertificates.OrderBy(c => c.EndDateTime).ToList();
            expiredCertificates = expiredCertificates.OrderByDescending(c => c.EndDateTime).ToList();

            _logger.LogInformation(
                "App credentials analysis complete: {ExpiringSecrets} expiring secrets, {ExpiredSecrets} expired secrets, " +
                "{ExpiringCerts} expiring certificates, {ExpiredCerts} expired certificates",
                expiringSecrets.Count, expiredSecrets.Count, expiringCertificates.Count, expiredCertificates.Count);

            return new AppCredentialStatusDto(
                TotalApps: totalApps,
                AppsWithExpiringSecrets: appsWithExpiringSecrets.Count,
                AppsWithExpiredSecrets: appsWithExpiredSecrets.Count,
                AppsWithExpiringCertificates: appsWithExpiringCertificates.Count,
                AppsWithExpiredCertificates: appsWithExpiredCertificates.Count,
                ThresholdDays: thresholdDays,
                ExpiringSecrets: expiringSecrets,
                ExpiredSecrets: expiredSecrets,
                ExpiringCertificates: expiringCertificates,
                ExpiredCertificates: expiredCertificates,
                LastUpdated: DateTime.UtcNow
            );
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching app registration credentials");
            throw;
        }
    }

    /// <summary>
    /// Gets all public Microsoft 365 groups with owner and member counts.
    /// Used for CIS Control 1.2.1 (L2) compliance - ensuring only managed public groups exist.
    /// </summary>
    public async Task<PublicGroupsReportDto> GetPublicGroupsAsync()
    {
        _logger.LogInformation("Fetching public Microsoft 365 groups...");
        
        try
        {
            var publicGroups = new List<PublicGroupDto>();
            var allGroups = new List<Group>();
            
            // Get all groups with necessary properties
            var groups = await _graphClient.Groups.GetAsync(requestConfiguration =>
            {
                requestConfiguration.QueryParameters.Select = new[] 
                { 
                    "id", "displayName", "createdDateTime", "visibility", 
                    "groupTypes", "description" 
                };
                requestConfiguration.QueryParameters.Top = 999;
            });

            if (groups?.Value != null)
            {
                allGroups.AddRange(groups.Value);
            }

            // Page through all results
            while (groups?.OdataNextLink != null)
            {
                groups = await _graphClient.Groups
                    .WithUrl(groups.OdataNextLink)
                    .GetAsync();
                if (groups?.Value != null)
                {
                    allGroups.AddRange(groups.Value);
                }
            }

            _logger.LogInformation("Found {Count} total groups, filtering for public groups...", allGroups.Count);

            // Filter for public groups only
            var publicGroupsRaw = allGroups
                .Where(g => g.Visibility?.Equals("Public", StringComparison.OrdinalIgnoreCase) == true)
                .OrderByDescending(g => g.CreatedDateTime)
                .ToList();

            _logger.LogInformation("Found {Count} public groups, fetching owner/member counts...", publicGroupsRaw.Count);

            // Get owner and member counts for each public group
            foreach (var group in publicGroupsRaw)
            {
                // Check if it's a Team by trying to access the team endpoint
                var isTeam = false;
                if (group.GroupTypes?.Contains("Unified") == true)
                {
                    try
                    {
                        var team = await _graphClient.Teams[group.Id].GetAsync();
                        isTeam = team != null;
                    }
                    catch
                    {
                        // Not a team or no access
                    }
                }
                
                var groupType = isTeam ? "Team" : "Microsoft 365 Group";

                int ownerCount = 0;
                int memberCount = 0;

                try
                {
                    // Get owners count
                    var owners = await _graphClient.Groups[group.Id].Owners.GetAsync(config =>
                    {
                        config.QueryParameters.Count = true;
                        config.Headers.Add("ConsistencyLevel", "eventual");
                    });
                    ownerCount = owners?.Value?.Count ?? 0;
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, "Could not fetch owners for group {GroupId}", group.Id);
                }

                try
                {
                    // Get members count
                    var members = await _graphClient.Groups[group.Id].Members.GetAsync(config =>
                    {
                        config.QueryParameters.Count = true;
                        config.Headers.Add("ConsistencyLevel", "eventual");
                    });
                    memberCount = members?.Value?.Count ?? 0;
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, "Could not fetch members for group {GroupId}", group.Id);
                }

                publicGroups.Add(new PublicGroupDto(
                    Id: group.Id ?? "",
                    DisplayName: group.DisplayName ?? "Unknown",
                    CreatedDateTime: group.CreatedDateTime?.DateTime,
                    GroupType: groupType,
                    IsTeam: isTeam,
                    OwnerCount: ownerCount,
                    MemberCount: memberCount,
                    Description: group.Description
                ));
            }

            var totalTeams = publicGroups.Count(g => g.IsTeam);
            var totalM365Groups = publicGroups.Count(g => !g.IsTeam);
            var groupsWithNoOwner = publicGroups.Count(g => g.OwnerCount == 0);
            var groupsWithSingleOwner = publicGroups.Count(g => g.OwnerCount == 1);

            _logger.LogInformation(
                "Public groups report complete: {Total} total, {Teams} teams, {M365} M365 groups, {NoOwner} without owners",
                publicGroups.Count, totalTeams, totalM365Groups, groupsWithNoOwner);

            return new PublicGroupsReportDto(
                TotalPublicGroups: publicGroups.Count,
                TotalTeams: totalTeams,
                TotalM365Groups: totalM365Groups,
                GroupsWithNoOwner: groupsWithNoOwner,
                GroupsWithSingleOwner: groupsWithSingleOwner,
                Groups: publicGroups,
                LastUpdated: DateTime.UtcNow
            );
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching public groups");
            throw;
        }
    }

    #endregion

    #region Stale Privileged Accounts

    /// <summary>
    /// Gets privileged users who have not signed in within the specified threshold.
    /// Used for MT.1029 compliance - ensuring stale accounts are not assigned to privileged roles.
    /// </summary>
    public async Task<StalePrivilegedAccountsReportDto> GetStalePrivilegedAccountsAsync(int inactiveDaysThreshold = 30)
    {
        _logger.LogInformation("Fetching stale privileged accounts with {Days} day threshold...", inactiveDaysThreshold);
        
        try
        {
            // Privileged roles to monitor (matching the PowerShell script)
            var privilegedRolesToMonitor = new List<string>
            {
                "Global Administrator",
                "Privileged Role Administrator",
                "Security Administrator",
                "Exchange Administrator",
                "SharePoint Administrator",
                "Intune Administrator",
                "Conditional Access Administrator",
                "User Administrator"
            };

            // Get all directory roles
            var directoryRoles = await _graphClient.DirectoryRoles.GetAsync(config =>
            {
                config.QueryParameters.Select = new[] { "id", "displayName" };
            });

            if (directoryRoles?.Value == null)
            {
                _logger.LogWarning("No directory roles found");
                return new StalePrivilegedAccountsReportDto(
                    TotalPrivilegedUsers: 0,
                    TotalStaleAccounts: 0,
                    AccountsNeverSignedIn: 0,
                    AccountsDisabled: 0,
                    InactiveDaysThreshold: inactiveDaysThreshold,
                    MonitoredRoles: privilegedRolesToMonitor,
                    StaleAccounts: new List<StalePrivilegedAccountDto>(),
                    AllPrivilegedUsers: new List<StalePrivilegedAccountDto>(),
                    LastUpdated: DateTime.UtcNow
                );
            }

            // Filter to privileged roles we care about
            var privilegedRoles = directoryRoles.Value
                .Where(r => privilegedRolesToMonitor.Contains(r.DisplayName ?? "", StringComparer.OrdinalIgnoreCase))
                .ToList();

            _logger.LogInformation("Found {Count} privileged roles to check", privilegedRoles.Count);

            var privilegedUsers = new List<StalePrivilegedAccountDto>();
            var processedUserIds = new HashSet<string>(); // Track to avoid duplicates

            foreach (var role in privilegedRoles)
            {
                if (string.IsNullOrEmpty(role.Id)) continue;

                try
                {
                    // Get members of this role
                    var members = await _graphClient.DirectoryRoles[role.Id].Members.GetAsync();
                    
                    if (members?.Value == null) continue;

                    foreach (var member in members.Value)
                    {
                        // Only process users (not service principals or groups)
                        if (member.OdataType != "#microsoft.graph.user") continue;
                        if (string.IsNullOrEmpty(member.Id)) continue;

                        // Skip if already processed (user may have multiple roles)
                        if (processedUserIds.Contains(member.Id))
                        {
                            // User already exists, just add this role if not already listed
                            var existingUser = privilegedUsers.FirstOrDefault(u => u.UserId == member.Id);
                            if (existingUser != null && !existingUser.Role.Contains(role.DisplayName ?? ""))
                            {
                                // Update with additional role
                                var index = privilegedUsers.IndexOf(existingUser);
                                privilegedUsers[index] = existingUser with 
                                { 
                                    Role = $"{existingUser.Role}, {role.DisplayName}" 
                                };
                            }
                            continue;
                        }

                        try
                        {
                            // Get user details including sign-in activity
                            var user = await _graphClient.Users[member.Id].GetAsync(config =>
                            {
                                config.QueryParameters.Select = new[] 
                                { 
                                    "id", "displayName", "userPrincipalName", 
                                    "signInActivity", "accountEnabled" 
                                };
                            });

                            if (user == null) continue;

                            var lastSignIn = user.SignInActivity?.LastSignInDateTime?.DateTime;
                            var lastNonInteractiveSignIn = user.SignInActivity?.LastNonInteractiveSignInDateTime?.DateTime;
                            var accountEnabled = user.AccountEnabled == true;

                            // Calculate days since last sign-in
                            int daysSinceLastSignIn;
                            if (lastSignIn.HasValue)
                            {
                                daysSinceLastSignIn = (int)(DateTime.UtcNow - lastSignIn.Value).TotalDays;
                            }
                            else
                            {
                                daysSinceLastSignIn = int.MaxValue; // Never signed in
                            }

                            privilegedUsers.Add(new StalePrivilegedAccountDto(
                                UserId: user.Id ?? "",
                                DisplayName: user.DisplayName ?? "Unknown",
                                UserPrincipalName: user.UserPrincipalName ?? "",
                                Role: role.DisplayName ?? "Unknown",
                                LastSignIn: lastSignIn,
                                LastNonInteractiveSignIn: lastNonInteractiveSignIn,
                                AccountStatus: accountEnabled ? "Enabled" : "Disabled",
                                DaysSinceLastSignIn: daysSinceLastSignIn == int.MaxValue ? -1 : daysSinceLastSignIn
                            ));

                            processedUserIds.Add(member.Id);
                        }
                        catch (Exception ex)
                        {
                            _logger.LogWarning(ex, "Could not fetch details for user {UserId}", member.Id);
                        }
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, "Could not fetch members for role {RoleId}", role.Id);
                }
            }

            _logger.LogInformation("Found {Count} privileged users, checking for stale accounts...", privilegedUsers.Count);

            // Filter stale accounts (inactive > threshold OR never signed in)
            var cutoffDate = DateTime.UtcNow.AddDays(-inactiveDaysThreshold);
            var staleAccounts = privilegedUsers
                .Where(u => !u.LastSignIn.HasValue || u.LastSignIn < cutoffDate)
                .OrderBy(u => u.LastSignIn ?? DateTime.MinValue)
                .ToList();

            var accountsNeverSignedIn = staleAccounts.Count(u => !u.LastSignIn.HasValue);
            var accountsDisabled = staleAccounts.Count(u => u.AccountStatus == "Disabled");

            _logger.LogInformation(
                "Stale privileged accounts report complete: {Total} total privileged, {Stale} stale, {Never} never signed in",
                privilegedUsers.Count, staleAccounts.Count, accountsNeverSignedIn);

            return new StalePrivilegedAccountsReportDto(
                TotalPrivilegedUsers: privilegedUsers.Count,
                TotalStaleAccounts: staleAccounts.Count,
                AccountsNeverSignedIn: accountsNeverSignedIn,
                AccountsDisabled: accountsDisabled,
                InactiveDaysThreshold: inactiveDaysThreshold,
                MonitoredRoles: privilegedRolesToMonitor,
                StaleAccounts: staleAccounts,
                AllPrivilegedUsers: privilegedUsers.OrderBy(u => u.DisplayName).ToList(),
                LastUpdated: DateTime.UtcNow
            );
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching stale privileged accounts");
            throw;
        }
    }

    #endregion

    #region Conditional Access Break Glass

    /// <summary>
    /// Resolves a user principal name to get user details.
    /// </summary>
    public async Task<BreakGlassAccountDto?> ResolveUserAsync(string userPrincipalName)
    {
        try
        {
            var user = await _graphClient.Users[userPrincipalName].GetAsync(config =>
            {
                config.QueryParameters.Select = new[] { "id", "displayName", "userPrincipalName" };
            });

            if (user == null) return null;

            return new BreakGlassAccountDto(
                UserPrincipalName: user.UserPrincipalName ?? userPrincipalName,
                DisplayName: user.DisplayName,
                ObjectId: user.Id,
                IsResolved: true
            );
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Could not resolve user {UPN}", userPrincipalName);
            return new BreakGlassAccountDto(
                UserPrincipalName: userPrincipalName,
                DisplayName: null,
                ObjectId: null,
                IsResolved: false
            );
        }
    }

    /// <summary>
    /// Gets Conditional Access policies and checks if break glass accounts are properly excluded.
    /// </summary>
    public async Task<CABreakGlassReportDto> GetCABreakGlassReportAsync(List<string> breakGlassUpns)
    {
        _logger.LogInformation("Fetching Conditional Access policies for break glass report...");
        
        try
        {
            // Get tenant info
            var organization = await _graphClient.Organization.GetAsync();
            var tenantName = organization?.Value?.FirstOrDefault()?.DisplayName ?? "Unknown";
            var tenantId = organization?.Value?.FirstOrDefault()?.Id ?? "Unknown";

            // Resolve break glass accounts to get their object IDs
            var breakGlassAccounts = new List<BreakGlassAccountDto>();
            var breakGlassObjectIds = new Dictionary<string, string>(); // UPN -> ObjectId

            foreach (var upn in breakGlassUpns.Where(u => !string.IsNullOrWhiteSpace(u)))
            {
                var account = await ResolveUserAsync(upn);
                if (account != null)
                {
                    breakGlassAccounts.Add(account);
                    if (account.IsResolved && !string.IsNullOrEmpty(account.ObjectId))
                    {
                        breakGlassObjectIds[upn] = account.ObjectId;
                    }
                }
            }

            // Get all Conditional Access policies
            var policies = await _graphClient.Identity.ConditionalAccess.Policies.GetAsync();
            
            if (policies?.Value == null || policies.Value.Count == 0)
            {
                _logger.LogInformation("No Conditional Access policies found");
                return new CABreakGlassReportDto(
                    TenantName: tenantName,
                    TenantId: tenantId,
                    TotalPolicies: 0,
                    PoliciesWithFullExclusion: 0,
                    PoliciesWithPartialExclusion: 0,
                    PoliciesWithNoExclusion: 0,
                    EnabledPolicies: 0,
                    DisabledPolicies: 0,
                    ReportOnlyPolicies: 0,
                    ConfiguredBreakGlassAccounts: breakGlassAccounts,
                    Policies: new List<ConditionalAccessPolicyDto>(),
                    LastUpdated: DateTime.UtcNow
                );
            }

            _logger.LogInformation("Found {Count} Conditional Access policies", policies.Value.Count);

            var policyDtos = new List<ConditionalAccessPolicyDto>();
            int enabledCount = 0, disabledCount = 0, reportOnlyCount = 0;
            int fullExclusionCount = 0, partialExclusionCount = 0, noExclusionCount = 0;

            foreach (var policy in policies.Value.OrderBy(p => p.DisplayName))
            {
                // Get policy state
                var state = policy.State?.ToString()?.ToLower() ?? "unknown";
                string displayState;
                switch (state)
                {
                    case "enabled":
                        displayState = "Enabled";
                        enabledCount++;
                        break;
                    case "disabled":
                        displayState = "Disabled";
                        disabledCount++;
                        break;
                    case "enabledforreportingbutnotenforced":
                        displayState = "Report-Only";
                        reportOnlyCount++;
                        break;
                    default:
                        displayState = state;
                        break;
                }

                // Check excluded users
                var excludedUserIds = policy.Conditions?.Users?.ExcludeUsers ?? new List<string>();
                
                var excludedBreakGlass = new List<string>();
                var missingBreakGlass = new List<string>();

                foreach (var upn in breakGlassUpns.Where(u => !string.IsNullOrWhiteSpace(u)))
                {
                    if (breakGlassObjectIds.TryGetValue(upn, out var objectId))
                    {
                        if (excludedUserIds.Contains(objectId))
                        {
                            excludedBreakGlass.Add(upn);
                        }
                        else
                        {
                            missingBreakGlass.Add(upn);
                        }
                    }
                    else
                    {
                        // User couldn't be resolved
                        missingBreakGlass.Add($"{upn} (not found)");
                    }
                }

                string exclusionStatus;
                if (breakGlassUpns.Count == 0 || breakGlassAccounts.Count == 0)
                {
                    exclusionStatus = "No break glass accounts configured";
                }
                else if (missingBreakGlass.Count == 0)
                {
                    exclusionStatus = "✅ All excluded";
                    fullExclusionCount++;
                }
                else if (excludedBreakGlass.Count == 0)
                {
                    exclusionStatus = "❌ None excluded";
                    noExclusionCount++;
                }
                else
                {
                    exclusionStatus = "⚠️ Partial";
                    partialExclusionCount++;
                }

                policyDtos.Add(new ConditionalAccessPolicyDto(
                    Id: policy.Id ?? "",
                    DisplayName: policy.DisplayName ?? "Unknown",
                    State: state,
                    DisplayState: displayState,
                    IsBreakGlassExcluded: missingBreakGlass.Count == 0 && breakGlassAccounts.Count > 0,
                    ExcludedBreakGlassAccounts: excludedBreakGlass,
                    MissingBreakGlassAccounts: missingBreakGlass,
                    ExclusionStatus: exclusionStatus
                ));
            }

            _logger.LogInformation(
                "CA Break Glass report complete: {Total} policies, {Full} fully excluded, {Partial} partial, {None} not excluded",
                policyDtos.Count, fullExclusionCount, partialExclusionCount, noExclusionCount);

            return new CABreakGlassReportDto(
                TenantName: tenantName,
                TenantId: tenantId,
                TotalPolicies: policyDtos.Count,
                PoliciesWithFullExclusion: fullExclusionCount,
                PoliciesWithPartialExclusion: partialExclusionCount,
                PoliciesWithNoExclusion: noExclusionCount,
                EnabledPolicies: enabledCount,
                DisabledPolicies: disabledCount,
                ReportOnlyPolicies: reportOnlyCount,
                ConfiguredBreakGlassAccounts: breakGlassAccounts,
                Policies: policyDtos,
                LastUpdated: DateTime.UtcNow
            );
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching CA break glass report");
            throw;
        }
    }

    #endregion
}
