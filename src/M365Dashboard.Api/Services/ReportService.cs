using System.Text;
using System.Text.Json;
using Microsoft.EntityFrameworkCore;
using M365Dashboard.Api.Data;
using M365Dashboard.Api.Models;
using M365Dashboard.Api.Models.Dtos;

namespace M365Dashboard.Api.Services;

public interface IReportService
{
    List<ReportDefinitionDto> GetReportDefinitions();
    Task<ReportResultDto> GenerateReportAsync(GenerateReportRequest request, string? userId = null, bool isScheduled = false, int? scheduledReportId = null);
    Task<string> ExportReportToCsvAsync(GenerateReportRequest request);
    Task<string> ExportReportToHtmlAsync(GenerateReportRequest request);
    Task<List<ScheduledReportDto>> GetScheduledReportsAsync(string userId);
    Task<ScheduledReportDto> CreateScheduledReportAsync(string userId, string? userEmail, CreateScheduledReportRequest request);
    Task<ScheduledReportDto?> UpdateScheduledReportAsync(string userId, string scheduleId, UpdateScheduledReportRequest request);
    Task<bool> DeleteScheduledReportAsync(string userId, string scheduleId);
    Task<List<ReportHistoryDto>> GetReportHistoryAsync(string userId, int take);
    Task<List<ScheduledReport>> GetDueScheduledReportsAsync();
    Task UpdateScheduledReportAfterRunAsync(int scheduleId, bool success, string? error = null);
}

public class ReportService : IReportService
{
    private readonly ApplicationDbContext _dbContext;
    private readonly IGraphService _graphService;
    private readonly ILogger<ReportService> _logger;
    private readonly IWebHostEnvironment _environment;

    private static readonly List<ReportDefinitionDto> _reportDefinitions = new()
    {
        // Usage Reports
        new ReportDefinitionDto(
            "user-activity",
            "User Activity Report",
            "Overview of user activity including sign-ins, license usage, and account status",
            "Usage",
            new List<string> { "json", "csv", "html" },
            true,
            true
        ),
        new ReportDefinitionDto(
            "license-usage",
            "License Usage Report",
            "Detailed breakdown of license assignments and utilization across all SKUs",
            "Usage",
            new List<string> { "json", "csv", "html" },
            false,
            true
        ),
        new ReportDefinitionDto(
            "mailbox-usage",
            "Mailbox Usage Report",
            "Email activity and mailbox statistics for all users",
            "Usage",
            new List<string> { "json", "csv", "html" },
            true,
            true
        ),
        new ReportDefinitionDto(
            "teams-usage",
            "Teams Usage Report",
            "Microsoft Teams activity including messages, calls, and meetings",
            "Usage",
            new List<string> { "json", "csv", "html" },
            true,
            true
        ),
        
        // Security Reports
        new ReportDefinitionDto(
            "mfa-status",
            "MFA Registration Status",
            "Multi-factor authentication registration status for all users",
            "Security",
            new List<string> { "json", "csv", "html" },
            false,
            true
        ),
        new ReportDefinitionDto(
            "risky-users",
            "Risky Users Report",
            "Users flagged by Identity Protection with risk details",
            "Security",
            new List<string> { "json", "csv", "html" },
            false,
            true
        ),
        new ReportDefinitionDto(
            "sign-in-failures",
            "Sign-in Failures Report",
            "Failed sign-in attempts with error details and locations",
            "Security",
            new List<string> { "json", "csv", "html" },
            true,
            true
        ),
        new ReportDefinitionDto(
            "secure-score",
            "Secure Score Report",
            "Microsoft Secure Score breakdown with improvement recommendations",
            "Security",
            new List<string> { "json", "csv", "html" },
            false,
            true
        ),
        
        // Device Reports
        new ReportDefinitionDto(
            "device-compliance",
            "Device Compliance Report",
            "Intune device compliance status and details",
            "Devices",
            new List<string> { "json", "csv", "html" },
            false,
            true
        ),
        new ReportDefinitionDto(
            "device-inventory",
            "Device Inventory Report",
            "Complete inventory of all managed devices with hardware details",
            "Devices",
            new List<string> { "json", "csv", "html" },
            false,
            true
        ),
        
        // Identity Reports
        new ReportDefinitionDto(
            "guest-users",
            "Guest Users Report",
            "External/guest user accounts in the tenant",
            "Identity",
            new List<string> { "json", "csv", "html" },
            false,
            true
        ),
        new ReportDefinitionDto(
            "inactive-users",
            "Inactive Users Report",
            "Users who haven't signed in within a specified period",
            "Identity",
            new List<string> { "json", "csv", "html" },
            true,
            true
        ),
        new ReportDefinitionDto(
            "group-membership",
            "Group Membership Report",
            "Groups with their members and owners",
            "Identity",
            new List<string> { "json", "csv", "html" },
            false,
            true
        ),
        new ReportDefinitionDto(
            "app-credentials",
            "App Registration Credentials Report",
            "Expiring and expired secrets and certificates for all app registrations",
            "Security",
            new List<string> { "json", "csv", "html" },
            false,
            true
        ),
        new ReportDefinitionDto(
            "public-groups",
            "Public Groups Report",
            "All public Microsoft 365 groups and Teams for CIS compliance review (Control 1.2.1)",
            "Security",
            new List<string> { "json", "csv", "html" },
            false,
            true
        ),
        new ReportDefinitionDto(
            "stale-privileged-accounts",
            "Stale Privileged Accounts Report",
            "Privileged users who have not signed in within the specified period (MT.1029 compliance)",
            "Security",
            new List<string> { "json", "csv", "html" },
            true,
            true
        ),
        new ReportDefinitionDto(
            "ca-breakglass",
            "Conditional Access Break Glass Report",
            "Verifies break glass accounts are excluded from all Conditional Access policies",
            "Security",
            new List<string> { "json", "csv", "html" },
            false,
            true
        ),

        // Executive Summary
        new ReportDefinitionDto(
            "executive-summary-pdf",
            "Executive Summary Report",
            "Full executive summary PDF including secure score, devices, users, MFA, domain security and app credentials",
            "Executive",
            new List<string> { "pdf" },
            false,
            true
        )
    };

    private readonly IExecutiveReportService _executiveReportService;

    public ReportService(ApplicationDbContext dbContext, IGraphService graphService, ILogger<ReportService> logger, IWebHostEnvironment environment, IExecutiveReportService executiveReportService)
    {
        _dbContext = dbContext;
        _graphService = graphService;
        _logger = logger;
        _environment = environment;
        _executiveReportService = executiveReportService;
    }

    // ─── Branding helpers ──────────────────────────────────────────────────────

    private ReportSettings LoadReportSettings()
    {
        try
        {
            var filePath = Path.Combine(_environment.ContentRootPath, "App_Data", "report-settings.json");
            if (!File.Exists(filePath)) return new ReportSettings();
            var json = File.ReadAllText(filePath);
            return JsonSerializer.Deserialize<ReportSettings>(json, new JsonSerializerOptions { PropertyNameCaseInsensitive = true })
                   ?? new ReportSettings();
        }
        catch
        {
            return new ReportSettings();
        }
    }

    /// <summary>
    /// Returns an HTML cover-banner block to inject at the top of every report body.
    /// Contains: logo (if set), company name, report title and generation timestamp.
    /// </summary>
    private static string GenerateBrandedHeader(ReportSettings s, string reportTitle, string subtitle)
    {
        var primary = System.Net.WebUtility.HtmlEncode(s.PrimaryColor);
        var accent  = System.Net.WebUtility.HtmlEncode(s.AccentColor);
        var company = System.Net.WebUtility.HtmlEncode(s.CompanyName);
        var title   = System.Net.WebUtility.HtmlEncode(reportTitle);
        var sub     = System.Net.WebUtility.HtmlEncode(subtitle);

        var logoHtml = !string.IsNullOrEmpty(s.LogoBase64)
            ? $"<img src='data:{s.LogoContentType ?? "image/png"};base64,{s.LogoBase64}' alt='{company} logo' style='max-height:48px;max-width:200px;object-fit:contain;' />"
            : $"<span style='font-size:20px;font-weight:700;color:white;letter-spacing:-0.5px;'>{company}</span>";

        return $"""
            <div class='report-header' style='background:linear-gradient(135deg,{primary} 0%,{accent} 100%);padding:28px 36px;margin:-30px -30px 30px -30px;border-radius:8px 8px 0 0;display:flex;align-items:center;justify-content:space-between;gap:20px;'>
                <div style='display:flex;align-items:center;gap:16px;'>
                    {logoHtml}
                    <div style='width:1px;height:40px;background:rgba(255,255,255,0.35);'></div>
                    <div>
                        <div style='color:rgba(255,255,255,0.85);font-size:12px;font-weight:500;text-transform:uppercase;letter-spacing:0.8px;margin-bottom:2px;'>{company}</div>
                        <div style='color:white;font-size:20px;font-weight:700;line-height:1.2;'>{title}</div>
                    </div>
                </div>
                <div style='text-align:right;color:rgba(255,255,255,0.8);font-size:12px;white-space:nowrap;'>
                    <div style='font-weight:600;margin-bottom:2px;'>Generated</div>
                    <div>{sub}</div>
                </div>
            </div>
            """;
    }

    /// <summary>
    /// Returns a branded footer string for injection into report HTML.
    /// </summary>
    private static string GenerateBrandedFooter(ReportSettings s, string? extraText = null)
    {
        var company  = System.Net.WebUtility.HtmlEncode(s.CompanyName);
        var footer   = System.Net.WebUtility.HtmlEncode(
            !string.IsNullOrEmpty(s.FooterText) ? s.FooterText : $"This report was generated by {s.CompanyName}.");
        var extra    = extraText != null ? $"<p>{System.Net.WebUtility.HtmlEncode(extraText)}</p>" : "";
        return $"<div class='footer'><p>{footer}</p>{extra}</div>";
    }

    // ──────────────────────────────────────────────────────────────────────────

    public List<ReportDefinitionDto> GetReportDefinitions() => _reportDefinitions;

    public async Task<ReportResultDto> GenerateReportAsync(GenerateReportRequest request, string? userId = null, bool isScheduled = false, int? scheduledReportId = null)
    {
        _logger.LogInformation("Generating report: {ReportType}", request.ReportType);

        var definition = _reportDefinitions.FirstOrDefault(d => d.ReportType == request.ReportType)
            ?? throw new ArgumentException($"Unknown report type: {request.ReportType}");

        object data;
        ReportSummaryDto? summary = null;
        int recordCount = 0;

        try
        {
            switch (request.ReportType)
            {
                case "user-activity":
                    var users = await _graphService.GetUsersAsync(null, "displayName", true, 999);
                    data = users.Users;
                    recordCount = users.TotalCount;
                    summary = new ReportSummaryDto(
                        users.TotalCount,
                        new Dictionary<string, object>
                        {
                            { "enabledUsers", users.Users.Count(u => u.AccountEnabled) },
                            { "disabledUsers", users.Users.Count(u => !u.AccountEnabled) },
                            { "mfaRegistered", users.Users.Count(u => u.IsMfaRegistered) }
                        }
                    );
                    break;

                case "license-usage":
                    var licenses = await _graphService.GetLicenseUsageAsync();
                    data = licenses.Licenses;
                    recordCount = licenses.Licenses.Count;
                    summary = new ReportSummaryDto(
                        licenses.Licenses.Count,
                        new Dictionary<string, object>
                        {
                            { "totalConsumed", licenses.TotalConsumed },
                            { "totalAvailable", licenses.TotalAvailable },
                            { "overallUtilization", licenses.OverallUtilization }
                        }
                    );
                    break;

                case "mailbox-usage":
                    var mailflow = await _graphService.GetMailflowSummaryAsync(GetDaysFromDateRange(request.DateRange));
                    data = new { mailflow.DailyTraffic, mailflow.TopSenders, mailflow.TopRecipients };
                    recordCount = mailflow.DailyTraffic.Count;
                    summary = new ReportSummaryDto(
                        mailflow.DailyTraffic.Count,
                        new Dictionary<string, object>
                        {
                            { "totalSent", mailflow.TotalMessagesSent },
                            { "totalReceived", mailflow.TotalMessagesReceived },
                            { "avgPerDay", mailflow.AverageMessagesPerDay }
                        }
                    );
                    break;

                case "teams-usage":
                    var teamsActivity = await _graphService.GetTeamsActivityAsync(
                        DateTime.UtcNow.AddDays(-GetDaysFromDateRange(request.DateRange)), 
                        DateTime.UtcNow);
                    data = teamsActivity.Trend;
                    recordCount = teamsActivity.Trend.Count;
                    summary = new ReportSummaryDto(
                        teamsActivity.Trend.Count,
                        new Dictionary<string, object>
                        {
                            { "totalMessages", teamsActivity.TotalMessages },
                            { "totalCalls", teamsActivity.TotalCalls },
                            { "totalMeetings", teamsActivity.TotalMeetings }
                        }
                    );
                    break;

                case "mfa-status":
                    var mfaDetails = await _graphService.GetMfaRegistrationDetailsAsync();
                    data = mfaDetails.Users;
                    recordCount = mfaDetails.TotalCount;
                    summary = new ReportSummaryDto(
                        mfaDetails.TotalCount,
                        new Dictionary<string, object>
                        {
                            { "registered", mfaDetails.MfaRegisteredCount },
                            { "notRegistered", mfaDetails.MfaNotRegisteredCount },
                            { "registrationPercentage", mfaDetails.MfaRegistrationPercentage }
                        }
                    );
                    break;

                case "risky-users":
                    var riskyUsers = await _graphService.GetRiskyUsersAsync();
                    data = riskyUsers;
                    recordCount = riskyUsers.Count;
                    summary = new ReportSummaryDto(
                        riskyUsers.Count,
                        new Dictionary<string, object>
                        {
                            { "high", riskyUsers.Count(u => u.RiskLevel.Equals("High", StringComparison.OrdinalIgnoreCase)) },
                            { "medium", riskyUsers.Count(u => u.RiskLevel.Equals("Medium", StringComparison.OrdinalIgnoreCase)) },
                            { "low", riskyUsers.Count(u => u.RiskLevel.Equals("Low", StringComparison.OrdinalIgnoreCase)) }
                        }
                    );
                    break;

                case "sign-in-failures":
                    var signInData = await _graphService.GetSignInAnalyticsAsync(
                        DateTime.UtcNow.AddDays(-GetDaysFromDateRange(request.DateRange)),
                        DateTime.UtcNow);
                    data = signInData.Trend;
                    recordCount = signInData.TotalSignIns;
                    summary = new ReportSummaryDto(
                        signInData.TotalSignIns,
                        new Dictionary<string, object>
                        {
                            { "successful", signInData.SuccessfulSignIns },
                            { "failed", signInData.FailedSignIns },
                            { "successRate", signInData.SuccessRate }
                        }
                    );
                    break;

                case "secure-score":
                    var secureScore = await _graphService.GetSecureScoreAsync();
                    if (secureScore == null)
                    {
                        data = new { message = "Secure Score not available" };
                        summary = new ReportSummaryDto(0, null);
                    }
                    else
                    {
                        data = secureScore.ControlScores;
                        recordCount = secureScore.ControlScores.Count;
                        summary = new ReportSummaryDto(
                            secureScore.ControlScores.Count,
                            new Dictionary<string, object>
                            {
                                { "currentScore", secureScore.CurrentScore },
                                { "maxScore", secureScore.MaxScore },
                                { "percentage", secureScore.PercentageScore }
                            }
                        );
                    }
                    break;

                case "device-compliance":
                    var compliance = await _graphService.GetDeviceComplianceAsync();
                    data = compliance.ByPlatform;
                    recordCount = compliance.TotalDevices;
                    summary = new ReportSummaryDto(
                        compliance.TotalDevices,
                        new Dictionary<string, object>
                        {
                            { "compliant", compliance.CompliantDevices },
                            { "nonCompliant", compliance.NonCompliantDevices },
                            { "complianceRate", compliance.ComplianceRate }
                        }
                    );
                    break;

                case "device-inventory":
                    var devices = await _graphService.GetDevicesAsync(null, "deviceName", true, 999);
                    data = devices.Devices;
                    recordCount = devices.TotalCount;
                    summary = new ReportSummaryDto(
                        devices.TotalCount,
                        new Dictionary<string, object>
                        {
                            { "windows", devices.Devices.Count(d => d.OperatingSystem?.Contains("Windows", StringComparison.OrdinalIgnoreCase) == true) },
                            { "ios", devices.Devices.Count(d => d.OperatingSystem?.Contains("iOS", StringComparison.OrdinalIgnoreCase) == true) },
                            { "android", devices.Devices.Count(d => d.OperatingSystem?.Contains("Android", StringComparison.OrdinalIgnoreCase) == true) },
                            { "macos", devices.Devices.Count(d => d.OperatingSystem?.Contains("macOS", StringComparison.OrdinalIgnoreCase) == true) }
                        }
                    );
                    break;

                case "guest-users":
                    var allUsers = await _graphService.GetUsersAsync(null, "displayName", true, 999);
                    var guests = allUsers.Users.Where(u => u.UserType == "Guest").ToList();
                    data = guests;
                    recordCount = guests.Count;
                    summary = new ReportSummaryDto(
                        guests.Count,
                        new Dictionary<string, object>
                        {
                            { "enabled", guests.Count(u => u.AccountEnabled) },
                            { "disabled", guests.Count(u => !u.AccountEnabled) }
                        }
                    );
                    break;

                case "inactive-users":
                    var usersForInactive = await _graphService.GetUsersAsync(null, "displayName", true, 999);
                    var cutoffDate = DateTime.UtcNow.AddDays(-GetDaysFromDateRange(request.DateRange));
                    var inactive = usersForInactive.Users
                        .Where(u => u.LastSignInDateTime == null || u.LastSignInDateTime < cutoffDate)
                        .ToList();
                    data = inactive;
                    recordCount = inactive.Count;
                    summary = new ReportSummaryDto(
                        inactive.Count,
                        new Dictionary<string, object>
                        {
                            { "neverSignedIn", inactive.Count(u => u.LastSignInDateTime == null) },
                            { "stale", inactive.Count(u => u.LastSignInDateTime != null) }
                        }
                    );
                    break;

                case "group-membership":
                    var groups = await _graphService.GetGroupsAsync(null, "displayName", true, 100);
                    data = groups.Groups;
                    recordCount = groups.TotalCount;
                    summary = new ReportSummaryDto(
                        groups.TotalCount,
                        new Dictionary<string, object>
                        {
                            { "m365Groups", groups.Groups.Count(g => g.GroupType == "Microsoft 365") },
                            { "securityGroups", groups.Groups.Count(g => g.GroupType == "Security") },
                            { "teamsEnabled", groups.Groups.Count(g => g.IsTeam) }
                        }
                    );
                    break;

                case "app-credentials":
                    var appCredentials = await _graphService.GetAppCredentialStatusAsync();
                    data = appCredentials;
                    recordCount = appCredentials.TotalApps;
                    summary = new ReportSummaryDto(
                        appCredentials.TotalApps,
                        new Dictionary<string, object>
                        {
                            { "totalApps", appCredentials.TotalApps },
                            { "expiringSecrets", appCredentials.ExpiringSecrets.Count },
                            { "expiredSecrets", appCredentials.ExpiredSecrets.Count },
                            { "expiringCertificates", appCredentials.ExpiringCertificates.Count },
                            { "expiredCertificates", appCredentials.ExpiredCertificates.Count },
                            { "appsWithExpiringSecrets", appCredentials.AppsWithExpiringSecrets },
                            { "appsWithExpiredSecrets", appCredentials.AppsWithExpiredSecrets },
                            { "appsWithExpiringCertificates", appCredentials.AppsWithExpiringCertificates },
                            { "appsWithExpiredCertificates", appCredentials.AppsWithExpiredCertificates }
                        }
                    );
                    break;

                case "public-groups":
                    var publicGroups = await _graphService.GetPublicGroupsAsync();
                    data = publicGroups;
                    recordCount = publicGroups.TotalPublicGroups;
                    summary = new ReportSummaryDto(
                        publicGroups.TotalPublicGroups,
                        new Dictionary<string, object>
                        {
                            { "totalPublicGroups", publicGroups.TotalPublicGroups },
                            { "totalTeams", publicGroups.TotalTeams },
                            { "totalM365Groups", publicGroups.TotalM365Groups },
                            { "groupsWithNoOwner", publicGroups.GroupsWithNoOwner },
                            { "groupsWithSingleOwner", publicGroups.GroupsWithSingleOwner }
                        }
                    );
                    break;

                case "stale-privileged-accounts":
                    var inactiveDays = GetDaysFromDateRange(request.DateRange);
                    var stalePrivileged = await _graphService.GetStalePrivilegedAccountsAsync(inactiveDays);
                    data = stalePrivileged;
                    recordCount = stalePrivileged.TotalStaleAccounts;
                    summary = new ReportSummaryDto(
                        stalePrivileged.TotalPrivilegedUsers,
                        new Dictionary<string, object>
                        {
                            { "totalPrivilegedUsers", stalePrivileged.TotalPrivilegedUsers },
                            { "totalStaleAccounts", stalePrivileged.TotalStaleAccounts },
                            { "accountsNeverSignedIn", stalePrivileged.AccountsNeverSignedIn },
                            { "accountsDisabled", stalePrivileged.AccountsDisabled },
                            { "inactiveDaysThreshold", stalePrivileged.InactiveDaysThreshold }
                        }
                    );
                    break;

                case "ca-breakglass":
                    // Get break glass accounts from tenant settings
                    var breakGlassUpns = await GetBreakGlassAccountsAsync();
                    var caBreakGlass = await _graphService.GetCABreakGlassReportAsync(breakGlassUpns);
                    data = caBreakGlass;
                    recordCount = caBreakGlass.TotalPolicies;
                    summary = new ReportSummaryDto(
                        caBreakGlass.TotalPolicies,
                        new Dictionary<string, object>
                        {
                            { "totalPolicies", caBreakGlass.TotalPolicies },
                            { "policiesWithFullExclusion", caBreakGlass.PoliciesWithFullExclusion },
                            { "policiesWithPartialExclusion", caBreakGlass.PoliciesWithPartialExclusion },
                            { "policiesWithNoExclusion", caBreakGlass.PoliciesWithNoExclusion },
                            { "enabledPolicies", caBreakGlass.EnabledPolicies },
                            { "configuredBreakGlassAccounts", caBreakGlass.ConfiguredBreakGlassAccounts.Count }
                        }
                    );
                    break;

                case "executive-summary-pdf":
                    // PDF is generated on demand via /api/reports/download — return a lightweight summary here
                    var execData = await _executiveReportService.GatherDataAsync();
                    data = new { message = "Use the Download PDF button to get the full report.", generatedAt = execData.GeneratedAt, reportMonth = execData.ReportMonth };
                    recordCount = 1;
                    summary = new ReportSummaryDto(1, new Dictionary<string, object>
                    {
                        { "reportMonth", execData.ReportMonth },
                        { "totalUsers", execData.UserStats?.TotalUsers ?? 0 },
                        { "totalDevices", execData.DeviceStats?.TotalDevices ?? 0 },
                        { "secureScore", execData.SecureScore?.PercentageScore ?? 0 }
                    });
                    break;

                default:
                    throw new ArgumentException($"Report type not implemented: {request.ReportType}");
            }

            // Log to history if we have a userId
            if (!string.IsNullOrEmpty(userId))
            {
                var history = new ReportHistory
                {
                    UserId = userId,
                    ReportType = request.ReportType,
                    DisplayName = definition.DisplayName,
                    GeneratedAt = DateTime.UtcNow,
                    Status = "success",
                    RecordCount = recordCount,
                    WasScheduled = isScheduled,
                    ScheduledReportId = scheduledReportId
                };
                _dbContext.ReportHistories.Add(history);
                await _dbContext.SaveChangesAsync();
            }

            return new ReportResultDto(
                request.ReportType,
                definition.DisplayName,
                DateTime.UtcNow,
                request.DateRange ?? "last30days",
                data,
                summary
            );
        }
        catch (Exception ex)
        {
            // Log failure to history
            if (!string.IsNullOrEmpty(userId))
            {
                var history = new ReportHistory
                {
                    UserId = userId,
                    ReportType = request.ReportType,
                    DisplayName = definition.DisplayName,
                    GeneratedAt = DateTime.UtcNow,
                    Status = "failed",
                    ErrorMessage = ex.Message,
                    WasScheduled = isScheduled,
                    ScheduledReportId = scheduledReportId
                };
                _dbContext.ReportHistories.Add(history);
                await _dbContext.SaveChangesAsync();
            }
            throw;
        }
    }

    public async Task<string> ExportReportToCsvAsync(GenerateReportRequest request)
    {
        var report = await GenerateReportAsync(request);
        return ConvertToCsv(report.Data, request.ReportType);
    }

    private static string ConvertToCsv(object data, string reportType)
    {
        var sb = new StringBuilder();

        // Handle special report types with nested data
        switch (reportType)
        {
            case "app-credentials":
                if (data is AppCredentialStatusDto appCreds)
                {
                    return ConvertAppCredentialsToCsv(appCreds);
                }
                break;
            case "public-groups":
                if (data is PublicGroupsReportDto publicGroups)
                {
                    return ConvertPublicGroupsToCsv(publicGroups);
                }
                break;
            case "stale-privileged-accounts":
                if (data is StalePrivilegedAccountsReportDto stalePrivileged)
                {
                    return ConvertStalePrivilegedToCsv(stalePrivileged);
                }
                break;
            case "ca-breakglass":
                if (data is CABreakGlassReportDto caBreakGlass)
                {
                    return ConvertCABreakGlassToCsv(caBreakGlass);
                }
                break;
        }

        // Generic CSV conversion for simple list data
        if (data is System.Collections.IEnumerable enumerable && !(data is string))
        {
            var items = enumerable.Cast<object>().ToList();
            if (items.Count == 0)
            {
                return "No data available";
            }

            var firstItem = items[0];
            var properties = firstItem.GetType().GetProperties()
                .Where(p => p.PropertyType != typeof(byte[]) && 
                           !typeof(System.Collections.IEnumerable).IsAssignableFrom(p.PropertyType) || p.PropertyType == typeof(string))
                .ToList();
            
            // Header
            sb.AppendLine(string.Join(",", properties.Select(p => EscapeCsvField(FormatPropertyName(p.Name)))));
            
            // Data rows
            foreach (var item in items)
            {
                var values = properties.Select(p => 
                {
                    var value = p.GetValue(item);
                    if (value is DateTime dt)
                        return EscapeCsvField(dt == DateTime.MinValue ? "" : dt.ToString("yyyy-MM-dd HH:mm:ss"));
                    if (value is DateTimeOffset dto)
                        return EscapeCsvField(dto.ToString("yyyy-MM-dd HH:mm:ss"));
                    if (value is bool b)
                        return b ? "Yes" : "No";
                    return EscapeCsvField(value?.ToString() ?? "");
                });
                sb.AppendLine(string.Join(",", values));
            }
        }
        else
        {
            sb.AppendLine("Data format not supported for CSV export");
        }

        return sb.ToString();
    }

    private static string GenerateSecureScoreHtmlReport(SecurityScoreDto data, ReportSettings branding)
    {
        var sb = new StringBuilder();
        
        sb.AppendLine("<!DOCTYPE html>");
        sb.AppendLine("<html lang='en'>");
        sb.AppendLine("<head>");
        sb.AppendLine("    <meta charset='UTF-8'>");
        sb.AppendLine("    <meta name='viewport' content='width=device-width, initial-scale=1.0'>");
        sb.AppendLine("    <title>Microsoft Secure Score Report</title>");
        sb.AppendLine("    <style>");
        sb.AppendLine("        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 0; padding: 20px; background: #f5f5f5; color: #333; }");
        sb.AppendLine("        .container { max-width: 1200px; margin: 0 auto; background: white; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); padding: 30px; }");
        sb.AppendLine("        h1 { color: #0078d4; margin-bottom: 10px; font-size: 28px; }");
        sb.AppendLine("        .subtitle { color: #666; margin-bottom: 30px; font-size: 14px; }");
        sb.AppendLine("        .score-overview { display: flex; align-items: center; gap: 40px; margin-bottom: 40px; padding: 30px; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); border-radius: 12px; color: white; }");
        sb.AppendLine("        .score-circle { width: 150px; height: 150px; border-radius: 50%; background: rgba(255,255,255,0.2); display: flex; flex-direction: column; align-items: center; justify-content: center; border: 4px solid white; }");
        sb.AppendLine("        .score-value { font-size: 42px; font-weight: bold; }");
        sb.AppendLine("        .score-label { font-size: 14px; opacity: 0.9; }");
        sb.AppendLine("        .score-details { flex: 1; }");
        sb.AppendLine("        .score-details h2 { margin: 0 0 15px 0; font-size: 24px; }");
        sb.AppendLine("        .score-breakdown { display: flex; gap: 30px; }");
        sb.AppendLine("        .score-item { text-align: center; }");
        sb.AppendLine("        .score-item-value { font-size: 28px; font-weight: bold; }");
        sb.AppendLine("        .score-item-label { font-size: 12px; opacity: 0.9; }");
        sb.AppendLine("        .progress-bar { height: 8px; background: rgba(255,255,255,0.3); border-radius: 4px; margin-top: 15px; overflow: hidden; }");
        sb.AppendLine("        .progress-fill { height: 100%; background: white; border-radius: 4px; transition: width 0.3s; }");
        sb.AppendLine("        .category-filter { margin-bottom: 10px; display: flex; gap: 10px; flex-wrap: wrap; }");
        sb.AppendLine("        .category-btn { padding: 8px 16px; border: 1px solid #ddd; border-radius: 20px; background: white; cursor: pointer; font-size: 13px; transition: all 0.2s; }");
        sb.AppendLine("        .category-btn:hover, .category-btn.active { background: #0078d4; color: white; border-color: #0078d4; }");
        sb.AppendLine("        .status-filter { margin-bottom: 20px; display: flex; gap: 10px; flex-wrap: wrap; align-items: center; }");
        sb.AppendLine("        .filter-label { font-weight: 600; color: #666; margin-right: 5px; }");
        sb.AppendLine("        .status-btn { padding: 6px 14px; border: 1px solid #ddd; border-radius: 20px; background: white; cursor: pointer; font-size: 12px; transition: all 0.2s; }");
        sb.AppendLine("        .status-btn:hover { background: #f0f0f0; }");
        sb.AppendLine("        .status-btn.active { background: #6c757d; color: white; border-color: #6c757d; }");
        sb.AppendLine("        h2 { color: #333; font-size: 20px; margin-top: 40px; margin-bottom: 15px; padding-bottom: 10px; border-bottom: 2px solid #eee; }");
        sb.AppendLine("        .controls-grid { display: grid; gap: 15px; }");
        sb.AppendLine("        .control-card { background: #f8f9fa; border-radius: 8px; padding: 20px; border-left: 4px solid #0078d4; }");
        sb.AppendLine("        .control-card.identity { border-left-color: #6264a7; }");
        sb.AppendLine("        .control-card.data { border-left-color: #28a745; }");
        sb.AppendLine("        .control-card.device { border-left-color: #fd7e14; }");
        sb.AppendLine("        .control-card.apps { border-left-color: #dc3545; }");
        sb.AppendLine("        .control-card.infrastructure { border-left-color: #17a2b8; }");
        sb.AppendLine("        .control-card.completed { opacity: 0.7; }");
        sb.AppendLine("        .control-card.completed .control-name::after { content: ' ✓'; color: #28a745; }");
        sb.AppendLine("        .control-header { display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 10px; }");
        sb.AppendLine("        .control-name { font-weight: 600; font-size: 15px; color: #333; margin: 0; }");
        sb.AppendLine("        .control-score { display: flex; align-items: center; gap: 8px; }");
        sb.AppendLine("        .control-score-value { font-weight: bold; font-size: 14px; }");
        sb.AppendLine("        .control-score-max { color: #666; font-size: 13px; }");
        sb.AppendLine("        .control-category { display: inline-block; padding: 3px 10px; border-radius: 12px; font-size: 11px; font-weight: 600; margin-bottom: 10px; }");
        sb.AppendLine("        .category-identity { background: #e8e8fc; color: #6264a7; }");
        sb.AppendLine("        .category-data { background: #d4edda; color: #155724; }");
        sb.AppendLine("        .category-device { background: #fff3cd; color: #856404; }");
        sb.AppendLine("        .category-apps { background: #f8d7da; color: #721c24; }");
        sb.AppendLine("        .category-infrastructure { background: #d1ecf1; color: #0c5460; }");
        sb.AppendLine("        .control-description { color: #555; font-size: 13px; line-height: 1.5; margin: 0; }");
        sb.AppendLine("        .control-meta { margin-top: 12px; display: flex; gap: 20px; font-size: 12px; color: #666; }");
        sb.AppendLine("        .control-mini-bar { flex: 1; max-width: 100px; height: 6px; background: #e0e0e0; border-radius: 3px; overflow: hidden; }");
        sb.AppendLine("        .control-mini-fill { height: 100%; background: #28a745; border-radius: 3px; }");
        sb.AppendLine("        .control-mini-fill.partial { background: #ffc107; }");
        sb.AppendLine("        .control-mini-fill.none { background: #dc3545; width: 0 !important; }");
        sb.AppendLine("        .empty-message { color: #666; padding: 40px; text-align: center; background: #f8f9fa; border-radius: 8px; }");
        sb.AppendLine("        .footer { margin-top: 40px; padding-top: 20px; border-top: 1px solid #eee; color: #666; font-size: 12px; text-align: center; }");
        sb.AppendLine("        @media print { body { background: white; } .container { box-shadow: none; } .score-overview { -webkit-print-color-adjust: exact; print-color-adjust: exact; } }");
        sb.AppendLine("        @media (max-width: 768px) { .score-overview { flex-direction: column; text-align: center; } .score-breakdown { justify-content: center; } }");
        sb.AppendLine("    </style>");
        sb.AppendLine("    <script>");
        sb.AppendLine("        var currentCategory = 'all';");
        sb.AppendLine("        var currentStatus = 'all';");
        sb.AppendLine("        ");
        sb.AppendLine("        function applyFilters() {");
        sb.AppendLine("            document.querySelectorAll('.control-card').forEach(card => {");
        sb.AppendLine("                var categoryMatch = (currentCategory === 'all' || card.dataset.category === currentCategory);");
        sb.AppendLine("                var statusMatch = (currentStatus === 'all' || card.dataset.status === currentStatus);");
        sb.AppendLine("                card.style.display = (categoryMatch && statusMatch) ? 'block' : 'none';");
        sb.AppendLine("            });");
        sb.AppendLine("        }");
        sb.AppendLine("        ");
        sb.AppendLine("        function filterByCategory(el, category) {");
        sb.AppendLine("            currentCategory = category;");
        sb.AppendLine("            document.querySelectorAll('.category-btn').forEach(btn => btn.classList.remove('active'));");
        sb.AppendLine("            el.classList.add('active');");
        sb.AppendLine("            applyFilters();");
        sb.AppendLine("        }");
        sb.AppendLine("        ");
        sb.AppendLine("        function filterByStatus(el, status) {");
        sb.AppendLine("            currentStatus = status;");
        sb.AppendLine("            document.querySelectorAll('.status-btn').forEach(btn => btn.classList.remove('active'));");
        sb.AppendLine("            el.classList.add('active');");
        sb.AppendLine("            applyFilters();");
        sb.AppendLine("        }");
        sb.AppendLine("    </script>");
        sb.AppendLine("</head>");
        sb.AppendLine("<body>");
        sb.AppendLine("<div class='container'>");

        sb.AppendLine(GenerateBrandedHeader(branding, "Microsoft Secure Score Report",
            data.LastUpdated.ToString("dddd, MMMM d, yyyy 'at' h:mm tt") + " UTC"));
        
        // Score Overview
        sb.AppendLine("    <div class='score-overview'>");
        sb.AppendLine("        <div class='score-circle'>");
        sb.AppendLine($"            <div class='score-value'>{data.PercentageScore:N0}%</div>");
        sb.AppendLine("            <div class='score-label'>Secure Score</div>");
        sb.AppendLine("        </div>");
        sb.AppendLine("        <div class='score-details'>");
        sb.AppendLine("            <h2>Your Security Posture</h2>");
        sb.AppendLine("            <div class='score-breakdown'>");
        sb.AppendLine($"                <div class='score-item'><div class='score-item-value'>{data.CurrentScore:N2}</div><div class='score-item-label'>Current Score</div></div>");
        sb.AppendLine($"                <div class='score-item'><div class='score-item-value'>{data.MaxScore:N2}</div><div class='score-item-label'>Max Score</div></div>");
        sb.AppendLine($"                <div class='score-item'><div class='score-item-value'>{data.ControlScores.Count}</div><div class='score-item-label'>Controls</div></div>");
        sb.AppendLine("            </div>");
        sb.AppendLine("            <div class='progress-bar'>");
        sb.AppendLine($"                <div class='progress-fill' style='width: {data.PercentageScore}%'></div>");
        sb.AppendLine("            </div>");
        sb.AppendLine("        </div>");
        sb.AppendLine("    </div>");
        
        // Category Filter
        var categories = data.ControlScores
            .Select(c => c.ControlCategory)
            .Where(c => !string.IsNullOrEmpty(c))
            .Distinct()
            .OrderBy(c => c)
            .ToList();
        
        sb.AppendLine("    <div class='category-filter'>");
        sb.AppendLine("        <button class='category-btn active' onclick=\"filterByCategory(this,'all')\">All Controls</button>");
        foreach (var category in categories)
        {
            sb.AppendLine($"        <button class='category-btn' onclick=\"filterByCategory(this,'{System.Net.WebUtility.HtmlEncode(category.ToLower())}')\">{ System.Net.WebUtility.HtmlEncode(category)}</button>");
        }
        sb.AppendLine("    </div>");
        
        // Status filter (completed/incomplete)
        var completedCount = data.ControlScores.Count(c => c.MaxScore > 0 && c.Score >= c.MaxScore);
        var incompleteCount = data.ControlScores.Count(c => c.MaxScore > 0 && c.Score < c.MaxScore);
        var notApplicableCount = data.ControlScores.Count(c => c.MaxScore == 0);
        
        sb.AppendLine("    <div class='status-filter'>");
        sb.AppendLine("        <label class='filter-label'>Status:</label>");
        sb.AppendLine($"        <button class='status-btn active' onclick=\"filterByStatus(this,'all')\">All ({data.ControlScores.Count})</button>");
        sb.AppendLine($"        <button class='status-btn' onclick=\"filterByStatus(this,'incomplete')\">Incomplete ({incompleteCount})</button>");
        sb.AppendLine($"        <button class='status-btn' onclick=\"filterByStatus(this,'completed')\">Completed ({completedCount})</button>");
        if (notApplicableCount > 0)
        {
            sb.AppendLine($"        <button class='status-btn' onclick=\"filterByStatus(this,'na')\">Not Scored ({notApplicableCount})</button>");
        }
        sb.AppendLine("    </div>");
        
        // Controls
        sb.AppendLine("    <h2>\ud83d\udcca Security Controls</h2>");
        
        if (data.ControlScores.Count == 0)
        {
            sb.AppendLine("    <div class='empty-message'>No security controls found.</div>");
        }
        else
        {
            sb.AppendLine("    <div class='controls-grid'>");
            
            foreach (var control in data.ControlScores.OrderByDescending(c => c.MaxScore - c.Score))
            {
                var categoryLower = (control.ControlCategory ?? "other").ToLower();
                var categoryClass = categoryLower switch
                {
                    "identity" => "identity",
                    "data" => "data",
                    "device" => "device",
                    "apps" => "apps",
                    "infrastructure" => "infrastructure",
                    _ => ""
                };
                
                var badgeClass = categoryLower switch
                {
                    "identity" => "category-identity",
                    "data" => "category-data",
                    "device" => "category-device",
                    "apps" => "category-apps",
                    "infrastructure" => "category-infrastructure",
                    _ => ""
                };
                
                var scorePercent = control.MaxScore > 0 ? (control.Score / control.MaxScore * 100) : 0;
                var barClass = scorePercent >= 100 ? "" : (scorePercent > 0 ? "partial" : "none");
                
                // Determine status for filtering
                var status = control.MaxScore == 0 ? "na" : (scorePercent >= 100 ? "completed" : "incomplete");
                
                // Clean the description - remove HTML tags
                var description = control.Description ?? "";
                description = System.Text.RegularExpressions.Regex.Replace(description, @"<[^>]+>", " ");
                description = System.Text.RegularExpressions.Regex.Replace(description, @"&nbsp;", " ");
                description = System.Text.RegularExpressions.Regex.Replace(description, @"&[a-zA-Z]+;", "");
                description = System.Text.RegularExpressions.Regex.Replace(description, @"\s+", " ").Trim();
                // Truncate long descriptions
                if (description.Length > 300)
                {
                    description = description.Substring(0, 297) + "...";
                }
                
                // Add completed class for visual styling
                var completedClass = status == "completed" ? " completed" : "";
                
                sb.AppendLine($"        <div class='control-card {categoryClass}{completedClass}' data-category='{categoryLower}' data-status='{status}'>");
                sb.AppendLine("            <div class='control-header'>");
                sb.AppendLine($"                <div>");
                sb.AppendLine($"                    <span class='control-category {badgeClass}'>{System.Net.WebUtility.HtmlEncode(control.ControlCategory ?? "Other")}</span>");
                sb.AppendLine($"                    <p class='control-name'>{System.Net.WebUtility.HtmlEncode(FormatSecureScoreControlName(control.ControlName))}</p>");
                sb.AppendLine($"                </div>");
                sb.AppendLine("                <div class='control-score'>");
                sb.AppendLine($"                    <div class='control-mini-bar'><div class='control-mini-fill {barClass}' style='width: {scorePercent:N0}%'></div></div>");
                sb.AppendLine($"                    <span class='control-score-value'>{control.Score:N2}</span>");
                sb.AppendLine($"                    <span class='control-score-max'>/ {control.MaxScore:N2}</span>");
                sb.AppendLine("                </div>");
                sb.AppendLine("            </div>");
                if (!string.IsNullOrEmpty(description))
                {
                    sb.AppendLine($"            <p class='control-description'>{System.Net.WebUtility.HtmlEncode(description)}</p>");
                }
                sb.AppendLine("        </div>");
            }
            
            sb.AppendLine("    </div>");
        }
        
        sb.AppendLine(GenerateBrandedFooter(branding,
            "Microsoft Secure Score helps you understand your organisation\u2019s security posture and provides recommendations for improvement."));

        sb.AppendLine("</div>");
        sb.AppendLine("</body>");
        sb.AppendLine("</html>");
        
        return sb.ToString();
    }

    private static string ConvertAppCredentialsToCsv(AppCredentialStatusDto data)
    {
        var sb = new StringBuilder();
        sb.AppendLine("Application Name,App ID,Credential Type,Credential Name,Start Date,Expiry Date,Days Until Expiry,Status");
        
        foreach (var cred in data.ExpiredSecrets)
        {
            sb.AppendLine($"{EscapeCsvField(cred.AppDisplayName)},{EscapeCsvField(cred.AppId)},Secret,{EscapeCsvField(cred.DisplayName ?? "(unnamed)")},{cred.StartDateTime:yyyy-MM-dd},{cred.EndDateTime:yyyy-MM-dd},{cred.DaysUntilExpiry},Expired");
        }
        foreach (var cred in data.ExpiringSecrets)
        {
            sb.AppendLine($"{EscapeCsvField(cred.AppDisplayName)},{EscapeCsvField(cred.AppId)},Secret,{EscapeCsvField(cred.DisplayName ?? "(unnamed)")},{cred.StartDateTime:yyyy-MM-dd},{cred.EndDateTime:yyyy-MM-dd},{cred.DaysUntilExpiry},{cred.Status}");
        }
        foreach (var cred in data.ExpiredCertificates)
        {
            sb.AppendLine($"{EscapeCsvField(cred.AppDisplayName)},{EscapeCsvField(cred.AppId)},Certificate,{EscapeCsvField(cred.DisplayName ?? "(unnamed)")},{cred.StartDateTime:yyyy-MM-dd},{cred.EndDateTime:yyyy-MM-dd},{cred.DaysUntilExpiry},Expired");
        }
        foreach (var cred in data.ExpiringCertificates)
        {
            sb.AppendLine($"{EscapeCsvField(cred.AppDisplayName)},{EscapeCsvField(cred.AppId)},Certificate,{EscapeCsvField(cred.DisplayName ?? "(unnamed)")},{cred.StartDateTime:yyyy-MM-dd},{cred.EndDateTime:yyyy-MM-dd},{cred.DaysUntilExpiry},{cred.Status}");
        }
        
        return sb.ToString();
    }

    private static string ConvertPublicGroupsToCsv(PublicGroupsReportDto data)
    {
        var sb = new StringBuilder();
        sb.AppendLine("Group Name,Group ID,Description,Group Type,Is Team,Owner Count,Member Count,Created Date");
        
        foreach (var group in data.Groups)
        {
            sb.AppendLine($"{EscapeCsvField(group.DisplayName)},{EscapeCsvField(group.Id)},{EscapeCsvField(group.Description ?? "")},{EscapeCsvField(group.GroupType)},{(group.IsTeam ? "Yes" : "No")},{group.OwnerCount},{group.MemberCount},{group.CreatedDateTime?.ToString("yyyy-MM-dd") ?? ""}");
        }
        
        return sb.ToString();
    }

    private static string ConvertStalePrivilegedToCsv(StalePrivilegedAccountsReportDto data)
    {
        var sb = new StringBuilder();
        sb.AppendLine("Display Name,User Principal Name,User ID,Role,Last Sign-In,Days Since Last Sign-In,Account Status");
        
        foreach (var account in data.StaleAccounts)
        {
            var lastSignIn = account.LastSignIn?.ToString("yyyy-MM-dd HH:mm:ss") ?? "Never";
            var daysSince = account.DaysSinceLastSignIn == -1 ? "Never" : account.DaysSinceLastSignIn.ToString();
            sb.AppendLine($"{EscapeCsvField(account.DisplayName)},{EscapeCsvField(account.UserPrincipalName)},{EscapeCsvField(account.UserId)},{EscapeCsvField(account.Role)},{lastSignIn},{daysSince},{account.AccountStatus}");
        }
        
        return sb.ToString();
    }

    private static string ConvertCABreakGlassToCsv(CABreakGlassReportDto data)
    {
        var sb = new StringBuilder();
        sb.AppendLine("Policy Name,Policy ID,State,Excluded Break Glass Accounts,Missing Break Glass Accounts,Has Full Exclusion");
        
        foreach (var policy in data.Policies)
        {
            var excluded = string.Join("; ", policy.ExcludedBreakGlassAccounts);
            var missing = string.Join("; ", policy.MissingBreakGlassAccounts);
            var hasFullExclusion = policy.MissingBreakGlassAccounts.Count == 0 && data.ConfiguredBreakGlassAccounts.Count > 0 ? "Yes" : "No";
            sb.AppendLine($"{EscapeCsvField(policy.DisplayName)},{EscapeCsvField(policy.Id)},{policy.DisplayState},{EscapeCsvField(excluded)},{EscapeCsvField(missing)},{hasFullExclusion}");
        }
        
        return sb.ToString();
    }

    public async Task<string> ExportReportToHtmlAsync(GenerateReportRequest request)
    {
        var branding = LoadReportSettings();

        // Special handling for mailbox-usage report - need to fetch data directly
        if (request.ReportType == "mailbox-usage")
        {
            var mailflow = await _graphService.GetMailflowSummaryAsync(GetDaysFromDateRange(request.DateRange));
            return GenerateMailboxUsageHtmlReport(mailflow, request.DateRange ?? "last30days", branding);
        }
        
        // Special handling for teams-usage report - need to fetch data directly
        if (request.ReportType == "teams-usage")
        {
            var days = GetDaysFromDateRange(request.DateRange);
            var teamsActivity = await _graphService.GetTeamsActivityAsync(
                DateTime.UtcNow.AddDays(-days), 
                DateTime.UtcNow);
            return GenerateTeamsUsageHtmlReport(teamsActivity, request.DateRange ?? "last30days", branding);
        }
        
        var report = await GenerateReportAsync(request);
        
        // Special handling for app-credentials report
        if (request.ReportType == "app-credentials" && report.Data is AppCredentialStatusDto appCreds)
        {
            return GenerateAppCredentialsHtmlReport(appCreds, branding);
        }
        
        // Special handling for public-groups report
        if (request.ReportType == "public-groups" && report.Data is PublicGroupsReportDto publicGroups)
        {
            return GeneratePublicGroupsHtmlReport(publicGroups, branding);
        }
        
        // Special handling for stale-privileged-accounts report
        if (request.ReportType == "stale-privileged-accounts" && report.Data is StalePrivilegedAccountsReportDto stalePrivileged)
        {
            return GenerateStalePrivilegedAccountsHtmlReport(stalePrivileged, branding);
        }
        
        // Special handling for inactive-users report
        if (request.ReportType == "inactive-users" && report.Data is List<TenantUserDto> inactiveUsers)
        {
            var days = GetDaysFromDateRange(request.DateRange);
            return GenerateInactiveUsersHtmlReport(inactiveUsers, days, request.DateRange ?? "last30days", branding);
        }
        
        // Special handling for group-membership report
        if (request.ReportType == "group-membership" && report.Data is List<TenantGroupDto> groups)
        {
            return GenerateGroupMembershipHtmlReport(groups, branding);
        }
        
        // Special handling for ca-breakglass report
        if (request.ReportType == "ca-breakglass" && report.Data is CABreakGlassReportDto caBreakGlass)
        {
            return GenerateCABreakGlassHtmlReport(caBreakGlass, branding);
        }
        
        // Special handling for secure-score report
        if (request.ReportType == "secure-score")
        {
            var secureScore = await _graphService.GetSecureScoreAsync();
            if (secureScore != null)
            {
                return GenerateSecureScoreHtmlReport(secureScore, branding);
            }
        }
        
        // Generic HTML conversion for other reports
        return ConvertToGenericHtml(report, branding);
    }

    private async Task<List<string>> GetBreakGlassAccountsAsync()
    {
        // Get from TenantSettings - use "default" tenant for now
        // In a full implementation, you'd get the tenant ID from the request context
        var setting = await _dbContext.TenantSettings
            .FirstOrDefaultAsync(s => s.SettingKey == "BreakGlassAccounts");

        if (setting == null)
        {
            return new List<string>();
        }

        try
        {
            return System.Text.Json.JsonSerializer.Deserialize<List<string>>(setting.SettingValue) ?? new List<string>();
        }
        catch
        {
            return new List<string>();
        }
    }

    private static string GenerateAppCredentialsHtmlReport(AppCredentialStatusDto data, ReportSettings branding)
    {
        var sb = new StringBuilder();
        
        sb.AppendLine("<!DOCTYPE html>");
        sb.AppendLine("<html lang='en'>");
        sb.AppendLine("<head>");
        sb.AppendLine("    <meta charset='UTF-8'>");
        sb.AppendLine("    <meta name='viewport' content='width=device-width, initial-scale=1.0'>");
        sb.AppendLine("    <title>App Registration Credentials Report</title>");
        sb.AppendLine("    <style>");
        sb.AppendLine("        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 0; padding: 20px; background: #f5f5f5; color: #333; }");
        sb.AppendLine("        .container { max-width: 1200px; margin: 0 auto; background: white; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); padding: 30px; }");
        sb.AppendLine("        h1 { color: #0078d4; margin-bottom: 10px; font-size: 28px; }");
        sb.AppendLine("        .subtitle { color: #666; margin-bottom: 30px; font-size: 14px; }");
        sb.AppendLine("        .summary { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 20px; margin-bottom: 40px; }");
        sb.AppendLine("        .summary-card { background: #f8f9fa; border-radius: 8px; padding: 20px; text-align: center; border-left: 4px solid #0078d4; cursor: pointer; transition: transform 0.2s, box-shadow 0.2s; }");
        sb.AppendLine("        .summary-card:hover { transform: translateY(-2px); box-shadow: 0 4px 8px rgba(0,0,0,0.1); }");
        sb.AppendLine("        .summary-card.active { box-shadow: 0 0 0 3px #0078d4; }");
        sb.AppendLine("        .summary-card.warning { border-left-color: #ffc107; }");
        sb.AppendLine("        .summary-card.danger { border-left-color: #dc3545; }");
        sb.AppendLine("        .summary-card.success { border-left-color: #28a745; }");
        sb.AppendLine("        .summary-value { font-size: 36px; font-weight: bold; color: #0078d4; }");
        sb.AppendLine("        .summary-card.warning .summary-value { color: #ffc107; }");
        sb.AppendLine("        .summary-card.danger .summary-value { color: #dc3545; }");
        sb.AppendLine("        .summary-card.success .summary-value { color: #28a745; }");
        sb.AppendLine("        .summary-label { color: #666; font-size: 14px; margin-top: 5px; }");
        sb.AppendLine("        h2 { color: #333; font-size: 20px; margin-top: 40px; margin-bottom: 15px; padding-bottom: 10px; border-bottom: 2px solid #eee; }");
        sb.AppendLine("        h2.warning { color: #856404; }");
        sb.AppendLine("        h2.danger { color: #721c24; }");
        sb.AppendLine("        table { width: 100%; border-collapse: collapse; margin-bottom: 30px; font-size: 14px; }");
        sb.AppendLine("        th { background: #0078d4; color: white; padding: 12px 15px; text-align: left; font-weight: 600; }");
        sb.AppendLine("        td { padding: 12px 15px; border-bottom: 1px solid #eee; }");
        sb.AppendLine("        tr:hover { background: #f8f9fa; }");
        sb.AppendLine("        .status-badge { padding: 4px 12px; border-radius: 20px; font-size: 12px; font-weight: 600; display: inline-block; }");
        sb.AppendLine("        .status-critical { background: #f8d7da; color: #721c24; }");
        sb.AppendLine("        .status-warning { background: #fff3cd; color: #856404; }");
        sb.AppendLine("        .status-expiring { background: #cce5ff; color: #004085; }");
        sb.AppendLine("        .status-expired { background: #f8d7da; color: #721c24; }");
        sb.AppendLine("        .empty-message { color: #28a745; padding: 20px; background: #d4edda; border-radius: 8px; text-align: center; margin-bottom: 30px; }");
        sb.AppendLine("        .footer { margin-top: 40px; padding-top: 20px; border-top: 1px solid #eee; color: #666; font-size: 12px; text-align: center; }");
        sb.AppendLine("        @media print { body { background: white; } .container { box-shadow: none; } }");
        sb.AppendLine("    </style>");
        sb.AppendLine("    <script>");
        sb.AppendLine("        function filterByType(el, type) {");
        sb.AppendLine("            document.querySelectorAll('.summary-card').forEach(card => card.classList.remove('active'));");
        sb.AppendLine("            if (type !== 'all') {");
        sb.AppendLine("                el.closest('.summary-card').classList.add('active');");
        sb.AppendLine("            }");
        sb.AppendLine("            document.querySelectorAll('tbody tr').forEach(row => {");
        sb.AppendLine("                if (type === 'all') {");
        sb.AppendLine("                    row.style.display = '';");
        sb.AppendLine("                } else if (type === 'never') {");
        sb.AppendLine("                    row.style.display = row.classList.contains('never-signed-in') ? '' : 'none';");
        sb.AppendLine("                } else if (type === 'stale') {");
        sb.AppendLine("                    row.style.display = (row.classList.contains('stale') || row.classList.contains('never-signed-in')) ? '' : 'none';");
        sb.AppendLine("                } else if (type === 'disabled') {");
        sb.AppendLine("                    row.style.display = row.dataset.status === 'Disabled' ? '' : 'none';");
        sb.AppendLine("                }");
        sb.AppendLine("            });");
        sb.AppendLine("            updateVisibleCount();");
        sb.AppendLine("        }");
        sb.AppendLine("        function updateVisibleCount() {");
        sb.AppendLine("            var visible = document.querySelectorAll('tbody tr:not([style*=\"display: none\"])').length;");
        sb.AppendLine("            var counter = document.getElementById('visible-count');");
        sb.AppendLine("            if (counter) counter.textContent = 'Showing ' + visible + ' account(s)';");
        sb.AppendLine("        }");
        sb.AppendLine("    </script>");
        sb.AppendLine("</head>");
        sb.AppendLine("<body>");
        sb.AppendLine("<div class='container'>");

        sb.AppendLine(GenerateBrandedHeader(branding, "App Registration Credentials Report",
            data.LastUpdated.ToString("dddd, MMMM d, yyyy 'at' h:mm tt") + $" UTC \u2022 Threshold: {data.ThresholdDays} days"));
        
        // Summary Cards
        sb.AppendLine("    <div class='summary'>");
        sb.AppendLine($"        <div class='summary-card'><div class='summary-value'>{data.TotalApps}</div><div class='summary-label'>Total App Registrations</div></div>");
        
        var expiredClass = (data.AppsWithExpiredSecrets + data.AppsWithExpiredCertificates) > 0 ? "danger" : "success";
        sb.AppendLine($"        <div class='summary-card {expiredClass}'><div class='summary-value'>{data.AppsWithExpiredSecrets + data.AppsWithExpiredCertificates}</div><div class='summary-label'>Apps with Expired Credentials</div></div>");
        
        var expiringClass = (data.AppsWithExpiringSecrets + data.AppsWithExpiringCertificates) > 0 ? "warning" : "success";
        sb.AppendLine($"        <div class='summary-card {expiringClass}'><div class='summary-value'>{data.AppsWithExpiringSecrets + data.AppsWithExpiringCertificates}</div><div class='summary-label'>Apps with Expiring Credentials</div></div>");
        
        sb.AppendLine($"        <div class='summary-card'><div class='summary-value'>{data.ExpiredSecrets.Count + data.ExpiredCertificates.Count}</div><div class='summary-label'>Total Expired Credentials</div></div>");
        sb.AppendLine($"        <div class='summary-card'><div class='summary-value'>{data.ExpiringSecrets.Count + data.ExpiringCertificates.Count}</div><div class='summary-label'>Total Expiring Credentials</div></div>");
        sb.AppendLine("    </div>");
        
        // Expired Secrets
        sb.AppendLine("    <h2 class='danger'>❌ Expired Secrets</h2>");
        if (data.ExpiredSecrets.Count == 0)
        {
            sb.AppendLine("    <div class='empty-message'>✓ No expired secrets found</div>");
        }
        else
        {
            AppendCredentialsTable(sb, data.ExpiredSecrets);
        }
        
        // Expiring Secrets
        sb.AppendLine("    <h2 class='warning'>⚠️ Expiring Secrets</h2>");
        if (data.ExpiringSecrets.Count == 0)
        {
            sb.AppendLine("    <div class='empty-message'>✓ No secrets expiring within the threshold period</div>");
        }
        else
        {
            AppendCredentialsTable(sb, data.ExpiringSecrets);
        }
        
        // Expired Certificates
        sb.AppendLine("    <h2 class='danger'>❌ Expired Certificates</h2>");
        if (data.ExpiredCertificates.Count == 0)
        {
            sb.AppendLine("    <div class='empty-message'>✓ No expired certificates found</div>");
        }
        else
        {
            AppendCredentialsTable(sb, data.ExpiredCertificates);
        }
        
        // Expiring Certificates
        sb.AppendLine("    <h2 class='warning'>⚠️ Expiring Certificates</h2>");
        if (data.ExpiringCertificates.Count == 0)
        {
            sb.AppendLine("    <div class='empty-message'>✓ No certificates expiring within the threshold period</div>");
        }
        else
        {
            AppendCredentialsTable(sb, data.ExpiringCertificates);
        }
        
        sb.AppendLine(GenerateBrandedFooter(branding, "For questions, contact your IT administrator."));

        sb.AppendLine("</div>");
        sb.AppendLine("</body>");
        sb.AppendLine("</html>");
        
        return sb.ToString();
    }

    private static void AppendCredentialsTable(StringBuilder sb, List<AppCredentialDetailDto> credentials)
    {
        sb.AppendLine("    <table>");
        sb.AppendLine("        <thead>");
        sb.AppendLine("            <tr>");
        sb.AppendLine("                <th>Application Name</th>");
        sb.AppendLine("                <th>App ID</th>");
        sb.AppendLine("                <th>Credential Name</th>");
        sb.AppendLine("                <th>Expiry Date</th>");
        sb.AppendLine("                <th>Days</th>");
        sb.AppendLine("                <th>Status</th>");
        sb.AppendLine("            </tr>");
        sb.AppendLine("        </thead>");
        sb.AppendLine("        <tbody>");
        
        foreach (var cred in credentials)
        {
            var statusClass = cred.Status switch
            {
                "Critical" => "status-critical",
                "Warning" => "status-warning",
                "Expiring" => "status-expiring",
                "Expired" => "status-expired",
                _ => ""
            };
            
            var daysDisplay = cred.DaysUntilExpiry < 0 
                ? $"{Math.Abs(cred.DaysUntilExpiry)} days ago" 
                : $"{cred.DaysUntilExpiry} days";
            
            sb.AppendLine("            <tr>");
            sb.AppendLine($"                <td><strong>{System.Net.WebUtility.HtmlEncode(cred.AppDisplayName)}</strong></td>");
            sb.AppendLine($"                <td style='font-family: monospace; font-size: 12px;'>{System.Net.WebUtility.HtmlEncode(cred.AppId)}</td>");
            sb.AppendLine($"                <td>{System.Net.WebUtility.HtmlEncode(cred.DisplayName ?? "(unnamed)")}</td>");
            sb.AppendLine($"                <td>{cred.EndDateTime:MMM d, yyyy}</td>");
            sb.AppendLine($"                <td>{daysDisplay}</td>");
            sb.AppendLine($"                <td><span class='status-badge {statusClass}'>{cred.Status}</span></td>");
            sb.AppendLine("            </tr>");
        }
        
        sb.AppendLine("        </tbody>");
        sb.AppendLine("    </table>");
    }

    private static string GeneratePublicGroupsHtmlReport(PublicGroupsReportDto data, ReportSettings branding)
    {
        var sb = new StringBuilder();
        
        sb.AppendLine("<!DOCTYPE html>");
        sb.AppendLine("<html lang='en'>");
        sb.AppendLine("<head>");
        sb.AppendLine("    <meta charset='UTF-8'>");
        sb.AppendLine("    <meta name='viewport' content='width=device-width, initial-scale=1.0'>");
        sb.AppendLine("    <title>Microsoft 365 Public Groups Report</title>");
        sb.AppendLine("    <style>");
        sb.AppendLine("        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 0; padding: 20px; background: #f5f5f5; color: #333; }");
        sb.AppendLine("        .container { max-width: 1200px; margin: 0 auto; background: white; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); padding: 30px; }");
        sb.AppendLine("        h1 { color: #0078d4; margin-bottom: 10px; font-size: 28px; }");
        sb.AppendLine("        .subtitle { color: #666; margin-bottom: 20px; font-size: 14px; }");
        sb.AppendLine("        .compliance-note { background: #fff3cd; border: 1px solid #ffc107; border-radius: 8px; padding: 15px; margin-bottom: 30px; }");
        sb.AppendLine("        .compliance-note h4 { color: #856404; margin: 0 0 10px 0; }");
        sb.AppendLine("        .compliance-note p { color: #856404; margin: 0; font-size: 14px; }");
        sb.AppendLine("        .summary { display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 20px; margin-bottom: 40px; }");
        sb.AppendLine("        .summary-card { background: #f8f9fa; border-radius: 8px; padding: 20px; text-align: center; border-left: 4px solid #0078d4; }");
        sb.AppendLine("        .summary-card.warning { border-left-color: #ffc107; }");
        sb.AppendLine("        .summary-card.danger { border-left-color: #dc3545; }");
        sb.AppendLine("        .summary-card.teams { border-left-color: #6264a7; }");
        sb.AppendLine("        .summary-card.groups { border-left-color: #0078d4; }");
        sb.AppendLine("        .summary-value { font-size: 36px; font-weight: bold; color: #0078d4; }");
        sb.AppendLine("        .summary-card.warning .summary-value { color: #ffc107; }");
        sb.AppendLine("        .summary-card.danger .summary-value { color: #dc3545; }");
        sb.AppendLine("        .summary-card.teams .summary-value { color: #6264a7; }");
        sb.AppendLine("        .summary-label { color: #666; font-size: 14px; margin-top: 5px; }");
        sb.AppendLine("        h2 { color: #333; font-size: 20px; margin-top: 40px; margin-bottom: 15px; padding-bottom: 10px; border-bottom: 2px solid #eee; }");
        sb.AppendLine("        table { width: 100%; border-collapse: collapse; margin-bottom: 30px; font-size: 14px; }");
        sb.AppendLine("        th { background: #0078d4; color: white; padding: 12px 15px; text-align: left; font-weight: 600; }");
        sb.AppendLine("        td { padding: 12px 15px; border-bottom: 1px solid #eee; }");
        sb.AppendLine("        tr:hover { background: #f8f9fa; }");
        sb.AppendLine("        .badge { padding: 4px 12px; border-radius: 20px; font-size: 12px; font-weight: 600; display: inline-block; }");
        sb.AppendLine("        .badge-team { background: #e8e8fc; color: #6264a7; }");
        sb.AppendLine("        .badge-group { background: #cce5ff; color: #004085; }");
        sb.AppendLine("        .badge-danger { background: #f8d7da; color: #721c24; }");
        sb.AppendLine("        .badge-warning { background: #fff3cd; color: #856404; }");
        sb.AppendLine("        .empty-message { color: #28a745; padding: 20px; background: #d4edda; border-radius: 8px; text-align: center; margin-bottom: 30px; }");
        sb.AppendLine("        .footer { margin-top: 40px; padding-top: 20px; border-top: 1px solid #eee; color: #666; font-size: 12px; text-align: center; }");
        sb.AppendLine("        @media print { body { background: white; } .container { box-shadow: none; } }");
        sb.AppendLine("    </style>");
        sb.AppendLine("</head>");
        sb.AppendLine("<body>");
        sb.AppendLine("<div class='container'>");

        sb.AppendLine(GenerateBrandedHeader(branding, "Microsoft 365 Public Groups Report",
            data.LastUpdated.ToString("dddd, MMMM d, yyyy 'at' h:mm tt") + " UTC"));

        // CIS Compliance Note
        sb.AppendLine("    <div class='compliance-note'>");
        sb.AppendLine("        <h4>\u26a0\ufe0f CIS Control 1.2.1 (L2) Compliance Review</h4>");
        sb.AppendLine("        <p>To ensure that only managed and approved public Microsoft 365 groups exist, in compliance with CIS Control 1.2.1 (L2).</p>");
        sb.AppendLine("        <p>Public groups must be reviewed to prevent unauthorised access and data exposure.</p>");
        sb.AppendLine("    </div>");
        
        // Summary Cards
        sb.AppendLine("    <div class='summary'>");
        sb.AppendLine($"        <div class='summary-card'><div class='summary-value'>{data.TotalPublicGroups}</div><div class='summary-label'>Total Public Groups</div></div>");
        sb.AppendLine($"        <div class='summary-card teams'><div class='summary-value'>{data.TotalTeams}</div><div class='summary-label'>Teams</div></div>");
        sb.AppendLine($"        <div class='summary-card groups'><div class='summary-value'>{data.TotalM365Groups}</div><div class='summary-label'>M365 Groups</div></div>");
        
        var noOwnerClass = data.GroupsWithNoOwner > 0 ? "danger" : "";
        sb.AppendLine($"        <div class='summary-card {noOwnerClass}'><div class='summary-value'>{data.GroupsWithNoOwner}</div><div class='summary-label'>Groups Without Owner</div></div>");
        
        var singleOwnerClass = data.GroupsWithSingleOwner > 0 ? "warning" : "";
        sb.AppendLine($"        <div class='summary-card {singleOwnerClass}'><div class='summary-value'>{data.GroupsWithSingleOwner}</div><div class='summary-label'>Single Owner Groups</div></div>");
        sb.AppendLine("    </div>");
        
        // Groups Table
        sb.AppendLine("    <h2>\ud83d\udcca Public Groups Details</h2>");
        
        if (data.Groups.Count == 0)
        {
            sb.AppendLine("    <div class='empty-message'>\u2713 No public groups found - all groups are private</div>");
        }
        else
        {
            sb.AppendLine("    <table>");
            sb.AppendLine("        <thead>");
            sb.AppendLine("            <tr>");
            sb.AppendLine("                <th>Group Name</th>");
            sb.AppendLine("                <th>Created Date</th>");
            sb.AppendLine("                <th>Type</th>");
            sb.AppendLine("                <th>Owners</th>");
            sb.AppendLine("                <th>Members</th>");
            sb.AppendLine("            </tr>");
            sb.AppendLine("        </thead>");
            sb.AppendLine("        <tbody>");
            
            foreach (var group in data.Groups)
            {
                var typeBadge = group.IsTeam 
                    ? "<span class='badge badge-team'>Team</span>" 
                    : "<span class='badge badge-group'>M365 Group</span>";
                
                var ownerDisplay = group.OwnerCount.ToString();
                if (group.OwnerCount == 0)
                {
                    ownerDisplay = "<span class='badge badge-danger'>0 - No Owner</span>";
                }
                else if (group.OwnerCount == 1)
                {
                    ownerDisplay = "<span class='badge badge-warning'>1 - Single Owner</span>";
                }
                
                sb.AppendLine("            <tr>");
                sb.AppendLine($"                <td><strong>{System.Net.WebUtility.HtmlEncode(group.DisplayName)}</strong></td>");
                sb.AppendLine($"                <td>{group.CreatedDateTime?.ToString("MMM d, yyyy") ?? "N/A"}</td>");
                sb.AppendLine($"                <td>{typeBadge}</td>");
                sb.AppendLine($"                <td>{ownerDisplay}</td>");
                sb.AppendLine($"                <td>{group.MemberCount}</td>");
                sb.AppendLine("            </tr>");
            }
            
            sb.AppendLine("        </tbody>");
            sb.AppendLine("    </table>");
        }
        
        sb.AppendLine(GenerateBrandedFooter(branding,
            "Public groups should be reviewed regularly to ensure compliance with your organisation\u2019s security policies."));

        sb.AppendLine("</div>");
        sb.AppendLine("</body>");
        sb.AppendLine("</html>");
        
        return sb.ToString();
    }

    private static string GenerateStalePrivilegedAccountsHtmlReport(StalePrivilegedAccountsReportDto data, ReportSettings branding)
    {
        var sb = new StringBuilder();
        
        sb.AppendLine("<!DOCTYPE html>");
        sb.AppendLine("<html lang='en'>");
        sb.AppendLine("<head>");
        sb.AppendLine("    <meta charset='UTF-8'>");
        sb.AppendLine("    <meta name='viewport' content='width=device-width, initial-scale=1.0'>");
        sb.AppendLine("    <title>Stale Privileged Accounts Report</title>");
        sb.AppendLine("    <style>");
        sb.AppendLine("        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 0; padding: 20px; background: #f5f5f5; color: #333; }");
        sb.AppendLine("        .container { max-width: 1200px; margin: 0 auto; background: white; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); padding: 30px; }");
        sb.AppendLine("        h1 { color: #0078d4; margin-bottom: 10px; font-size: 28px; }");
        sb.AppendLine("        .subtitle { color: #666; margin-bottom: 20px; font-size: 14px; }");
        sb.AppendLine("        .compliance-note { background: #fff3cd; border: 1px solid #ffc107; border-radius: 8px; padding: 15px; margin-bottom: 30px; }");
        sb.AppendLine("        .compliance-note h4 { color: #856404; margin: 0 0 10px 0; }");
        sb.AppendLine("        .compliance-note p { color: #856404; margin: 0; font-size: 14px; }");
        sb.AppendLine("        .summary { display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 20px; margin-bottom: 40px; }");
        sb.AppendLine("        .summary-card { background: #f8f9fa; border-radius: 8px; padding: 20px; text-align: center; border-left: 4px solid #0078d4; }");
        sb.AppendLine("        .summary-card.warning { border-left-color: #ffc107; }");
        sb.AppendLine("        .summary-card.danger { border-left-color: #dc3545; }");
        sb.AppendLine("        .summary-card.success { border-left-color: #28a745; }");
        sb.AppendLine("        .summary-value { font-size: 36px; font-weight: bold; color: #0078d4; }");
        sb.AppendLine("        .summary-card.warning .summary-value { color: #ffc107; }");
        sb.AppendLine("        .summary-card.danger .summary-value { color: #dc3545; }");
        sb.AppendLine("        .summary-card.success .summary-value { color: #28a745; }");
        sb.AppendLine("        .summary-label { color: #666; font-size: 14px; margin-top: 5px; }");
        sb.AppendLine("        h2 { color: #333; font-size: 20px; margin-top: 40px; margin-bottom: 15px; padding-bottom: 10px; border-bottom: 2px solid #eee; }");
        sb.AppendLine("        h2.danger { color: #721c24; }");
        sb.AppendLine("        table { width: 100%; border-collapse: collapse; margin-bottom: 30px; font-size: 14px; }");
        sb.AppendLine("        th { background: #0078d4; color: white; padding: 12px 15px; text-align: left; font-weight: 600; }");
        sb.AppendLine("        td { padding: 12px 15px; border-bottom: 1px solid #eee; }");
        sb.AppendLine("        tr:hover { background: #f8f9fa; }");
        sb.AppendLine("        tr.stale { background: #fff3cd; }");
        sb.AppendLine("        tr.never-signed-in { background: #f8d7da; }");
        sb.AppendLine("        .badge { padding: 4px 12px; border-radius: 20px; font-size: 12px; font-weight: 600; display: inline-block; }");
        sb.AppendLine("        .badge-enabled { background: #d4edda; color: #155724; }");
        sb.AppendLine("        .badge-disabled { background: #f8d7da; color: #721c24; }");
        sb.AppendLine("        .badge-danger { background: #f8d7da; color: #721c24; }");
        sb.AppendLine("        .badge-warning { background: #fff3cd; color: #856404; }");
        sb.AppendLine("        .empty-message { color: #28a745; padding: 20px; background: #d4edda; border-radius: 8px; text-align: center; margin-bottom: 30px; }");
        sb.AppendLine("        .roles-list { background: #f8f9fa; border-radius: 8px; padding: 15px; margin-bottom: 30px; }");
        sb.AppendLine("        .roles-list h4 { margin: 0 0 10px 0; color: #333; }");
        sb.AppendLine("        .roles-list ul { margin: 0; padding-left: 20px; }");
        sb.AppendLine("        .roles-list li { margin: 5px 0; }");
        sb.AppendLine("        .footer { margin-top: 40px; padding-top: 20px; border-top: 1px solid #eee; color: #666; font-size: 12px; text-align: center; }");
        sb.AppendLine("        @media print { body { background: white; } .container { box-shadow: none; } }");
        sb.AppendLine("    </style>");
        sb.AppendLine("    <script>");
        sb.AppendLine("        function filterByType(el, type) {");
        sb.AppendLine("            document.querySelectorAll('.summary-card').forEach(function(c) { c.classList.remove('active'); });");
        sb.AppendLine("            if (type !== 'all') el.classList.add('active');");
        sb.AppendLine("            document.querySelectorAll('tbody tr').forEach(function(row) {");
        sb.AppendLine("                if (type === 'all') { row.style.display = ''; }");
        sb.AppendLine("                else if (type === 'stale') { row.style.display = row.classList.contains('stale') ? '' : 'none'; }");
        sb.AppendLine("                else if (type === 'never') { row.style.display = row.classList.contains('never-signed-in') ? '' : 'none'; }");
        sb.AppendLine("                else if (type === 'disabled') { row.style.display = row.dataset.status === 'Disabled' ? '' : 'none'; }");
        sb.AppendLine("            });");
        sb.AppendLine("        }");
        sb.AppendLine("    </script>");
        sb.AppendLine("</head>");
        sb.AppendLine("<body>");
        sb.AppendLine("<div class='container'>");

        sb.AppendLine(GenerateBrandedHeader(branding, "Stale Privileged Accounts Report",
            data.LastUpdated.ToString("dddd, MMMM d, yyyy 'at' h:mm tt") + $" UTC \u2022 Threshold: {data.InactiveDaysThreshold} days"));

        // MT.1029 Compliance Note
        sb.AppendLine("    <div class='compliance-note'>");
        sb.AppendLine("        <h4>\u26a0\ufe0f MT.1029 Compliance Review</h4>");
        sb.AppendLine("        <p>This report helps ensure compliance with MT.1029: Stale accounts are not assigned to privileged roles.</p>");
        sb.AppendLine("        <p>Stale accounts pose a security risk as they can be compromised and exploited for unauthorized access.</p>");
        sb.AppendLine("    </div>");
        
        // Summary Cards
        sb.AppendLine("    <p style='color: #666; font-size: 13px; margin-bottom: 10px;'>💡 Click on the cards below to filter the table</p>");
        sb.AppendLine("    <div class='summary'>");
        sb.AppendLine($"        <div class='summary-card' onclick=\"filterByType(this,'all')\"><div class='summary-value'>{data.TotalPrivilegedUsers}</div><div class='summary-label'>Total Privileged Users</div></div>");
        
        var staleClass = data.TotalStaleAccounts > 0 ? "danger" : "success";
        sb.AppendLine($"        <div class='summary-card {staleClass}' onclick=\"filterByType(this,'stale')\"><div class='summary-value'>{data.TotalStaleAccounts}</div><div class='summary-label'>Stale Accounts</div></div>");
        
        var neverClass = data.AccountsNeverSignedIn > 0 ? "danger" : "success";
        sb.AppendLine($"        <div class='summary-card {neverClass}' onclick=\"filterByType(this,'never')\"><div class='summary-value'>{data.AccountsNeverSignedIn}</div><div class='summary-label'>Never Signed In</div></div>");
        
        var disabledClass = data.AccountsDisabled > 0 ? "warning" : "";
        sb.AppendLine($"        <div class='summary-card {disabledClass}' onclick=\"filterByType(this,'disabled')\"><div class='summary-value'>{data.AccountsDisabled}</div><div class='summary-label'>Disabled Accounts</div></div>");
        sb.AppendLine("    </div>");
        
        // Monitored Roles
        sb.AppendLine("    <div class='roles-list'>");
        sb.AppendLine("        <h4>\ud83d\udc65 Monitored Privileged Roles</h4>");
        sb.AppendLine("        <ul>");
        foreach (var role in data.MonitoredRoles)
        {
            sb.AppendLine($"            <li>{System.Net.WebUtility.HtmlEncode(role)}</li>");
        }
        sb.AppendLine("        </ul>");
        sb.AppendLine("    </div>");
        
        // Stale Accounts Table
        sb.AppendLine("    <h2 class='danger'>\u26a0\ufe0f Stale Privileged Accounts</h2>");
        sb.AppendLine($"    <p id='visible-count' style='color: #666; margin-bottom: 15px;'>Showing {data.StaleAccounts.Count} account(s)</p>");
        
        if (data.StaleAccounts.Count == 0)
        {
            sb.AppendLine("    <div class='empty-message'>\u2713 No stale privileged accounts found - all accounts are active</div>");
        }
        else
        {
            sb.AppendLine("    <table>");
            sb.AppendLine("        <thead>");
            sb.AppendLine("            <tr>");
            sb.AppendLine("                <th>Display Name</th>");
            sb.AppendLine("                <th>User Principal Name</th>");
            sb.AppendLine("                <th>Role(s)</th>");
            sb.AppendLine("                <th>Last Sign-In</th>");
            sb.AppendLine("                <th>Days Inactive</th>");
            sb.AppendLine("                <th>Status</th>");
            sb.AppendLine("            </tr>");
            sb.AppendLine("        </thead>");
            sb.AppendLine("        <tbody>");
            
            foreach (var account in data.StaleAccounts)
            {
                var rowClass = !account.LastSignIn.HasValue ? "never-signed-in" : "stale";
                var statusBadge = account.AccountStatus == "Enabled" 
                    ? "<span class='badge badge-enabled'>Enabled</span>" 
                    : "<span class='badge badge-disabled'>Disabled</span>";
                
                var lastSignInDisplay = account.LastSignIn.HasValue 
                    ? account.LastSignIn.Value.ToString("MMM d, yyyy HH:mm") 
                    : "<span class='badge badge-danger'>Never</span>";
                    
                var daysInactiveDisplay = account.DaysSinceLastSignIn == -1 
                    ? "<span class='badge badge-danger'>Never</span>" 
                    : $"<span class='badge badge-warning'>{account.DaysSinceLastSignIn} days</span>";
                
                sb.AppendLine($"            <tr class='{rowClass}' data-status='{account.AccountStatus}'>");
                sb.AppendLine($"                <td><strong>{System.Net.WebUtility.HtmlEncode(account.DisplayName)}</strong></td>");
                sb.AppendLine($"                <td>{System.Net.WebUtility.HtmlEncode(account.UserPrincipalName)}</td>");
                sb.AppendLine($"                <td>{System.Net.WebUtility.HtmlEncode(account.Role)}</td>");
                sb.AppendLine($"                <td>{lastSignInDisplay}</td>");
                sb.AppendLine($"                <td>{daysInactiveDisplay}</td>");
                sb.AppendLine($"                <td>{statusBadge}</td>");
                sb.AppendLine("            </tr>");
            }
            
            sb.AppendLine("        </tbody>");
            sb.AppendLine("    </table>");
        }
        
        sb.AppendLine(GenerateBrandedFooter(branding,
            "Stale privileged accounts should be reviewed and remediated promptly. Recommended actions: disable or remove stale accounts from privileged roles, or ensure the user signs in to confirm account activity."));

        sb.AppendLine("</div>");
        sb.AppendLine("</body>");
        sb.AppendLine("</html>");
        
        return sb.ToString();
    }

    private static string GenerateTeamsUsageHtmlReport(TeamsActivityDto data, string dateRange, ReportSettings branding)
    {
        var sb = new StringBuilder();
        
        sb.AppendLine("<!DOCTYPE html>");
        sb.AppendLine("<html lang='en'>");
        sb.AppendLine("<head>");
        sb.AppendLine("    <meta charset='UTF-8'>");
        sb.AppendLine("    <meta name='viewport' content='width=device-width, initial-scale=1.0'>");
        sb.AppendLine("    <title>Teams Usage Report</title>");
        sb.AppendLine("    <style>");
        sb.AppendLine("        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 0; padding: 20px; background: #f5f5f5; color: #333; }");
        sb.AppendLine("        .container { max-width: 1200px; margin: 0 auto; background: white; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); padding: 30px; }");
        sb.AppendLine("        h1 { color: #6264a7; margin-bottom: 10px; font-size: 28px; }");
        sb.AppendLine("        .subtitle { color: #666; margin-bottom: 30px; font-size: 14px; }");
        sb.AppendLine("        .summary { display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 20px; margin-bottom: 40px; }");
        sb.AppendLine("        .summary-card { background: #f8f9fa; border-radius: 8px; padding: 20px; text-align: center; border-left: 4px solid #6264a7; }");
        sb.AppendLine("        .summary-card.messages { border-left-color: #6264a7; }");
        sb.AppendLine("        .summary-card.calls { border-left-color: #28a745; }");
        sb.AppendLine("        .summary-card.meetings { border-left-color: #0078d4; }");
        sb.AppendLine("        .summary-card.users { border-left-color: #fd7e14; }");
        sb.AppendLine("        .summary-value { font-size: 36px; font-weight: bold; color: #6264a7; }");
        sb.AppendLine("        .summary-card.messages .summary-value { color: #6264a7; }");
        sb.AppendLine("        .summary-card.calls .summary-value { color: #28a745; }");
        sb.AppendLine("        .summary-card.meetings .summary-value { color: #0078d4; }");
        sb.AppendLine("        .summary-card.users .summary-value { color: #fd7e14; }");
        sb.AppendLine("        .summary-label { color: #666; font-size: 14px; margin-top: 5px; }");
        sb.AppendLine("        h2 { color: #333; font-size: 20px; margin-top: 40px; margin-bottom: 15px; padding-bottom: 10px; border-bottom: 2px solid #eee; }");
        sb.AppendLine("        table { width: 100%; border-collapse: collapse; margin-bottom: 30px; font-size: 14px; }");
        sb.AppendLine("        th { background: #6264a7; color: white; padding: 12px 15px; text-align: left; font-weight: 600; }");
        sb.AppendLine("        td { padding: 12px 15px; border-bottom: 1px solid #eee; }");
        sb.AppendLine("        tr:hover { background: #f8f9fa; }");
        sb.AppendLine("        tr:nth-child(even) { background: #fafafa; }");
        sb.AppendLine("        .empty-message { color: #666; padding: 20px; background: #f8f9fa; border-radius: 8px; text-align: center; }");
        sb.AppendLine("        .footer { margin-top: 40px; padding-top: 20px; border-top: 1px solid #eee; color: #666; font-size: 12px; text-align: center; }");
        sb.AppendLine("        @media print { body { background: white; } .container { box-shadow: none; } }");
        sb.AppendLine("    </style>");
        sb.AppendLine("</head>");
        sb.AppendLine("<body>");
        sb.AppendLine("<div class='container'>");

        sb.AppendLine(GenerateBrandedHeader(branding, "Teams Usage Report",
            data.LastUpdated.ToString("dddd, MMMM d, yyyy 'at' h:mm tt") + $" UTC \u2022 {dateRange}"));
        
        // Summary Cards
        sb.AppendLine("    <div class='summary'>");
        sb.AppendLine($"        <div class='summary-card messages'><div class='summary-value'>{data.TotalMessages:N0}</div><div class='summary-label'>Total Messages</div></div>");
        sb.AppendLine($"        <div class='summary-card calls'><div class='summary-value'>{data.TotalCalls:N0}</div><div class='summary-label'>Total Calls</div></div>");
        sb.AppendLine($"        <div class='summary-card meetings'><div class='summary-value'>{data.TotalMeetings:N0}</div><div class='summary-label'>Total Meetings</div></div>");
        sb.AppendLine("    </div>");
        
        // Daily Activity Table
        sb.AppendLine("    <h2>\ud83d\udcc8 Daily Teams Activity</h2>");
        if (data.Trend.Count == 0)
        {
            sb.AppendLine("    <div class='empty-message'>No Teams activity data available for this period.</div>");
        }
        else
        {
            sb.AppendLine("    <table>");
            sb.AppendLine("        <thead>");
            sb.AppendLine("            <tr>");
            sb.AppendLine("                <th>Date</th>");
            sb.AppendLine("                <th>Messages</th>");
            sb.AppendLine("                <th>Calls</th>");
            sb.AppendLine("                <th>Meetings</th>");
            sb.AppendLine("            </tr>");
            sb.AppendLine("        </thead>");
            sb.AppendLine("        <tbody>");
            
            foreach (var day in data.Trend.OrderByDescending(d => d.Date))
            {
                sb.AppendLine("            <tr>");
                sb.AppendLine($"                <td>{day.Date:ddd, MMM d, yyyy}</td>");
                sb.AppendLine($"                <td>{day.Messages:N0}</td>");
                sb.AppendLine($"                <td>{day.Calls:N0}</td>");
                sb.AppendLine($"                <td>{day.Meetings:N0}</td>");
                sb.AppendLine("            </tr>");
            }
            
            sb.AppendLine("        </tbody>");
            sb.AppendLine("    </table>");
        }
        
        sb.AppendLine(GenerateBrandedFooter(branding, "Teams activity data is provided by Microsoft Graph reporting APIs."));

        sb.AppendLine("</div>");
        sb.AppendLine("</body>");
        sb.AppendLine("</html>");
        
        return sb.ToString();
    }

    private static string GenerateInactiveUsersHtmlReport(List<TenantUserDto> users, int inactiveDays, string dateRange, ReportSettings branding)
    {
        var sb = new StringBuilder();
        
        var neverSignedIn = users.Count(u => u.LastSignInDateTime == null);
        var staleUsers = users.Count(u => u.LastSignInDateTime != null);
        var disabledUsers = users.Count(u => !u.AccountEnabled);
        var enabledUsers = users.Count(u => u.AccountEnabled);
        var guestUsers = users.Count(u => u.UserType == "Guest");
        var memberUsers = users.Count(u => u.UserType != "Guest");
        
        sb.AppendLine("<!DOCTYPE html>");
        sb.AppendLine("<html lang='en'>");
        sb.AppendLine("<head>");
        sb.AppendLine("    <meta charset='UTF-8'>");
        sb.AppendLine("    <meta name='viewport' content='width=device-width, initial-scale=1.0'>");
        sb.AppendLine("    <title>Inactive Users Report</title>");
        sb.AppendLine("    <style>");
        sb.AppendLine("        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 0; padding: 20px; background: #f5f5f5; color: #333; }");
        sb.AppendLine("        .container { max-width: 1400px; margin: 0 auto; background: white; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); padding: 30px; }");
        sb.AppendLine("        h1 { color: #0078d4; margin-bottom: 10px; font-size: 28px; }");
        sb.AppendLine("        .subtitle { color: #666; margin-bottom: 20px; font-size: 14px; }");
        sb.AppendLine("        .summary { display: grid; grid-template-columns: repeat(auto-fit, minmax(150px, 1fr)); gap: 15px; margin-bottom: 30px; }");
        sb.AppendLine("        .summary-card { background: #f8f9fa; border-radius: 8px; padding: 20px; text-align: center; border-left: 4px solid #0078d4; cursor: pointer; transition: transform 0.2s, box-shadow 0.2s; }");
        sb.AppendLine("        .summary-card:hover { transform: translateY(-2px); box-shadow: 0 4px 8px rgba(0,0,0,0.1); }");
        sb.AppendLine("        .summary-card.active { box-shadow: 0 0 0 3px #0078d4; }");
        sb.AppendLine("        .summary-card.warning { border-left-color: #ffc107; }");
        sb.AppendLine("        .summary-card.danger { border-left-color: #dc3545; }");
        sb.AppendLine("        .summary-card.success { border-left-color: #28a745; }");
        sb.AppendLine("        .summary-card.info { border-left-color: #17a2b8; }");
        sb.AppendLine("        .summary-value { font-size: 32px; font-weight: bold; color: #0078d4; }");
        sb.AppendLine("        .summary-card.warning .summary-value { color: #ffc107; }");
        sb.AppendLine("        .summary-card.danger .summary-value { color: #dc3545; }");
        sb.AppendLine("        .summary-card.success .summary-value { color: #28a745; }");
        sb.AppendLine("        .summary-card.info .summary-value { color: #17a2b8; }");
        sb.AppendLine("        .summary-label { color: #666; font-size: 12px; margin-top: 5px; }");
        sb.AppendLine("        h2 { color: #333; font-size: 20px; margin-top: 30px; margin-bottom: 15px; padding-bottom: 10px; border-bottom: 2px solid #eee; }");
        sb.AppendLine("        table { width: 100%; border-collapse: collapse; margin-bottom: 30px; font-size: 14px; }");
        sb.AppendLine("        th { background: #0078d4; color: white; padding: 12px 15px; text-align: left; font-weight: 600; position: sticky; top: 0; }");
        sb.AppendLine("        td { padding: 10px 15px; border-bottom: 1px solid #eee; }");
        sb.AppendLine("        tr:hover { background: #f8f9fa; }");
        sb.AppendLine("        tr.never-signed-in { background: #f8d7da; }");
        sb.AppendLine("        tr.stale { background: #fff3cd; }");
        sb.AppendLine("        .badge { padding: 4px 10px; border-radius: 12px; font-size: 11px; font-weight: 600; display: inline-block; }");
        sb.AppendLine("        .badge-success { background: #d4edda; color: #155724; }");
        sb.AppendLine("        .badge-warning { background: #fff3cd; color: #856404; }");
        sb.AppendLine("        .badge-danger { background: #f8d7da; color: #721c24; }");
        sb.AppendLine("        .badge-info { background: #cce5ff; color: #004085; }");
        sb.AppendLine("        .badge-secondary { background: #e2e3e5; color: #383d41; }");
        sb.AppendLine("        .empty-message { color: #28a745; padding: 20px; background: #d4edda; border-radius: 8px; text-align: center; }");
        sb.AppendLine("        .footer { margin-top: 40px; padding-top: 20px; border-top: 1px solid #eee; color: #666; font-size: 12px; text-align: center; }");
        sb.AppendLine("        .table-wrapper { overflow-x: auto; }");
        sb.AppendLine("        @media print { body { background: white; } .container { box-shadow: none; } }");
        sb.AppendLine("    </style>");
        sb.AppendLine("    <script>");
        sb.AppendLine("        function filterByType(el, type) {");
        sb.AppendLine("            document.querySelectorAll('.summary-card').forEach(card => card.classList.remove('active'));");
        sb.AppendLine("            if (type !== 'all') {");
        sb.AppendLine("                el.closest('.summary-card').classList.add('active');");
        sb.AppendLine("            }");
        sb.AppendLine("            document.querySelectorAll('tbody tr').forEach(row => {");
        sb.AppendLine("                if (type === 'all') {");
        sb.AppendLine("                    row.style.display = '';");
        sb.AppendLine("                } else if (type === 'never') {");
        sb.AppendLine("                    row.style.display = row.classList.contains('never-signed-in') ? '' : 'none';");
        sb.AppendLine("                } else if (type === 'stale') {");
        sb.AppendLine("                    row.style.display = row.classList.contains('stale') ? '' : 'none';");
        sb.AppendLine("                } else if (type === 'disabled') {");
        sb.AppendLine("                    row.style.display = row.dataset.status === 'Disabled' ? '' : 'none';");
        sb.AppendLine("                } else if (type === 'enabled') {");
        sb.AppendLine("                    row.style.display = row.dataset.status === 'Enabled' ? '' : 'none';");
        sb.AppendLine("                } else if (type === 'guest') {");
        sb.AppendLine("                    row.style.display = row.dataset.usertype === 'Guest' ? '' : 'none';");
        sb.AppendLine("                } else if (type === 'member') {");
        sb.AppendLine("                    row.style.display = row.dataset.usertype !== 'Guest' ? '' : 'none';");
        sb.AppendLine("                }");
        sb.AppendLine("            });");
        sb.AppendLine("            updateVisibleCount();");
        sb.AppendLine("        }");
        sb.AppendLine("        function updateVisibleCount() {");
        sb.AppendLine("            var visible = document.querySelectorAll('tbody tr:not([style*=\"display: none\"])').length;");
        sb.AppendLine("            var counter = document.getElementById('visible-count');");
        sb.AppendLine("            if (counter) counter.textContent = 'Showing ' + visible + ' user(s)';");
        sb.AppendLine("        }");
        sb.AppendLine("    </script>");
        sb.AppendLine("</head>");
        sb.AppendLine("<body>");
        sb.AppendLine("<div class='container'>");

        sb.AppendLine(GenerateBrandedHeader(branding, "Inactive Users Report",
            DateTime.UtcNow.ToString("dddd, MMMM d, yyyy 'at' h:mm tt") + $" UTC \u2022 Threshold: {inactiveDays} days"));
        
        // Summary Cards
        sb.AppendLine("    <p style='color: #666; font-size: 13px; margin-bottom: 10px;'>\ud83d\udca1 Click on the cards below to filter the table</p>");
        sb.AppendLine("    <div class='summary'>");
        sb.AppendLine($"        <div class='summary-card' onclick=\"filterByType(this,'all')\"><div class='summary-value'>{users.Count}</div><div class='summary-label'>Total Inactive</div></div>");
        
        var neverClass = neverSignedIn > 0 ? "danger" : "success";
        sb.AppendLine($"        <div class='summary-card {neverClass}' onclick=\"filterByType(this,'never')\"><div class='summary-value'>{neverSignedIn}</div><div class='summary-label'>Never Signed In</div></div>");
        
        var staleClass = staleUsers > 0 ? "warning" : "success";
        sb.AppendLine($"        <div class='summary-card {staleClass}' onclick=\"filterByType(this,'stale')\"><div class='summary-value'>{staleUsers}</div><div class='summary-label'>Stale Users</div></div>");
        
        var disabledClass = disabledUsers > 0 ? "info" : "";
        sb.AppendLine($"        <div class='summary-card {disabledClass}' onclick=\"filterByType(this,'disabled')\"><div class='summary-value'>{disabledUsers}</div><div class='summary-label'>Disabled</div></div>");
        
        sb.AppendLine($"        <div class='summary-card' onclick=\"filterByType(this,'enabled')\"><div class='summary-value'>{enabledUsers}</div><div class='summary-label'>Enabled</div></div>");
        
        if (guestUsers > 0)
        {
            sb.AppendLine($"        <div class='summary-card info' onclick=\"filterByType(this,'guest')\"><div class='summary-value'>{guestUsers}</div><div class='summary-label'>Guest Users</div></div>");
        }
        
        sb.AppendLine("    </div>");
        
        // Users Table
        sb.AppendLine("    <h2>\ud83d\udcca Inactive Users</h2>");
        sb.AppendLine($"    <p id='visible-count' style='color: #666; margin-bottom: 15px;'>Showing {users.Count} user(s)</p>");
        
        if (users.Count == 0)
        {
            sb.AppendLine("    <div class='empty-message'>\u2713 No inactive users found - all users are active</div>");
        }
        else
        {
            sb.AppendLine("    <div class='table-wrapper'>");
            sb.AppendLine("    <table>");
            sb.AppendLine("        <thead>");
            sb.AppendLine("            <tr>");
            sb.AppendLine("                <th>Display Name</th>");
            sb.AppendLine("                <th>User Principal Name</th>");
            sb.AppendLine("                <th>User Type</th>");
            sb.AppendLine("                <th>Last Sign-In</th>");
            sb.AppendLine("                <th>Days Inactive</th>");
            sb.AppendLine("                <th>Status</th>");
            sb.AppendLine("            </tr>");
            sb.AppendLine("        </thead>");
            sb.AppendLine("        <tbody>");
            
            foreach (var user in users.OrderByDescending(u => u.LastSignInDateTime == null).ThenBy(u => u.LastSignInDateTime))
            {
                var rowClass = user.LastSignInDateTime == null ? "never-signed-in" : "stale";
                var accountStatus = user.AccountEnabled ? "Enabled" : "Disabled";
                var statusBadge = user.AccountEnabled 
                    ? "<span class='badge badge-success'>Enabled</span>" 
                    : "<span class='badge badge-danger'>Disabled</span>";
                
                var userTypeBadge = user.UserType == "Guest"
                    ? "<span class='badge badge-info'>Guest</span>"
                    : "<span class='badge badge-secondary'>Member</span>";
                
                var lastSignInDisplay = user.LastSignInDateTime.HasValue 
                    ? user.LastSignInDateTime.Value.ToString("MMM d, yyyy HH:mm") 
                    : "<span class='badge badge-danger'>Never</span>";
                
                var daysInactive = user.LastSignInDateTime.HasValue
                    ? (int)(DateTime.UtcNow - user.LastSignInDateTime.Value).TotalDays
                    : -1;
                    
                var daysInactiveDisplay = daysInactive == -1 
                    ? "<span class='badge badge-danger'>Never</span>" 
                    : $"<span class='badge badge-warning'>{daysInactive} days</span>";
                
                sb.AppendLine($"            <tr class='{rowClass}' data-status='{accountStatus}' data-usertype='{user.UserType ?? "Member"}'>");
                sb.AppendLine($"                <td><strong>{System.Net.WebUtility.HtmlEncode(user.DisplayName ?? "")}</strong></td>");
                sb.AppendLine($"                <td>{System.Net.WebUtility.HtmlEncode(user.UserPrincipalName ?? "")}</td>");
                sb.AppendLine($"                <td>{userTypeBadge}</td>");
                sb.AppendLine($"                <td>{lastSignInDisplay}</td>");
                sb.AppendLine($"                <td>{daysInactiveDisplay}</td>");
                sb.AppendLine($"                <td>{statusBadge}</td>");
                sb.AppendLine("            </tr>");
            }
            
            sb.AppendLine("        </tbody>");
            sb.AppendLine("    </table>");
            sb.AppendLine("    </div>");
        }
        
        sb.AppendLine(GenerateBrandedFooter(branding,
            "Users who haven\u2019t signed in within the specified period should be reviewed for potential disabling or removal."));

        sb.AppendLine("</div>");
        sb.AppendLine("</body>");
        sb.AppendLine("</html>");
        
        return sb.ToString();
    }

    private static string GenerateMailboxUsageHtmlReport(MailflowSummaryDto data, string dateRange, ReportSettings branding)
    {
        var sb = new StringBuilder();
        
        sb.AppendLine("<!DOCTYPE html>");
        sb.AppendLine("<html lang='en'>");
        sb.AppendLine("<head>");
        sb.AppendLine("    <meta charset='UTF-8'>");
        sb.AppendLine("    <meta name='viewport' content='width=device-width, initial-scale=1.0'>");
        sb.AppendLine("    <title>Mailbox Usage Report</title>");
        sb.AppendLine("    <style>");
        sb.AppendLine("        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 0; padding: 20px; background: #f5f5f5; color: #333; }");
        sb.AppendLine("        .container { max-width: 1200px; margin: 0 auto; background: white; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); padding: 30px; }");
        sb.AppendLine("        h1 { color: #0078d4; margin-bottom: 10px; font-size: 28px; }");
        sb.AppendLine("        .subtitle { color: #666; margin-bottom: 30px; font-size: 14px; }");
        sb.AppendLine("        .summary { display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 20px; margin-bottom: 40px; }");
        sb.AppendLine("        .summary-card { background: #f8f9fa; border-radius: 8px; padding: 20px; text-align: center; border-left: 4px solid #0078d4; }");
        sb.AppendLine("        .summary-card.sent { border-left-color: #28a745; }");
        sb.AppendLine("        .summary-card.received { border-left-color: #17a2b8; }");
        sb.AppendLine("        .summary-card.avg { border-left-color: #6f42c1; }");
        sb.AppendLine("        .summary-value { font-size: 36px; font-weight: bold; color: #0078d4; }");
        sb.AppendLine("        .summary-card.sent .summary-value { color: #28a745; }");
        sb.AppendLine("        .summary-card.received .summary-value { color: #17a2b8; }");
        sb.AppendLine("        .summary-card.avg .summary-value { color: #6f42c1; }");
        sb.AppendLine("        .summary-label { color: #666; font-size: 14px; margin-top: 5px; }");
        sb.AppendLine("        h2 { color: #333; font-size: 20px; margin-top: 40px; margin-bottom: 15px; padding-bottom: 10px; border-bottom: 2px solid #eee; }");
        sb.AppendLine("        table { width: 100%; border-collapse: collapse; margin-bottom: 30px; font-size: 14px; }");
        sb.AppendLine("        th { background: #0078d4; color: white; padding: 12px 15px; text-align: left; font-weight: 600; }");
        sb.AppendLine("        td { padding: 12px 15px; border-bottom: 1px solid #eee; }");
        sb.AppendLine("        tr:hover { background: #f8f9fa; }");
        sb.AppendLine("        tr:nth-child(even) { background: #fafafa; }");
        sb.AppendLine("        .section-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 30px; }");
        sb.AppendLine("        @media (max-width: 768px) { .section-grid { grid-template-columns: 1fr; } }");
        sb.AppendLine("        .empty-message { color: #666; padding: 20px; background: #f8f9fa; border-radius: 8px; text-align: center; }");
        sb.AppendLine("        .footer { margin-top: 40px; padding-top: 20px; border-top: 1px solid #eee; color: #666; font-size: 12px; text-align: center; }");
        sb.AppendLine("        @media print { body { background: white; } .container { box-shadow: none; } }");
        sb.AppendLine("    </style>");
        sb.AppendLine("</head>");
        sb.AppendLine("<body>");
        sb.AppendLine("<div class='container'>");

        sb.AppendLine(GenerateBrandedHeader(branding, "Mailbox Usage Report",
            data.LastUpdated.ToString("dddd, MMMM d, yyyy 'at' h:mm tt") + $" UTC \u2022 {dateRange}"));
        
        // Summary Cards
        sb.AppendLine("    <div class='summary'>");
        sb.AppendLine($"        <div class='summary-card'><div class='summary-value'>{data.DailyTraffic.Count:N0}</div><div class='summary-label'>Days of Data</div></div>");
        sb.AppendLine($"        <div class='summary-card sent'><div class='summary-value'>{data.TotalMessagesSent:N0}</div><div class='summary-label'>Messages Sent</div></div>");
        sb.AppendLine($"        <div class='summary-card received'><div class='summary-value'>{data.TotalMessagesReceived:N0}</div><div class='summary-label'>Messages Received</div></div>");
        sb.AppendLine($"        <div class='summary-card avg'><div class='summary-value'>{data.AverageMessagesPerDay:N0}</div><div class='summary-label'>Avg Messages/Day</div></div>");
        sb.AppendLine("    </div>");
        
        // Daily Traffic Table
        sb.AppendLine("    <h2>\ud83d\udcc8 Daily Email Traffic</h2>");
        if (data.DailyTraffic.Count == 0)
        {
            sb.AppendLine("    <div class='empty-message'>No email traffic data available for this period.</div>");
        }
        else
        {
            sb.AppendLine("    <table>");
            sb.AppendLine("        <thead>");
            sb.AppendLine("            <tr>");
            sb.AppendLine("                <th>Date</th>");
            sb.AppendLine("                <th>Messages Sent</th>");
            sb.AppendLine("                <th>Messages Received</th>");
            sb.AppendLine("                <th>Total</th>");
            sb.AppendLine("            </tr>");
            sb.AppendLine("        </thead>");
            sb.AppendLine("        <tbody>");
            
            foreach (var day in data.DailyTraffic.OrderByDescending(d => d.Date))
            {
                sb.AppendLine("            <tr>");
                sb.AppendLine($"                <td>{day.Date:ddd, MMM d, yyyy}</td>");
                sb.AppendLine($"                <td>{day.MessagesSent:N0}</td>");
                sb.AppendLine($"                <td>{day.MessagesReceived:N0}</td>");
                sb.AppendLine($"                <td><strong>{(day.MessagesSent + day.MessagesReceived):N0}</strong></td>");
                sb.AppendLine("            </tr>");
            }
            
            sb.AppendLine("        </tbody>");
            sb.AppendLine("    </table>");
        }
        
        // Top Senders and Recipients
        sb.AppendLine("    <div class='section-grid'>");
        
        // Top Senders
        sb.AppendLine("    <div>");
        sb.AppendLine("    <h2>\ud83d\udce4 Top Senders</h2>");
        if (data.TopSenders.Count == 0)
        {
            sb.AppendLine("    <div class='empty-message'>No sender data available.</div>");
        }
        else
        {
            sb.AppendLine("    <table>");
            sb.AppendLine("        <thead>");
            sb.AppendLine("            <tr>");
            sb.AppendLine("                <th>User</th>");
            sb.AppendLine("                <th>Messages Sent</th>");
            sb.AppendLine("            </tr>");
            sb.AppendLine("        </thead>");
            sb.AppendLine("        <tbody>");
            
            foreach (var sender in data.TopSenders.Take(10))
            {
                var displayName = !string.IsNullOrEmpty(sender.DisplayName) ? sender.DisplayName : sender.UserPrincipalName;
                sb.AppendLine("            <tr>");
                sb.AppendLine($"                <td>{System.Net.WebUtility.HtmlEncode(displayName)}</td>");
                sb.AppendLine($"                <td>{sender.MessageCount:N0}</td>");
                sb.AppendLine("            </tr>");
            }
            
            sb.AppendLine("        </tbody>");
            sb.AppendLine("    </table>");
        }
        sb.AppendLine("    </div>");
        
        // Top Recipients
        sb.AppendLine("    <div>");
        sb.AppendLine("    <h2>\ud83d\udce5 Top Recipients</h2>");
        if (data.TopRecipients.Count == 0)
        {
            sb.AppendLine("    <div class='empty-message'>No recipient data available.</div>");
        }
        else
        {
            sb.AppendLine("    <table>");
            sb.AppendLine("        <thead>");
            sb.AppendLine("            <tr>");
            sb.AppendLine("                <th>User</th>");
            sb.AppendLine("                <th>Messages Received</th>");
            sb.AppendLine("            </tr>");
            sb.AppendLine("        </thead>");
            sb.AppendLine("        <tbody>");
            
            foreach (var recipient in data.TopRecipients.Take(10))
            {
                var displayName = !string.IsNullOrEmpty(recipient.DisplayName) ? recipient.DisplayName : recipient.UserPrincipalName;
                sb.AppendLine("            <tr>");
                sb.AppendLine($"                <td>{System.Net.WebUtility.HtmlEncode(displayName)}</td>");
                sb.AppendLine($"                <td>{recipient.MessageCount:N0}</td>");
                sb.AppendLine("            </tr>");
            }
            
            sb.AppendLine("        </tbody>");
            sb.AppendLine("    </table>");
        }
        sb.AppendLine("    </div>");
        
        sb.AppendLine("    </div>"); // End section-grid
        
        sb.AppendLine(GenerateBrandedFooter(branding));

        sb.AppendLine("</div>");
        sb.AppendLine("</body>");
        sb.AppendLine("</html>");
        
        return sb.ToString();
    }

    private static string GenerateCABreakGlassHtmlReport(CABreakGlassReportDto data, ReportSettings branding)
    {
        var sb = new StringBuilder();
        
        sb.AppendLine("<!DOCTYPE html>");
        sb.AppendLine("<html lang='en'>");
        sb.AppendLine("<head>");
        sb.AppendLine("    <meta charset='UTF-8'>");
        sb.AppendLine("    <meta name='viewport' content='width=device-width, initial-scale=1.0'>");
        sb.AppendLine("    <title>Conditional Access Break Glass Report</title>");
        sb.AppendLine("    <style>");
        sb.AppendLine("        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 0; padding: 20px; background: #f5f5f5; color: #333; }");
        sb.AppendLine("        .container { max-width: 1200px; margin: 0 auto; background: white; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); padding: 30px; }");
        sb.AppendLine("        h1 { color: #0078d4; margin-bottom: 10px; font-size: 28px; }");
        sb.AppendLine("        .subtitle { color: #666; margin-bottom: 20px; font-size: 14px; }");
        sb.AppendLine("        .tenant-info { background: #e3f2fd; border-radius: 8px; padding: 15px; margin-bottom: 20px; }");
        sb.AppendLine("        .tenant-info p { margin: 5px 0; color: #1565c0; }");
        sb.AppendLine("        .compliance-note { background: #e8f5e9; border: 1px solid #4caf50; border-radius: 8px; padding: 15px; margin-bottom: 30px; }");
        sb.AppendLine("        .compliance-note h4 { color: #2e7d32; margin: 0 0 10px 0; }");
        sb.AppendLine("        .compliance-note p { color: #2e7d32; margin: 5px 0; font-size: 14px; }");
        sb.AppendLine("        .compliance-note a { color: #1565c0; }");
        sb.AppendLine("        .warning-note { background: #fff3cd; border: 1px solid #ffc107; }");
        sb.AppendLine("        .warning-note h4, .warning-note p { color: #856404; }");
        sb.AppendLine("        .summary { display: grid; grid-template-columns: repeat(auto-fit, minmax(150px, 1fr)); gap: 15px; margin-bottom: 40px; }");
        sb.AppendLine("        .summary-card { background: #f8f9fa; border-radius: 8px; padding: 20px; text-align: center; border-left: 4px solid #0078d4; }");
        sb.AppendLine("        .summary-card.success { border-left-color: #28a745; }");
        sb.AppendLine("        .summary-card.warning { border-left-color: #ffc107; }");
        sb.AppendLine("        .summary-card.danger { border-left-color: #dc3545; }");
        sb.AppendLine("        .summary-card.info { border-left-color: #17a2b8; }");
        sb.AppendLine("        .summary-value { font-size: 32px; font-weight: bold; color: #0078d4; }");
        sb.AppendLine("        .summary-card.success .summary-value { color: #28a745; }");
        sb.AppendLine("        .summary-card.warning .summary-value { color: #ffc107; }");
        sb.AppendLine("        .summary-card.danger .summary-value { color: #dc3545; }");
        sb.AppendLine("        .summary-card.info .summary-value { color: #17a2b8; }");
        sb.AppendLine("        .summary-label { color: #666; font-size: 12px; margin-top: 5px; }");
        sb.AppendLine("        .breakglass-accounts { background: #f8f9fa; border-radius: 8px; padding: 15px; margin-bottom: 30px; }");
        sb.AppendLine("        .breakglass-accounts h4 { margin: 0 0 10px 0; color: #333; }");
        sb.AppendLine("        .breakglass-accounts ul { margin: 0; padding-left: 20px; }");
        sb.AppendLine("        .breakglass-accounts li { margin: 5px 0; }");
        sb.AppendLine("        .breakglass-accounts .resolved { color: #28a745; }");
        sb.AppendLine("        .breakglass-accounts .unresolved { color: #dc3545; }");
        sb.AppendLine("        h2 { color: #333; font-size: 20px; margin-top: 40px; margin-bottom: 15px; padding-bottom: 10px; border-bottom: 2px solid #eee; }");
        sb.AppendLine("        table { width: 100%; border-collapse: collapse; margin-bottom: 30px; font-size: 14px; }");
        sb.AppendLine("        th { background: #0078d4; color: white; padding: 12px 15px; text-align: left; font-weight: 600; }");
        sb.AppendLine("        td { padding: 12px 15px; border-bottom: 1px solid #eee; }");
        sb.AppendLine("        tr:hover { background: #f8f9fa; }");
        sb.AppendLine("        tr.row-warning { background: #fff3cd !important; }");
        sb.AppendLine("        tr.row-danger { background: #f8d7da !important; }");
        sb.AppendLine("        .badge { padding: 4px 10px; border-radius: 12px; font-size: 11px; font-weight: 600; display: inline-block; }");
        sb.AppendLine("        .badge-success { background: #d4edda; color: #155724; }");
        sb.AppendLine("        .badge-warning { background: #fff3cd; color: #856404; }");
        sb.AppendLine("        .badge-danger { background: #f8d7da; color: #721c24; }");
        sb.AppendLine("        .badge-enabled { background: #28a745; color: white; }");
        sb.AppendLine("        .badge-disabled { background: #6c757d; color: white; }");
        sb.AppendLine("        .badge-report { background: #fd7e14; color: white; }");
        sb.AppendLine("        .filter-buttons { margin-bottom: 20px; }");
        sb.AppendLine("        .filter-buttons label { margin-right: 20px; cursor: pointer; }");
        sb.AppendLine("        .empty-message { color: #28a745; padding: 20px; background: #d4edda; border-radius: 8px; text-align: center; margin-bottom: 30px; }");
        sb.AppendLine("        .footer { margin-top: 40px; padding-top: 20px; border-top: 1px solid #eee; color: #666; font-size: 12px; text-align: center; }");
        sb.AppendLine("        @media print { body { background: white; } .container { box-shadow: none; } .filter-buttons { display: none; } }");
        sb.AppendLine("    </style>");
        sb.AppendLine("    <script>");
        sb.AppendLine("        function filterTable(mode) {");
        sb.AppendLine("            let rows = document.querySelectorAll('tbody tr');");
        sb.AppendLine("            rows.forEach(row => {");
        sb.AppendLine("                let filterValue = row.getAttribute('data-filter');");
        sb.AppendLine("                row.style.display = (mode === 'all' || filterValue === mode) ? '' : 'none';");
        sb.AppendLine("            });");
        sb.AppendLine("        }");
        sb.AppendLine("    </script>");
        sb.AppendLine("</head>");
        sb.AppendLine("<body>");
        sb.AppendLine("<div class='container'>");

        sb.AppendLine(GenerateBrandedHeader(branding, "Conditional Access Break Glass Report",
            data.LastUpdated.ToString("dddd, MMMM d, yyyy 'at' h:mm tt") + " UTC"));

        // Tenant Info
        sb.AppendLine("    <div class='tenant-info'>");
        sb.AppendLine($"        <p><strong>Tenant:</strong> {System.Net.WebUtility.HtmlEncode(data.TenantName)}</p>");
        sb.AppendLine($"        <p><strong>Tenant ID:</strong> {System.Net.WebUtility.HtmlEncode(data.TenantId)}</p>");
        sb.AppendLine("    </div>");
        
        // Microsoft Guidance Note
        sb.AppendLine("    <div class='compliance-note'>");
        sb.AppendLine("        <h4>\ud83d\udee1\ufe0f Why Break Glass Accounts Should Be Excluded</h4>");
        sb.AppendLine("        <p>Break glass or emergency access accounts are critical for maintaining administrative access to your environment in case of accidental lockouts, such as a misconfigured Conditional Access policy or MFA failure.</p>");
        sb.AppendLine("        <p>These accounts should be cloud-only, excluded from all Conditional Access policies, and stored securely.</p>");
        sb.AppendLine("        <p>See <a href='https://learn.microsoft.com/en-us/entra/identity/role-based-access-control/security-emergency-access' target='_blank'>Microsoft's official guidance</a> for more information.</p>");
        sb.AppendLine("    </div>");
        
        // Warning if no break glass accounts configured
        if (data.ConfiguredBreakGlassAccounts.Count == 0)
        {
            sb.AppendLine("    <div class='compliance-note warning-note'>");
            sb.AppendLine("        <h4>\u26a0\ufe0f No Break Glass Accounts Configured</h4>");
            sb.AppendLine("        <p>Please configure break glass account UPNs in Settings > Report Settings before running this report.</p>");
            sb.AppendLine("    </div>");
        }
        
        // Summary Cards
        sb.AppendLine("    <div class='summary'>");
        sb.AppendLine($"        <div class='summary-card info'><div class='summary-value'>{data.TotalPolicies}</div><div class='summary-label'>Total Policies</div></div>");
        
        var fullClass = data.PoliciesWithFullExclusion > 0 ? "success" : "";
        sb.AppendLine($"        <div class='summary-card {fullClass}'><div class='summary-value'>{data.PoliciesWithFullExclusion}</div><div class='summary-label'>Fully Excluded</div></div>");
        
        var partialClass = data.PoliciesWithPartialExclusion > 0 ? "warning" : "success";
        sb.AppendLine($"        <div class='summary-card {partialClass}'><div class='summary-value'>{data.PoliciesWithPartialExclusion}</div><div class='summary-label'>Partial Exclusion</div></div>");
        
        var noneClass = data.PoliciesWithNoExclusion > 0 ? "danger" : "success";
        sb.AppendLine($"        <div class='summary-card {noneClass}'><div class='summary-value'>{data.PoliciesWithNoExclusion}</div><div class='summary-label'>Not Excluded</div></div>");
        
        sb.AppendLine($"        <div class='summary-card'><div class='summary-value'>{data.EnabledPolicies}</div><div class='summary-label'>Enabled</div></div>");
        sb.AppendLine($"        <div class='summary-card'><div class='summary-value'>{data.DisabledPolicies}</div><div class='summary-label'>Disabled</div></div>");
        sb.AppendLine($"        <div class='summary-card'><div class='summary-value'>{data.ReportOnlyPolicies}</div><div class='summary-label'>Report-Only</div></div>");
        sb.AppendLine("    </div>");
        
        // Configured Break Glass Accounts
        sb.AppendLine("    <div class='breakglass-accounts'>");
        sb.AppendLine("        <h4>\ud83d\udd11 Configured Break Glass Accounts</h4>");
        if (data.ConfiguredBreakGlassAccounts.Count == 0)
        {
            sb.AppendLine("        <p><em>No break glass accounts configured.</em></p>");
        }
        else
        {
            sb.AppendLine("        <ul>");
            foreach (var account in data.ConfiguredBreakGlassAccounts)
            {
                var statusClass = account.IsResolved ? "resolved" : "unresolved";
                var displayName = !string.IsNullOrEmpty(account.DisplayName) ? $" ({System.Net.WebUtility.HtmlEncode(account.DisplayName)})" : "";
                var status = account.IsResolved ? "\u2713" : "\u2717 Not found";
                sb.AppendLine($"            <li class='{statusClass}'>{System.Net.WebUtility.HtmlEncode(account.UserPrincipalName)}{displayName} - {status}</li>");
            }
            sb.AppendLine("        </ul>");
        }
        sb.AppendLine("    </div>");
        
        // Filter buttons
        sb.AppendLine("    <div class='filter-buttons'>");
        sb.AppendLine("        <label><input type='radio' name='filter' onclick=\"filterTable('all')\" checked> All</label>");
        sb.AppendLine("        <label><input type='radio' name='filter' onclick=\"filterTable('on')\"> Enabled</label>");
        sb.AppendLine("        <label><input type='radio' name='filter' onclick=\"filterTable('off')\"> Disabled</label>");
        sb.AppendLine("        <label><input type='radio' name='filter' onclick=\"filterTable('report-only')\"> Report-Only</label>");
        sb.AppendLine("    </div>");
        
        // Policies Table
        sb.AppendLine("    <h2>\ud83d\udcca Policy Details</h2>");
        
        if (data.Policies.Count == 0)
        {
            sb.AppendLine("    <div class='empty-message'>No Conditional Access policies found.</div>");
        }
        else
        {
            sb.AppendLine("    <table>");
            sb.AppendLine("        <thead>");
            sb.AppendLine("            <tr>");
            sb.AppendLine("                <th>Policy Name</th>");
            sb.AppendLine("                <th>State</th>");
            sb.AppendLine("                <th>Break Glass Status</th>");
            sb.AppendLine("            </tr>");
            sb.AppendLine("        </thead>");
            sb.AppendLine("        <tbody>");
            
            foreach (var policy in data.Policies)
            {
                var filterState = policy.State switch
                {
                    "enabled" => "on",
                    "disabled" => "off",
                    "enabledforreportingbutnotenforced" => "report-only",
                    _ => policy.State
                };
                
                var stateBadge = policy.DisplayState switch
                {
                    "Enabled" => "<span class='badge badge-enabled'>Enabled</span>",
                    "Disabled" => "<span class='badge badge-disabled'>Disabled</span>",
                    "Report-Only" => "<span class='badge badge-report'>Report-Only</span>",
                    _ => $"<span class='badge'>{System.Net.WebUtility.HtmlEncode(policy.DisplayState)}</span>"
                };
                
                var rowClass = "";
                string breakGlassHtml;
                if (data.ConfiguredBreakGlassAccounts.Count == 0)
                {
                    breakGlassHtml = "<span class='badge badge-warning'>No accounts configured</span>";
                }
                else if (policy.MissingBreakGlassAccounts.Count == 0)
                {
                    breakGlassHtml = $"<span class='badge badge-success'>\u2705 All excluded</span>";
                    if (policy.ExcludedBreakGlassAccounts.Count > 0)
                    {
                        breakGlassHtml += $"<br><small>Excluded: {string.Join(", ", policy.ExcludedBreakGlassAccounts.Select(u => System.Net.WebUtility.HtmlEncode(u)))}</small>";
                    }
                }
                else if (policy.ExcludedBreakGlassAccounts.Count == 0)
                {
                    rowClass = "row-danger";
                    breakGlassHtml = $"<span class='badge badge-danger'>\u274c None excluded</span>";
                    breakGlassHtml += $"<br><small>Missing: {string.Join(", ", policy.MissingBreakGlassAccounts.Select(u => System.Net.WebUtility.HtmlEncode(u)))}</small>";
                }
                else
                {
                    rowClass = "row-warning";
                    breakGlassHtml = $"<span class='badge badge-warning'>\u26a0\ufe0f Partial</span>";
                    breakGlassHtml += $"<br><small>Excluded: {string.Join(", ", policy.ExcludedBreakGlassAccounts.Select(u => System.Net.WebUtility.HtmlEncode(u)))}</small>";
                    breakGlassHtml += $"<br><small style='color:#dc3545'>Missing: {string.Join(", ", policy.MissingBreakGlassAccounts.Select(u => System.Net.WebUtility.HtmlEncode(u)))}</small>";
                }
                
                sb.AppendLine($"            <tr class='{rowClass}' data-filter='{filterState}'>");
                sb.AppendLine($"                <td><strong>{System.Net.WebUtility.HtmlEncode(policy.DisplayName)}</strong><br><small style='color:#666'>{System.Net.WebUtility.HtmlEncode(policy.Id)}</small></td>");
                sb.AppendLine($"                <td>{stateBadge}</td>");
                sb.AppendLine($"                <td>{breakGlassHtml}</td>");
                sb.AppendLine("            </tr>");
            }
            
            sb.AppendLine("        </tbody>");
            sb.AppendLine("    </table>");
        }
        
        sb.AppendLine(GenerateBrandedFooter(branding,
            "Ensure all Conditional Access policies exclude your break glass accounts to prevent lockouts during emergencies."));

        sb.AppendLine("</div>");
        sb.AppendLine("</body>");
        sb.AppendLine("</html>");
        
        return sb.ToString();
    }

    private static string ConvertToGenericHtml(ReportResultDto report, ReportSettings branding)
    {
        var sb = new StringBuilder();
        
        sb.AppendLine("<!DOCTYPE html>");
        sb.AppendLine("<html lang='en'>");
        sb.AppendLine("<head>");
        sb.AppendLine("    <meta charset='UTF-8'>");
        sb.AppendLine("    <meta name='viewport' content='width=device-width, initial-scale=1.0'>");
        sb.AppendLine($"    <title>{System.Net.WebUtility.HtmlEncode(report.DisplayName)}</title>");
        sb.AppendLine("    <style>");
        sb.AppendLine("        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 0; padding: 20px; background: #f5f5f5; color: #333; }");
        sb.AppendLine("        .container { max-width: 1400px; margin: 0 auto; background: white; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); padding: 30px; }");
        sb.AppendLine("        h1 { color: #0078d4; margin-bottom: 10px; font-size: 28px; }");
        sb.AppendLine("        .subtitle { color: #666; margin-bottom: 30px; font-size: 14px; }");
        sb.AppendLine("        .summary { display: grid; grid-template-columns: repeat(auto-fit, minmax(150px, 1fr)); gap: 15px; margin-bottom: 30px; }");
        sb.AppendLine("        .summary-card { background: #f8f9fa; border-radius: 8px; padding: 15px; text-align: center; border-left: 4px solid #0078d4; }");
        sb.AppendLine("        .summary-value { font-size: 28px; font-weight: bold; color: #0078d4; }");
        sb.AppendLine("        .summary-label { color: #666; font-size: 12px; margin-top: 5px; }");
        sb.AppendLine("        table { width: 100%; border-collapse: collapse; margin-bottom: 30px; font-size: 14px; }");
        sb.AppendLine("        th { background: #0078d4; color: white; padding: 12px 15px; text-align: left; font-weight: 600; position: sticky; top: 0; }");
        sb.AppendLine("        td { padding: 10px 15px; border-bottom: 1px solid #eee; }");
        sb.AppendLine("        tr:hover { background: #f8f9fa; }");
        sb.AppendLine("        tr:nth-child(even) { background: #fafafa; }");
        sb.AppendLine("        tr:nth-child(even):hover { background: #f0f0f0; }");
        sb.AppendLine("        .badge { padding: 3px 8px; border-radius: 12px; font-size: 11px; font-weight: 600; display: inline-block; }");
        sb.AppendLine("        .badge-success { background: #d4edda; color: #155724; }");
        sb.AppendLine("        .badge-warning { background: #fff3cd; color: #856404; }");
        sb.AppendLine("        .badge-danger { background: #f8d7da; color: #721c24; }");
        sb.AppendLine("        .badge-info { background: #cce5ff; color: #004085; }");
        sb.AppendLine("        .empty-message { color: #666; padding: 40px; text-align: center; background: #f8f9fa; border-radius: 8px; }");
        sb.AppendLine("        .footer { margin-top: 40px; padding-top: 20px; border-top: 1px solid #eee; color: #666; font-size: 12px; text-align: center; }");
        sb.AppendLine("        .table-wrapper { overflow-x: auto; }");
        sb.AppendLine("        @media print { body { background: white; } .container { box-shadow: none; } }");
        sb.AppendLine("    </style>");
        sb.AppendLine("</head>");
        sb.AppendLine("<body>");
        sb.AppendLine("<div class='container'>");

        sb.AppendLine(GenerateBrandedHeader(branding, report.DisplayName,
            report.GeneratedAt.ToString("dddd, MMMM d, yyyy 'at' h:mm tt") + $" UTC \u2022 {report.DateRange}"));
        
        // Summary Cards (if available)
        if (report.Summary?.Highlights != null && report.Summary.Highlights.Count > 0)
        {
            sb.AppendLine("    <div class='summary'>");
            sb.AppendLine($"        <div class='summary-card'><div class='summary-value'>{report.Summary.TotalRecords:N0}</div><div class='summary-label'>Total Records</div></div>");
            foreach (var detail in report.Summary.Highlights.Take(5)) // Limit to 5 summary cards
            {
                var label = FormatPropertyName(detail.Key);
                var value = detail.Value;
                sb.AppendLine($"        <div class='summary-card'><div class='summary-value'>{FormatValue(value)}</div><div class='summary-label'>{System.Net.WebUtility.HtmlEncode(label)}</div></div>");
            }
            sb.AppendLine("    </div>");
        }
        
        // Data Table
        sb.AppendLine("    <div class='table-wrapper'>");
        
        if (report.Data is System.Collections.IEnumerable enumerable && !(report.Data is string))
        {
            var items = enumerable.Cast<object>().ToList();
            if (items.Count > 0)
            {
                var firstItem = items[0];
                var properties = firstItem.GetType().GetProperties()
                    .Where(p => !p.Name.Equals("Id", StringComparison.OrdinalIgnoreCase) && 
                                !p.Name.EndsWith("Id", StringComparison.OrdinalIgnoreCase) &&
                                p.PropertyType != typeof(byte[]))
                    .Take(12) // Limit columns for readability
                    .ToList();
                
                sb.AppendLine("    <table>");
                sb.AppendLine("        <thead>");
                sb.AppendLine("            <tr>");
                foreach (var prop in properties)
                {
                    sb.AppendLine($"                <th>{System.Net.WebUtility.HtmlEncode(FormatPropertyName(prop.Name))}</th>");
                }
                sb.AppendLine("            </tr>");
                sb.AppendLine("        </thead>");
                sb.AppendLine("        <tbody>");
                
                foreach (var item in items.Take(1000)) // Limit to 1000 rows for performance
                {
                    sb.AppendLine("            <tr>");
                    foreach (var prop in properties)
                    {
                        var value = prop.GetValue(item);
                        var formattedValue = FormatCellValue(value, prop.Name);
                        sb.AppendLine($"                <td>{formattedValue}</td>");
                    }
                    sb.AppendLine("            </tr>");
                }
                sb.AppendLine("        </tbody>");
                sb.AppendLine("    </table>");
                
                if (items.Count > 1000)
                {
                    sb.AppendLine($"    <p class='empty-message'>Showing first 1,000 of {items.Count:N0} records. Export to CSV for complete data.</p>");
                }
            }
            else
            {
                sb.AppendLine("    <p class='empty-message'>No data available for this report.</p>");
            }
        }
        else
        {
            sb.AppendLine("    <p class='empty-message'>Report data format not supported for table display.</p>");
        }
        
        sb.AppendLine("    </div>");
        
        sb.AppendLine(GenerateBrandedFooter(branding));

        sb.AppendLine("</div>");
        sb.AppendLine("</body>");
        sb.AppendLine("</html>");
        
        return sb.ToString();
    }

    private static string GenerateGroupMembershipHtmlReport(List<TenantGroupDto> groups, ReportSettings branding)
    {
        var sb = new StringBuilder();
        
        var totalGroups = groups.Count;
        var m365Groups = groups.Count(g => g.GroupType == "Microsoft 365");
        var securityGroups = groups.Count(g => g.GroupType == "Security");
        var distributionGroups = groups.Count(g => g.GroupType == "Distribution");
        var teamsEnabled = groups.Count(g => g.IsTeam);
        var publicGroups = groups.Count(g => g.Visibility?.Equals("Public", StringComparison.OrdinalIgnoreCase) == true);
        var privateGroups = groups.Count(g => g.Visibility?.Equals("Private", StringComparison.OrdinalIgnoreCase) == true);
        var noOwnerGroups = groups.Count(g => g.OwnerCount == 0);
        
        sb.AppendLine("<!DOCTYPE html>");
        sb.AppendLine("<html lang='en'>");
        sb.AppendLine("<head>");
        sb.AppendLine("    <meta charset='UTF-8'>");
        sb.AppendLine("    <meta name='viewport' content='width=device-width, initial-scale=1.0'>");
        sb.AppendLine("    <title>Group Membership Report</title>");
        sb.AppendLine("    <style>");
        sb.AppendLine("        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 0; padding: 20px; background: #f5f5f5; color: #333; }");
        sb.AppendLine("        .container { max-width: 1400px; margin: 0 auto; background: white; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); padding: 30px; }");
        sb.AppendLine("        h1 { color: #0078d4; margin-bottom: 10px; font-size: 28px; }");
        sb.AppendLine("        .subtitle { color: #666; margin-bottom: 20px; font-size: 14px; }");
        sb.AppendLine("        .summary { display: grid; grid-template-columns: repeat(auto-fit, minmax(140px, 1fr)); gap: 15px; margin-bottom: 30px; }");
        sb.AppendLine("        .summary-card { background: #f8f9fa; border-radius: 8px; padding: 20px; text-align: center; border-left: 4px solid #0078d4; cursor: pointer; transition: transform 0.2s, box-shadow 0.2s; }");
        sb.AppendLine("        .summary-card:hover { transform: translateY(-2px); box-shadow: 0 4px 8px rgba(0,0,0,0.1); }");
        sb.AppendLine("        .summary-card.active { box-shadow: 0 0 0 3px #0078d4; }");
        sb.AppendLine("        .summary-card.m365 { border-left-color: #0078d4; }");
        sb.AppendLine("        .summary-card.security { border-left-color: #28a745; }");
        sb.AppendLine("        .summary-card.distribution { border-left-color: #6f42c1; }");
        sb.AppendLine("        .summary-card.teams { border-left-color: #6264a7; }");
        sb.AppendLine("        .summary-card.public { border-left-color: #ffc107; }");
        sb.AppendLine("        .summary-card.private { border-left-color: #17a2b8; }");
        sb.AppendLine("        .summary-card.danger { border-left-color: #dc3545; }");
        sb.AppendLine("        .summary-value { font-size: 32px; font-weight: bold; color: #0078d4; }");
        sb.AppendLine("        .summary-card.m365 .summary-value { color: #0078d4; }");
        sb.AppendLine("        .summary-card.security .summary-value { color: #28a745; }");
        sb.AppendLine("        .summary-card.distribution .summary-value { color: #6f42c1; }");
        sb.AppendLine("        .summary-card.teams .summary-value { color: #6264a7; }");
        sb.AppendLine("        .summary-card.public .summary-value { color: #ffc107; }");
        sb.AppendLine("        .summary-card.private .summary-value { color: #17a2b8; }");
        sb.AppendLine("        .summary-card.danger .summary-value { color: #dc3545; }");
        sb.AppendLine("        .summary-label { color: #666; font-size: 12px; margin-top: 5px; }");
        sb.AppendLine("        h2 { color: #333; font-size: 20px; margin-top: 30px; margin-bottom: 15px; padding-bottom: 10px; border-bottom: 2px solid #eee; }");
        sb.AppendLine("        table { width: 100%; border-collapse: collapse; margin-bottom: 30px; font-size: 14px; }");
        sb.AppendLine("        th { background: #0078d4; color: white; padding: 12px 15px; text-align: left; font-weight: 600; position: sticky; top: 0; }");
        sb.AppendLine("        td { padding: 10px 15px; border-bottom: 1px solid #eee; }");
        sb.AppendLine("        tr:hover { background: #f8f9fa; }");
        sb.AppendLine("        tr.no-owner { background: #f8d7da; }");
        sb.AppendLine("        .badge { padding: 4px 10px; border-radius: 12px; font-size: 11px; font-weight: 600; display: inline-block; }");
        sb.AppendLine("        .badge-m365 { background: #cce5ff; color: #004085; }");
        sb.AppendLine("        .badge-security { background: #d4edda; color: #155724; }");
        sb.AppendLine("        .badge-distribution { background: #e2d5f1; color: #4a2c7a; }");
        sb.AppendLine("        .badge-teams { background: #e8e8fc; color: #6264a7; }");
        sb.AppendLine("        .badge-public { background: #fff3cd; color: #856404; }");
        sb.AppendLine("        .badge-private { background: #d1ecf1; color: #0c5460; }");
        sb.AppendLine("        .badge-danger { background: #f8d7da; color: #721c24; }");
        sb.AppendLine("        .badge-success { background: #d4edda; color: #155724; }");
        sb.AppendLine("        .empty-message { color: #28a745; padding: 20px; background: #d4edda; border-radius: 8px; text-align: center; }");
        sb.AppendLine("        .footer { margin-top: 40px; padding-top: 20px; border-top: 1px solid #eee; color: #666; font-size: 12px; text-align: center; }");
        sb.AppendLine("        .table-wrapper { overflow-x: auto; }");
        sb.AppendLine("        @media print { body { background: white; } .container { box-shadow: none; } }");
        sb.AppendLine("    </style>");
        sb.AppendLine("    <script>");
        sb.AppendLine("        function filterByType(el, type) {");
        sb.AppendLine("            document.querySelectorAll('.summary-card').forEach(card => card.classList.remove('active'));");
        sb.AppendLine("            if (type !== 'all') {");
        sb.AppendLine("                el.closest('.summary-card').classList.add('active');");
        sb.AppendLine("            }");
        sb.AppendLine("            document.querySelectorAll('tbody tr').forEach(row => {");
        sb.AppendLine("                if (type === 'all') {");
        sb.AppendLine("                    row.style.display = '';");
        sb.AppendLine("                } else if (type === 'm365') {");
        sb.AppendLine("                    row.style.display = row.dataset.grouptype === 'Microsoft 365' ? '' : 'none';");
        sb.AppendLine("                } else if (type === 'security') {");
        sb.AppendLine("                    row.style.display = row.dataset.grouptype === 'Security' ? '' : 'none';");
        sb.AppendLine("                } else if (type === 'distribution') {");
        sb.AppendLine("                    row.style.display = row.dataset.grouptype === 'Distribution' ? '' : 'none';");
        sb.AppendLine("                } else if (type === 'teams') {");
        sb.AppendLine("                    row.style.display = row.dataset.isteam === 'true' ? '' : 'none';");
        sb.AppendLine("                } else if (type === 'public') {");
        sb.AppendLine("                    row.style.display = row.dataset.visibility === 'Public' ? '' : 'none';");
        sb.AppendLine("                } else if (type === 'private') {");
        sb.AppendLine("                    row.style.display = row.dataset.visibility === 'Private' ? '' : 'none';");
        sb.AppendLine("                } else if (type === 'noowner') {");
        sb.AppendLine("                    row.style.display = row.dataset.ownercount === '0' ? '' : 'none';");
        sb.AppendLine("                }");
        sb.AppendLine("            });");
        sb.AppendLine("            updateVisibleCount();");
        sb.AppendLine("        }");
        sb.AppendLine("        function updateVisibleCount() {");
        sb.AppendLine("            var visible = document.querySelectorAll('tbody tr:not([style*=\"display: none\"])').length;");
        sb.AppendLine("            var counter = document.getElementById('visible-count');");
        sb.AppendLine("            if (counter) counter.textContent = 'Showing ' + visible + ' group(s)';");
        sb.AppendLine("        }");
        sb.AppendLine("    </script>");
        sb.AppendLine("</head>");
        sb.AppendLine("<body>");
        sb.AppendLine("<div class='container'>");

        sb.AppendLine(GenerateBrandedHeader(branding, "Group Membership Report",
            DateTime.UtcNow.ToString("dddd, MMMM d, yyyy 'at' h:mm tt") + " UTC"));
        
        // Summary Cards
        sb.AppendLine("    <p style='color: #666; font-size: 13px; margin-bottom: 10px;'>\ud83d\udca1 Click on the cards below to filter the table</p>");
        sb.AppendLine("    <div class='summary'>");
        sb.AppendLine($"        <div class='summary-card' onclick=\"filterByType(this,'all')\"><div class='summary-value'>{totalGroups}</div><div class='summary-label'>Total Groups</div></div>");
        sb.AppendLine($"        <div class='summary-card m365' onclick=\"filterByType(this,'m365')\"><div class='summary-value'>{m365Groups}</div><div class='summary-label'>Microsoft 365</div></div>");
        sb.AppendLine($"        <div class='summary-card security' onclick=\"filterByType(this,'security')\"><div class='summary-value'>{securityGroups}</div><div class='summary-label'>Security</div></div>");
        
        if (distributionGroups > 0)
        {
            sb.AppendLine($"        <div class='summary-card distribution' onclick=\"filterByType(this,'distribution')\"><div class='summary-value'>{distributionGroups}</div><div class='summary-label'>Distribution</div></div>");
        }
        
        sb.AppendLine($"        <div class='summary-card teams' onclick=\"filterByType(this,'teams')\"><div class='summary-value'>{teamsEnabled}</div><div class='summary-label'>Teams Enabled</div></div>");
        sb.AppendLine($"        <div class='summary-card public' onclick=\"filterByType(this,'public')\"><div class='summary-value'>{publicGroups}</div><div class='summary-label'>Public</div></div>");
        sb.AppendLine($"        <div class='summary-card private' onclick=\"filterByType(this,'private')\"><div class='summary-value'>{privateGroups}</div><div class='summary-label'>Private</div></div>");
        
        if (noOwnerGroups > 0)
        {
            sb.AppendLine($"        <div class='summary-card danger' onclick=\"filterByType(this,'noowner')\"><div class='summary-value'>{noOwnerGroups}</div><div class='summary-label'>No Owner</div></div>");
        }
        
        sb.AppendLine("    </div>");
        
        // Groups Table
        sb.AppendLine("    <h2>\ud83d\udcca Groups</h2>");
        sb.AppendLine($"    <p id='visible-count' style='color: #666; margin-bottom: 15px;'>Showing {groups.Count} group(s)</p>");
        
        if (groups.Count == 0)
        {
            sb.AppendLine("    <div class='empty-message'>\u2713 No groups found</div>");
        }
        else
        {
            sb.AppendLine("    <div class='table-wrapper'>");
            sb.AppendLine("    <table>");
            sb.AppendLine("        <thead>");
            sb.AppendLine("            <tr>");
            sb.AppendLine("                <th>Display Name</th>");
            sb.AppendLine("                <th>Type</th>");
            sb.AppendLine("                <th>Visibility</th>");
            sb.AppendLine("                <th>Teams</th>");
            sb.AppendLine("                <th>Members</th>");
            sb.AppendLine("                <th>Owners</th>");
            sb.AppendLine("                <th>Created</th>");
            sb.AppendLine("            </tr>");
            sb.AppendLine("        </thead>");
            sb.AppendLine("        <tbody>");
            
            foreach (var group in groups.OrderBy(g => g.DisplayName))
            {
                var rowClass = group.OwnerCount == 0 ? "no-owner" : "";
                var visibility = group.Visibility ?? "Unknown";
                
                var typeBadge = group.GroupType switch
                {
                    "Microsoft 365" => "<span class='badge badge-m365'>Microsoft 365</span>",
                    "Security" => "<span class='badge badge-security'>Security</span>",
                    "Distribution" => "<span class='badge badge-distribution'>Distribution</span>",
                    _ => $"<span class='badge'>{System.Net.WebUtility.HtmlEncode(group.GroupType)}</span>"
                };
                
                var visibilityBadge = visibility switch
                {
                    "Public" => "<span class='badge badge-public'>Public</span>",
                    "Private" => "<span class='badge badge-private'>Private</span>",
                    _ => $"<span class='badge'>{System.Net.WebUtility.HtmlEncode(visibility)}</span>"
                };
                
                var teamsBadge = group.IsTeam
                    ? "<span class='badge badge-teams'>\u2713 Teams</span>"
                    : "<span style='color:#999'>\u2014</span>";
                
                var ownerDisplay = group.OwnerCount == 0
                    ? "<span class='badge badge-danger'>0 - No Owner</span>"
                    : group.OwnerCount.ToString();
                
                var createdDate = group.CreatedDateTime?.ToString("MMM d, yyyy") ?? "<span style='color:#999'>\u2014</span>";
                
                sb.AppendLine($"            <tr class='{rowClass}' data-grouptype='{System.Net.WebUtility.HtmlEncode(group.GroupType)}' data-visibility='{System.Net.WebUtility.HtmlEncode(visibility)}' data-isteam='{group.IsTeam.ToString().ToLower()}' data-ownercount='{group.OwnerCount}'>");
                sb.AppendLine($"                <td><strong>{System.Net.WebUtility.HtmlEncode(group.DisplayName)}</strong></td>");
                sb.AppendLine($"                <td>{typeBadge}</td>");
                sb.AppendLine($"                <td>{visibilityBadge}</td>");
                sb.AppendLine($"                <td>{teamsBadge}</td>");
                sb.AppendLine($"                <td>{group.MemberCount}</td>");
                sb.AppendLine($"                <td>{ownerDisplay}</td>");
                sb.AppendLine($"                <td>{createdDate}</td>");
                sb.AppendLine("            </tr>");
            }
            
            sb.AppendLine("        </tbody>");
            sb.AppendLine("    </table>");
            sb.AppendLine("    </div>");
        }
        
        sb.AppendLine(GenerateBrandedFooter(branding,
            "Groups without owners should be reviewed and assigned appropriate ownership."));

        sb.AppendLine("</div>");
        sb.AppendLine("</body>");
        sb.AppendLine("</html>");
        
        return sb.ToString();
    }

    private static string FormatPropertyName(string name)
    {
        // Convert camelCase or PascalCase to Title Case with spaces
        var result = new StringBuilder();
        foreach (char c in name)
        {
            if (char.IsUpper(c) && result.Length > 0)
            {
                result.Append(' ');
            }
            result.Append(c);
        }
        return result.ToString();
    }

    /// <summary>
    /// Formats Microsoft Secure Score control names into human-readable titles
    /// </summary>
    private static string FormatSecureScoreControlName(string controlName)
    {
        if (string.IsNullOrEmpty(controlName)) return "Unknown Control";
        
        // Known control name mappings for common Microsoft Secure Score controls
        var knownMappings = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            // Microsoft Defender for Office 365 (MDO) controls
            { "mdo_atpprotection", "Enable ATP Protection" },
            { "mdo_safedocuments", "Enable Safe Documents" },
            { "mdo_zapspam", "Enable Zero-hour Auto Purge for Spam" },
            { "mdo_zapmalware", "Enable Zero-hour Auto Purge for Malware" },
            { "mdo_zapphish", "Enable Zero-hour Auto Purge for Phishing" },
            { "mdo_safelinks", "Enable Safe Links" },
            { "mdo_safeattachments", "Enable Safe Attachments" },
            { "mdo_antiphishpolicy", "Configure Anti-Phishing Policy" },
            { "mdo_impersonationprotection", "Enable Impersonation Protection" },
            
            // Exchange Online controls
            { "exo_dkim", "Enable DKIM Signing" },
            { "exo_auditing", "Enable Mailbox Auditing" },
            { "exo_mailtips", "Configure MailTips" },
            { "exo_externalforwarding", "Block External Email Forwarding" },
            
            // Azure AD / Entra ID controls
            { "aad_mfaregistration", "Require MFA Registration" },
            { "aad_selfservicepasswordreset", "Enable Self-Service Password Reset" },
            { "aad_passwordprotection", "Enable Password Protection" },
            { "aad_blocklegacyauth", "Block Legacy Authentication" },
            { "aad_conditionalaccess", "Configure Conditional Access" },
            { "aad_privilegedidentity", "Enable Privileged Identity Management" },
            { "aad_riskyusers", "Monitor Risky Users" },
            { "aad_riskysignins", "Monitor Risky Sign-ins" },
            
            // Microsoft Teams controls
            { "meeting_restrictanonymousjoin_v1", "Restrict Anonymous Meeting Join" },
            { "teams_externaldomain", "Restrict External Domain Access" },
            { "teams_guestaccess", "Configure Guest Access" },
            
            // SharePoint controls
            { "sharepoint_externalsharing", "Configure External Sharing" },
            { "sharepoint_anonymouslinks", "Restrict Anonymous Links" },
            
            // Intune / Device controls
            { "mdm_windowsdeviceencryption", "Require Windows Device Encryption" },
            { "mdm_devicecompliance", "Enable Device Compliance Policies" },
            { "mdm_appprotection", "Configure App Protection Policies" },
            
            // Microsoft Cloud App Security (MCAS) / Defender for Cloud Apps
            { "mcasfirefalllogupload", "Enable Cloud App Security Log Upload" },
            { "McasFirewallLogUpload", "Enable Cloud App Security Log Upload" },
            
            // Security Center / Defender controls
            { "securitydefaults", "Enable Security Defaults" },
            { "defenderforendpoint", "Enable Defender for Endpoint" },
        };
        
        // Check for exact match first
        if (knownMappings.TryGetValue(controlName, out var friendlyName))
        {
            return friendlyName;
        }
        
        // Handle scid_ prefixed controls (Secure Score Control IDs)
        if (controlName.StartsWith("scid_", StringComparison.OrdinalIgnoreCase))
        {
            // Remove prefix and format the rest
            var remainder = controlName.Substring(5);
            return $"Security Control {remainder}";
        }
        
        // Apply transformations for unknown controls
        var result = controlName;
        
        // Remove common prefixes
        var prefixes = new[] { "mdo_", "exo_", "aad_", "mdm_", "mcas_", "teams_", "sharepoint_", "meeting_" };
        foreach (var prefix in prefixes)
        {
            if (result.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
            {
                result = result.Substring(prefix.Length);
                break;
            }
        }
        
        // Remove version suffixes like _v1, _v2
        result = System.Text.RegularExpressions.Regex.Replace(result, @"_v\d+$", "", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        
        // Replace underscores with spaces
        result = result.Replace("_", " ");
        
        // Insert spaces before capital letters (for camelCase/PascalCase)
        result = System.Text.RegularExpressions.Regex.Replace(result, @"(\p{Ll})(\p{Lu})", "$1 $2");
        
        // Title case each word
        var words = result.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        for (int i = 0; i < words.Length; i++)
        {
            if (words[i].Length > 0)
            {
                // Keep acronyms uppercase (2-4 letter all-caps words)
                if (words[i].Length <= 4 && words[i].All(char.IsUpper))
                {
                    continue;
                }
                words[i] = char.ToUpper(words[i][0]) + words[i].Substring(1).ToLower();
            }
        }
        
        return string.Join(" ", words);
    }

    private static string FormatValue(object? value)
    {
        if (value == null) return "0";
        if (value is double d) return d.ToString("N2");
        if (value is decimal dec) return dec.ToString("N2");
        if (value is int i) return i.ToString("N0");
        if (value is long l) return l.ToString("N0");
        return value.ToString() ?? "0";
    }

    private static string FormatCellValue(object? value, string propertyName)
    {
        if (value == null) return "<span style='color:#999'>—</span>";
        
        // Handle List<string> - join with commas
        if (value is IEnumerable<string> stringList)
        {
            var items = stringList.ToList();
            if (items.Count == 0) return "<span style='color:#999'>—</span>";
            return System.Net.WebUtility.HtmlEncode(string.Join(", ", items));
        }
        
        // Handle other IEnumerable types (but not string)
        if (value is System.Collections.IEnumerable enumerable && !(value is string))
        {
            var items = enumerable.Cast<object>().ToList();
            if (items.Count == 0) return "<span style='color:#999'>—</span>";
            return System.Net.WebUtility.HtmlEncode(string.Join(", ", items.Select(i => i?.ToString() ?? "")));
        }
        
        var strValue = value.ToString() ?? "";
        
        // Handle dates
        if (value is DateTime dt)
        {
            return dt == DateTime.MinValue ? "<span style='color:#999'>Never</span>" : dt.ToString("MMM d, yyyy HH:mm");
        }
        if (value is DateTimeOffset dto)
        {
            return dto.ToString("MMM d, yyyy HH:mm");
        }
        
        // Handle booleans with badges
        if (value is bool b)
        {
            return b 
                ? "<span class='badge badge-success'>Yes</span>" 
                : "<span class='badge badge-danger'>No</span>";
        }
        
        // Handle specific property patterns
        var lowerName = propertyName.ToLower();
        
        // Status-like fields
        if (lowerName.Contains("status") || lowerName.Contains("state") || lowerName.Contains("enabled"))
        {
            var lowerValue = strValue.ToLower();
            if (lowerValue == "enabled" || lowerValue == "active" || lowerValue == "compliant" || lowerValue == "true" || lowerValue == "yes")
                return $"<span class='badge badge-success'>{System.Net.WebUtility.HtmlEncode(strValue)}</span>";
            if (lowerValue == "disabled" || lowerValue == "inactive" || lowerValue == "noncompliant" || lowerValue == "false" || lowerValue == "no")
                return $"<span class='badge badge-danger'>{System.Net.WebUtility.HtmlEncode(strValue)}</span>";
            if (lowerValue.Contains("warning") || lowerValue.Contains("pending"))
                return $"<span class='badge badge-warning'>{System.Net.WebUtility.HtmlEncode(strValue)}</span>";
        }
        
        // Risk levels
        if (lowerName.Contains("risk"))
        {
            var lowerValue = strValue.ToLower();
            if (lowerValue == "high") return "<span class='badge badge-danger'>High</span>";
            if (lowerValue == "medium") return "<span class='badge badge-warning'>Medium</span>";
            if (lowerValue == "low") return "<span class='badge badge-info'>Low</span>";
        }
        
        // Percentages
        if (value is double dbl && (lowerName.Contains("percent") || lowerName.Contains("rate")))
        {
            return $"{dbl:N1}%";
        }
        
        // Numbers
        if (value is int || value is long || value is double || value is decimal)
        {
            return FormatValue(value);
        }
        
        // Default: escape HTML
        return System.Net.WebUtility.HtmlEncode(strValue);
    }

    public async Task<List<ScheduledReportDto>> GetScheduledReportsAsync(string userId)
    {
        var schedules = await _dbContext.ScheduledReports
            .Where(s => s.UserId == userId)
            .OrderBy(s => s.DisplayName)
            .ToListAsync();

        return schedules.Select(MapToDto).ToList();
    }

    public async Task<ScheduledReportDto> CreateScheduledReportAsync(string userId, string? userEmail, CreateScheduledReportRequest request)
    {
        var definition = _reportDefinitions.FirstOrDefault(d => d.ReportType == request.ReportType)
            ?? throw new ArgumentException($"Unknown report type: {request.ReportType}");

        var recipients = request.Recipients?.ToList() ?? new List<string>();
        if (!string.IsNullOrEmpty(userEmail) && !recipients.Contains(userEmail, StringComparer.OrdinalIgnoreCase))
        {
            recipients.Insert(0, userEmail);
        }

        var schedule = new ScheduledReport
        {
            UserId = userId,
            UserEmail = userEmail,
            ReportType = request.ReportType,
            DisplayName = definition.DisplayName,
            Frequency = request.Frequency,
            TimeOfDay = request.Time ?? "08:00",
            DayOfWeek = request.Frequency.Equals("weekly", StringComparison.OrdinalIgnoreCase) ? request.DayOfWeek : null,
            DayOfMonth = request.Frequency.Equals("monthly", StringComparison.OrdinalIgnoreCase) ? request.DayOfMonth : null,
            Recipients = string.Join(",", recipients),
            DateRange = request.DateRange,
            IsEnabled = true,
            NextRunAt = CalculateNextRun(request.Frequency, request.Time, request.DayOfWeek, request.DayOfMonth),
            CreatedAt = DateTime.UtcNow,
            UpdatedAt = DateTime.UtcNow
        };

        _dbContext.ScheduledReports.Add(schedule);
        await _dbContext.SaveChangesAsync();

        return MapToDto(schedule);
    }

    public async Task<ScheduledReportDto?> UpdateScheduledReportAsync(string userId, string scheduleId, UpdateScheduledReportRequest request)
    {
        if (!int.TryParse(scheduleId, out var id))
        {
            return null;
        }

        var schedule = await _dbContext.ScheduledReports
            .FirstOrDefaultAsync(s => s.Id == id && s.UserId == userId);

        if (schedule == null)
        {
            return null;
        }

        if (request.Frequency != null)
        {
            schedule.Frequency = request.Frequency;
        }
        if (request.Time != null)
        {
            schedule.TimeOfDay = request.Time;
        }
        if (request.DayOfWeek.HasValue)
        {
            schedule.DayOfWeek = request.DayOfWeek;
        }
        if (request.DayOfMonth.HasValue)
        {
            schedule.DayOfMonth = request.DayOfMonth;
        }
        if (request.Recipients != null)
        {
            schedule.Recipients = string.Join(",", request.Recipients);
        }
        if (request.DateRange != null)
        {
            schedule.DateRange = request.DateRange;
        }
        if (request.IsEnabled.HasValue)
        {
            schedule.IsEnabled = request.IsEnabled.Value;
        }

        schedule.NextRunAt = CalculateNextRun(schedule.Frequency, schedule.TimeOfDay, schedule.DayOfWeek, schedule.DayOfMonth);
        schedule.UpdatedAt = DateTime.UtcNow;

        await _dbContext.SaveChangesAsync();

        return MapToDto(schedule);
    }

    public async Task<bool> DeleteScheduledReportAsync(string userId, string scheduleId)
    {
        if (!int.TryParse(scheduleId, out var id))
        {
            return false;
        }

        var schedule = await _dbContext.ScheduledReports
            .FirstOrDefaultAsync(s => s.Id == id && s.UserId == userId);

        if (schedule == null)
        {
            return false;
        }

        _dbContext.ScheduledReports.Remove(schedule);
        await _dbContext.SaveChangesAsync();

        return true;
    }

    public async Task<List<ReportHistoryDto>> GetReportHistoryAsync(string userId, int take)
    {
        var history = await _dbContext.ReportHistories
            .Where(h => h.UserId == userId)
            .OrderByDescending(h => h.GeneratedAt)
            .Take(take)
            .ToListAsync();

        return history.Select(h => new ReportHistoryDto(
            h.Id.ToString(),
            h.ReportType,
            h.DisplayName,
            h.GeneratedAt,
            h.Status,
            h.ErrorMessage,
            h.RecordCount,
            h.WasScheduled
        )).ToList();
    }

    public async Task<List<ScheduledReport>> GetDueScheduledReportsAsync()
    {
        var now = DateTime.UtcNow;
        return await _dbContext.ScheduledReports
            .Where(s => s.IsEnabled && s.NextRunAt <= now)
            .ToListAsync();
    }

    public async Task UpdateScheduledReportAfterRunAsync(int scheduleId, bool success, string? error = null)
    {
        var schedule = await _dbContext.ScheduledReports.FindAsync(scheduleId);
        if (schedule == null) return;

        schedule.LastRunAt = DateTime.UtcNow;
        schedule.LastRunStatus = success ? "success" : "failed";
        schedule.LastRunError = error;
        schedule.NextRunAt = CalculateNextRun(schedule.Frequency, schedule.TimeOfDay, schedule.DayOfWeek, schedule.DayOfMonth);
        schedule.UpdatedAt = DateTime.UtcNow;

        await _dbContext.SaveChangesAsync();
    }

    private static ScheduledReportDto MapToDto(ScheduledReport schedule)
    {
        return new ScheduledReportDto(
            Id: schedule.Id.ToString(),
            ReportType: schedule.ReportType,
            DisplayName: schedule.DisplayName,
            Schedule: BuildScheduleDescription(schedule.Frequency, schedule.TimeOfDay, schedule.DayOfWeek, schedule.DayOfMonth),
            Frequency: schedule.Frequency,
            Recipients: schedule.Recipients.Split(',', StringSplitOptions.RemoveEmptyEntries).ToList(),
            DateRange: schedule.DateRange,
            IsEnabled: schedule.IsEnabled,
            LastRunAt: schedule.LastRunAt,
            NextRunAt: schedule.NextRunAt,
            CreatedAt: schedule.CreatedAt
        );
    }

    private static int GetDaysFromDateRange(string? dateRange) => dateRange switch
    {
        "last7days" => 7,
        "last30days" => 30,
        "last90days" => 90,
        "thismonth" => DateTime.UtcNow.Day,
        "lastmonth" => DateTime.DaysInMonth(DateTime.UtcNow.Year, DateTime.UtcNow.Month - 1),
        _ => 30
    };

    private static string BuildScheduleDescription(string frequency, string? time, int? dayOfWeek, int? dayOfMonth)
    {
        var timeStr = time ?? "08:00";
        return frequency.ToLower() switch
        {
            "daily" => $"Daily at {timeStr} UTC",
            "weekly" => $"Weekly on {(DayOfWeek)(dayOfWeek ?? 1)} at {timeStr} UTC",
            "monthly" => $"Monthly on day {dayOfMonth ?? 1} at {timeStr} UTC",
            _ => "Unknown schedule"
        };
    }

    private static DateTime CalculateNextRun(string frequency, string? time, int? dayOfWeek, int? dayOfMonth)
    {
        var now = DateTime.UtcNow;
        var timeOfDay = TimeSpan.TryParse(time ?? "08:00", out var t) ? t : TimeSpan.FromHours(8);

        switch (frequency.ToLower())
        {
            case "daily":
                var nextDaily = now.Date.Add(timeOfDay);
                if (nextDaily <= now) nextDaily = nextDaily.AddDays(1);
                return nextDaily;

            case "weekly":
                var targetDay = (DayOfWeek)(dayOfWeek ?? 1);
                var daysUntil = ((int)targetDay - (int)now.DayOfWeek + 7) % 7;
                if (daysUntil == 0 && now.TimeOfDay >= timeOfDay) daysUntil = 7;
                return now.Date.AddDays(daysUntil).Add(timeOfDay);

            case "monthly":
                var day = Math.Min(dayOfMonth ?? 1, DateTime.DaysInMonth(now.Year, now.Month));
                var nextMonthly = new DateTime(now.Year, now.Month, day).Add(timeOfDay);
                if (nextMonthly <= now)
                {
                    nextMonthly = nextMonthly.AddMonths(1);
                    day = Math.Min(dayOfMonth ?? 1, DateTime.DaysInMonth(nextMonthly.Year, nextMonthly.Month));
                    nextMonthly = new DateTime(nextMonthly.Year, nextMonthly.Month, day).Add(timeOfDay);
                }
                return nextMonthly;

            default:
                return now.AddDays(1);
        }
    }

    private static string EscapeCsvField(string field)
    {
        if (string.IsNullOrEmpty(field)) return "";
        if (field.Contains(',') || field.Contains('"') || field.Contains('\n'))
        {
            return $"\"{field.Replace("\"", "\"\"")}\"";
        }
        return field;
    }
}
