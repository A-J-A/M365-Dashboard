namespace M365Dashboard.Api.Configuration;

public class CacheOptions
{
    public int DefaultTtlMinutes { get; set; } = 15;
    public int SignInDataTtlMinutes { get; set; } = 5;
    public int LicenseDataTtlMinutes { get; set; } = 60;
    public int ReportDataTtlMinutes { get; set; } = 30;
}

public class DashboardOptions
{
    public int DefaultRefreshIntervalSeconds { get; set; } = 300;
    public int MinRefreshIntervalSeconds { get; set; } = 60;
    public int MaxWidgetsPerDashboard { get; set; } = 20;
}

public static class WidgetTypes
{
    public const string ActiveUsers = "active-users";
    public const string SignInAnalytics = "sign-in-analytics";
    public const string LicenseUsage = "license-usage";
    public const string DeviceCompliance = "device-compliance";
    public const string MailActivity = "mail-activity";
    public const string TeamsActivity = "teams-activity";

    public static readonly Dictionary<string, WidgetDefinition> Definitions = new()
    {
        [ActiveUsers] = new WidgetDefinition
        {
            Type = ActiveUsers,
            Name = "Active Users",
            Description = "Daily, weekly, and monthly active user trends",
            Category = "Usage",
            RequiredPermissions = ["Reports.Read.All"],
            DefaultWidth = 2,
            DefaultHeight = 1
        },
        [SignInAnalytics] = new WidgetDefinition
        {
            Type = SignInAnalytics,
            Name = "Sign-in Analytics",
            Description = "Success/failure rates and risky sign-in detection",
            Category = "Security",
            RequiredPermissions = ["AuditLog.Read.All"],
            DefaultWidth = 2,
            DefaultHeight = 2
        },
        [LicenseUsage] = new WidgetDefinition
        {
            Type = LicenseUsage,
            Name = "License Usage",
            Description = "License consumption by SKU",
            Category = "Licensing",
            RequiredPermissions = ["Directory.Read.All"],
            DefaultWidth = 2,
            DefaultHeight = 1
        },
        [DeviceCompliance] = new WidgetDefinition
        {
            Type = DeviceCompliance,
            Name = "Device Compliance",
            Description = "Intune device compliance status",
            Category = "Devices",
            RequiredPermissions = ["DeviceManagementManagedDevices.Read.All"],
            DefaultWidth = 1,
            DefaultHeight = 1
        },
        [MailActivity] = new WidgetDefinition
        {
            Type = MailActivity,
            Name = "Mail Activity",
            Description = "Email sent/received trends",
            Category = "Usage",
            RequiredPermissions = ["Reports.Read.All"],
            DefaultWidth = 2,
            DefaultHeight = 1
        },
        [TeamsActivity] = new WidgetDefinition
        {
            Type = TeamsActivity,
            Name = "Teams Activity",
            Description = "Messages, calls, and meetings summary",
            Category = "Usage",
            RequiredPermissions = ["Reports.Read.All"],
            DefaultWidth = 2,
            DefaultHeight = 1
        }
    };
}

public class WidgetDefinition
{
    public string Type { get; set; } = string.Empty;
    public string Name { get; set; } = string.Empty;
    public string Description { get; set; } = string.Empty;
    public string Category { get; set; } = string.Empty;
    public string[] RequiredPermissions { get; set; } = [];
    public int DefaultWidth { get; set; } = 1;
    public int DefaultHeight { get; set; } = 1;
}

public static class DateRangePresets
{
    public const string Last7Days = "last7days";
    public const string Last30Days = "last30days";
    public const string Last90Days = "last90days";
    public const string ThisMonth = "thismonth";
    public const string LastMonth = "lastmonth";
    public const string Custom = "custom";

    public static (DateTime Start, DateTime End) GetDateRange(string preset)
    {
        var now = DateTime.UtcNow;
        return preset switch
        {
            Last7Days => (now.AddDays(-7), now),
            Last30Days => (now.AddDays(-30), now),
            Last90Days => (now.AddDays(-90), now),
            ThisMonth => (new DateTime(now.Year, now.Month, 1), now),
            LastMonth => (new DateTime(now.Year, now.Month, 1).AddMonths(-1), 
                         new DateTime(now.Year, now.Month, 1).AddDays(-1)),
            _ => (now.AddDays(-30), now)
        };
    }
}
