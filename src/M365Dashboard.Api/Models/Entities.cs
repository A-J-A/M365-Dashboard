using System.ComponentModel.DataAnnotations;

namespace M365Dashboard.Api.Models;

public class UserSettings
{
    public int Id { get; set; }
    
    [Required]
    [MaxLength(100)]
    public string UserId { get; set; } = string.Empty;
    
    [MaxLength(20)]
    public string Theme { get; set; } = "system";
    
    public int RefreshIntervalSeconds { get; set; } = 300;
    
    [MaxLength(20)]
    public string DateRangePreference { get; set; } = "last30days";
    
    public bool ShowWelcomeMessage { get; set; } = true;
    
    public bool CompactMode { get; set; } = false;
    
    public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
    
    public DateTime UpdatedAt { get; set; } = DateTime.UtcNow;
}

public class WidgetConfiguration
{
    public int Id { get; set; }
    
    [Required]
    [MaxLength(100)]
    public string UserId { get; set; } = string.Empty;
    
    [Required]
    [MaxLength(50)]
    public string WidgetType { get; set; } = string.Empty;
    
    public bool IsEnabled { get; set; } = true;
    
    public int DisplayOrder { get; set; }
    
    public int GridColumn { get; set; }
    
    public int GridRow { get; set; }
    
    public int GridWidth { get; set; } = 1;
    
    public int GridHeight { get; set; } = 1;
    
    /// <summary>
    /// JSON string containing widget-specific settings
    /// </summary>
    public string? CustomSettings { get; set; }
    
    public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
    
    public DateTime UpdatedAt { get; set; } = DateTime.UtcNow;
}

public class CachedMetric
{
    public int Id { get; set; }
    
    [Required]
    [MaxLength(100)]
    public string MetricType { get; set; } = string.Empty;
    
    [Required]
    [MaxLength(100)]
    public string TenantId { get; set; } = string.Empty;
    
    /// <summary>
    /// JSON serialized metric data
    /// </summary>
    [Required]
    public string Data { get; set; } = string.Empty;
    
    public DateTime CachedAt { get; set; } = DateTime.UtcNow;
    
    public DateTime ExpiresAt { get; set; }
}

public class DashboardLayout
{
    public int Id { get; set; }
    
    [Required]
    [MaxLength(100)]
    public string UserId { get; set; } = string.Empty;
    
    [Required]
    [MaxLength(100)]
    public string Name { get; set; } = string.Empty;
    
    public bool IsDefault { get; set; }
    
    /// <summary>
    /// JSON containing the full layout configuration
    /// </summary>
    [Required]
    public string LayoutJson { get; set; } = string.Empty;
    
    public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
    
    public DateTime UpdatedAt { get; set; } = DateTime.UtcNow;
}

public class AuditLog
{
    public int Id { get; set; }
    
    [Required]
    [MaxLength(100)]
    public string UserId { get; set; } = string.Empty;
    
    [Required]
    [MaxLength(100)]
    public string Action { get; set; } = string.Empty;
    
    public string? Details { get; set; }
    
    public string? IpAddress { get; set; }
    
    public DateTime Timestamp { get; set; } = DateTime.UtcNow;
}

public class ScheduledReport
{
    public int Id { get; set; }
    
    [Required]
    [MaxLength(100)]
    public string UserId { get; set; } = string.Empty;
    
    [MaxLength(256)]
    public string? UserEmail { get; set; }
    
    [Required]
    [MaxLength(50)]
    public string ReportType { get; set; } = string.Empty;
    
    [Required]
    [MaxLength(100)]
    public string DisplayName { get; set; } = string.Empty;
    
    /// <summary>
    /// Frequency: daily, weekly, monthly
    /// </summary>
    [Required]
    [MaxLength(20)]
    public string Frequency { get; set; } = "weekly";
    
    /// <summary>
    /// Time of day to run (HH:mm format, UTC)
    /// </summary>
    [MaxLength(10)]
    public string TimeOfDay { get; set; } = "08:00";
    
    /// <summary>
    /// Day of week for weekly reports (0=Sunday, 1=Monday, etc.)
    /// </summary>
    public int? DayOfWeek { get; set; }
    
    /// <summary>
    /// Day of month for monthly reports (1-28)
    /// </summary>
    public int? DayOfMonth { get; set; }
    
    /// <summary>
    /// Comma-separated list of email recipients
    /// </summary>
    [Required]
    public string Recipients { get; set; } = string.Empty;
    
    /// <summary>
    /// Date range for the report (last7days, last30days, etc.)
    /// </summary>
    [MaxLength(20)]
    public string? DateRange { get; set; }
    
    public bool IsEnabled { get; set; } = true;
    
    public DateTime? LastRunAt { get; set; }
    
    public DateTime? NextRunAt { get; set; }
    
    [MaxLength(50)]
    public string? LastRunStatus { get; set; }
    
    public string? LastRunError { get; set; }
    
    public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
    
    public DateTime UpdatedAt { get; set; } = DateTime.UtcNow;
}

public class ReportHistory
{
    public int Id { get; set; }
    
    [Required]
    [MaxLength(100)]
    public string UserId { get; set; } = string.Empty;
    
    [Required]
    [MaxLength(50)]
    public string ReportType { get; set; } = string.Empty;
    
    [Required]
    [MaxLength(100)]
    public string DisplayName { get; set; } = string.Empty;
    
    public DateTime GeneratedAt { get; set; } = DateTime.UtcNow;
    
    [Required]
    [MaxLength(20)]
    public string Status { get; set; } = "success";
    
    public string? ErrorMessage { get; set; }
    
    public int? RecordCount { get; set; }
    
    public bool WasScheduled { get; set; }
    
    /// <summary>
    /// Reference to the scheduled report if this was a scheduled run
    /// </summary>
    public int? ScheduledReportId { get; set; }
}

/// <summary>
/// Tenant-level settings for reports and configurations
/// </summary>
public class TenantSettings
{
    public int Id { get; set; }
    
    /// <summary>
    /// The tenant ID this setting applies to (defaults to app's tenant)
    /// </summary>
    [Required]
    [MaxLength(100)]
    public string TenantId { get; set; } = string.Empty;
    
    /// <summary>
    /// Setting key (e.g., "BreakGlassAccounts")
    /// </summary>
    [Required]
    [MaxLength(100)]
    public string SettingKey { get; set; } = string.Empty;
    
    /// <summary>
    /// JSON serialized setting value
    /// </summary>
    [Required]
    public string SettingValue { get; set; } = string.Empty;
    
    /// <summary>
    /// User-friendly description of this setting
    /// </summary>
    [MaxLength(500)]
    public string? Description { get; set; }
    
    /// <summary>
    /// User who last modified this setting
    /// </summary>
    [MaxLength(100)]
    public string? LastModifiedBy { get; set; }
    
    public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
    
    public DateTime UpdatedAt { get; set; } = DateTime.UtcNow;
}
