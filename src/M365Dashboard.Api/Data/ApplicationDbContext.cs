using Microsoft.EntityFrameworkCore;
using M365Dashboard.Api.Models;

namespace M365Dashboard.Api.Data;

public class ApplicationDbContext : DbContext
{
    public ApplicationDbContext(DbContextOptions<ApplicationDbContext> options)
        : base(options)
    {
    }

    public DbSet<UserSettings> UserSettings => Set<UserSettings>();
    public DbSet<WidgetConfiguration> WidgetConfigurations => Set<WidgetConfiguration>();
    public DbSet<CachedMetric> CachedMetrics => Set<CachedMetric>();
    public DbSet<DashboardLayout> DashboardLayouts => Set<DashboardLayout>();
    public DbSet<AuditLog> AuditLogs => Set<AuditLog>();
    public DbSet<ScheduledReport> ScheduledReports => Set<ScheduledReport>();
    public DbSet<ReportHistory> ReportHistories => Set<ReportHistory>();
    public DbSet<TenantSettings> TenantSettings => Set<TenantSettings>();

    protected override void OnModelCreating(ModelBuilder modelBuilder)
    {
        base.OnModelCreating(modelBuilder);

        // UserSettings configuration
        modelBuilder.Entity<UserSettings>(entity =>
        {
            entity.HasKey(e => e.Id);
            entity.HasIndex(e => e.UserId).IsUnique();
            entity.Property(e => e.UserId).HasMaxLength(100).IsRequired();
            entity.Property(e => e.Theme).HasMaxLength(20).HasDefaultValue("system");
            entity.Property(e => e.DateRangePreference).HasMaxLength(20).HasDefaultValue("last30days");
            entity.Property(e => e.CreatedAt).HasDefaultValueSql("GETUTCDATE()");
        });

        // WidgetConfiguration configuration
        modelBuilder.Entity<WidgetConfiguration>(entity =>
        {
            entity.HasKey(e => e.Id);
            entity.HasIndex(e => new { e.UserId, e.WidgetType }).IsUnique();
            entity.Property(e => e.UserId).HasMaxLength(100).IsRequired();
            entity.Property(e => e.WidgetType).HasMaxLength(50).IsRequired();
            entity.Property(e => e.CustomSettings).HasColumnType("nvarchar(max)");
        });

        // CachedMetric configuration
        modelBuilder.Entity<CachedMetric>(entity =>
        {
            entity.HasKey(e => e.Id);
            entity.HasIndex(e => new { e.MetricType, e.TenantId });
            entity.HasIndex(e => e.ExpiresAt);
            entity.Property(e => e.MetricType).HasMaxLength(100).IsRequired();
            entity.Property(e => e.TenantId).HasMaxLength(100).IsRequired();
            entity.Property(e => e.Data).HasColumnType("nvarchar(max)").IsRequired();
        });

        // DashboardLayout configuration
        modelBuilder.Entity<DashboardLayout>(entity =>
        {
            entity.HasKey(e => e.Id);
            entity.HasIndex(e => e.UserId);
            entity.Property(e => e.UserId).HasMaxLength(100).IsRequired();
            entity.Property(e => e.Name).HasMaxLength(100).IsRequired();
            entity.Property(e => e.LayoutJson).HasColumnType("nvarchar(max)").IsRequired();
        });

        // AuditLog configuration
        modelBuilder.Entity<AuditLog>(entity =>
        {
            entity.HasKey(e => e.Id);
            entity.HasIndex(e => e.UserId);
            entity.HasIndex(e => e.Timestamp);
            entity.Property(e => e.UserId).HasMaxLength(100).IsRequired();
            entity.Property(e => e.Action).HasMaxLength(100).IsRequired();
            entity.Property(e => e.Details).HasColumnType("nvarchar(max)");
        });

        // ScheduledReport configuration
        modelBuilder.Entity<ScheduledReport>(entity =>
        {
            entity.HasKey(e => e.Id);
            entity.HasIndex(e => e.UserId);
            entity.HasIndex(e => e.NextRunAt);
            entity.HasIndex(e => new { e.IsEnabled, e.NextRunAt });
            entity.Property(e => e.UserId).HasMaxLength(100).IsRequired();
            entity.Property(e => e.UserEmail).HasMaxLength(256);
            entity.Property(e => e.ReportType).HasMaxLength(50).IsRequired();
            entity.Property(e => e.DisplayName).HasMaxLength(100).IsRequired();
            entity.Property(e => e.Frequency).HasMaxLength(20).IsRequired();
            entity.Property(e => e.TimeOfDay).HasMaxLength(10).HasDefaultValue("08:00");
            entity.Property(e => e.Recipients).HasColumnType("nvarchar(max)").IsRequired();
            entity.Property(e => e.DateRange).HasMaxLength(20);
            entity.Property(e => e.LastRunStatus).HasMaxLength(50);
            entity.Property(e => e.LastRunError).HasColumnType("nvarchar(max)");
            entity.Property(e => e.CreatedAt).HasDefaultValueSql("GETUTCDATE()");
        });

        // ReportHistory configuration
        modelBuilder.Entity<ReportHistory>(entity =>
        {
            entity.HasKey(e => e.Id);
            entity.HasIndex(e => e.UserId);
            entity.HasIndex(e => e.GeneratedAt);
            entity.Property(e => e.UserId).HasMaxLength(100).IsRequired();
            entity.Property(e => e.ReportType).HasMaxLength(50).IsRequired();
            entity.Property(e => e.DisplayName).HasMaxLength(100).IsRequired();
            entity.Property(e => e.Status).HasMaxLength(20).IsRequired();
            entity.Property(e => e.ErrorMessage).HasColumnType("nvarchar(max)");
            entity.Property(e => e.GeneratedAt).HasDefaultValueSql("GETUTCDATE()");
        });

        // TenantSettings configuration
        modelBuilder.Entity<TenantSettings>(entity =>
        {
            entity.HasKey(e => e.Id);
            entity.HasIndex(e => new { e.TenantId, e.SettingKey }).IsUnique();
            entity.Property(e => e.TenantId).HasMaxLength(100).IsRequired();
            entity.Property(e => e.SettingKey).HasMaxLength(100).IsRequired();
            entity.Property(e => e.SettingValue).HasColumnType("nvarchar(max)").IsRequired();
            entity.Property(e => e.Description).HasMaxLength(500);
            entity.Property(e => e.LastModifiedBy).HasMaxLength(100);
            entity.Property(e => e.CreatedAt).HasDefaultValueSql("GETUTCDATE()");
        });

        // Seed default widget types
        SeedWidgetTypes(modelBuilder);
    }

    private static void SeedWidgetTypes(ModelBuilder modelBuilder)
    {
        // We don't seed user-specific data, but we could seed lookup tables here
    }
}
