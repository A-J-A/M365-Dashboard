using Microsoft.EntityFrameworkCore;
using M365Dashboard.Api.Data;
using M365Dashboard.Api.Models;
using M365Dashboard.Api.Models.Dtos;
using M365Dashboard.Api.Configuration;
using System.Text.Json;

namespace M365Dashboard.Api.Services;

public interface IUserSettingsService
{
    Task<UserSettingsDto> GetUserSettingsAsync(string userId);
    Task<UserSettingsDto> UpdateUserSettingsAsync(string userId, UpdateUserSettingsDto settings);
    Task<List<WidgetConfigurationDto>> GetWidgetConfigurationsAsync(string userId);
    Task<WidgetConfigurationDto> UpdateWidgetConfigurationAsync(string userId, string widgetType, UpdateWidgetConfigurationDto config);
    Task<List<WidgetConfigurationDto>> ResetWidgetsToDefaultAsync(string userId);
    Task<DashboardLayoutDto?> GetDefaultLayoutAsync(string userId);
    Task<DashboardLayoutDto> SaveLayoutAsync(string userId, CreateDashboardLayoutDto layout);
}

public class UserSettingsService : IUserSettingsService
{
    private readonly ApplicationDbContext _context;
    private readonly ILogger<UserSettingsService> _logger;

    public UserSettingsService(ApplicationDbContext context, ILogger<UserSettingsService> logger)
    {
        _context = context;
        _logger = logger;
    }

    public async Task<UserSettingsDto> GetUserSettingsAsync(string userId)
    {
        var settings = await _context.UserSettings
            .FirstOrDefaultAsync(s => s.UserId == userId);

        if (settings == null)
        {
            settings = await CreateDefaultUserSettingsAsync(userId);
        }

        return new UserSettingsDto(
            settings.Theme,
            settings.RefreshIntervalSeconds,
            settings.DateRangePreference,
            settings.ShowWelcomeMessage,
            settings.CompactMode
        );
    }

    public async Task<UserSettingsDto> UpdateUserSettingsAsync(string userId, UpdateUserSettingsDto dto)
    {
        var settings = await _context.UserSettings
            .FirstOrDefaultAsync(s => s.UserId == userId);

        if (settings == null)
        {
            settings = await CreateDefaultUserSettingsAsync(userId);
        }

        if (dto.Theme != null) settings.Theme = dto.Theme;
        if (dto.RefreshIntervalSeconds.HasValue) settings.RefreshIntervalSeconds = dto.RefreshIntervalSeconds.Value;
        if (dto.DateRangePreference != null) settings.DateRangePreference = dto.DateRangePreference;
        if (dto.ShowWelcomeMessage.HasValue) settings.ShowWelcomeMessage = dto.ShowWelcomeMessage.Value;
        if (dto.CompactMode.HasValue) settings.CompactMode = dto.CompactMode.Value;

        settings.UpdatedAt = DateTime.UtcNow;

        await _context.SaveChangesAsync();

        _logger.LogInformation("Updated settings for user {UserId}", userId);

        return new UserSettingsDto(
            settings.Theme,
            settings.RefreshIntervalSeconds,
            settings.DateRangePreference,
            settings.ShowWelcomeMessage,
            settings.CompactMode
        );
    }

    public async Task<List<WidgetConfigurationDto>> GetWidgetConfigurationsAsync(string userId)
    {
        var configs = await _context.WidgetConfigurations
            .Where(w => w.UserId == userId)
            .OrderBy(w => w.DisplayOrder)
            .ToListAsync();

        if (configs.Count == 0)
        {
            configs = await CreateDefaultWidgetConfigurationsAsync(userId);
        }

        return configs.Select(MapToDto).ToList();
    }

    public async Task<WidgetConfigurationDto> UpdateWidgetConfigurationAsync(
        string userId, 
        string widgetType, 
        UpdateWidgetConfigurationDto dto)
    {
        var config = await _context.WidgetConfigurations
            .FirstOrDefaultAsync(w => w.UserId == userId && w.WidgetType == widgetType);

        if (config == null)
        {
            config = new WidgetConfiguration
            {
                UserId = userId,
                WidgetType = widgetType
            };
            _context.WidgetConfigurations.Add(config);
        }

        if (dto.IsEnabled.HasValue) config.IsEnabled = dto.IsEnabled.Value;
        if (dto.DisplayOrder.HasValue) config.DisplayOrder = dto.DisplayOrder.Value;
        if (dto.GridColumn.HasValue) config.GridColumn = dto.GridColumn.Value;
        if (dto.GridRow.HasValue) config.GridRow = dto.GridRow.Value;
        if (dto.GridWidth.HasValue) config.GridWidth = dto.GridWidth.Value;
        if (dto.GridHeight.HasValue) config.GridHeight = dto.GridHeight.Value;
        if (dto.CustomSettings != null) config.CustomSettings = JsonSerializer.Serialize(dto.CustomSettings);

        config.UpdatedAt = DateTime.UtcNow;

        await _context.SaveChangesAsync();

        return MapToDto(config);
    }

    public async Task<List<WidgetConfigurationDto>> ResetWidgetsToDefaultAsync(string userId)
    {
        var existing = await _context.WidgetConfigurations
            .Where(w => w.UserId == userId)
            .ToListAsync();

        _context.WidgetConfigurations.RemoveRange(existing);
        await _context.SaveChangesAsync();

        var defaults = await CreateDefaultWidgetConfigurationsAsync(userId);
        return defaults.Select(MapToDto).ToList();
    }

    public async Task<DashboardLayoutDto?> GetDefaultLayoutAsync(string userId)
    {
        var layout = await _context.DashboardLayouts
            .FirstOrDefaultAsync(l => l.UserId == userId && l.IsDefault);

        if (layout == null) return null;

        var widgets = await GetWidgetConfigurationsAsync(userId);

        return new DashboardLayoutDto(
            layout.Id,
            layout.Name,
            layout.IsDefault,
            widgets,
            layout.CreatedAt,
            layout.UpdatedAt
        );
    }

    public async Task<DashboardLayoutDto> SaveLayoutAsync(string userId, CreateDashboardLayoutDto dto)
    {
        if (dto.IsDefault)
        {
            // Reset other layouts to non-default
            var existingDefaults = await _context.DashboardLayouts
                .Where(l => l.UserId == userId && l.IsDefault)
                .ToListAsync();

            foreach (var existing in existingDefaults)
            {
                existing.IsDefault = false;
            }
        }

        var layout = new DashboardLayout
        {
            UserId = userId,
            Name = dto.Name,
            IsDefault = dto.IsDefault,
            LayoutJson = JsonSerializer.Serialize(dto.Widgets)
        };

        _context.DashboardLayouts.Add(layout);
        await _context.SaveChangesAsync();

        return new DashboardLayoutDto(
            layout.Id,
            layout.Name,
            layout.IsDefault,
            dto.Widgets,
            layout.CreatedAt,
            layout.UpdatedAt
        );
    }

    private async Task<UserSettings> CreateDefaultUserSettingsAsync(string userId)
    {
        var settings = new UserSettings
        {
            UserId = userId,
            Theme = "system",
            RefreshIntervalSeconds = 300,
            DateRangePreference = "last30days",
            ShowWelcomeMessage = true,
            CompactMode = false
        };

        _context.UserSettings.Add(settings);
        await _context.SaveChangesAsync();

        _logger.LogInformation("Created default settings for user {UserId}", userId);

        return settings;
    }

    private async Task<List<WidgetConfiguration>> CreateDefaultWidgetConfigurationsAsync(string userId)
    {
        var defaultWidgets = new List<WidgetConfiguration>
        {
            new()
            {
                UserId = userId,
                WidgetType = WidgetTypes.ActiveUsers,
                IsEnabled = true,
                DisplayOrder = 0,
                GridColumn = 0,
                GridRow = 0,
                GridWidth = 2,
                GridHeight = 1
            },
            new()
            {
                UserId = userId,
                WidgetType = WidgetTypes.SignInAnalytics,
                IsEnabled = true,
                DisplayOrder = 1,
                GridColumn = 2,
                GridRow = 0,
                GridWidth = 2,
                GridHeight = 2
            },
            new()
            {
                UserId = userId,
                WidgetType = WidgetTypes.LicenseUsage,
                IsEnabled = true,
                DisplayOrder = 2,
                GridColumn = 0,
                GridRow = 1,
                GridWidth = 2,
                GridHeight = 1
            },
            new()
            {
                UserId = userId,
                WidgetType = WidgetTypes.DeviceCompliance,
                IsEnabled = true,
                DisplayOrder = 3,
                GridColumn = 0,
                GridRow = 2,
                GridWidth = 1,
                GridHeight = 1
            },
            new()
            {
                UserId = userId,
                WidgetType = WidgetTypes.MailActivity,
                IsEnabled = true,
                DisplayOrder = 4,
                GridColumn = 1,
                GridRow = 2,
                GridWidth = 2,
                GridHeight = 1
            },
            new()
            {
                UserId = userId,
                WidgetType = WidgetTypes.TeamsActivity,
                IsEnabled = true,
                DisplayOrder = 5,
                GridColumn = 3,
                GridRow = 2,
                GridWidth = 1,
                GridHeight = 1
            }
        };

        _context.WidgetConfigurations.AddRange(defaultWidgets);
        await _context.SaveChangesAsync();

        _logger.LogInformation("Created default widget configurations for user {UserId}", userId);

        return defaultWidgets;
    }

    private static WidgetConfigurationDto MapToDto(WidgetConfiguration config)
    {
        Dictionary<string, object>? customSettings = null;
        if (!string.IsNullOrEmpty(config.CustomSettings))
        {
            customSettings = JsonSerializer.Deserialize<Dictionary<string, object>>(config.CustomSettings);
        }

        return new WidgetConfigurationDto(
            config.Id,
            config.WidgetType,
            config.IsEnabled,
            config.DisplayOrder,
            config.GridColumn,
            config.GridRow,
            config.GridWidth,
            config.GridHeight,
            customSettings
        );
    }
}
