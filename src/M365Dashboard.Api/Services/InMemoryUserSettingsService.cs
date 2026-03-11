using M365Dashboard.Api.Models.Dtos;
using M365Dashboard.Api.Configuration;

namespace M365Dashboard.Api.Services;

/// <summary>
/// In-memory implementation of IUserSettingsService for development/testing
/// when database is not available
/// </summary>
public class InMemoryUserSettingsService : IUserSettingsService
{
    private readonly Dictionary<string, UserSettingsDto> _settings = new();
    private readonly Dictionary<string, List<WidgetConfigurationDto>> _widgets = new();
    private readonly ILogger<InMemoryUserSettingsService> _logger;

    public InMemoryUserSettingsService(ILogger<InMemoryUserSettingsService> logger)
    {
        _logger = logger;
    }

    public Task<UserSettingsDto> GetUserSettingsAsync(string userId)
    {
        if (!_settings.TryGetValue(userId, out var settings))
        {
            settings = new UserSettingsDto(
                Theme: "system",
                RefreshIntervalSeconds: 300,
                DateRangePreference: "last30days",
                ShowWelcomeMessage: true,
                CompactMode: false
            );
            _settings[userId] = settings;
        }

        return Task.FromResult(settings);
    }

    public Task<UserSettingsDto> UpdateUserSettingsAsync(string userId, UpdateUserSettingsDto dto)
    {
        var current = _settings.GetValueOrDefault(userId) ?? new UserSettingsDto(
            "system", 300, "last30days", true, false
        );

        var updated = new UserSettingsDto(
            dto.Theme ?? current.Theme,
            dto.RefreshIntervalSeconds ?? current.RefreshIntervalSeconds,
            dto.DateRangePreference ?? current.DateRangePreference,
            dto.ShowWelcomeMessage ?? current.ShowWelcomeMessage,
            dto.CompactMode ?? current.CompactMode
        );

        _settings[userId] = updated;
        return Task.FromResult(updated);
    }

    public Task<List<WidgetConfigurationDto>> GetWidgetConfigurationsAsync(string userId)
    {
        if (!_widgets.TryGetValue(userId, out var widgets))
        {
            widgets = GetDefaultWidgets();
            _widgets[userId] = widgets;
        }

        return Task.FromResult(widgets);
    }

    public Task<WidgetConfigurationDto> UpdateWidgetConfigurationAsync(
        string userId,
        string widgetType,
        UpdateWidgetConfigurationDto dto)
    {
        var widgets = _widgets.GetValueOrDefault(userId) ?? GetDefaultWidgets();
        var widget = widgets.FirstOrDefault(w => w.WidgetType == widgetType);

        if (widget != null)
        {
            var index = widgets.IndexOf(widget);
            var updated = new WidgetConfigurationDto(
                widget.Id,
                widget.WidgetType,
                dto.IsEnabled ?? widget.IsEnabled,
                dto.DisplayOrder ?? widget.DisplayOrder,
                dto.GridColumn ?? widget.GridColumn,
                dto.GridRow ?? widget.GridRow,
                dto.GridWidth ?? widget.GridWidth,
                dto.GridHeight ?? widget.GridHeight,
                dto.CustomSettings ?? widget.CustomSettings
            );
            widgets[index] = updated;
            _widgets[userId] = widgets;
            return Task.FromResult(updated);
        }

        throw new KeyNotFoundException($"Widget {widgetType} not found");
    }

    public Task<List<WidgetConfigurationDto>> ResetWidgetsToDefaultAsync(string userId)
    {
        var widgets = GetDefaultWidgets();
        _widgets[userId] = widgets;
        return Task.FromResult(widgets);
    }

    public Task<DashboardLayoutDto?> GetDefaultLayoutAsync(string userId)
    {
        return Task.FromResult<DashboardLayoutDto?>(null);
    }

    public Task<DashboardLayoutDto> SaveLayoutAsync(string userId, CreateDashboardLayoutDto dto)
    {
        var layout = new DashboardLayoutDto(
            1,
            dto.Name,
            dto.IsDefault,
            dto.Widgets,
            DateTime.UtcNow,
            DateTime.UtcNow
        );
        return Task.FromResult(layout);
    }

    private static List<WidgetConfigurationDto> GetDefaultWidgets()
    {
        return new List<WidgetConfigurationDto>
        {
            new(1, WidgetTypes.ActiveUsers, true, 0, 0, 0, 2, 1, null),
            new(2, WidgetTypes.SignInAnalytics, true, 1, 2, 0, 2, 2, null),
            new(3, WidgetTypes.LicenseUsage, true, 2, 0, 1, 2, 1, null),
            new(4, WidgetTypes.DeviceCompliance, true, 3, 0, 2, 1, 1, null),
            new(5, WidgetTypes.MailActivity, true, 4, 1, 2, 2, 1, null),
            new(6, WidgetTypes.TeamsActivity, true, 5, 3, 2, 1, 1, null),
        };
    }
}
