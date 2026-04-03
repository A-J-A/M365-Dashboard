using M365Dashboard.Api.Models.Dtos;
using M365Dashboard.Api.Configuration;
using Microsoft.Extensions.Options;

namespace M365Dashboard.Api.Services;

public interface IWidgetDataService
{
    Task<ActiveUsersDataDto> GetActiveUsersDataAsync(string dateRange);
    Task<SignInAnalyticsDto> GetSignInAnalyticsDataAsync(string dateRange);
    Task<LicenseUsageDto> GetLicenseUsageDataAsync();
    Task<DeviceComplianceDto> GetDeviceComplianceDataAsync();
    Task<MailActivityDto> GetMailActivityDataAsync(string dateRange);
    Task<TeamsActivityDto> GetTeamsActivityDataAsync(string dateRange);
    Task<DashboardSummaryDto> GetDashboardSummaryAsync();
    Task RefreshCacheAsync(string metricType);
}

public class WidgetDataService : IWidgetDataService
{
    private readonly IGraphService _graphService;
    private readonly ICacheService _cacheService;
    private readonly ILogger<WidgetDataService> _logger;
    private readonly CacheOptions _cacheOptions;

    public WidgetDataService(
        IGraphService graphService,
        ICacheService cacheService,
        ILogger<WidgetDataService> logger,
        IOptions<CacheOptions> cacheOptions)
    {
        _graphService = graphService;
        _cacheService = cacheService;
        _logger = logger;
        _cacheOptions = cacheOptions.Value;
    }

    public async Task<ActiveUsersDataDto> GetActiveUsersDataAsync(string dateRange)
    {
        var cacheKey = $"active-users:{dateRange}";
        
        var cached = await _cacheService.GetAsync<ActiveUsersDataDto>(cacheKey);
        if (cached != null)
        {
            _logger.LogDebug("Cache hit for {CacheKey}", cacheKey);
            return cached;
        }

        _logger.LogDebug("Cache miss for {CacheKey}, fetching from Graph", cacheKey);
        
        var (startDate, endDate) = DateRangePresets.GetDateRange(dateRange);
        var data = await _graphService.GetActiveUsersAsync(startDate, endDate);
        
        await _cacheService.SetAsync(cacheKey, data, TimeSpan.FromMinutes(_cacheOptions.ReportDataTtlMinutes));
        
        return data;
    }

    public async Task<SignInAnalyticsDto> GetSignInAnalyticsDataAsync(string dateRange)
    {
        var cacheKey = $"sign-in-analytics:{dateRange}";
        
        var cached = await _cacheService.GetAsync<SignInAnalyticsDto>(cacheKey);
        if (cached != null)
        {
            return cached;
        }

        var (startDate, endDate) = DateRangePresets.GetDateRange(dateRange);
        var data = await _graphService.GetSignInAnalyticsAsync(startDate, endDate);
        
        await _cacheService.SetAsync(cacheKey, data, TimeSpan.FromMinutes(_cacheOptions.SignInDataTtlMinutes));
        
        return data;
    }

    public async Task<LicenseUsageDto> GetLicenseUsageDataAsync()
    {
        const string cacheKey = "license-usage";
        
        var cached = await _cacheService.GetAsync<LicenseUsageDto>(cacheKey);
        if (cached != null)
        {
            return cached;
        }

        var data = await _graphService.GetLicenseUsageAsync();
        
        await _cacheService.SetAsync(cacheKey, data, TimeSpan.FromMinutes(_cacheOptions.LicenseDataTtlMinutes));
        
        return data;
    }

    public async Task<DeviceComplianceDto> GetDeviceComplianceDataAsync()
    {
        const string cacheKey = "device-compliance";
        
        var cached = await _cacheService.GetAsync<DeviceComplianceDto>(cacheKey);
        if (cached != null)
        {
            return cached;
        }

        var data = await _graphService.GetDeviceComplianceAsync();
        
        await _cacheService.SetAsync(cacheKey, data, TimeSpan.FromMinutes(_cacheOptions.DefaultTtlMinutes));
        
        return data;
    }

    public async Task<MailActivityDto> GetMailActivityDataAsync(string dateRange)
    {
        var cacheKey = $"mail-activity:{dateRange}";
        
        var cached = await _cacheService.GetAsync<MailActivityDto>(cacheKey);
        if (cached != null)
        {
            return cached;
        }

        var (startDate, endDate) = DateRangePresets.GetDateRange(dateRange);
        var data = await _graphService.GetMailActivityAsync(startDate, endDate);
        
        await _cacheService.SetAsync(cacheKey, data, TimeSpan.FromMinutes(_cacheOptions.ReportDataTtlMinutes));
        
        return data;
    }

    public async Task<TeamsActivityDto> GetTeamsActivityDataAsync(string dateRange)
    {
        var cacheKey = $"teams-activity:{dateRange}";
        
        var cached = await _cacheService.GetAsync<TeamsActivityDto>(cacheKey);
        if (cached != null)
        {
            return cached;
        }

        var (startDate, endDate) = DateRangePresets.GetDateRange(dateRange);
        var data = await _graphService.GetTeamsActivityAsync(startDate, endDate);
        
        await _cacheService.SetAsync(cacheKey, data, TimeSpan.FromMinutes(_cacheOptions.ReportDataTtlMinutes));
        
        return data;
    }

    public async Task<DashboardSummaryDto> GetDashboardSummaryAsync()
    {
        const string cacheKey = "dashboard-summary";
        
        var cached = await _cacheService.GetAsync<DashboardSummaryDto>(cacheKey);
        if (cached != null)
        {
            return cached;
        }

        // Fetch all data in parallel
        var activeUsersTask = GetActiveUsersDataAsync(DateRangePresets.Last30Days);
        var signInTask = GetSignInAnalyticsDataAsync(DateRangePresets.Last7Days);
        var licenseTask = GetLicenseUsageDataAsync();
        var complianceTask = GetDeviceComplianceDataAsync();

        await Task.WhenAll(activeUsersTask, signInTask, licenseTask, complianceTask);

        var summary = new DashboardSummaryDto(
            ActiveUsers: activeUsersTask.Result.MonthlyActiveUsers,
            SignInSuccessRate: signInTask.Result.SuccessRate,
            LicenseUtilization: licenseTask.Result.OverallUtilization,
            DeviceComplianceRate: complianceTask.Result.ComplianceRate,
            LastUpdated: DateTime.UtcNow
        );

        await _cacheService.SetAsync(cacheKey, summary, TimeSpan.FromMinutes(_cacheOptions.DefaultTtlMinutes));

        return summary;
    }

    public async Task RefreshCacheAsync(string metricType)
    {
        _logger.LogInformation("Refreshing cache for metric type: {MetricType}", metricType);

        switch (metricType.ToLower())
        {
            case "active-users":
                await _cacheService.RemoveByPrefixAsync("active-users:");
                break;
            case "sign-in-analytics":
                await _cacheService.RemoveByPrefixAsync("sign-in-analytics:");
                break;
            case "license-usage":
                await _cacheService.RemoveAsync("license-usage");
                break;
            case "device-compliance":
                await _cacheService.RemoveAsync("device-compliance");
                break;
            case "mail-activity":
                await _cacheService.RemoveByPrefixAsync("mail-activity:");
                break;
            case "teams-activity":
                await _cacheService.RemoveByPrefixAsync("teams-activity:");
                break;
            case "all":
                await _cacheService.ClearAsync();
                break;
            default:
                _logger.LogWarning("Unknown metric type: {MetricType}", metricType);
                break;
        }
    }
}
