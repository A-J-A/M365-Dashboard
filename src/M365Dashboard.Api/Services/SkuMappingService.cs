using System.Collections.Concurrent;
using System.Globalization;
using CsvHelper;
using CsvHelper.Configuration;

namespace M365Dashboard.Api.Services;

public interface ISkuMappingService
{
    string GetFriendlyName(string skuPartNumber);
    bool IsFreeTrial(string skuPartNumber);
    Task RefreshMappingsAsync();
    SkuMappingStatus GetStatus();
}

public class SkuMappingStatus
{
    public DateTime? LastRefreshed { get; set; }
    public int TotalMappings { get; set; }
    public int FreeTrialSkusCount { get; set; }
    public bool IsRefreshing { get; set; }
    public string? LastError { get; set; }
    public DateTime? NextScheduledRefresh { get; set; }
}

public class SkuMappingService : ISkuMappingService, IHostedService, IDisposable
{
    private readonly ILogger<SkuMappingService> _logger;
    private readonly IHttpClientFactory _httpClientFactory;
    private readonly ConcurrentDictionary<string, string> _skuMappings = new(StringComparer.OrdinalIgnoreCase);
    private readonly HashSet<string> _freeTrialSkus = new(StringComparer.OrdinalIgnoreCase);
    private Timer? _timer;
    private readonly SemaphoreSlim _refreshLock = new(1, 1);
    
    // Status tracking
    private DateTime? _lastRefreshed;
    private DateTime? _nextScheduledRefresh;
    private bool _isRefreshing;
    private string? _lastError;

    private const string MicrosoftCsvUrl = "https://download.microsoft.com/download/e/3/e/e3e9faf2-f28b-490a-9ada-c6089a1fc5b0/Product%20names%20and%20service%20plan%20identifiers%20for%20licensing.csv";

    // Fallback mappings for when CSV is unavailable
    private static readonly Dictionary<string, string> FallbackMappings = new(StringComparer.OrdinalIgnoreCase)
    {
        // Microsoft 365 / Office 365 Suites
        { "ENTERPRISEPACK", "Office 365 E3" },
        { "ENTERPRISEPREMIUM", "Office 365 E5" },
        { "SPE_E3", "Microsoft 365 E3" },
        { "SPE_E5", "Microsoft 365 E5" },
        { "SPB", "Microsoft 365 Business Premium" },
        { "O365_BUSINESS_ESSENTIALS", "Microsoft 365 Business Basic" },
        { "O365_BUSINESS_PREMIUM", "Microsoft 365 Business Standard" },
        { "OFFICESUBSCRIPTION", "Microsoft 365 Apps for Enterprise" },
        { "M365_F1", "Microsoft 365 F1" },
        { "DESKLESSPACK", "Office 365 F3" },
        { "STANDARDPACK", "Office 365 E1" },
        
        // Teams
        { "MCOMEETADV", "Microsoft 365 Audio Conferencing" },
        { "MCOEV", "Microsoft 365 Phone System" },
        { "MCOPSTN1", "Microsoft 365 Domestic Calling Plan" },
        { "MCOPSTN2", "Microsoft 365 Domestic and International Calling Plan" },
        { "MCOPSTNC", "Communications Credits" },
        { "PHONESYSTEM_VIRTUALUSER", "Microsoft 365 Phone System - Virtual User" },
        { "TEAMS_EXPLORATORY", "Microsoft Teams Exploratory" },
        
        // Power Platform
        { "FLOW_FREE", "Power Automate Free" },
        { "POWERAUTOMATE_VIRAL", "Power Automate Free" },
        { "POWER_BI_STANDARD", "Power BI (Free)" },
        { "POWER_BI_PRO", "Power BI Pro" },
        
        // Security
        { "EMS", "Enterprise Mobility + Security E3" },
        { "EMSPREMIUM", "Enterprise Mobility + Security E5" },
        { "AAD_PREMIUM", "Microsoft Entra ID P1" },
        { "AAD_PREMIUM_P2", "Microsoft Entra ID P2" },
        { "INTUNE_A", "Microsoft Intune Plan 1" },
        
        // Rights Management
        { "RMSBASIC", "Rights Management Service Basic Content Protection" },
        { "RIGHTSMANAGEMENT_ADHOC", "Rights Management Adhoc" },
        
        // Other
        { "WINDOWS_STORE", "Windows Store for Business" },
        { "STREAM", "Microsoft Stream" },
        { "Microsoft_365_Copilot", "Microsoft 365 Copilot" },
    };

    // Known free/trial patterns
    private static readonly HashSet<string> KnownFreeTrialSkus = new(StringComparer.OrdinalIgnoreCase)
    {
        "FLOW_FREE", "POWERAUTOMATE_VIRAL", "POWERAPPS_VIRAL", "POWER_BI_STANDARD",
        "TEAMS_EXPLORATORY", "TEAMS_FREE", "CCIBOTS_PRIVPREV_VIRAL", "RIGHTSMANAGEMENT_ADHOC",
        "WINDOWS_STORE", "STREAM", "MCOPSTNC", "MICROSOFT_BUSINESS_CENTER", "RMSBASIC",
        "EXCHANGEDESKLESS", "SMB_APPS"
    };

    public SkuMappingService(ILogger<SkuMappingService> logger, IHttpClientFactory httpClientFactory)
    {
        _logger = logger;
        _httpClientFactory = httpClientFactory;
        
        // Initialize with fallback mappings
        foreach (var mapping in FallbackMappings)
        {
            _skuMappings[mapping.Key] = mapping.Value;
        }
        foreach (var sku in KnownFreeTrialSkus)
        {
            _freeTrialSkus.Add(sku);
        }
    }

    public string GetFriendlyName(string skuPartNumber)
    {
        if (string.IsNullOrEmpty(skuPartNumber))
            return "Unknown";

        if (_skuMappings.TryGetValue(skuPartNumber, out var friendlyName))
            return friendlyName;

        // Fallback: make the SKU part number more readable
        var name = skuPartNumber.Replace("_", " ").Replace("-", " ");
        return CultureInfo.CurrentCulture.TextInfo.ToTitleCase(name.ToLower());
    }

    public bool IsFreeTrial(string skuPartNumber)
    {
        if (string.IsNullOrEmpty(skuPartNumber))
            return false;

        return _freeTrialSkus.Contains(skuPartNumber) ||
               skuPartNumber.Contains("TRIAL", StringComparison.OrdinalIgnoreCase) ||
               skuPartNumber.Contains("FREE", StringComparison.OrdinalIgnoreCase) ||
               skuPartNumber.Contains("VIRAL", StringComparison.OrdinalIgnoreCase) ||
               skuPartNumber.Contains("_DEV", StringComparison.OrdinalIgnoreCase);
    }

    public async Task RefreshMappingsAsync()
    {
        if (!await _refreshLock.WaitAsync(TimeSpan.FromSeconds(5)))
        {
            _logger.LogWarning("SKU mapping refresh already in progress, skipping");
            return;
        }

        _isRefreshing = true;
        
        try
        {
            _logger.LogInformation("Refreshing SKU mappings from Microsoft...");
            
            var client = _httpClientFactory.CreateClient();
            client.Timeout = TimeSpan.FromSeconds(30);
            
            var response = await client.GetAsync(MicrosoftCsvUrl);
            
            if (!response.IsSuccessStatusCode)
            {
                _lastError = $"Failed to download CSV: {response.StatusCode}";
                _logger.LogWarning("Failed to download SKU mappings CSV: {StatusCode}", response.StatusCode);
                return;
            }

            var csvContent = await response.Content.ReadAsStringAsync();
            ParseCsvMappings(csvContent);
            
            _lastRefreshed = DateTime.UtcNow;
            _lastError = null;
            _logger.LogInformation("Successfully refreshed SKU mappings. Total mappings: {Count}", _skuMappings.Count);
        }
        catch (Exception ex)
        {
            _lastError = ex.Message;
            _logger.LogError(ex, "Error refreshing SKU mappings from Microsoft");
        }
        finally
        {
            _isRefreshing = false;
            _refreshLock.Release();
        }
    }

    public SkuMappingStatus GetStatus()
    {
        return new SkuMappingStatus
        {
            LastRefreshed = _lastRefreshed,
            TotalMappings = _skuMappings.Count,
            FreeTrialSkusCount = _freeTrialSkus.Count,
            IsRefreshing = _isRefreshing,
            LastError = _lastError,
            NextScheduledRefresh = _nextScheduledRefresh
        };
    }

    private void ParseCsvMappings(string csvContent)
    {
        try
        {
            using var reader = new StringReader(csvContent);
            using var csv = new CsvReader(reader, new CsvConfiguration(CultureInfo.InvariantCulture)
            {
                HasHeaderRecord = true,
                MissingFieldFound = null,
                BadDataFound = null
            });

            // Read header to find column indices
            csv.Read();
            csv.ReadHeader();
            
            var productNameIndex = csv.GetFieldIndex("Product_Display_Name", isTryGet: true);
            var stringIdIndex = csv.GetFieldIndex("String_Id", isTryGet: true);
            
            // Fallback column names (Microsoft changes these sometimes)
            if (productNameIndex < 0)
                productNameIndex = csv.GetFieldIndex("Product name", isTryGet: true);
            if (stringIdIndex < 0)
                stringIdIndex = csv.GetFieldIndex("String ID", isTryGet: true);

            if (productNameIndex < 0 || stringIdIndex < 0)
            {
                _logger.LogWarning("Could not find required columns in CSV. Headers: {Headers}", 
                    string.Join(", ", csv.HeaderRecord ?? Array.Empty<string>()));
                return;
            }

            var newMappings = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            
            while (csv.Read())
            {
                try
                {
                    var stringId = csv.GetField(stringIdIndex)?.Trim();
                    var productName = csv.GetField(productNameIndex)?.Trim();

                    if (!string.IsNullOrEmpty(stringId) && !string.IsNullOrEmpty(productName))
                    {
                        newMappings[stringId] = productName;
                        
                        // Detect free/trial SKUs
                        if (productName.Contains("Free", StringComparison.OrdinalIgnoreCase) ||
                            productName.Contains("Trial", StringComparison.OrdinalIgnoreCase) ||
                            productName.Contains("Viral", StringComparison.OrdinalIgnoreCase))
                        {
                            _freeTrialSkus.Add(stringId);
                        }
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogDebug(ex, "Error parsing CSV row");
                }
            }

            // Merge new mappings (don't replace existing fallbacks that might be better)
            foreach (var mapping in newMappings)
            {
                _skuMappings[mapping.Key] = mapping.Value;
            }
            
            _logger.LogInformation("Parsed {Count} SKU mappings from CSV", newMappings.Count);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error parsing SKU mappings CSV");
        }
    }

    public Task StartAsync(CancellationToken cancellationToken)
    {
        _logger.LogInformation("SKU Mapping Service starting...");
        
        // Refresh immediately on startup
        _ = RefreshMappingsAsync();
        
        // Then refresh daily at 3 AM
        var now = DateTime.Now;
        var nextRun = now.Date.AddDays(1).AddHours(3);
        var initialDelay = nextRun - now;
        
        _nextScheduledRefresh = nextRun.ToUniversalTime();
        
        _timer = new Timer(
            async _ => 
            {
                await RefreshMappingsAsync();
                _nextScheduledRefresh = DateTime.UtcNow.Date.AddDays(1).AddHours(3);
            },
            null,
            initialDelay,
            TimeSpan.FromDays(1)
        );
        
        _logger.LogInformation("SKU mappings will refresh daily at 3 AM. Next refresh in {Hours:F1} hours", 
            initialDelay.TotalHours);
        
        return Task.CompletedTask;
    }

    public Task StopAsync(CancellationToken cancellationToken)
    {
        _logger.LogInformation("SKU Mapping Service stopping...");
        _timer?.Change(Timeout.Infinite, 0);
        return Task.CompletedTask;
    }

    public void Dispose()
    {
        _timer?.Dispose();
        _refreshLock.Dispose();
    }
}
