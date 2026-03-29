using Microsoft.Graph;
using M365Dashboard.Api.Models;
using M365Dashboard.Api.Services;
using M365Dashboard.Api.Controllers;
using System.Text.Json;

namespace M365Dashboard.Api.Services;

/// <summary>
/// Extracts executive report data gathering and PDF generation so both
/// ExecutiveReportController and ReportService can reuse it.
/// </summary>
public interface IExecutiveReportService
{
    Task<ExecutiveReportData> GatherDataAsync();
    Task<byte[]> GeneratePdfAsync();
}

public class ExecutiveReportService : IExecutiveReportService
{
    private readonly IGraphService _graphService;
    private readonly GraphServiceClient _graphClient;
    private readonly IConfiguration _configuration;
    private readonly IDomainSecurityService _domainSecurityService;
    private readonly IOsVersionService _osVersionService;
    private readonly PdfReportGenerator _pdfReportGenerator;
    private readonly WordReportGenerator _wordReportGenerator;
    private readonly ITenantSettingsService _tenantSettingsService;
    private readonly ILogger<ExecutiveReportService> _logger;

    public ExecutiveReportService(
        IGraphService graphService,
        GraphServiceClient graphClient,
        IConfiguration configuration,
        IDomainSecurityService domainSecurityService,
        IOsVersionService osVersionService,
        PdfReportGenerator pdfReportGenerator,
        WordReportGenerator wordReportGenerator,
        ITenantSettingsService tenantSettingsService,
        ILogger<ExecutiveReportService> logger)
    {
        _graphService = graphService;
        _graphClient = graphClient;
        _configuration = configuration;
        _domainSecurityService = domainSecurityService;
        _osVersionService = osVersionService;
        _pdfReportGenerator = pdfReportGenerator;
        _wordReportGenerator = wordReportGenerator;
        _tenantSettingsService = tenantSettingsService;
        _logger = logger;
    }

    public async Task<ExecutiveReportData> GatherDataAsync()
    {
        var now = DateTime.UtcNow;
        var reportDate = now.AddMonths(-1);
        var startDate = new DateTime(reportDate.Year, reportDate.Month, 1);
        var endDate = startDate.AddMonths(1).AddDays(-1);

        var reportData = new ExecutiveReportData
        {
            ReportMonth = startDate.ToString("MMMM yyyy"),
            GeneratedAt = now,
            StartDate = startDate,
            EndDate = endDate
        };

        var tasks = new List<Task>();

        tasks.Add(Task.Run(async () =>
        {
            try
            {
                var score = await _graphService.GetSecureScoreAsync();
                if (score != null)
                    reportData.SecureScore = new SecureScoreData
                    {
                        CurrentScore = score.CurrentScore,
                        MaxScore = score.MaxScore,
                        PercentageScore = score.MaxScore > 0 ? Math.Round((double)score.CurrentScore / score.MaxScore * 100, 1) : 0,
                    };
            }
            catch (Exception ex) { _logger.LogWarning(ex, "Failed to get Secure Score"); }
        }));

        tasks.Add(Task.Run(async () =>
        {
            try
            {
                var deviceStats = await _graphService.GetDeviceStatsAsync();
                reportData.DeviceStats = new DeviceStatsData
                {
                    TotalDevices = deviceStats.TotalDevices,
                    WindowsDevices = deviceStats.WindowsDevices,
                    MacOsDevices = deviceStats.MacOsDevices,
                    IosDevices = deviceStats.IosDevices,
                    AndroidDevices = deviceStats.AndroidDevices,
                    CompliantDevices = deviceStats.CompliantDevices,
                    NonCompliantDevices = deviceStats.NonCompliantDevices,
                    ComplianceRate = deviceStats.TotalDevices > 0
                        ? Math.Round((double)deviceStats.CompliantDevices / deviceStats.TotalDevices * 100, 1) : 0
                };
            }
            catch (Exception ex) { _logger.LogWarning(ex, "Failed to get Device Stats"); }
        }));

        tasks.Add(Task.Run(async () =>
        {
            try
            {
                var userStats = await _graphService.GetUserStatsAsync();
                reportData.UserStats = new UserStatsData
                {
                    TotalUsers = userStats.TotalUsers,
                    GuestUsers = userStats.GuestUsers,
                    DeletedUsers = userStats.DeletedUsers,
                    MfaRegistered = userStats.MfaRegistered,
                    MfaNotRegistered = userStats.MfaNotRegistered
                };
            }
            catch (Exception ex) { _logger.LogWarning(ex, "Failed to get User Stats"); }
        }));

        tasks.Add(Task.Run(async () =>
        {
            try { reportData.DefenderStats = await GetDefenderStatsAsync(); }
            catch (Exception ex) { _logger.LogWarning(ex, "Failed to get Defender Stats"); }
        }));

        tasks.Add(Task.Run(async () =>
        {
            try
            {
                var riskyUsers = await _graphService.GetRiskyUsersAsync();
                reportData.RiskyUsersCount = riskyUsers?.Count ?? 0;
                reportData.HighRiskUsers = riskyUsers?.Where(u => u.RiskLevel == "high")
                    .Select(u => u.DisplayName ?? u.UserPrincipalName).ToList() ?? new List<string>();
            }
            catch (Exception ex) { _logger.LogWarning(ex, "Failed to get Risky Users"); }
        }));

        tasks.Add(Task.Run(async () =>
        {
            try { reportData.UserSignInDetails = await GetUserSignInDetailsAsync(); }
            catch (Exception ex) { _logger.LogWarning(ex, "Failed to get User Sign-in Details"); }
        }));

        tasks.Add(Task.Run(async () =>
        {
            try { reportData.DeletedUsersInPeriod = await GetDeletedUsersInPeriodAsync(startDate, endDate); }
            catch (Exception ex) { _logger.LogWarning(ex, "Failed to get Deleted Users"); }
        }));

        tasks.Add(Task.Run(async () =>
        {
            try { reportData.MailboxDetails = await GetMailboxDetailsAsync(); }
            catch (Exception ex) { _logger.LogWarning(ex, "Failed to get Mailbox Details"); }
        }));

        tasks.Add(Task.Run(async () =>
        {
            try { reportData.DeviceDetails = await GetDeviceDetailsAsync(); }
            catch (Exception ex) { _logger.LogWarning(ex, "Failed to get Device Details"); }
        }));

        tasks.Add(Task.Run(async () =>
        {
            try { reportData.AppCredentialStatus = await GetAppCredentialStatusAsync(); }
            catch (Exception ex) { _logger.LogWarning(ex, "Failed to get App Credential Status"); }
        }));

        tasks.Add(Task.Run(async () =>
        {
            try { reportData.SignInLocations = await GetSignInLocationsAsync(); }
            catch (Exception ex) { _logger.LogWarning(ex, "Failed to get Sign-in Locations"); }
        }));

        tasks.Add(Task.Run(async () =>
        {
            try
            {
                var tenantDomains = await _graphClient.Domains.GetAsync(config =>
                    config.QueryParameters.Select = new[] { "id", "isVerified", "isDefault" });

                if (tenantDomains?.Value != null)
                {
                    var domainNames = tenantDomains.Value
                        .Where(d => d.IsVerified == true)
                        .Select(d => d.Id)
                        .Where(id => !string.IsNullOrEmpty(id))
                        .ToArray();

                    if (domainNames.Length > 0)
                    {
                        var results = await _domainSecurityService.CheckDomainsAsync(domainNames!);
                        var summary = await _domainSecurityService.GetSecuritySummaryAsync(results);
                        reportData.DomainSecurityResults = results;
                        reportData.DomainSecuritySummary = summary;
                    }
                }
            }
            catch (Exception ex) { _logger.LogWarning(ex, "Failed to get Domain Security data"); }
        }));

        await Task.WhenAll(tasks);
        return reportData;
    }

    public async Task<byte[]> GeneratePdfAsync()
    {
        var tenantId = _configuration["AzureAd:TenantId"] ?? "default";
        var settings = await _tenantSettingsService.GetReportSettingsAsync(tenantId);
        var data = await GatherDataAsync();

        // ── Sign-in locations map (orange pins) ──────────────────────────────
        try { data.SignInMapImageBytes = await GenerateSignInMapAsync(data.SignInLocations ?? new()); }
        catch (Exception ex) { _logger.LogWarning(ex, "Failed to generate sign-in map image"); }

        // ── Failed sign-in locations map (red pins) ──────────────────────────
        try
        {
            var since = DateTime.UtcNow.AddDays(-30).ToString("yyyy-MM-ddTHH:mm:ssZ");
            var failedSignIns = await _graphClient.AuditLogs.SignIns.GetAsync(config =>
            {
                config.QueryParameters.Filter = $"createdDateTime ge {since} and status/errorCode ne 0";
                config.QueryParameters.Select = new[] { "location", "createdDateTime", "status" };
                config.QueryParameters.Top = 1000;
            });

            var byCountry = new Dictionary<string, (int count, double lat, double lon)>(StringComparer.OrdinalIgnoreCase);
            foreach (var s in failedSignIns?.Value ?? new())
            {
                var country = s.Location?.CountryOrRegion;
                if (string.IsNullOrEmpty(country)) continue;
                var lat = s.Location?.GeoCoordinates?.Latitude ?? 20.0;
                var lon = s.Location?.GeoCoordinates?.Longitude ?? 0.0;
                if (byCountry.TryGetValue(country, out var ex))
                    byCountry[country] = (ex.count + 1, ex.lat, ex.lon);
                else
                    byCountry[country] = (1, lat, lon);
            }

            data.FailedSignInLocations = byCountry.Select(kvp => new SignInLocationData
            {
                Country = kvp.Key, CountryCode = kvp.Key,
                Latitude = kvp.Value.lat, Longitude = kvp.Value.lon,
                SignInCount = kvp.Value.count
            }).OrderByDescending(l => l.SignInCount).ToList();

            data.FailedSignInMapImageBytes = await GenerateFailedSignInMapAsync(data.FailedSignInLocations);
        }
        catch (Exception ex) { _logger.LogWarning(ex, "Failed to generate failed sign-in map"); }

        try
        {
            return _pdfReportGenerator.GenerateReport(data, settings);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "PDF generation failed, falling back to Word");
            return _wordReportGenerator.GenerateReport(data, settings);
        }
    }

    // ── Private data helpers (mirrors ExecutiveReportController) ──────────────

    private async Task<DefenderStatsData?> GetDefenderStatsAsync()
    {
        try
        {
            var tenantId = _configuration["AzureAd:TenantId"];
            var clientId = _configuration["AzureAd:ClientId"];
            var clientSecret = _configuration["AzureAd:ClientSecret"];

            var credential = new Azure.Identity.ClientSecretCredential(tenantId, clientId, clientSecret);
            var scopes = new[] { "https://api.securitycenter.microsoft.com/.default" };
            var token = await credential.GetTokenAsync(new Azure.Core.TokenRequestContext(scopes));

            using var httpClient = new HttpClient();
            httpClient.DefaultRequestHeaders.Authorization =
                new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token.Token);

            var result = new DefenderStatsData();

            try
            {
                var resp = await httpClient.GetAsync("https://api.securitycenter.microsoft.com/api/exposureScore");
                if (resp.IsSuccessStatusCode)
                {
                    var doc = JsonDocument.Parse(await resp.Content.ReadAsStringAsync());
                    if (doc.RootElement.TryGetProperty("score", out var el))
                    {
                        var s = el.GetDouble();
                        result.ExposureScore = s switch { <= 30 => "Low", <= 70 => "Medium", _ => "High" };
                        result.ExposureScoreNumeric = Math.Round(s, 1);
                    }
                }
            }
            catch (Exception ex) { _logger.LogWarning(ex, "Error fetching exposure score"); }

            return result;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Could not connect to Defender API");
            return new DefenderStatsData { ExposureScore = "N/A" };
        }
    }

    private async Task<List<UserSignInDetailData>> GetUserSignInDetailsAsync()
    {
        var result = new List<UserSignInDetailData>();
        try
        {
            var users = await _graphService.GetUsersAsync(null, "displayName", true, 999);
            result = users.Users
                .Where(u => !string.Equals(u.UserType, "Guest", StringComparison.OrdinalIgnoreCase))
                .Select(u => new UserSignInDetailData
                {
                    DisplayName = u.DisplayName,
                    UserPrincipalName = u.UserPrincipalName,
                    LastInteractiveSignIn = u.LastSignInDateTime,
                    LastNonInteractiveSignIn = u.LastNonInteractiveSignInDateTime,
                    DefaultMfaMethod = u.DefaultMfaMethod,
                    IsMfaRegistered = u.IsMfaRegistered,
                    AccountEnabled = u.AccountEnabled
                })
                .OrderBy(u => u.DisplayName)
                .ToList();
        }
        catch (Exception ex) { _logger.LogWarning(ex, "Error fetching user sign-in details"); }
        return result;
    }

    private async Task<List<DeletedUserData>> GetDeletedUsersInPeriodAsync(DateTime startDate, DateTime endDate)
    {
        var result = new List<DeletedUserData>();
        try
        {
            var tenantId = _configuration["AzureAd:TenantId"];
            var clientId = _configuration["AzureAd:ClientId"];
            var clientSecret = _configuration["AzureAd:ClientSecret"];
            var credential = new Azure.Identity.ClientSecretCredential(tenantId, clientId, clientSecret);
            var graphClient = new GraphServiceClient(credential, new[] { "https://graph.microsoft.com/.default" });

            var deletedUsers = await graphClient.Directory.DeletedItems.GraphUser.GetAsync(config =>
            {
                config.QueryParameters.Select = new[] { "id", "displayName", "userPrincipalName", "mail", "deletedDateTime", "jobTitle", "department" };
                config.QueryParameters.Top = 999;
            });

            while (deletedUsers?.Value != null)
            {
                foreach (var user in deletedUsers.Value)
                {
                    if (user.DeletedDateTime.HasValue)
                    {
                        var d = user.DeletedDateTime.Value.DateTime;
                        if (d >= startDate && d <= endDate.AddDays(1))
                            result.Add(new DeletedUserData
                            {
                                DisplayName = user.DisplayName,
                                UserPrincipalName = user.UserPrincipalName,
                                Mail = user.Mail,
                                DeletedDateTime = d,
                                JobTitle = user.JobTitle,
                                Department = user.Department
                            });
                    }
                }
                if (deletedUsers.OdataNextLink == null) break;
                deletedUsers = await graphClient.Directory.DeletedItems.GraphUser
                    .WithUrl(deletedUsers.OdataNextLink).GetAsync();
            }

            result = result.OrderByDescending(u => u.DeletedDateTime).ToList();
        }
        catch (Exception ex) { _logger.LogWarning(ex, "Error fetching deleted users"); }
        return result;
    }

    private async Task<List<MailboxDetailData>> GetMailboxDetailsAsync()
    {
        var result = new List<MailboxDetailData>();
        try
        {
            var resp = await _graphClient.Reports.GetMailboxUsageDetailWithPeriod("D30").GetAsync();
            if (resp == null) return result;

            using var reader = new StreamReader(resp);
            var lines = (await reader.ReadToEndAsync()).Split('\n');
            var header = lines.FirstOrDefault()?.Split(',') ?? Array.Empty<string>();

            int Idx(string name) => Array.FindIndex(header, h => h.Trim().Equals(name, StringComparison.OrdinalIgnoreCase));

            var dnIdx      = Idx("Display Name");
            var upnIdx     = Idx("User Principal Name");
            var storIdx    = Idx("Storage Used (Byte)");
            var quotaIdx   = Idx("Prohibit Send Quota (Byte)");
            var itemIdx    = Idx("Item Count");
            var delIdx     = Idx("Is Deleted");

            foreach (var line in lines.Skip(1).Where(l => !string.IsNullOrWhiteSpace(l)))
            {
                var parts = line.Split(',');
                if (delIdx >= 0 && delIdx < parts.Length &&
                    parts[delIdx]?.Trim().Equals("TRUE", StringComparison.OrdinalIgnoreCase) == true) continue;

                long stor = 0; if (storIdx >= 0) long.TryParse(parts.ElementAtOrDefault(storIdx)?.Trim(), out stor);
                long? quota = null; if (quotaIdx >= 0 && long.TryParse(parts.ElementAtOrDefault(quotaIdx)?.Trim(), out var q)) quota = q;
                int? items = null; if (itemIdx >= 0 && int.TryParse(parts.ElementAtOrDefault(itemIdx)?.Trim(), out var ic)) items = ic;

                result.Add(new MailboxDetailData
                {
                    DisplayName = dnIdx >= 0 ? parts.ElementAtOrDefault(dnIdx)?.Trim().Trim('"') : null,
                    UserPrincipalName = upnIdx >= 0 ? parts.ElementAtOrDefault(upnIdx)?.Trim().Trim('"') : null,
                    StorageUsedBytes = stor,
                    StorageUsedGB = Math.Round(stor / 1073741824.0, 2),
                    QuotaBytes = quota,
                    QuotaGB = quota.HasValue ? Math.Round(quota.Value / 1073741824.0, 2) : null,
                    PercentUsed = quota is > 0 ? Math.Round((double)stor / quota.Value * 100, 1) : null,
                    ItemCount = items
                });
            }

            result = result.OrderByDescending(m => m.StorageUsedBytes).ToList();
        }
        catch (Exception ex) { _logger.LogWarning(ex, "Error fetching mailbox details"); }
        return result;
    }

    private async Task<DeviceDetailsData> GetDeviceDetailsAsync()
    {
        var result = new DeviceDetailsData();
        try
        {
            try { await _osVersionService.GetLatestVersionsAsync(); } catch { }

            var devices = await _graphService.GetDevicesAsync(null, "deviceName", true, 1000);
            foreach (var device in devices.Devices)
            {
                var os = device.OperatingSystem?.ToLower() ?? "";
                if (os.Contains("windows"))
                {
                    var vs = _osVersionService.CheckWindowsVersion(device.OsVersion);
                    result.WindowsDevices.Add(new WindowsDeviceDetailData
                    {
                        DeviceName = device.DeviceName, OsVersion = device.OsVersion,
                        ComplianceState = FormatCompliance(device.ComplianceState),
                        LastCheckIn = device.LastSyncDateTime, SkuFamily = device.Model,
                        OsVersionStatus = vs.Status, OsVersionStatusMessage = vs.Message, LatestVersion = vs.LatestVersion
                    });
                }
                else if (os.Contains("ios") || os.Contains("ipad"))
                {
                    var vs = _osVersionService.CheckiOSVersion(device.OsVersion);
                    result.IosDevices.Add(new IosDeviceDetailData
                    {
                        DeviceName = device.DeviceName, OsVersion = device.OsVersion,
                        ComplianceState = FormatCompliance(device.ComplianceState),
                        LastCheckIn = device.LastSyncDateTime,
                        OsVersionStatus = vs.Status, OsVersionStatusMessage = vs.Message, LatestVersion = vs.LatestVersion
                    });
                }
                else if (os.Contains("android"))
                {
                    var vs = _osVersionService.CheckAndroidVersion(device.OsVersion, null);
                    result.AndroidDevices.Add(new AndroidDeviceDetailData
                    {
                        DeviceName = device.DeviceName, OsVersion = device.OsVersion,
                        ComplianceState = FormatCompliance(device.ComplianceState),
                        LastCheckIn = device.LastSyncDateTime,
                        OsVersionStatus = vs.Status, OsVersionStatusMessage = vs.Message, LatestVersion = vs.LatestVersion
                    });
                }
                else if (os.Contains("mac"))
                {
                    var vs = _osVersionService.CheckMacOSVersion(device.OsVersion);
                    result.MacDevices.Add(new MacDeviceDetailData
                    {
                        DeviceName = device.DeviceName, OsVersion = device.OsVersion,
                        ComplianceState = FormatCompliance(device.ComplianceState),
                        LastCheckIn = device.LastSyncDateTime,
                        OsVersionStatus = vs.Status, OsVersionStatusMessage = vs.Message, LatestVersion = vs.LatestVersion
                    });
                }
            }

            result.WindowsDevices = result.WindowsDevices.OrderBy(d => d.DeviceName).ToList();
            result.IosDevices     = result.IosDevices.OrderBy(d => d.DeviceName).ToList();
            result.AndroidDevices = result.AndroidDevices.OrderBy(d => d.DeviceName).ToList();
            result.MacDevices     = result.MacDevices.OrderBy(d => d.DeviceName).ToList();
        }
        catch (Exception ex) { _logger.LogWarning(ex, "Error fetching device details"); }
        return result;
    }

    private async Task<AppCredentialStatusData> GetAppCredentialStatusAsync()
    {
        var result = new AppCredentialStatusData { ThresholdDays = 45 };
        try
        {
            var today = DateTime.UtcNow;
            var threshold = today.AddDays(result.ThresholdDays);
            var expiredSet = new HashSet<string>(); var expiringSet = new HashSet<string>();
            var expiredCertSet = new HashSet<string>(); var expiringCertSet = new HashSet<string>();

            var apps = await _graphClient.Applications.GetAsync(config =>
            {
                config.QueryParameters.Select = new[] { "id", "appId", "displayName", "passwordCredentials", "keyCredentials" };
                config.QueryParameters.Top = 999;
            });

            while (apps?.Value != null)
            {
                foreach (var app in apps.Value)
                {
                    result.TotalApps++;
                    foreach (var s in app.PasswordCredentials ?? new())
                    {
                        if (s.EndDateTime == null) continue;
                        var exp = s.EndDateTime.Value.DateTime;
                        var days = (int)(exp - today).TotalDays;
                        if (exp < today) { expiredSet.Add(app.Id ?? ""); result.ExpiredSecrets.Add(new AppCredentialDetail { AppName = app.DisplayName, AppId = app.AppId, CredentialType = "Secret", ExpiryDate = exp, DaysUntilExpiry = days, Status = "Expired" }); }
                        else if (exp < threshold) { expiringSet.Add(app.Id ?? ""); result.ExpiringSecrets.Add(new AppCredentialDetail { AppName = app.DisplayName, AppId = app.AppId, CredentialType = "Secret", ExpiryDate = exp, DaysUntilExpiry = days, Status = $"Expires in {days} days" }); }
                    }
                    foreach (var k in app.KeyCredentials ?? new())
                    {
                        if (k.EndDateTime == null) continue;
                        var exp = k.EndDateTime.Value.DateTime;
                        var days = (int)(exp - today).TotalDays;
                        if (exp < today) { expiredCertSet.Add(app.Id ?? ""); result.ExpiredCertificates.Add(new AppCredentialDetail { AppName = app.DisplayName, AppId = app.AppId, CredentialType = "Certificate", ExpiryDate = exp, DaysUntilExpiry = days, Status = "Expired" }); }
                        else if (exp < threshold) { expiringCertSet.Add(app.Id ?? ""); result.ExpiringCertificates.Add(new AppCredentialDetail { AppName = app.DisplayName, AppId = app.AppId, CredentialType = "Certificate", ExpiryDate = exp, DaysUntilExpiry = days, Status = $"Expires in {days} days" }); }
                    }
                }
                if (apps.OdataNextLink == null) break;
                apps = await _graphClient.Applications.WithUrl(apps.OdataNextLink).GetAsync();
            }

            result.AppsWithExpiredSecrets      = expiredSet.Count;
            result.AppsWithExpiringSecrets     = expiringSet.Count;
            result.AppsWithExpiredCertificates = expiredCertSet.Count;
            result.AppsWithExpiringCertificates = expiringCertSet.Count;
        }
        catch (Exception ex) { _logger.LogWarning(ex, "Error fetching app credential status"); }
        return result;
    }

    private async Task<List<SignInLocationData>> GetSignInLocationsAsync()
    {
        var result = new List<SignInLocationData>();
        try
        {
            var since = DateTime.UtcNow.AddDays(-30).ToString("yyyy-MM-ddTHH:mm:ssZ");

            _logger.LogInformation("Fetching sign-in locations since {Since}", since);

            var signIns = await _graphClient.AuditLogs.SignIns.GetAsync(config =>
            {
                config.QueryParameters.Filter = $"createdDateTime ge {since}";
                config.QueryParameters.Select = new[] { "location", "createdDateTime" };
                config.QueryParameters.Top = 1000;
            });

            _logger.LogInformation("Sign-ins returned: {Count}", signIns?.Value?.Count ?? 0);

            // Aggregate by country — use centroid coordinates per country
            var byCountry = new Dictionary<string, (int count, double lat, double lon, string code)>(StringComparer.OrdinalIgnoreCase);

            foreach (var signIn in signIns?.Value ?? new List<Microsoft.Graph.Models.SignIn>())
            {
                var country = signIn.Location?.CountryOrRegion;
                if (string.IsNullOrEmpty(country)) continue;

                // Use Graph-provided coordinates if available, otherwise use country centroid
                var lat = signIn.Location?.GeoCoordinates?.Latitude ?? GetCountryCentroid(country).lat;
                var lon = signIn.Location?.GeoCoordinates?.Longitude ?? GetCountryCentroid(country).lon;
                var code = signIn.Location?.CountryOrRegion ?? "";

                if (byCountry.TryGetValue(country, out var existing))
                    byCountry[country] = (existing.count + 1, existing.lat, existing.lon, existing.code);
                else
                    byCountry[country] = (1, lat, lon, code);
            }

            result = byCountry.Select(kvp => new SignInLocationData
            {
                Country     = kvp.Key,
                CountryCode = kvp.Value.code,
                Latitude    = kvp.Value.lat,
                Longitude   = kvp.Value.lon,
                SignInCount = kvp.Value.count
            }).OrderByDescending(l => l.SignInCount).ToList();

            _logger.LogInformation("Sign-in locations aggregated: {Count} countries", result.Count);
        }
        catch (Exception ex) { _logger.LogWarning(ex, "Error fetching sign-in locations"); }
        return result;
    }

    private static (double lat, double lon) GetCountryCentroid(string country) => country.ToUpperInvariant() switch
    {
        "UNITED KINGDOM" or "UK" or "GB"         => (55.3781, -3.4360),
        "UNITED STATES" or "USA" or "US"         => (37.0902, -95.7129),
        "GERMANY" or "DE"                        => (51.1657,  10.4515),
        "FRANCE" or "FR"                         => (46.2276,   2.2137),
        "NETHERLANDS" or "NL"                    => (52.1326,   5.2913),
        "IRELAND" or "IE"                        => (53.4129,  -8.2439),
        "AUSTRALIA" or "AU"                      => (-25.2744, 133.7751),
        "CANADA" or "CA"                         => (56.1304, -106.3468),
        "INDIA" or "IN"                          => (20.5937,  78.9629),
        "JAPAN" or "JP"                          => (36.2048, 138.2529),
        "BRAZIL" or "BR"                         => (-14.2350, -51.9253),
        "SOUTH AFRICA" or "ZA"                   => (-30.5595,  22.9375),
        "CHINA" or "CN"                          => (35.8617, 104.1954),
        "SPAIN" or "ES"                          => (40.4637,  -3.7492),
        "ITALY" or "IT"                          => (41.8719,  12.5674),
        "POLAND" or "PL"                         => (51.9194,  19.1451),
        "SWEDEN" or "SE"                         => (60.1282,  18.6435),
        "NORWAY" or "NO"                         => (60.4720,   8.4689),
        "DENMARK" or "DK"                        => (56.2639,   9.5018),
        "SWITZERLAND" or "CH"                    => (46.8182,   8.2275),
        "BELGIUM" or "BE"                        => (50.5039,   4.4699),
        "PORTUGAL" or "PT"                       => (39.3999,  -8.2245),
        "NEW ZEALAND" or "NZ"                    => (-40.9006, 174.8860),
        "SINGAPORE" or "SG"                      => (1.3521,  103.8198),
        "UAE" or "UNITED ARAB EMIRATES" or "AE"  => (23.4241,  53.8478),
        _                                        => (20.0, 0.0)  // Default to mid-Atlantic
    };

    private async Task<byte[]?> GenerateSignInMapAsync(List<SignInLocationData> locations)
    {
        var mapsKey = _configuration["AzureMaps:SubscriptionKey"];
        _logger.LogInformation("AzureMaps key present: {HasKey}, locations: {Count}",
            !string.IsNullOrEmpty(mapsKey), locations.Count);

        if (string.IsNullOrEmpty(mapsKey)) return null;
        // Continue even with no locations — we still render a plain world map

        try
        {
            var validPins = locations
                .Where(l => l.Latitude != 0 || l.Longitude != 0)
                .Take(50)
                .ToList();

            // Azure Maps Render v2 static image API
            // Correct pin format: default||lon lat||  (space between lon/lat, double pipe before AND after coords)
            // All pins for a given style go in ONE &pins= parameter, coords separated by |
            // e.g. &pins=default|coFF6600||0.0 51.5||-3.4 55.3||
            var inv = System.Globalization.CultureInfo.InvariantCulture;

            // Build the URL string manually — do NOT let HttpClient encode pipes/spaces
            var url = "https://atlas.microsoft.com/map/static" +
                      "?api-version=2024-04-01" +
                      $"&subscription-key={Uri.EscapeDataString(mapsKey)}" +
                      "&zoom=1" +
                      "&width=800" +
                      "&height=400" +
                      "&tilesetId=microsoft.base.road" +
                      "&center=0,20";

            // Add pins only if we have locations with valid coordinates
            if (validPins.Any())
            {
                // Format: default|coFF6600||lon1 lat1|lon2 lat2||  (trailing || required)
                var coords = string.Join("|", validPins.Select(l =>
                    $"{l.Longitude.ToString("F4", inv)} {l.Latitude.ToString("F4", inv)}"));
                url += $"&pins=default|coFF6600||{coords}||";
            }

            _logger.LogInformation("Calling Azure Maps static image (pins: {Count})", validPins.Count);

            using var http = new HttpClient { Timeout = TimeSpan.FromSeconds(20) };
            var response = await http.GetAsync(url);

            if (response.IsSuccessStatusCode)
            {
                var bytes = await response.Content.ReadAsByteArrayAsync();
                _logger.LogInformation("Azure Maps returned {Bytes} bytes", bytes.Length);
                return bytes;
            }

            var errorBody = await response.Content.ReadAsStringAsync();
            _logger.LogWarning("Azure Maps static image returned {Status}: {Body}", response.StatusCode, errorBody);
            return null;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Error generating sign-in map image");
            return null;
        }
    }

    private async Task<byte[]?> GenerateFailedSignInMapAsync(List<SignInLocationData> locations)
    {
        var mapsKey = _configuration["AzureMaps:SubscriptionKey"];
        if (string.IsNullOrEmpty(mapsKey)) return null;

        try
        {
            var validPins = locations.Where(l => l.Latitude != 0 || l.Longitude != 0).Take(50).ToList();
            var inv = System.Globalization.CultureInfo.InvariantCulture;

            var url = "https://atlas.microsoft.com/map/static" +
                      "?api-version=2024-04-01" +
                      $"&subscription-key={Uri.EscapeDataString(mapsKey)}" +
                      "&zoom=1&width=800&height=400" +
                      "&tilesetId=microsoft.base.road&center=0,20";

            if (validPins.Any())
            {
                var coords = string.Join("|", validPins.Select(l =>
                    $"{l.Longitude.ToString("F4", inv)} {l.Latitude.ToString("F4", inv)}"));
                url += $"&pins=default|coFF0000||{coords}||";
            }

            using var http = new HttpClient { Timeout = TimeSpan.FromSeconds(20) };
            var response = await http.GetAsync(url);
            if (response.IsSuccessStatusCode)
                return await response.Content.ReadAsByteArrayAsync();

            _logger.LogWarning("Azure Maps (failed sign-ins) returned {Status}", response.StatusCode);
            return null;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Error generating failed sign-in map image");
            return null;
        }
    }

    private static string FormatCompliance(string? state) => state?.ToLower() switch
    {
        "compliant"      => "Compliant",
        "noncompliant"   => "Non-Compliant",
        "conflict"       => "Conflict",
        "error"          => "Error",
        "ingraceperiod"  => "In Grace Period",
        _                => state ?? "Unknown"
    };
}
