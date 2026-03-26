using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Graph;
using M365Dashboard.Api.Services;
using M365Dashboard.Api.Models;
using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace M365Dashboard.Api.Controllers;

[ApiController]
[Route("api/[controller]")]
[Authorize]
public class ExecutiveReportController : ControllerBase
{
    private readonly IGraphService _graphService;
    private readonly GraphServiceClient _graphClient;
    private readonly IConfiguration _configuration;
    private readonly ILogger<ExecutiveReportController> _logger;
    private readonly IDomainSecurityService _domainSecurityService;
    private readonly IWebHostEnvironment _environment;
    private readonly IOsVersionService _osVersionService;
    private readonly PdfReportGenerator _pdfReportGenerator;
    private readonly ITenantSettingsService _tenantSettingsService;
    private readonly WordReportGenerator _wordReportGenerator;

    public ExecutiveReportController(
        IGraphService graphService, 
        GraphServiceClient graphClient,
        IConfiguration configuration,
        ILogger<ExecutiveReportController> logger,
        IDomainSecurityService domainSecurityService,
        IWebHostEnvironment environment,
        IOsVersionService osVersionService,
        ITenantSettingsService tenantSettingsService,
        WordReportGenerator wordReportGenerator,
        PdfReportGenerator pdfReportGenerator)
    {
        _graphService = graphService;
        _graphClient = graphClient;
        _configuration = configuration;
        _logger = logger;
        _domainSecurityService = domainSecurityService;
        _environment = environment;
        _osVersionService = osVersionService;
        _tenantSettingsService = tenantSettingsService;
        _wordReportGenerator = wordReportGenerator;
        _pdfReportGenerator = pdfReportGenerator;
    }

    private async Task<ReportSettings> LoadReportSettingsAsync()
    {
        try
        {
            var tenantId = _configuration["AzureAd:TenantId"] ?? "default";
            return await _tenantSettingsService.GetReportSettingsAsync(tenantId);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Could not load report settings, using defaults");
            return new ReportSettings();
        }
    }

    // Convert "#rrggbb" to "RRGGBB" (no hash, uppercase) for OpenXml Color elements
    private static string ToOxmlColor(string hex) =>
        hex.TrimStart('#').ToUpperInvariant();

    /// <summary>
    /// Test endpoint to check domain security configuration and run a quick test
    /// </summary>
    [HttpGet("domain-security-test")]
    public async Task<IActionResult> TestDomainSecurity()
    {
        try
        {
            // Get domains from the connected Microsoft 365 tenant via Graph API
            var tenantDomains = await _graphClient.Domains.GetAsync(config =>
            {
                config.QueryParameters.Select = new[] { "id", "isVerified", "isDefault" };
            });
            
            var allDomains = tenantDomains?.Value?.Select(d => new { d.Id, d.IsVerified, d.IsDefault }).ToList();
            var verifiedDomains = tenantDomains?.Value?
                .Where(d => d.IsVerified == true)
                .Select(d => d.Id)
                .Where(id => !string.IsNullOrEmpty(id))
                .ToArray() ?? Array.Empty<string>();
            
            if (verifiedDomains.Length > 0)
            {
                // Test with first domain only
                _logger.LogInformation("Testing domain security with {Domain}", verifiedDomains[0]);
                var testDomain = await _domainSecurityService.CheckDomainAsync(verifiedDomains[0]!);
                return Ok(new
                {
                    Source = "Microsoft Graph API - Tenant Domains",
                    TotalDomainsInTenant = allDomains?.Count ?? 0,
                    VerifiedDomains = verifiedDomains.Length,
                    AllDomains = allDomains,
                    TestResult = testDomain
                });
            }
            
            return Ok(new
            {
                Source = "Microsoft Graph API - Tenant Domains",
                TotalDomainsInTenant = allDomains?.Count ?? 0,
                VerifiedDomains = 0,
                AllDomains = allDomains,
                Message = "No verified domains found in tenant"
            });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error testing domain security");
            return Ok(new
            {
                Error = ex.Message,
                StackTrace = ex.StackTrace
            });
        }
    }

    /// <summary>
    /// Get executive summary report data
    /// </summary>
    [HttpGet("data")]
    public async Task<IActionResult> GetReportData([FromQuery] int month = 0, [FromQuery] int year = 0)
    {
        try
        {
            // Default to previous month if not specified
            var reportDate = month == 0 || year == 0 
                ? DateTime.UtcNow.AddMonths(-1) 
                : new DateTime(year, month, 1);
            
            var startDate = new DateTime(reportDate.Year, reportDate.Month, 1);
            var endDate = startDate.AddMonths(1).AddDays(-1);
            var reportMonth = startDate.ToString("MMMM yyyy");

            _logger.LogInformation("Generating executive report for {Month}", reportMonth);

            var reportData = new ExecutiveReportData
            {
                ReportMonth = reportMonth,
                GeneratedAt = DateTime.UtcNow,
                StartDate = startDate,
                EndDate = endDate
            };

            // Gather all data in parallel where possible
            var tasks = new List<Task>();

            // 1. Microsoft Secure Score
            tasks.Add(Task.Run(async () =>
            {
                try
                {
                    var score = await _graphService.GetSecureScoreAsync();
                    if (score != null)
                    {
                        reportData.SecureScore = new SecureScoreData
                        {
                            CurrentScore = score.CurrentScore,
                            MaxScore = score.MaxScore,
                            PercentageScore = score.MaxScore > 0 ? Math.Round((double)score.CurrentScore / score.MaxScore * 100, 1) : 0
                        };
                    }
                }
                catch (Exception ex) { _logger.LogWarning(ex, "Failed to get Secure Score"); }
            }));

            // 2. Identity Secure Score - removed as it doesn't have a public API

            // 3. Device Stats
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

            // 4. User Stats
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

            // 5. Vulnerability/Exposure Score (Defender)
            tasks.Add(Task.Run(async () =>
            {
                try
                {
                    reportData.DefenderStats = await GetDefenderStatsAsync();
                }
                catch (Exception ex) { _logger.LogWarning(ex, "Failed to get Defender Stats"); }
            }));

            // 6. Mailbox Usage
            tasks.Add(Task.Run(async () =>
            {
                try
                {
                    var mailboxStats = await _graphService.GetMailboxStatsAsync();
                    reportData.MailboxStats = new MailboxStatsData
                    {
                        TotalMailboxes = mailboxStats.TotalMailboxes,
                        ActiveMailboxes = mailboxStats.ActiveMailboxes,
                        TotalStorageUsedGB = Math.Round(mailboxStats.TotalStorageUsedBytes / 1073741824.0, 2),
                        AverageStorageGB = mailboxStats.TotalMailboxes > 0 
                            ? Math.Round(mailboxStats.TotalStorageUsedBytes / 1073741824.0 / mailboxStats.TotalMailboxes, 2) 
                            : 0
                    };
                }
                catch (Exception ex) { _logger.LogWarning(ex, "Failed to get Mailbox Stats"); }
            }));

            // 7. SharePoint Usage
            tasks.Add(Task.Run(async () =>
            {
                try
                {
                    reportData.SharePointStats = await GetSharePointStatsAsync();
                }
                catch (Exception ex) { _logger.LogWarning(ex, "Failed to get SharePoint Stats"); }
            }));

            // 8. Attack Simulation (if available)
            tasks.Add(Task.Run(async () =>
            {
                try
                {
                    reportData.AttackSimulation = await GetAttackSimulationStatsAsync();
                }
                catch (Exception ex) { _logger.LogWarning(ex, "Failed to get Attack Simulation Stats"); }
            }));

            // 9. Email Security (Threat Protection Stats from Defender for Office 365)
            tasks.Add(Task.Run(async () =>
            {
                try
                {
                    reportData.EmailSecurity = await GetEmailSecurityStatsAsync();
                }
                catch (Exception ex) { _logger.LogWarning(ex, "Failed to get Email Security Stats"); }
            }));

            // 10. Risky Users
            tasks.Add(Task.Run(async () =>
            {
                try
                {
                    var riskyUsers = await _graphService.GetRiskyUsersAsync();
                    reportData.RiskyUsersCount = riskyUsers?.Count ?? 0;
                    reportData.HighRiskUsers = riskyUsers?.Where(u => u.RiskLevel == "high").Select(u => u.DisplayName ?? u.UserPrincipalName).ToList() ?? new List<string>();
                }
                catch (Exception ex) { _logger.LogWarning(ex, "Failed to get Risky Users"); }
            }));

            // 11. (Windows Update Compliance removed - was using Intune compliance as proxy, not actual patch data)

            // 12. Cloud App Discovery (Shadow IT)
            tasks.Add(Task.Run(async () =>
            {
                try
                {
                    reportData.CloudAppDiscovery = await GetCloudAppDiscoveryAsync();
                }
                catch (Exception ex) { _logger.LogWarning(ex, "Failed to get Cloud App Discovery"); }
            }));

            // 13. User Sign-in and MFA Details
            tasks.Add(Task.Run(async () =>
            {
                try
                {
                    reportData.UserSignInDetails = await GetUserSignInDetailsAsync();
                }
                catch (Exception ex) { _logger.LogWarning(ex, "Failed to get User Sign-in Details"); }
            }));

            // 14. Deleted Users in Report Period
            tasks.Add(Task.Run(async () =>
            {
                try
                {
                    reportData.DeletedUsersInPeriod = await GetDeletedUsersInPeriodAsync(startDate, endDate);
                }
                catch (Exception ex) { _logger.LogWarning(ex, "Failed to get Deleted Users"); }
            }));

            // 15. Mailbox Details with Sizes
            tasks.Add(Task.Run(async () =>
            {
                try
                {
                    reportData.MailboxDetails = await GetMailboxDetailsAsync();
                }
                catch (Exception ex) { _logger.LogWarning(ex, "Failed to get Mailbox Details"); }
            }));

            // 16. Device Details by Platform
            tasks.Add(Task.Run(async () =>
            {
                try
                {
                    reportData.DeviceDetails = await GetDeviceDetailsAsync();
                }
                catch (Exception ex) { _logger.LogWarning(ex, "Failed to get Device Details"); }
            }));

            // 17. App Registration Secrets & Certificates Expiry
            tasks.Add(Task.Run(async () =>
            {
                try
                {
                    reportData.AppCredentialStatus = await GetAppCredentialStatusAsync();
                }
                catch (Exception ex) { _logger.LogWarning(ex, "Failed to get App Credential Status"); }
            }));

            // 18. Domain Email Security Check - fetches domains from connected tenant
            tasks.Add(Task.Run(async () =>
            {
                try
                {
                    // Get domains from the connected Microsoft 365 tenant via Graph API
                    var tenantDomains = await _graphClient.Domains.GetAsync(config =>
                    {
                        config.QueryParameters.Select = new[] { "id", "isVerified", "isDefault" };
                    });
                    
                    if (tenantDomains?.Value != null && tenantDomains.Value.Any())
                    {
                        // Get all verified domain names
                        var domainNames = tenantDomains.Value
                            .Where(d => d.IsVerified == true)
                            .Select(d => d.Id)
                            .Where(id => !string.IsNullOrEmpty(id))
                            .ToArray();
                        
                        if (domainNames.Length > 0)
                        {
                            _logger.LogInformation("Checking email security for {Count} tenant domains", domainNames.Length);
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

            return Ok(reportData);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error generating executive report data");
            return StatusCode(500, new { error = "Failed to generate report data", message = ex.Message });
        }
    }

    /// <summary>
    /// Generate and download the executive summary report as a Word document
    /// </summary>
    [HttpGet("download")]
    public async Task<IActionResult> DownloadReport([FromQuery] int month = 0, [FromQuery] int year = 0)
    {
        try
        {
            var reportSettings = await LoadReportSettingsAsync();

            // Get report data
            var dataResult = await GetReportData(month, year);
            if (dataResult is not OkObjectResult okResult || okResult.Value is not ExecutiveReportData reportData)
            {
                return BadRequest(new { error = "Failed to generate report data" });
            }

            try
            {
                var pdfBytes = _pdfReportGenerator.GenerateReport(reportData, reportSettings);
                var pdfFileName = $"{reportSettings.CompanyName.Replace(" ", "_")}_M365_Report_{reportData.ReportMonth.Replace(" ", "_")}_{reportData.GeneratedAt:yyyy-MM-dd}.pdf";
                return File(pdfBytes, "application/pdf", pdfFileName);
            }
            catch (Exception pdfEx)
            {
                _logger.LogError(pdfEx, "PDF generation failed: {Message}", pdfEx.Message);
                // Fall back to Word if PDF fails
                var documentBytes = _wordReportGenerator.GenerateReport(reportData, reportSettings);
                var fileName = $"{reportSettings.CompanyName.Replace(" ", "_")}_M365_Report_{reportData.ReportMonth.Replace(" ", "_")}_{reportData.GeneratedAt:yyyy-MM-dd}.docx";
                return File(documentBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", fileName);
            }
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error generating executive report document");
            return StatusCode(500, new { error = "Failed to generate report document", message = ex.Message });
        }
    }

    /// <summary>
    /// View the executive summary report as HTML in browser
    /// </summary>
    [HttpGet("html")]
    public async Task<IActionResult> ViewHtmlReport([FromQuery] int month = 0, [FromQuery] int year = 0)
    {
        try
        {
            var reportSettings = await LoadReportSettingsAsync();

            // Get report data
            var dataResult = await GetReportData(month, year);
            if (dataResult is not OkObjectResult okResult || okResult.Value is not ExecutiveReportData reportData)
            {
                return BadRequest(new { error = "Failed to generate report data" });
            }

            // Generate HTML report
            var html = GenerateHtmlReport(reportData, reportSettings);

            
            return Content(html, "text/html");
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error generating HTML report");
            return StatusCode(500, new { error = "Failed to generate HTML report", message = ex.Message });
        }
    }

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
            httpClient.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token.Token);
            
            var result = new DefenderStatsData();
            
            // Get Exposure Score
            try
            {
                var exposureResponse = await httpClient.GetAsync("https://api.securitycenter.microsoft.com/api/exposureScore");
                if (exposureResponse.IsSuccessStatusCode)
                {
                    var exposureJson = await exposureResponse.Content.ReadAsStringAsync();
                    var exposureDoc = JsonDocument.Parse(exposureJson);
                    
                    if (exposureDoc.RootElement.TryGetProperty("score", out var scoreElement))
                    {
                        var score = scoreElement.GetDouble();
                        // Convert numeric score to descriptive level
                        result.ExposureScore = score switch
                        {
                            <= 30 => "Low",
                            <= 70 => "Medium",
                            _ => "High"
                        };
                        result.ExposureScoreNumeric = Math.Round(score, 1);
                    }
                    
                    _logger.LogInformation("Retrieved Defender exposure score: {Score}", result.ExposureScore);
                }
                else
                {
                    _logger.LogWarning("Failed to get exposure score: {Status}", exposureResponse.StatusCode);
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "Error fetching exposure score");
            }
            
            // Get Vulnerabilities
            try
            {
                var vulnResponse = await httpClient.GetAsync("https://api.securitycenter.microsoft.com/api/vulnerabilities");
                if (vulnResponse.IsSuccessStatusCode)
                {
                    var vulnJson = await vulnResponse.Content.ReadAsStringAsync();
                    var vulnDoc = JsonDocument.Parse(vulnJson);
                    
                    if (vulnDoc.RootElement.TryGetProperty("value", out var vulnerabilities))
                    {
                        var vulnList = vulnerabilities.EnumerateArray().ToList();
                        result.VulnerabilitiesDetected = vulnList.Count;
                        
                        // Count by severity
                        int critical = 0, high = 0, medium = 0, low = 0;
                        
                        foreach (var vuln in vulnList)
                        {
                            if (vuln.TryGetProperty("severity", out var severity))
                            {
                                var severityStr = severity.GetString()?.ToLower();
                                switch (severityStr)
                                {
                                    case "critical": critical++; break;
                                    case "high": high++; break;
                                    case "medium": medium++; break;
                                    case "low": low++; break;
                                }
                            }
                        }
                        
                        result.CriticalVulnerabilities = critical;
                        result.HighVulnerabilities = high;
                        result.MediumVulnerabilities = medium;
                        result.LowVulnerabilities = low;
                        
                        _logger.LogInformation("Retrieved {Total} vulnerabilities: Critical={Critical}, High={High}, Medium={Medium}, Low={Low}",
                            vulnList.Count, critical, high, medium, low);
                    }
                }
                else
                {
                    _logger.LogWarning("Failed to get vulnerabilities: {Status}", vulnResponse.StatusCode);
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "Error fetching vulnerabilities");
            }
            

            
            return result;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Could not connect to Defender for Endpoint API");
            
            // Fall back to basic security stats from Graph
            try
            {
                var securityStats = await _graphService.GetSecurityStatsAsync();
                return new DefenderStatsData
                {
                    ExposureScore = "N/A",
                    VulnerabilitiesDetected = 0,
                    CriticalVulnerabilities = 0,
                    Note = "Defender for Endpoint API not accessible. Grant WindowsDefenderATP permissions in Azure AD."
                };
            }
            catch
            {
                return null;
            }
        }
    }

    private async Task<SharePointStatsData?> GetSharePointStatsAsync()
    {
        try
        {
            // Get SharePoint site usage from reports
            var period = "D7";
            var sites = await _graphClient.Reports.GetSharePointSiteUsageDetailWithPeriod(period).GetAsync();
            
            if (sites == null) return null;

            using var reader = new StreamReader(sites);
            var csv = await reader.ReadToEndAsync();
            var lines = csv.Split('\n').Where(l => !string.IsNullOrWhiteSpace(l)).ToList();
            
            if (lines.Count == 0) return null;

            // Parse header to find column indices
            var header = lines[0].Split(',');
            
            var storageUsedIdx = Array.FindIndex(header, h => 
                h.Trim().Equals("Storage Used (Byte)", StringComparison.OrdinalIgnoreCase));
            var lastActivityIdx = Array.FindIndex(header, h => 
                h.Trim().Equals("Last Activity Date", StringComparison.OrdinalIgnoreCase));
            var isDeletedIdx = Array.FindIndex(header, h => 
                h.Trim().Equals("Is Deleted", StringComparison.OrdinalIgnoreCase));

            long totalStorageUsed = 0;
            int totalSites = 0;
            int activeSites = 0;

            foreach (var line in lines.Skip(1))
            {
                var columns = line.Split(',');
                
                // Skip deleted sites
                if (isDeletedIdx >= 0 && isDeletedIdx < columns.Length &&
                    columns[isDeletedIdx]?.Trim().Equals("TRUE", StringComparison.OrdinalIgnoreCase) == true)
                    continue;
                
                totalSites++;
                
                // Get storage used
                if (storageUsedIdx >= 0 && storageUsedIdx < columns.Length)
                {
                    var storageStr = columns[storageUsedIdx]?.Trim();
                    if (long.TryParse(storageStr, out var storage))
                    {
                        totalStorageUsed += storage;
                    }
                }
                
                // Check if site is active (has last activity date)
                if (lastActivityIdx >= 0 && lastActivityIdx < columns.Length &&
                    !string.IsNullOrWhiteSpace(columns[lastActivityIdx]))
                {
                    activeSites++;
                }
            }

            var storageUsedGB = Math.Round(totalStorageUsed / 1073741824.0, 2);

            _logger.LogInformation("SharePoint stats: {Total} sites, {Active} active, {Used} GB used", 
                totalSites, activeSites, storageUsedGB);

            return new SharePointStatsData
            {
                TotalSites = totalSites,
                ActiveSites = activeSites,
                TotalStorageUsedGB = storageUsedGB
            };
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Error fetching SharePoint stats");
            return null;
        }
    }

    private async Task<AttackSimulationData?> GetAttackSimulationStatsAsync()
    {
        try
        {
            var tenantId = _configuration["AzureAd:TenantId"];
            var clientId = _configuration["AzureAd:ClientId"];
            var clientSecret = _configuration["AzureAd:ClientSecret"];
            
            var credential = new Azure.Identity.ClientSecretCredential(tenantId, clientId, clientSecret);
            var betaClient = new GraphServiceClient(credential, new[] { "https://graph.microsoft.com/.default" }, "https://graph.microsoft.com/beta");
            
            var requestInfo = new Microsoft.Kiota.Abstractions.RequestInformation
            {
                HttpMethod = Microsoft.Kiota.Abstractions.Method.GET,
                URI = new Uri("https://graph.microsoft.com/beta/security/attackSimulation/simulations")
            };
            
            var response = await betaClient.RequestAdapter.SendPrimitiveAsync<Stream>(requestInfo);
            
            if (response == null) return null;

            using var reader = new StreamReader(response);
            var json = await reader.ReadToEndAsync();
            var jsonDoc = JsonDocument.Parse(json);

            int totalSimulations = 0;
            int completedSimulations = 0;
            double avgCompromiseRate = 0;

            if (jsonDoc.RootElement.TryGetProperty("value", out var simulations))
            {
                var simulationList = simulations.EnumerateArray().ToList();
                totalSimulations = simulationList.Count;
                
                foreach (var sim in simulationList)
                {
                    if (sim.TryGetProperty("status", out var status) && status.GetString() == "completed")
                    {
                        completedSimulations++;
                        if (sim.TryGetProperty("report", out var report) && 
                            report.TryGetProperty("simulationEventsContent", out var events) &&
                            events.TryGetProperty("compromisedRate", out var rate))
                        {
                            avgCompromiseRate += rate.GetDouble();
                        }
                    }
                }

                if (completedSimulations > 0)
                    avgCompromiseRate /= completedSimulations;
            }

            return new AttackSimulationData
            {
                TotalSimulations = totalSimulations,
                CompletedSimulations = completedSimulations,
                AverageCompromiseRate = Math.Round(avgCompromiseRate * 100, 1)
            };
        }
        catch (Exception ex)
        {
            _logger.LogDebug(ex, "Attack simulation data not available");
            return new AttackSimulationData
            {
                Note = "Attack Simulation Training data requires Microsoft Defender for Office 365 Plan 2"
            };
        }
    }

    private async Task<WindowsUpdateStatsData?> GetWindowsUpdateStatsAsync()
    {
        try
        {
            var tenantId = _configuration["AzureAd:TenantId"];
            var clientId = _configuration["AzureAd:ClientId"];
            var clientSecret = _configuration["AzureAd:ClientSecret"];
            
            var credential = new Azure.Identity.ClientSecretCredential(tenantId, clientId, clientSecret);
            var betaClient = new GraphServiceClient(credential, new[] { "https://graph.microsoft.com/.default" }, "https://graph.microsoft.com/beta");

            // Try to get Windows Update for Business reports
            var requestInfo = new Microsoft.Kiota.Abstractions.RequestInformation
            {
                HttpMethod = Microsoft.Kiota.Abstractions.Method.GET,
                URI = new Uri("https://graph.microsoft.com/beta/deviceManagement/reports/cachedReportConfigurations")
            };

            // Get device stats for Windows devices
            var deviceStats = await _graphService.GetDeviceStatsAsync();
            var windowsCount = deviceStats.WindowsDevices;

            // Try to get software update status from Intune
            var softwareUpdateRequest = new Microsoft.Kiota.Abstractions.RequestInformation
            {
                HttpMethod = Microsoft.Kiota.Abstractions.Method.GET,
                URI = new Uri("https://graph.microsoft.com/beta/deviceManagement/deviceCompliancePolicySettingStateSummaries")
            };

            try
            {
                var response = await betaClient.RequestAdapter.SendPrimitiveAsync<Stream>(softwareUpdateRequest);
                
                if (response != null)
                {
                    using var reader = new StreamReader(response);
                    var json = await reader.ReadToEndAsync();
                    var jsonDoc = JsonDocument.Parse(json);
                    
                    // Look for Windows Update related policies
                    int compliant = 0;
                    int nonCompliant = 0;
                    int unknown = 0;
                    
                    if (jsonDoc.RootElement.TryGetProperty("value", out var summaries) &&
                        summaries.ValueKind == JsonValueKind.Array)
                    {
                        foreach (var summary in summaries.EnumerateArray())
                        {
                            var settingName = summary.TryGetProperty("settingName", out var name) ? name.GetString() : "";
                            
                            // Look for update-related settings
                            if (settingName?.Contains("Update", StringComparison.OrdinalIgnoreCase) == true ||
                                settingName?.Contains("Patch", StringComparison.OrdinalIgnoreCase) == true)
                            {
                                if (summary.TryGetProperty("compliantDeviceCount", out var comp))
                                    compliant += comp.GetInt32();
                                if (summary.TryGetProperty("nonCompliantDeviceCount", out var nonComp))
                                    nonCompliant += nonComp.GetInt32();
                                if (summary.TryGetProperty("unknownDeviceCount", out var unk))
                                    unknown += unk.GetInt32();
                            }
                        }
                    }
                    
                    if (compliant > 0 || nonCompliant > 0)
                    {
                        var total = compliant + nonCompliant;
                        return new WindowsUpdateStatsData
                        {
                            TotalWindowsDevices = windowsCount,
                            UpToDate = compliant,
                            NeedsUpdate = nonCompliant,
                            ComplianceRate = total > 0 ? Math.Round((double)compliant / total * 100, 1) : 0
                        };
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogDebug(ex, "Could not fetch compliance policy summaries");
            }

            // Fallback: Get Windows devices and check compliance state
            var devices = await _graphService.GetDevicesAsync(null, "deviceName", true, 1000);
            
            var windowsDevices = devices.Devices.Where(d => 
                d.OperatingSystem?.Contains("Windows", StringComparison.OrdinalIgnoreCase) == true).ToList();

            var compliantDevices = windowsDevices.Count(d => 
                d.ComplianceState?.Equals("compliant", StringComparison.OrdinalIgnoreCase) == true);
            var nonCompliantDevices = windowsDevices.Count(d => 
                d.ComplianceState?.Equals("noncompliant", StringComparison.OrdinalIgnoreCase) == true);

            return new WindowsUpdateStatsData
            {
                TotalWindowsDevices = windowsDevices.Count,
                UpToDate = compliantDevices,
                NeedsUpdate = nonCompliantDevices,
                ComplianceRate = windowsDevices.Count > 0 
                    ? Math.Round((double)compliantDevices / windowsDevices.Count * 100, 1) : 0,
                Note = "Based on Intune device compliance. For detailed patch status, enable Windows Update for Business reports."
            };
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Error fetching Windows Update stats");
            return null;
        }
    }

    private Task<CloudAppDiscoveryData?> GetCloudAppDiscoveryAsync()
    {
        return Task.FromResult<CloudAppDiscoveryData?>(new CloudAppDiscoveryData
        {
            DiscoveredApps = 0,
            SanctionedApps = 0,
            UnsanctionedApps = 0,
            Note = "Shadow IT discovery requires Microsoft Defender for Cloud Apps"
        });
    }

    private async Task<EmailSecurityData?> GetEmailSecurityStatsAsync()
    {
        try
        {
            _logger.LogInformation("Fetching email security stats");
            
            // Get total message count from Graph reports
            int totalMessages = 0;
            try
            {
                var mailflowSummary = await _graphService.GetMailflowSummaryAsync(30);
                totalMessages = (int)(mailflowSummary.TotalMessagesSent + mailflowSummary.TotalMessagesReceived);
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "Could not get total message count from mailflow summary");
            }
            
            // Note: Detailed email threat protection statistics (spam, malware, phishing blocked)
            // are not available through the standard Microsoft Graph API.
            // These statistics require:
            // 1. Microsoft Defender for Office 365 Plan 2 with the Security API
            // 2. Advanced Hunting queries via the Security.ThreatHunting.Read.All permission
            // 3. Or viewing directly in the Microsoft 365 Defender portal
            //
            // The threatAssessmentRequests API only shows user-submitted reports, not automatic detections.
            // The emailThreatSubmission API (beta) also only handles submissions, not detection stats.
            //
            // For now, we return the total message count and direct users to the Defender portal
            // for detailed threat protection statistics.
            
            return new EmailSecurityData
            {
                TotalMessages = totalMessages,
                SpamMessages = 0,
                MalwareMessages = 0,
                PhishingMessages = 0,
                BulkMessages = 0,
                Note = "Threat protection statistics (spam, malware, phishing blocked) are available in the Microsoft 365 Defender portal at security.microsoft.com under Email & collaboration > Reports."
            };
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Error fetching email security stats");
            return new EmailSecurityData
            {
                TotalMessages = 0,
                Note = "Could not retrieve email data. View threat statistics in the Microsoft 365 Defender portal."
            };
        }
    }

    private async Task<List<UserSignInDetailData>> GetUserSignInDetailsAsync()
    {
        var result = new List<UserSignInDetailData>();
        
        try
        {
            // Get users with sign-in activity
            var users = await _graphService.GetUsersAsync(null, "displayName", true, 999);
            
            foreach (var user in users.Users)
            {
                // Only include member users (not guests)
                if (user.UserType?.Equals("Guest", StringComparison.OrdinalIgnoreCase) == true)
                    continue;
                    
                result.Add(new UserSignInDetailData
                {
                    DisplayName = user.DisplayName,
                    UserPrincipalName = user.UserPrincipalName,
                    LastInteractiveSignIn = user.LastSignInDateTime,
                    LastNonInteractiveSignIn = user.LastNonInteractiveSignInDateTime,
                    DefaultMfaMethod = user.DefaultMfaMethod,
                    IsMfaRegistered = user.IsMfaRegistered,
                    AccountEnabled = user.AccountEnabled
                });
            }
            
            // Sort by display name
            result = result.OrderBy(u => u.DisplayName).ToList();
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Error fetching user sign-in details");
        }
        
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
            
            // Get deleted users from directory (without sorting - deletedDateTime sort not supported)
            var deletedUsers = await graphClient.Directory.DeletedItems.GraphUser.GetAsync(config =>
            {
                config.QueryParameters.Select = new[] { "id", "displayName", "userPrincipalName", "mail", "deletedDateTime", "jobTitle", "department" };
                config.QueryParameters.Top = 999;
            });
            
            if (deletedUsers?.Value != null)
            {
                foreach (var user in deletedUsers.Value)
                {
                    // Filter by deletion date within the report period
                    if (user.DeletedDateTime.HasValue)
                    {
                        var deletedDate = user.DeletedDateTime.Value.DateTime;
                        if (deletedDate >= startDate && deletedDate <= endDate.AddDays(1))
                        {
                            result.Add(new DeletedUserData
                            {
                                DisplayName = user.DisplayName,
                                UserPrincipalName = user.UserPrincipalName,
                                Mail = user.Mail,
                                DeletedDateTime = deletedDate,
                                JobTitle = user.JobTitle,
                                Department = user.Department
                            });
                        }
                    }
                }
            }
            
            // Page through if needed
            while (deletedUsers?.OdataNextLink != null)
            {
                deletedUsers = await graphClient.Directory.DeletedItems.GraphUser
                    .WithUrl(deletedUsers.OdataNextLink)
                    .GetAsync();
                    
                if (deletedUsers?.Value != null)
                {
                    foreach (var user in deletedUsers.Value)
                    {
                        if (user.DeletedDateTime.HasValue)
                        {
                            var deletedDate = user.DeletedDateTime.Value.DateTime;
                            if (deletedDate >= startDate && deletedDate <= endDate.AddDays(1))
                            {
                                result.Add(new DeletedUserData
                                {
                                    DisplayName = user.DisplayName,
                                    UserPrincipalName = user.UserPrincipalName,
                                    Mail = user.Mail,
                                    DeletedDateTime = deletedDate,
                                    JobTitle = user.JobTitle,
                                    Department = user.Department
                                });
                            }
                        }
                    }
                }
            }
            
            // Sort by deletion date descending (in memory since API doesn't support it)
            result = result.OrderByDescending(u => u.DeletedDateTime).ToList();
            
            _logger.LogInformation("Found {Count} deleted users in period {Start} to {End}", result.Count, startDate, endDate);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Error fetching deleted users");
        }
        
        return result;
    }

    private async Task<List<MailboxDetailData>> GetMailboxDetailsAsync()
    {
        var result = new List<MailboxDetailData>();
        
        try
        {
            _logger.LogInformation("Fetching mailbox details with sizes");

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
                
                var displayNameIdx = Array.FindIndex(header, h => 
                    h.Trim().Equals("Display Name", StringComparison.OrdinalIgnoreCase));
                var upnIdx = Array.FindIndex(header, h => 
                    h.Trim().Equals("User Principal Name", StringComparison.OrdinalIgnoreCase));
                var storageUsedIdx = Array.FindIndex(header, h => 
                    h.Trim().Equals("Storage Used (Byte)", StringComparison.OrdinalIgnoreCase));
                var recipientTypeIdx = Array.FindIndex(header, h => 
                    h.Trim().Equals("Recipient Type", StringComparison.OrdinalIgnoreCase));
                var lastActivityIdx = Array.FindIndex(header, h => 
                    h.Trim().Equals("Last Activity Date", StringComparison.OrdinalIgnoreCase));
                var prohibitSendQuotaIdx = Array.FindIndex(header, h => 
                    h.Trim().Equals("Prohibit Send Quota (Byte)", StringComparison.OrdinalIgnoreCase));
                var itemCountIdx = Array.FindIndex(header, h => 
                    h.Trim().Equals("Item Count", StringComparison.OrdinalIgnoreCase));
                var isDeletedIdx = Array.FindIndex(header, h => 
                    h.Trim().Equals("Is Deleted", StringComparison.OrdinalIgnoreCase));

                foreach (var line in lines.Skip(1).Where(l => !string.IsNullOrWhiteSpace(l)))
                {
                    var parts = line.Split(',');
                    
                    // Check if deleted
                    var isDeleted = isDeletedIdx >= 0 && isDeletedIdx < parts.Length &&
                        parts[isDeletedIdx]?.Trim().Equals("TRUE", StringComparison.OrdinalIgnoreCase) == true;
                    
                    if (isDeleted) continue;

                    var displayName = displayNameIdx >= 0 && displayNameIdx < parts.Length 
                        ? parts[displayNameIdx]?.Trim().Trim('"') : null;
                    var upn = upnIdx >= 0 && upnIdx < parts.Length 
                        ? parts[upnIdx]?.Trim().Trim('"') : null;
                    var recipientType = recipientTypeIdx >= 0 && recipientTypeIdx < parts.Length 
                        ? parts[recipientTypeIdx]?.Trim().Trim('"') : null;
                    
                    long storageUsed = 0;
                    if (storageUsedIdx >= 0 && storageUsedIdx < parts.Length)
                        long.TryParse(parts[storageUsedIdx]?.Trim(), out storageUsed);
                    
                    long? quota = null;
                    if (prohibitSendQuotaIdx >= 0 && prohibitSendQuotaIdx < parts.Length)
                    {
                        if (long.TryParse(parts[prohibitSendQuotaIdx]?.Trim(), out var q))
                            quota = q;
                    }
                    
                    int? itemCount = null;
                    if (itemCountIdx >= 0 && itemCountIdx < parts.Length)
                    {
                        if (int.TryParse(parts[itemCountIdx]?.Trim(), out var ic))
                            itemCount = ic;
                    }
                    
                    DateTime? lastActivity = null;
                    if (lastActivityIdx >= 0 && lastActivityIdx < parts.Length)
                    {
                        if (DateTime.TryParse(parts[lastActivityIdx]?.Trim(), out var la))
                            lastActivity = la;
                    }

                    var storageGB = Math.Round(storageUsed / 1073741824.0, 2);
                    var quotaGB = quota.HasValue ? Math.Round(quota.Value / 1073741824.0, 2) : (double?)null;
                    var percentUsed = quota.HasValue && quota.Value > 0 
                        ? Math.Round((double)storageUsed / quota.Value * 100, 1) 
                        : (double?)null;

                    result.Add(new MailboxDetailData
                    {
                        DisplayName = displayName,
                        UserPrincipalName = upn,
                        RecipientType = recipientType,
                        StorageUsedBytes = storageUsed,
                        StorageUsedGB = storageGB,
                        QuotaBytes = quota,
                        QuotaGB = quotaGB,
                        PercentUsed = percentUsed,
                        LastActivityDate = lastActivity,
                        ItemCount = itemCount
                    });
                }
            }

            // Sort by storage used descending
            result = result.OrderByDescending(m => m.StorageUsedBytes).ToList();
            
            _logger.LogInformation("Retrieved {Count} mailbox details", result.Count);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Error fetching mailbox details");
        }
        
        return result;
    }

    private async Task<byte[]> GenerateWordDocument(ExecutiveReportData data, ReportSettings settings)
    {
        using var stream = new MemoryStream();
        
        using (var wordDocument = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true))
        {
            // Add main document part
            var mainPart = wordDocument.AddMainDocumentPart();
            mainPart.Document = new Document();
            var body = mainPart.Document.AppendChild(new Body());

            // Add styles
            AddStyles(mainPart, settings);

            // --- Cover Page ---
            AddCoverPage(body, mainPart, data, settings);

            // Page break after cover
            body.Append(new Paragraph(new Run(new Break() { Type = BreakValues.Page })));

            // Title
            AddHeading(body, settings.ReportTitle, 1, true, settings);
            
            // Report info
            AddParagraph(body, $"Report Period: {data.ReportMonth}");
            AddParagraph(body, $"Prepared for: {settings.CompanyName}");
            AddParagraph(body, $"Generated: {data.GeneratedAt:dd MMMM yyyy HH:mm} UTC");
            AddParagraph(body, ""); // Empty line

            // Security Score Section
            AddHeading(body, "Security Score", 2);
            var securityTable = CreateTable(body, new[] { "Metric", "Score", "Max", "Percentage" });
            AddTableRow(securityTable, new[] { 
                "Microsoft Secure Score", 
                $"{data.SecureScore?.CurrentScore ?? 0}", 
                $"{data.SecureScore?.MaxScore ?? 0}", 
                $"{data.SecureScore?.PercentageScore ?? 0}%" 
            });

            // Intune Managed Devices
            AddHeading(body, "Intune Managed Devices", 2);
            var deviceTable = CreateTable(body, new[] { "Platform", "Count" });
            AddTableRow(deviceTable, new[] { "Total Devices", $"{data.DeviceStats?.TotalDevices ?? 0}" });
            AddTableRow(deviceTable, new[] { "Windows", $"{data.DeviceStats?.WindowsDevices ?? 0}" });
            AddTableRow(deviceTable, new[] { "macOS", $"{data.DeviceStats?.MacOsDevices ?? 0}" });
            AddTableRow(deviceTable, new[] { "iOS/iPadOS", $"{data.DeviceStats?.IosDevices ?? 0}" });
            AddTableRow(deviceTable, new[] { "Android", $"{data.DeviceStats?.AndroidDevices ?? 0}" });
            AddTableRow(deviceTable, new[] { "Compliant", $"{data.DeviceStats?.CompliantDevices ?? 0}" });
            AddTableRow(deviceTable, new[] { "Non-Compliant", $"{data.DeviceStats?.NonCompliantDevices ?? 0}" });
            AddTableRow(deviceTable, new[] { "Compliance Rate", $"{data.DeviceStats?.ComplianceRate ?? 0}%" });

            // Windows Patch Status
            AddHeading(body, "Windows Patch Status", 2);
            var patchTable = CreateTable(body, new[] { "Status", "Count" });
            AddTableRow(patchTable, new[] { "Total Windows Devices", $"{data.WindowsUpdateStats?.TotalWindowsDevices ?? 0}" });
            AddTableRow(patchTable, new[] { "Up to Date", $"{data.WindowsUpdateStats?.UpToDate ?? 0}" });
            AddTableRow(patchTable, new[] { "Needs Update", $"{data.WindowsUpdateStats?.NeedsUpdate ?? 0}" });
            AddTableRow(patchTable, new[] { "Compliance Rate", $"{data.WindowsUpdateStats?.ComplianceRate ?? 0}%" });
            if (!string.IsNullOrEmpty(data.WindowsUpdateStats?.Note))
                AddParagraph(body, data.WindowsUpdateStats.Note, true);

            // Microsoft Defender
            AddHeading(body, "Microsoft Defender for Endpoint", 2);
            var defenderTable = CreateTable(body, new[] { "Metric", "Value" });
            AddTableRow(defenderTable, new[] { "Exposure Score", data.DefenderStats?.ExposureScore ?? "N/A" });
            if (data.DefenderStats?.ExposureScoreNumeric.HasValue == true)
                AddTableRow(defenderTable, new[] { "Exposure Score (Numeric)", $"{data.DefenderStats.ExposureScoreNumeric}" });
            if (data.DefenderStats?.OnboardedMachines.HasValue == true)
                AddTableRow(defenderTable, new[] { "Onboarded Machines", $"{data.DefenderStats.OnboardedMachines}" });
            AddTableRow(defenderTable, new[] { "Total Vulnerabilities", $"{data.DefenderStats?.VulnerabilitiesDetected ?? 0}" });
            AddTableRow(defenderTable, new[] { "Critical", $"{data.DefenderStats?.CriticalVulnerabilities ?? 0}" });
            AddTableRow(defenderTable, new[] { "High", $"{data.DefenderStats?.HighVulnerabilities ?? 0}" });
            AddTableRow(defenderTable, new[] { "Medium", $"{data.DefenderStats?.MediumVulnerabilities ?? 0}" });
            AddTableRow(defenderTable, new[] { "Low", $"{data.DefenderStats?.LowVulnerabilities ?? 0}" });
            if (!string.IsNullOrEmpty(data.DefenderStats?.Note))
                AddParagraph(body, data.DefenderStats.Note, true);

            // User Accounts
            AddHeading(body, "User Accounts", 2);
            var userTable = CreateTable(body, new[] { "Type", "Count" });
            AddTableRow(userTable, new[] { "Total Users", $"{data.UserStats?.TotalUsers ?? 0}" });
            AddTableRow(userTable, new[] { "Guest Users", $"{data.UserStats?.GuestUsers ?? 0}" });
            AddTableRow(userTable, new[] { "Deleted Users (Soft)", $"{data.UserStats?.DeletedUsers ?? 0}" });
            AddTableRow(userTable, new[] { "MFA Registered", $"{data.UserStats?.MfaRegistered ?? 0}" });
            AddTableRow(userTable, new[] { "MFA Not Registered", $"{data.UserStats?.MfaNotRegistered ?? 0}" });
            AddParagraph(body, $"Risky Users: {data.RiskyUsersCount}");
            if (data.HighRiskUsers?.Any() == true)
                AddParagraph(body, $"High Risk: {string.Join(", ", data.HighRiskUsers)}", false, true);

            // Attack Simulation Training
            AddHeading(body, "Attack Simulation Training", 2);
            var attackTable = CreateTable(body, new[] { "Metric", "Value" });
            AddTableRow(attackTable, new[] { "Total Simulations", $"{data.AttackSimulation?.TotalSimulations ?? 0}" });
            AddTableRow(attackTable, new[] { "Completed", $"{data.AttackSimulation?.CompletedSimulations ?? 0}" });
            AddTableRow(attackTable, new[] { "Average Compromise Rate", $"{data.AttackSimulation?.AverageCompromiseRate ?? 0}%" });
            if (!string.IsNullOrEmpty(data.AttackSimulation?.Note))
                AddParagraph(body, data.AttackSimulation.Note, true);

            // Mailbox Usage
            AddHeading(body, "Mailbox Usage", 2);
            var mailboxTable = CreateTable(body, new[] { "Metric", "Value" });
            AddTableRow(mailboxTable, new[] { "Total Mailboxes", $"{data.MailboxStats?.TotalMailboxes ?? 0}" });
            AddTableRow(mailboxTable, new[] { "Active Mailboxes", $"{data.MailboxStats?.ActiveMailboxes ?? 0}" });
            AddTableRow(mailboxTable, new[] { "Total Storage Used", $"{data.MailboxStats?.TotalStorageUsedGB ?? 0} GB" });
            AddTableRow(mailboxTable, new[] { "Average Storage", $"{data.MailboxStats?.AverageStorageGB ?? 0} GB" });

            // SharePoint Usage
            AddHeading(body, "SharePoint Usage", 2);
            var spTable = CreateTable(body, new[] { "Metric", "Value" });
            AddTableRow(spTable, new[] { "Total Sites", $"{data.SharePointStats?.TotalSites ?? 0}" });
            AddTableRow(spTable, new[] { "Active Sites", $"{data.SharePointStats?.ActiveSites ?? 0}" });
            AddTableRow(spTable, new[] { "Storage Used", $"{data.SharePointStats?.TotalStorageUsedGB ?? 0} GB" });

            // Email Security
            AddHeading(body, "Email Security (Last 30 Days)", 2);
            var emailTable = CreateTable(body, new[] { "Metric", "Value" });
            AddTableRow(emailTable, new[] { "Total Messages Processed", $"{data.EmailSecurity?.TotalMessages ?? 0:N0}" });
            if (!string.IsNullOrEmpty(data.EmailSecurity?.Note))
                AddParagraph(body, data.EmailSecurity.Note, true);

            // Device Details - Windows
            if (data.DeviceDetails?.WindowsDevices?.Any() == true)
            {
                AddHeading(body, $"Windows Devices ({data.DeviceDetails.WindowsDevices.Count})", 2);
                var winTable = CreateTable(body, new[] { "Device Name", "OS Version", "Update Status", "Compliance", "Model", "Last Check-in" });
                foreach (var device in data.DeviceDetails.WindowsDevices)
                {
                    AddTableRow(winTable, new[] {
                        device.DeviceName ?? "-",
                        device.OsVersion ?? "-",
                        GetVersionStatusDisplay(device.OsVersionStatus, device.OsVersionStatusMessage),
                        device.ComplianceState ?? "-",
                        device.SkuFamily ?? "-",
                        device.LastCheckIn?.ToString("dd MMM yyyy") ?? "Never"
                    });
                }
            }

            // Device Details - macOS
            if (data.DeviceDetails?.MacDevices?.Any() == true)
            {
                AddHeading(body, $"macOS Devices ({data.DeviceDetails.MacDevices.Count})", 2);
                var macTable = CreateTable(body, new[] { "Device Name", "OS Version", "Update Status", "Compliance", "Ownership", "Last Check-in" });
                foreach (var device in data.DeviceDetails.MacDevices)
                {
                    AddTableRow(macTable, new[] {
                        device.DeviceName ?? "-",
                        device.OsVersion ?? "-",
                        GetVersionStatusDisplay(device.OsVersionStatus, device.OsVersionStatusMessage),
                        device.ComplianceState ?? "-",
                        device.Ownership ?? "-",
                        device.LastCheckIn?.ToString("dd MMM yyyy") ?? "Never"
                    });
                }
            }

            // Device Details - iOS
            if (data.DeviceDetails?.IosDevices?.Any() == true)
            {
                AddHeading(body, $"iOS/iPadOS Devices ({data.DeviceDetails.IosDevices.Count})", 2);
                var iosTable = CreateTable(body, new[] { "Device Name", "OS Version", "Update Status", "Compliance", "Ownership", "Last Check-in" });
                foreach (var device in data.DeviceDetails.IosDevices)
                {
                    AddTableRow(iosTable, new[] {
                        device.DeviceName ?? "-",
                        device.OsVersion ?? "-",
                        GetVersionStatusDisplay(device.OsVersionStatus, device.OsVersionStatusMessage),
                        device.ComplianceState ?? "-",
                        device.Ownership ?? "-",
                        device.LastCheckIn?.ToString("dd MMM yyyy") ?? "Never"
                    });
                }
            }

            // Device Details - Android
            if (data.DeviceDetails?.AndroidDevices?.Any() == true)
            {
                AddHeading(body, $"Android Devices ({data.DeviceDetails.AndroidDevices.Count})", 2);
                var androidTable = CreateTable(body, new[] { "Device Name", "OS Version", "Update Status", "Compliance", "Last Check-in" });
                foreach (var device in data.DeviceDetails.AndroidDevices)
                {
                    AddTableRow(androidTable, new[] {
                        device.DeviceName ?? "-",
                        device.OsVersion ?? "-",
                        GetVersionStatusDisplay(device.OsVersionStatus, device.OsVersionStatusMessage),
                        device.ComplianceState ?? "-",
                        device.LastCheckIn?.ToString("dd MMM yyyy") ?? "Never"
                    });
                }
            }

            // User Sign-in & MFA Details
            if (data.UserSignInDetails?.Any() == true)
            {
                AddHeading(body, $"User Sign-in & MFA Details ({data.UserSignInDetails.Count} users)", 2);
                var signInTable = CreateTable(body, new[] { "Display Name", "Email", "Last Interactive", "Last Non-Interactive", "Default MFA", "MFA", "Enabled" });
                foreach (var user in data.UserSignInDetails)
                {
                    AddTableRow(signInTable, new[] {
                        user.DisplayName ?? "-",
                        user.UserPrincipalName ?? "-",
                        user.LastInteractiveSignIn?.ToString("dd MMM yyyy") ?? "Never",
                        user.LastNonInteractiveSignIn?.ToString("dd MMM yyyy") ?? "Never",
                        user.DefaultMfaMethod ?? "None",
                        user.IsMfaRegistered ? "Yes" : "No",
                        user.AccountEnabled ? "Yes" : "No"
                    });
                }
            }

            // Deleted Users in Period
            AddHeading(body, $"Deleted Users in Period ({data.DeletedUsersInPeriod?.Count ?? 0} users)", 2);
            if (data.DeletedUsersInPeriod?.Any() == true)
            {
                var deletedTable = CreateTable(body, new[] { "Display Name", "Email", "Deleted Date", "Job Title", "Department" });
                foreach (var user in data.DeletedUsersInPeriod)
                {
                    AddTableRow(deletedTable, new[] {
                        user.DisplayName ?? "-",
                        user.UserPrincipalName ?? user.Mail ?? "-",
                        user.DeletedDateTime?.ToString("dd MMM yyyy") ?? "-",
                        user.JobTitle ?? "-",
                        user.Department ?? "-"
                    });
                }
            }
            else
            {
                AddParagraph(body, "No users were deleted during this period.");
            }

            // Mailbox Details with Sizes
            if (data.MailboxDetails?.Any() == true)
            {
                AddHeading(body, $"Mailbox Storage Details ({data.MailboxDetails.Count} mailboxes)", 2);
                var mailboxDetailTable = CreateTable(body, new[] { "Display Name", "Email", "Type", "Size (GB)", "Quota (GB)", "% Used", "Items", "Last Activity" });
                foreach (var mailbox in data.MailboxDetails)
                {
                    AddTableRow(mailboxDetailTable, new[] {
                        mailbox.DisplayName ?? "-",
                        mailbox.UserPrincipalName ?? "-",
                        mailbox.RecipientType ?? "User",
                        $"{mailbox.StorageUsedGB:F2}",
                        mailbox.QuotaGB.HasValue ? $"{mailbox.QuotaGB:F0}" : "-",
                        mailbox.PercentUsed.HasValue ? $"{mailbox.PercentUsed:F1}%" : "-",
                        mailbox.ItemCount?.ToString("N0") ?? "-",
                        mailbox.LastActivityDate?.ToString("dd MMM yyyy") ?? "Never"
                    });
                }
            }

            // Domain Email Security
            if (data.DomainSecuritySummary != null)
            {
                AddHeading(body, "Domain Email Security", 2);
                
                // Summary table
                var domainSummaryTable = CreateTable(body, new[] { "Metric", "Count" });
                AddTableRow(domainSummaryTable, new[] { "Total Domains Checked", $"{data.DomainSecuritySummary.TotalDomains}" });
                AddTableRow(domainSummaryTable, new[] { "Domains with MX Records", $"{data.DomainSecuritySummary.DomainsWithMx}" });
                AddTableRow(domainSummaryTable, new[] { "Domains with SPF", $"{data.DomainSecuritySummary.DomainsWithSpf}" });
                AddTableRow(domainSummaryTable, new[] { "Domains with DMARC", $"{data.DomainSecuritySummary.DomainsWithDmarc}" });
                AddTableRow(domainSummaryTable, new[] { "Domains with DKIM", $"{data.DomainSecuritySummary.DomainsWithDkim}" });
                AddTableRow(domainSummaryTable, new[] { "Domains with MTA-STS", $"{data.DomainSecuritySummary.DomainsWithMtaSts}" });
                
                // DMARC Policy Distribution
                AddParagraph(body, "");
                AddParagraph(body, "DMARC Policy Distribution:");
                var dmarcTable = CreateTable(body, new[] { "Policy", "Count" });
                AddTableRow(dmarcTable, new[] { "Reject (Full Protection)", $"{data.DomainSecuritySummary.DmarcRejectCount}" });
                AddTableRow(dmarcTable, new[] { "Quarantine", $"{data.DomainSecuritySummary.DmarcQuarantineCount}" });
                AddTableRow(dmarcTable, new[] { "None (Monitor Only)", $"{data.DomainSecuritySummary.DmarcNoneCount}" });
                
                // Security Grade Distribution
                AddParagraph(body, "");
                AddParagraph(body, "Security Grade Distribution:");
                var gradeTable = CreateTable(body, new[] { "Grade", "Count" });
                AddTableRow(gradeTable, new[] { "A (90-100)", $"{data.DomainSecuritySummary.GradeACount}" });
                AddTableRow(gradeTable, new[] { "B (80-89)", $"{data.DomainSecuritySummary.GradeBCount}" });
                AddTableRow(gradeTable, new[] { "C (70-79)", $"{data.DomainSecuritySummary.GradeCCount}" });
                AddTableRow(gradeTable, new[] { "D (60-69)", $"{data.DomainSecuritySummary.GradeDCount}" });
                AddTableRow(gradeTable, new[] { "F (Below 60)", $"{data.DomainSecuritySummary.GradeFCount}" });
                
                if (data.DomainSecuritySummary.CriticalIssuesCount > 0)
                    AddParagraph(body, $"Critical: {data.DomainSecuritySummary.CriticalIssuesCount} domains require immediate attention (Grade D or F)", false, true);
            }
            
            // Domain Security Details
            if (data.DomainSecurityResults?.Any() == true)
            {
                AddHeading(body, $"Domain Security Details ({data.DomainSecurityResults.Count} domains)", 2);
                var domainTable = CreateTable(body, new[] { "Domain", "MX", "SPF", "DMARC", "DKIM", "MTA-STS" });
                foreach (var domain in data.DomainSecurityResults.OrderByDescending(d => d.SecurityScore))
                {
                    AddTableRow(domainTable, new[] {
                        domain.Domain,
                        domain.HasMx ? "✓" : "✗",
                        domain.HasSpf ? (domain.SpfPolicy == "-all" ? "✓ Hard" : "~ Soft") : "✗",
                        domain.HasDmarc ? domain.DmarcPolicy ?? "-" : "✗",
                        domain.HasDkim ? "✓" : "✗",
                        domain.HasMtaSts ? "✓" : "✗"
                    });
                }
                
                // Critical domains needing attention
                var criticalDomains = data.DomainSecurityResults.Where(d => d.SecurityGrade == "D" || d.SecurityGrade == "F").ToList();
                if (criticalDomains.Any())
                {
                    AddParagraph(body, "");
                    AddParagraph(body, "Domains Requiring Immediate Attention:", false, true);
                    foreach (var domain in criticalDomains)
                    {
                        var issues = string.Join(", ", domain.Issues ?? new List<string>());
                        AddParagraph(body, $"• {domain.Domain}: {issues}", true);
                    }
                }
            }

            // Footer
            AddParagraph(body, "");
            AddParagraph(body, "This report was automatically generated by M365 Dashboard.", true);
            AddParagraph(body, "Some metrics may require additional licensing or API permissions.", true);

            mainPart.Document.Save();
        }

        return stream.ToArray();
    }

    private async Task<DeviceDetailsData> GetDeviceDetailsAsync()
    {
        var result = new DeviceDetailsData();
        
        try
        {
            _logger.LogInformation("Fetching Intune managed devices with OS version status");
            
            // Pre-fetch latest versions (caches for 24 hours) - don't fail if this doesn't work
            try
            {
                await _osVersionService.GetLatestVersionsAsync();
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "Failed to fetch latest OS versions, will use fallbacks");
            }
            
            var devices = await _graphService.GetDevicesAsync(null, "deviceName", true, 1000);
            
            foreach (var device in devices.Devices)
            {
                var os = device.OperatingSystem?.ToLower() ?? "";
                
                if (os.Contains("windows"))
                {
                    var versionStatus = _osVersionService.CheckWindowsVersion(device.OsVersion);
                    result.WindowsDevices.Add(new WindowsDeviceDetailData
                    {
                        DeviceName = device.DeviceName,
                        LastCheckIn = device.LastSyncDateTime,
                        OsVersion = device.OsVersion,
                        ComplianceState = FormatComplianceState(device.ComplianceState),
                        ManagementAgent = device.ManagementAgent,
                        Ownership = device.ManagedDeviceOwnerType,
                        SkuFamily = device.Model,
                        OsVersionStatus = versionStatus.Status,
                        OsVersionStatusMessage = versionStatus.Message,
                        LatestVersion = versionStatus.LatestVersion
                    });
                }
                else if (os.Contains("ios") || os.Contains("ipad"))
                {
                    var versionStatus = _osVersionService.CheckiOSVersion(device.OsVersion);
                    result.IosDevices.Add(new IosDeviceDetailData
                    {
                        DeviceName = device.DeviceName,
                        ComplianceState = FormatComplianceState(device.ComplianceState),
                        ManagementAgent = device.ManagementAgent,
                        Ownership = device.ManagedDeviceOwnerType,
                        Os = device.OperatingSystem,
                        OsVersion = device.OsVersion,
                        LastCheckIn = device.LastSyncDateTime,
                        OsVersionStatus = versionStatus.Status,
                        OsVersionStatusMessage = versionStatus.Message,
                        LatestVersion = versionStatus.LatestVersion
                    });
                }
                else if (os.Contains("android"))
                {
                    // Note: Graph API doesn't return securityPatchLevel in standard endpoint
                    var versionStatus = _osVersionService.CheckAndroidVersion(device.OsVersion, null);
                    result.AndroidDevices.Add(new AndroidDeviceDetailData
                    {
                        DeviceName = device.DeviceName,
                        ComplianceState = FormatComplianceState(device.ComplianceState),
                        ManagementAgent = device.ManagementAgent,
                        Os = device.OperatingSystem,
                        OsVersion = device.OsVersion,
                        LastCheckIn = device.LastSyncDateTime,
                        OsVersionStatus = versionStatus.Status,
                        OsVersionStatusMessage = versionStatus.Message,
                        LatestVersion = versionStatus.LatestVersion
                    });
                }
                else if (os.Contains("mac"))
                {
                    var versionStatus = _osVersionService.CheckMacOSVersion(device.OsVersion);
                    result.MacDevices.Add(new MacDeviceDetailData
                    {
                        DeviceName = device.DeviceName,
                        LastCheckIn = device.LastSyncDateTime,
                        OsVersion = device.OsVersion,
                        ComplianceState = FormatComplianceState(device.ComplianceState),
                        ManagementAgent = device.ManagementAgent,
                        Ownership = device.ManagedDeviceOwnerType,
                        OsVersionStatus = versionStatus.Status,
                        OsVersionStatusMessage = versionStatus.Message,
                        LatestVersion = versionStatus.LatestVersion
                    });
                }
            }
            
            // Sort each list by device name
            result.WindowsDevices = result.WindowsDevices.OrderBy(d => d.DeviceName).ToList();
            result.IosDevices = result.IosDevices.OrderBy(d => d.DeviceName).ToList();
            result.AndroidDevices = result.AndroidDevices.OrderBy(d => d.DeviceName).ToList();
            result.MacDevices = result.MacDevices.OrderBy(d => d.DeviceName).ToList();
            
            // Log summary with version status counts
            var criticalCount = result.WindowsDevices.Count(d => d.OsVersionStatus == VersionStatus.Critical)
                + result.IosDevices.Count(d => d.OsVersionStatus == VersionStatus.Critical)
                + result.AndroidDevices.Count(d => d.OsVersionStatus == VersionStatus.Critical)
                + result.MacDevices.Count(d => d.OsVersionStatus == VersionStatus.Critical);
            
            var warningCount = result.WindowsDevices.Count(d => d.OsVersionStatus == VersionStatus.Warning)
                + result.IosDevices.Count(d => d.OsVersionStatus == VersionStatus.Warning)
                + result.AndroidDevices.Count(d => d.OsVersionStatus == VersionStatus.Warning)
                + result.MacDevices.Count(d => d.OsVersionStatus == VersionStatus.Warning);
            
            _logger.LogInformation(
                "Device details: {Windows} Windows, {Mac} macOS, {iOS} iOS, {Android} Android. Version status: {Critical} critical, {Warning} warning",
                result.WindowsDevices.Count, result.MacDevices.Count, result.IosDevices.Count, result.AndroidDevices.Count,
                criticalCount, warningCount);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Error fetching device details");
        }
        
        return result;
    }

    private string FormatComplianceState(string? state)
    {
        if (string.IsNullOrEmpty(state)) return "Unknown";
        
        return state.ToLower() switch
        {
            "compliant" => "Compliant",
            "noncompliant" => "Non-Compliant",
            "conflict" => "Conflict",
            "error" => "Error",
            "ingraceperiod" => "In Grace Period",
            "configmanager" => "Config Manager",
            _ => state
        };
    }

    private string GetVersionStatusDisplay(VersionStatus status, string? message)
    {
        var prefix = status switch
        {
            VersionStatus.Current => "✓",
            VersionStatus.Warning => "⚠",
            VersionStatus.Critical => "❌",
            _ => "?"
        };
        return $"{prefix} {message ?? status.ToString()}";
    }

    private async Task<AppCredentialStatusData> GetAppCredentialStatusAsync()
    {
        var result = new AppCredentialStatusData
        {
            ThresholdDays = 45
        };
        
        try
        {
            _logger.LogInformation("Fetching app registration credentials status");
            
            var today = DateTime.UtcNow;
            var expirationThreshold = today.AddDays(result.ThresholdDays);
            
            // Get all applications with their credentials
            var apps = await _graphClient.Applications.GetAsync(config =>
            {
                config.QueryParameters.Select = new[] { "id", "appId", "displayName", "passwordCredentials", "keyCredentials" };
                config.QueryParameters.Top = 999;
            });
            
            var appsWithExpiringSecrets = new HashSet<string>();
            var appsWithExpiredSecrets = new HashSet<string>();
            var appsWithExpiringCerts = new HashSet<string>();
            var appsWithExpiredCerts = new HashSet<string>();
            
            if (apps?.Value != null)
            {
                foreach (var app in apps.Value)
                {
                    result.TotalApps++;
                    
                    // Check password credentials (secrets)
                    if (app.PasswordCredentials != null)
                    {
                        foreach (var secret in app.PasswordCredentials)
                        {
                            if (secret.EndDateTime == null) continue;
                            
                            var expiryDate = secret.EndDateTime.Value.DateTime;
                            var daysUntilExpiry = (int)(expiryDate - today).TotalDays;
                            
                            if (expiryDate < today)
                            {
                                // Expired
                                appsWithExpiredSecrets.Add(app.Id ?? "");
                                result.ExpiredSecrets.Add(new AppCredentialDetail
                                {
                                    AppName = app.DisplayName,
                                    AppId = app.AppId,
                                    CredentialType = "Secret",
                                    Description = secret.DisplayName,
                                    ExpiryDate = expiryDate,
                                    DaysUntilExpiry = daysUntilExpiry,
                                    Status = "Expired"
                                });
                            }
                            else if (expiryDate < expirationThreshold)
                            {
                                // Expiring soon
                                appsWithExpiringSecrets.Add(app.Id ?? "");
                                result.ExpiringSecrets.Add(new AppCredentialDetail
                                {
                                    AppName = app.DisplayName,
                                    AppId = app.AppId,
                                    CredentialType = "Secret",
                                    Description = secret.DisplayName,
                                    ExpiryDate = expiryDate,
                                    DaysUntilExpiry = daysUntilExpiry,
                                    Status = $"Expires in {daysUntilExpiry} days"
                                });
                            }
                        }
                    }
                    
                    // Check key credentials (certificates)
                    if (app.KeyCredentials != null)
                    {
                        foreach (var cert in app.KeyCredentials)
                        {
                            if (cert.EndDateTime == null) continue;
                            
                            var expiryDate = cert.EndDateTime.Value.DateTime;
                            var daysUntilExpiry = (int)(expiryDate - today).TotalDays;
                            
                            if (expiryDate < today)
                            {
                                // Expired
                                appsWithExpiredCerts.Add(app.Id ?? "");
                                result.ExpiredCertificates.Add(new AppCredentialDetail
                                {
                                    AppName = app.DisplayName,
                                    AppId = app.AppId,
                                    CredentialType = "Certificate",
                                    Description = cert.DisplayName,
                                    ExpiryDate = expiryDate,
                                    DaysUntilExpiry = daysUntilExpiry,
                                    Status = "Expired"
                                });
                            }
                            else if (expiryDate < expirationThreshold)
                            {
                                // Expiring soon
                                appsWithExpiringCerts.Add(app.Id ?? "");
                                result.ExpiringCertificates.Add(new AppCredentialDetail
                                {
                                    AppName = app.DisplayName,
                                    AppId = app.AppId,
                                    CredentialType = "Certificate",
                                    Description = cert.DisplayName,
                                    ExpiryDate = expiryDate,
                                    DaysUntilExpiry = daysUntilExpiry,
                                    Status = $"Expires in {daysUntilExpiry} days"
                                });
                            }
                        }
                    }
                }
                
                // Page through if needed
                while (apps.OdataNextLink != null)
                {
                    apps = await _graphClient.Applications.WithUrl(apps.OdataNextLink).GetAsync();
                    
                    if (apps?.Value == null) break;
                    
                    foreach (var app in apps.Value)
                    {
                        result.TotalApps++;
                        
                        // Check password credentials (secrets)
                        if (app.PasswordCredentials != null)
                        {
                            foreach (var secret in app.PasswordCredentials)
                            {
                                if (secret.EndDateTime == null) continue;
                                
                                var expiryDate = secret.EndDateTime.Value.DateTime;
                                var daysUntilExpiry = (int)(expiryDate - today).TotalDays;
                                
                                if (expiryDate < today)
                                {
                                    appsWithExpiredSecrets.Add(app.Id ?? "");
                                    result.ExpiredSecrets.Add(new AppCredentialDetail
                                    {
                                        AppName = app.DisplayName,
                                        AppId = app.AppId,
                                        CredentialType = "Secret",
                                        Description = secret.DisplayName,
                                        ExpiryDate = expiryDate,
                                        DaysUntilExpiry = daysUntilExpiry,
                                        Status = "Expired"
                                    });
                                }
                                else if (expiryDate < expirationThreshold)
                                {
                                    appsWithExpiringSecrets.Add(app.Id ?? "");
                                    result.ExpiringSecrets.Add(new AppCredentialDetail
                                    {
                                        AppName = app.DisplayName,
                                        AppId = app.AppId,
                                        CredentialType = "Secret",
                                        Description = secret.DisplayName,
                                        ExpiryDate = expiryDate,
                                        DaysUntilExpiry = daysUntilExpiry,
                                        Status = $"Expires in {daysUntilExpiry} days"
                                    });
                                }
                            }
                        }
                        
                        // Check key credentials (certificates)
                        if (app.KeyCredentials != null)
                        {
                            foreach (var cert in app.KeyCredentials)
                            {
                                if (cert.EndDateTime == null) continue;
                                
                                var expiryDate = cert.EndDateTime.Value.DateTime;
                                var daysUntilExpiry = (int)(expiryDate - today).TotalDays;
                                
                                if (expiryDate < today)
                                {
                                    appsWithExpiredCerts.Add(app.Id ?? "");
                                    result.ExpiredCertificates.Add(new AppCredentialDetail
                                    {
                                        AppName = app.DisplayName,
                                        AppId = app.AppId,
                                        CredentialType = "Certificate",
                                        Description = cert.DisplayName,
                                        ExpiryDate = expiryDate,
                                        DaysUntilExpiry = daysUntilExpiry,
                                        Status = "Expired"
                                    });
                                }
                                else if (expiryDate < expirationThreshold)
                                {
                                    appsWithExpiringCerts.Add(app.Id ?? "");
                                    result.ExpiringCertificates.Add(new AppCredentialDetail
                                    {
                                        AppName = app.DisplayName,
                                        AppId = app.AppId,
                                        CredentialType = "Certificate",
                                        Description = cert.DisplayName,
                                        ExpiryDate = expiryDate,
                                        DaysUntilExpiry = daysUntilExpiry,
                                        Status = $"Expires in {daysUntilExpiry} days"
                                    });
                                }
                            }
                        }
                    }
                }
            }
            
            result.AppsWithExpiringSecrets = appsWithExpiringSecrets.Count;
            result.AppsWithExpiredSecrets = appsWithExpiredSecrets.Count;
            result.AppsWithExpiringCertificates = appsWithExpiringCerts.Count;
            result.AppsWithExpiredCertificates = appsWithExpiredCerts.Count;
            
            // Sort by expiry date
            result.ExpiringSecrets = result.ExpiringSecrets.OrderBy(s => s.ExpiryDate).ToList();
            result.ExpiredSecrets = result.ExpiredSecrets.OrderByDescending(s => s.ExpiryDate).ToList();
            result.ExpiringCertificates = result.ExpiringCertificates.OrderBy(c => c.ExpiryDate).ToList();
            result.ExpiredCertificates = result.ExpiredCertificates.OrderByDescending(c => c.ExpiryDate).ToList();
            
            _logger.LogInformation("App credentials: {Total} apps, {ExpiringSecrets} expiring secrets, {ExpiredSecrets} expired secrets, {ExpiringCerts} expiring certs, {ExpiredCerts} expired certs",
                result.TotalApps, result.AppsWithExpiringSecrets, result.AppsWithExpiredSecrets, result.AppsWithExpiringCertificates, result.AppsWithExpiredCertificates);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Error fetching app credential status");
        }
        
        return result;
    }

    private void AddStyles(MainDocumentPart mainPart, ReportSettings settings)
    {
        var primary = ToOxmlColor(settings.PrimaryColor);
        var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
        var styles = new Styles();
        
        // Heading 1 style
        var heading1 = new Style() { Type = StyleValues.Paragraph, StyleId = "Heading1" };
        heading1.Append(new StyleName() { Val = "Heading 1" });
        heading1.Append(new StyleRunProperties(
            new Bold(),
            new FontSize() { Val = "48" },
            new Color() { Val = primary }
        ));
        styles.Append(heading1);

        // Heading 2 style
        var heading2 = new Style() { Type = StyleValues.Paragraph, StyleId = "Heading2" };
        heading2.Append(new StyleName() { Val = "Heading 2" });
        heading2.Append(new StyleRunProperties(
            new Bold(),
            new FontSize() { Val = "28" },
            new Color() { Val = "323130" }
        ));
        heading2.Append(new StyleParagraphProperties(
            new SpacingBetweenLines() { Before = "200", After = "100" }
        ));
        styles.Append(heading2);

        stylesPart.Styles = styles;
        stylesPart.Styles.Save();
    }

    /// <summary>Adds a branded cover page to the Word document body.</summary>
    private void AddCoverPage(Body body, MainDocumentPart mainPart, ExecutiveReportData data, ReportSettings settings)
    {
        var primary = ToOxmlColor(settings.PrimaryColor);

        // Full-width shaded block as cover background
        var coverPara = new Paragraph();
        var coverParaProps = new ParagraphProperties(
            new Shading() { Fill = primary, Color = "auto" },
            new SpacingBetweenLines() { Before = "0", After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto },
            new Justification() { Val = JustificationValues.Center }
        );
        coverPara.Append(coverParaProps);

        // Logo (if set)
        if (!string.IsNullOrEmpty(settings.LogoBase64))
        {
            try
            {
                var logoBytes = Convert.FromBase64String(settings.LogoBase64);
                var imagePart = mainPart.AddImagePart(settings.LogoContentType?.Contains("png") == true
                    ? ImagePartType.Png : ImagePartType.Jpeg);
                using var imgStream = new MemoryStream(logoBytes);
                imagePart.FeedData(imgStream);

                var relationshipId = mainPart.GetIdOfPart(imagePart);
                // ~2 inch wide logo (1440000 EMUs per inch)
                const long cx = 1440000L * 2;
                const long cy = 720000L;

                var logoRun = new Run(new Drawing(
                    new DW.Inline(
                        new DW.Extent() { Cx = cx, Cy = cy },
                        new DW.EffectExtent() { LeftEdge = 0, TopEdge = 0, RightEdge = 0, BottomEdge = 0 },
                        new DW.DocProperties() { Id = 1U, Name = "Logo" },
                        new DW.NonVisualGraphicFrameDrawingProperties(
                            new A.GraphicFrameLocks() { NoChangeAspect = true }),
                        new A.Graphic(
                            new A.GraphicData(
                                new PIC.Picture(
                                    new PIC.NonVisualPictureProperties(
                                        new PIC.NonVisualDrawingProperties() { Id = 0U, Name = "logo" },
                                        new PIC.NonVisualPictureDrawingProperties()),
                                    new PIC.BlipFill(
                                        new A.Blip() { Embed = relationshipId },
                                        new A.Stretch(new A.FillRectangle())),
                                    new PIC.ShapeProperties(
                                        new A.Transform2D(
                                            new A.Offset() { X = 0, Y = 0 },
                                            new A.Extents() { Cx = cx, Cy = cy }),
                                        new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }
                                    )))
                            { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                    ) { DistanceFromTop = 0U, DistanceFromBottom = 0U, DistanceFromLeft = 0U, DistanceFromRight = 0U }));
                coverPara.Append(logoRun);
            }
            catch { /* logo embedding failed - skip silently */ }
        }

        body.Append(coverPara);

        // Company name
        var companyPara = new Paragraph();
        companyPara.Append(new ParagraphProperties(
            new Justification() { Val = JustificationValues.Center },
            new SpacingBetweenLines() { Before = "480", After = "120" }
        ));
        var companyRun = new Run();
        companyRun.Append(new RunProperties(
            new Bold(),
            new FontSize() { Val = "36" },
            new Color() { Val = primary }
        ));
        companyRun.Append(new Text(settings.CompanyName));
        companyPara.Append(companyRun);
        body.Append(companyPara);

        // Report title
        var titlePara = new Paragraph();
        titlePara.Append(new ParagraphProperties(
            new Justification() { Val = JustificationValues.Center },
            new SpacingBetweenLines() { Before = "120", After = "120" }
        ));
        var titleRun = new Run();
        titleRun.Append(new RunProperties(
            new Bold(),
            new FontSize() { Val = "56" },
            new Color() { Val = primary }
        ));
        titleRun.Append(new Text(settings.ReportTitle));
        titlePara.Append(titleRun);
        body.Append(titlePara);

        // Period & generated date
        var infoPara = new Paragraph();
        infoPara.Append(new ParagraphProperties(
            new Justification() { Val = JustificationValues.Center },
            new SpacingBetweenLines() { Before = "240", After = "120" }
        ));
        var infoRun = new Run();
        infoRun.Append(new RunProperties(new FontSize() { Val = "24" }, new Color() { Val = "555555" }));
        infoRun.Append(new Text($"{data.ReportMonth}   |   Generated {data.GeneratedAt:dd MMMM yyyy} UTC"));
        infoPara.Append(infoRun);
        body.Append(infoPara);
    }

    private void AddHeading(Body body, string text, int level, bool isTitle = false, ReportSettings? settings = null)
    {
        var primary = settings != null ? ToOxmlColor(settings.PrimaryColor) : "0078D4";
        var para = new Paragraph();
        var run = new Run();
        var runProps = new RunProperties();
        
        if (level == 1)
        {
            runProps.Append(new Bold());
            runProps.Append(new FontSize() { Val = isTitle ? "56" : "48" });
            runProps.Append(new Color() { Val = primary });
        }
        else
        {
            runProps.Append(new Bold());
            runProps.Append(new FontSize() { Val = "28" });
            runProps.Append(new Color() { Val = "323130" });
            
            // Add bottom border for H2
            var paraProps = new ParagraphProperties();
            paraProps.Append(new SpacingBetweenLines() { Before = "300", After = "100" });
            para.Append(paraProps);
        }
        
        run.Append(runProps);
        run.Append(new Text(text));
        para.Append(run);
        body.Append(para);
    }

    private void AddParagraph(Body body, string text, bool isItalic = false, bool isRed = false)
    {
        var para = new Paragraph();
        var run = new Run();
        var runProps = new RunProperties();
        
        if (isItalic)
            runProps.Append(new Italic());
        if (isRed)
            runProps.Append(new Color() { Val = "D13438" });
        
        runProps.Append(new FontSize() { Val = "22" });
        
        run.Append(runProps);
        run.Append(new Text(text));
        para.Append(run);
        body.Append(para);
    }

    private Table CreateTable(Body body, string[] headers, ReportSettings? settings = null)
    {
        var table = new Table();
        
        // Table properties
        var tableProps = new TableProperties(
            new TableBorders(
                new TopBorder() { Val = BorderValues.Single, Size = 4, Color = "DDDDDD" },
                new BottomBorder() { Val = BorderValues.Single, Size = 4, Color = "DDDDDD" },
                new LeftBorder() { Val = BorderValues.Single, Size = 4, Color = "DDDDDD" },
                new RightBorder() { Val = BorderValues.Single, Size = 4, Color = "DDDDDD" },
                new InsideHorizontalBorder() { Val = BorderValues.Single, Size = 4, Color = "DDDDDD" },
                new InsideVerticalBorder() { Val = BorderValues.Single, Size = 4, Color = "DDDDDD" }
            ),
            new TableWidth() { Type = TableWidthUnitValues.Pct, Width = "5000" }
        );
        table.Append(tableProps);

        // Header row
        var headerRow = new TableRow();
        foreach (var header in headers)
        {
            var cell = new TableCell();
            var cellProps = new TableCellProperties(
                new Shading() { Fill = "0078D4" },
                new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center }
            );
            cell.Append(cellProps);
            
            var para = new Paragraph();
            var run = new Run();
            run.Append(new RunProperties(
                new Bold(),
                new Color() { Val = "FFFFFF" },
                new FontSize() { Val = "20" }
            ));
            run.Append(new Text(header));
            para.Append(run);
            cell.Append(para);
            headerRow.Append(cell);
        }
        table.Append(headerRow);

        body.Append(table);
        body.Append(new Paragraph()); // Add spacing after table
        
        return table;
    }

    private void AddTableRow(Table table, string[] values)
    {
        var row = new TableRow();
        foreach (var value in values)
        {
            var cell = new TableCell();
            var cellProps = new TableCellProperties(
                new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center }
            );
            cell.Append(cellProps);
            
            var para = new Paragraph();
            var run = new Run();
            run.Append(new RunProperties(new FontSize() { Val = "20" }));
            run.Append(new Text(value));
            para.Append(run);
            cell.Append(para);
            row.Append(cell);
        }
        
        // Insert before the last element (spacing paragraph)
        var lastRow = table.Elements<TableRow>().LastOrDefault();
        if (lastRow != null)
            lastRow.InsertAfterSelf(row);
        else
            table.Append(row);
    }

    private static List<ReportQuote> PickRandomQuotesForHtml(ReportSettings settings, int count)
    {
        var pool = settings.Quotes
            .Where(q => q.Enabled &&
                        !string.IsNullOrWhiteSpace(q.BigNumber) &&
                        !string.IsNullOrWhiteSpace(q.Line1))
            .ToList();
        if (pool.Count == 0) pool = ReportSettings.DefaultQuotes();
        if (pool.Count <= count) return pool;
        var rng = new Random();
        for (int i = pool.Count - 1; i > 0; i--)
        {
            int j = rng.Next(i + 1);
            (pool[i], pool[j]) = (pool[j], pool[i]);
        }
        return pool.Take(count).ToList();
    }

    private static string RenderHtmlInfoGraphic(ReportQuote q) =>
        $"<div class='infographic page-break'><div class='big-number'>{q.BigNumber}</div><div class='line1'>{q.Line1}</div><div class='line2'>{q.Line2}</div><div class='source'>{q.Source}</div></div>";

    private string GenerateHtmlReport(ExecutiveReportData data, ReportSettings settings)
    {
        var primaryColor = settings.PrimaryColor?.TrimStart('#') ?? "1E3A5F";
        var accentColor = settings.AccentColor?.TrimStart('#') ?? "E07C3A";
        var showQuotes = settings.ShowInfoGraphics && settings.ShowQuotes;
        var quotes = showQuotes ? PickRandomQuotesForHtml(settings, 3) : new List<ReportQuote>();
        
        return $@"
<!DOCTYPE html>
<html>
<head>
    <title>{settings.ReportTitle} - {data.ReportMonth}</title>
    <style>
        * {{ box-sizing: border-box; }}
        body {{ font-family: 'Segoe UI', Arial, sans-serif; margin: 0; padding: 0; color: #374151; font-size: 11pt; }}
        
        /* Print styles */
        @media print {{
            .no-print {{ display: none !important; }}
            .page-break {{ page-break-before: always; }}
            .cover-page {{ height: 100vh; }}
            .infographic {{ height: 100vh; page-break-inside: avoid; }}
            body {{ print-color-adjust: exact; -webkit-print-color-adjust: exact; }}
        }}
        
        /* Cover Page */
        .cover-page {{
            background: #{primaryColor};
            color: white;
            padding: 60px 50px;
            min-height: 500px;
            display: flex;
            flex-direction: column;
            justify-content: flex-start;
        }}
        .cover-page .subtitle {{ color: #{accentColor}; font-size: 20pt; letter-spacing: 3px; margin-top: 100px; }}
        .cover-page .title {{ font-size: 36pt; font-weight: 300; letter-spacing: 2px; margin-top: 10px; }}
        
        .cover-info {{
            background: white;
            padding: 50px;
            text-align: center;
        }}
        .cover-info .company {{ font-size: 18pt; font-weight: bold; color: #{primaryColor}; }}
        .cover-info .date {{ font-size: 11pt; color: #666; margin-top: 10px; }}
        .cover-info .logo {{ max-height: 60px; margin-top: 30px; }}
        
        /* Infographic Pages */
        .infographic {{
            background: #{primaryColor};
            color: white;
            padding: 60px;
            display: flex;
            flex-direction: column;
            justify-content: center;
            align-items: center;
            text-align: center;
            min-height: 600px;
        }}
        .infographic .big-number {{ font-size: 100pt; font-weight: 300; }}
        .infographic .line1 {{ font-size: 20pt; margin-top: 20px; }}
        .infographic .line2 {{ font-size: 20pt; font-weight: bold; color: #{accentColor}; }}
        .infographic .source {{ font-size: 9pt; color: #aaa; font-style: italic; margin-top: 60px; }}
        
        /* Content Pages */
        .content {{ padding: 40px 50px; }}
        
        h1 {{ color: #{primaryColor}; font-size: 28pt; font-weight: 300; margin-bottom: 20px; border: none; }}
        h2 {{ color: #{primaryColor}; font-size: 16pt; font-weight: 600; margin-top: 30px; margin-bottom: 10px; border: none; }}
        h3 {{ color: #374151; font-size: 12pt; font-weight: 600; margin-top: 20px; }}
        
        .intro {{ color: #666; font-size: 10pt; margin-bottom: 15px; line-height: 1.5; }}
        
        /* KPI Cards */
        .kpi-row {{ display: flex; gap: 15px; margin: 20px 0; }}
        .kpi-card {{ flex: 1; background: #F3F4F6; padding: 20px; text-align: center; }}
        .kpi-card .value {{ font-size: 32pt; font-weight: 300; color: #{primaryColor}; }}
        .kpi-card .label {{ font-size: 11pt; font-weight: 600; margin-top: 5px; }}
        .kpi-card .sublabel {{ font-size: 9pt; color: #666; }}
        
        /* Tables */
        .section {{ margin-bottom: 25px; }}
        table {{ border-collapse: collapse; width: 100%; margin-top: 10px; font-size: 10pt; }}
        th, td {{ border: 1px solid #E5E7EB; padding: 8px 10px; text-align: left; }}
        th {{ background-color: #F9FAFB; color: #374151; font-weight: 600; }}
        tr:nth-child(even) {{ background-color: #FAFAFA; }}
        
        /* Status colors */
        .good, .compliant {{ color: #107C6C; }}
        .warning {{ color: #F59E0B; }}
        .critical {{ color: #DC2626; }}
        
        /* Footer */
        .footer {{ margin-top: 40px; padding-top: 20px; border-top: 2px solid #{primaryColor}; font-size: 9pt; color: #666; font-style: italic; }}
        
        /* Print button */
        .print-btn {{
            position: fixed;
            top: 20px;
            right: 20px;
            background: #{primaryColor};
            color: white;
            border: none;
            padding: 12px 24px;
            font-size: 14px;
            cursor: pointer;
            border-radius: 5px;
            z-index: 1000;
        }}
        .print-btn:hover {{ opacity: 0.9; }}
    </style>
</head>
<body>
    <button class='print-btn no-print' onclick='window.print()'>Save as PDF</button>
    
    <!-- Cover Page -->
    <div class='cover-page'>
        <div class='subtitle'>MICROSOFT 365</div>
        <div class='title'>{settings.ReportTitle.ToUpper().Replace("MICROSOFT 365 ", "")}</div>
    </div>
    <div class='cover-info'>
        <div class='company'>{settings.CompanyName}</div>
        <div class='date'>Generated on {data.GeneratedAt:d MMMM yyyy}</div>
        {(settings.LogoBase64 != null ? $"<img src='data:{settings.LogoContentType};base64,{settings.LogoBase64}' class='logo' alt='Logo' />" : "")}
    </div>
    
    <!-- Executive Summary -->
    <div class='content page-break'>
        <h1>Executive Summary</h1>
        
        <div class='kpi-row'>
            <div class='kpi-card'>
                <div class='value'>{data.UserStats?.TotalUsers ?? 0}</div>
                <div class='label'>Total Users</div>
                <div class='sublabel'>Including {data.UserStats?.GuestUsers ?? 0} guest users</div>
            </div>
            <div class='kpi-card'>
                <div class='value'>{(data.UserStats?.TotalUsers ?? 0) - (data.UserStats?.GuestUsers ?? 0)}</div>
                <div class='label'>Licensed Users</div>
                <div class='sublabel'>{data.UserStats?.GuestUsers ?? 0} unlicensed users</div>
            </div>
            <div class='kpi-card'>
                <div class='value'>{data.UserStats?.MfaRegistered ?? 0}</div>
                <div class='label'>MFA Registered</div>
                <div class='sublabel'>{data.UserStats?.MfaNotRegistered ?? 0} not registered</div>
            </div>
        </div>
        
        <p class='intro'>This {settings.ReportTitle} for {data.GeneratedAt:MMMM yyyy} provides a comprehensive analysis of your organization's security configuration across key Microsoft 365 services, including Entra ID (Azure AD), Exchange Online, Intune, SharePoint, and Teams.</p>
        
        <p class='intro'>The aim of this review is to provide a clear and actionable understanding of your current security posture within Microsoft 365, helping to mitigate potential risks, safeguard sensitive data, and ensure compliance with leading security benchmarks.</p>
    </div>
    
    {(showQuotes && quotes.Count > 0 ? RenderHtmlInfoGraphic(quotes[0]) : "")}
    
    <!-- Security Metrics -->
    <div class='content page-break'>
    <div class='section'>
        <h2>Microsoft Secure Score</h2>
        <p class='intro'>Microsoft Secure Score is a measurement of an organization's security posture, with a higher number indicating more improvement actions taken.</p>
        <table>
            <tr><th>Metric</th><th>Score</th><th>Max</th><th>Percentage</th></tr>
            <tr>
                <td>Microsoft Secure Score</td>
                <td>{data.SecureScore?.CurrentScore ?? 0}</td>
                <td>{data.SecureScore?.MaxScore ?? 0}</td>
                <td class='{GetScoreClass(data.SecureScore?.PercentageScore ?? 0)}'>{data.SecureScore?.PercentageScore ?? 0}%</td>
            </tr>
        </table>
    </div>

    <div class='section'>
        <h2>Intune Managed Devices</h2>
        <table>
            <tr><th>Platform</th><th>Count</th></tr>
            <tr><td>Total Devices</td><td>{data.DeviceStats?.TotalDevices ?? 0}</td></tr>
            <tr><td>Windows</td><td>{data.DeviceStats?.WindowsDevices ?? 0}</td></tr>
            <tr><td>macOS</td><td>{data.DeviceStats?.MacOsDevices ?? 0}</td></tr>
            <tr><td>iOS/iPadOS</td><td>{data.DeviceStats?.IosDevices ?? 0}</td></tr>
            <tr><td>Android</td><td>{data.DeviceStats?.AndroidDevices ?? 0}</td></tr>
            <tr><td>Compliant</td><td class='good'>{data.DeviceStats?.CompliantDevices ?? 0}</td></tr>
            <tr><td>Non-Compliant</td><td class='critical'>{data.DeviceStats?.NonCompliantDevices ?? 0}</td></tr>
            <tr><td>Compliance Rate</td><td>{data.DeviceStats?.ComplianceRate ?? 0}%</td></tr>
        </table>
    </div>

    <div class='section'>
        <h2>Windows Patch Status</h2>
        <table>
            <tr><th>Status</th><th>Count</th></tr>
            <tr><td>Total Windows Devices</td><td>{data.WindowsUpdateStats?.TotalWindowsDevices ?? 0}</td></tr>
            <tr><td>Up to Date</td><td class='good'>{data.WindowsUpdateStats?.UpToDate ?? 0}</td></tr>
            <tr><td>Needs Update</td><td class='warning'>{data.WindowsUpdateStats?.NeedsUpdate ?? 0}</td></tr>
            <tr><td>Compliance Rate</td><td>{data.WindowsUpdateStats?.ComplianceRate ?? 0}%</td></tr>
        </table>
        {(data.WindowsUpdateStats?.Note != null ? $"<p><em>{data.WindowsUpdateStats.Note}</em></p>" : "")}
    </div>

    <div class='section'>
        <h2>Microsoft Defender for Endpoint</h2>
        <table>
            <tr><th>Metric</th><th>Value</th></tr>
            <tr><td>Exposure Score</td><td>{data.DefenderStats?.ExposureScore ?? "N/A"}{(data.DefenderStats?.ExposureScoreNumeric.HasValue == true ? $" ({data.DefenderStats.ExposureScoreNumeric})" : "")}</td></tr>
            <tr><td>Onboarded Machines</td><td>{data.DefenderStats?.OnboardedMachines ?? 0}</td></tr>
            <tr><td>Total Vulnerabilities</td><td>{data.DefenderStats?.VulnerabilitiesDetected ?? 0}</td></tr>
            <tr><td>Critical</td><td class='critical'>{data.DefenderStats?.CriticalVulnerabilities ?? 0}</td></tr>
            <tr><td>High</td><td class='critical'>{data.DefenderStats?.HighVulnerabilities ?? 0}</td></tr>
            <tr><td>Medium</td><td class='warning'>{data.DefenderStats?.MediumVulnerabilities ?? 0}</td></tr>
            <tr><td>Low</td><td class='good'>{data.DefenderStats?.LowVulnerabilities ?? 0}</td></tr>
        </table>
        {(data.DefenderStats?.Note != null ? $"<p><em>{data.DefenderStats.Note}</em></p>" : "")}
    </div>

    <div class='section'>
        <h2>User Accounts</h2>
        <table>
            <tr><th>Type</th><th>Count</th></tr>
            <tr><td>Total Users</td><td>{data.UserStats?.TotalUsers ?? 0}</td></tr>
            <tr><td>Guest Users</td><td>{data.UserStats?.GuestUsers ?? 0}</td></tr>
            <tr><td>Deleted Users (Soft)</td><td>{data.UserStats?.DeletedUsers ?? 0}</td></tr>
            <tr><td>MFA Registered</td><td class='good'>{data.UserStats?.MfaRegistered ?? 0}</td></tr>
            <tr><td>MFA Not Registered</td><td class='warning'>{data.UserStats?.MfaNotRegistered ?? 0}</td></tr>
        </table>
        <p><strong>Risky Users:</strong> {data.RiskyUsersCount}</p>
        {(data.HighRiskUsers?.Any() == true ? $"<p class='critical'>High Risk: {string.Join(", ", data.HighRiskUsers)}</p>" : "")}
    </div>

    <div class='section'>
        <h2>Attack Simulation Training</h2>
        <table>
            <tr><th>Metric</th><th>Value</th></tr>
            <tr><td>Total Simulations</td><td>{data.AttackSimulation?.TotalSimulations ?? 0}</td></tr>
            <tr><td>Completed</td><td>{data.AttackSimulation?.CompletedSimulations ?? 0}</td></tr>
            <tr><td>Average Compromise Rate</td><td>{data.AttackSimulation?.AverageCompromiseRate ?? 0}%</td></tr>
        </table>
        {(data.AttackSimulation?.Note != null ? $"<p><em>{data.AttackSimulation.Note}</em></p>" : "")}
    </div>

    <div class='section'>
        <h2>Shadow IT</h2>
        <table>
            <tr><th>Metric</th><th>Count</th></tr>
            <tr><td>Discovered Apps</td><td>{data.CloudAppDiscovery?.DiscoveredApps ?? 0}</td></tr>
            <tr><td>Sanctioned</td><td class='good'>{data.CloudAppDiscovery?.SanctionedApps ?? 0}</td></tr>
            <tr><td>Unsanctioned</td><td class='warning'>{data.CloudAppDiscovery?.UnsanctionedApps ?? 0}</td></tr>
        </table>
        {(data.CloudAppDiscovery?.Note != null ? $"<p><em>{data.CloudAppDiscovery.Note}</em></p>" : "")}
    </div>

    <div class='section'>
        <h2>Mailbox Usage</h2>
        <table>
            <tr><th>Metric</th><th>Value</th></tr>
            <tr><td>Total Mailboxes</td><td>{data.MailboxStats?.TotalMailboxes ?? 0}</td></tr>
            <tr><td>Active Mailboxes</td><td>{data.MailboxStats?.ActiveMailboxes ?? 0}</td></tr>
            <tr><td>Total Storage Used</td><td>{data.MailboxStats?.TotalStorageUsedGB ?? 0} GB</td></tr>
            <tr><td>Average Storage</td><td>{data.MailboxStats?.AverageStorageGB ?? 0} GB</td></tr>
        </table>
    </div>

    <div class='section'>
        <h2>SharePoint Usage</h2>
        <table>
            <tr><th>Metric</th><th>Value</th></tr>
            <tr><td>Total Sites</td><td>{data.SharePointStats?.TotalSites ?? 0}</td></tr>
            <tr><td>Active Sites</td><td>{data.SharePointStats?.ActiveSites ?? 0}</td></tr>
            <tr><td>Total Storage Used</td><td>{data.SharePointStats?.TotalStorageUsedGB ?? 0} GB</td></tr>
        </table>
    </div>

    <div class='section'>
        <h2>Email Security (Last 30 Days)</h2>
        <table>
            <tr><th>Metric</th><th>Value</th></tr>
            <tr><td>Total Messages Processed</td><td>{data.EmailSecurity?.TotalMessages ?? 0:N0}</td></tr>
        </table>
        <p><em>{data.EmailSecurity?.Note ?? "View detailed threat protection statistics in the Microsoft 365 Defender portal."}</em></p>
    </div>

    </div>
    
    {(showQuotes && quotes.Count > 1 ? RenderHtmlInfoGraphic(quotes[1]) : "")}
    
    <div class='content page-break'>
    {GenerateUserSignInTable(data.UserSignInDetails)}

    {GenerateDeletedUsersTable(data.DeletedUsersInPeriod)}
    </div>
    
    {(showQuotes && quotes.Count > 2 ? RenderHtmlInfoGraphic(quotes[2]) : "")}
    
    <div class='content page-break'>
    {GenerateDomainSecuritySection(data.DomainSecuritySummary, data.DomainSecurityResults, settings.ExcludedDomains)}

    <div class='footer'>
        {(settings.FooterText != null ? $"<p>{settings.FooterText}</p>" : "")}
        <p>This report was automatically generated by M365 Dashboard.</p>
        <p>Some metrics may require additional licensing or API permissions.</p>
    </div>
    </div>
</body>
</html>";
    }

    private string GetScoreClass(double score)
    {
        if (score >= 70) return "good";
        if (score >= 50) return "warning";
        return "critical";
    }

    private string GenerateUserSignInTable(List<UserSignInDetailData>? users)
    {
        if (users == null || users.Count == 0)
            return "";

        var sb = new System.Text.StringBuilder();
        sb.AppendLine("<div class='section'>");
        sb.AppendLine($"<h2>User Sign-in & MFA Details ({users.Count} users)</h2>");
        sb.AppendLine("<table>");
        sb.AppendLine("<tr><th>Display Name</th><th>Email</th><th>Last Interactive Sign-in</th><th>Last Non-Interactive Sign-in</th><th>Default MFA Method</th><th>MFA Registered</th><th>Enabled</th></tr>");
        
        foreach (var user in users)
        {
            var lastInteractive = user.LastInteractiveSignIn?.ToString("dd MMM yyyy") ?? "Never";
            var lastNonInteractive = user.LastNonInteractiveSignIn?.ToString("dd MMM yyyy") ?? "Never";
            var mfaMethod = user.DefaultMfaMethod ?? "None";
            var mfaClass = user.IsMfaRegistered ? "good" : "critical";
            var mfaText = user.IsMfaRegistered ? "Yes" : "No";
            var enabledClass = user.AccountEnabled ? "good" : "";
            var enabledText = user.AccountEnabled ? "Yes" : "No";
            
            sb.AppendLine($"<tr>");
            sb.AppendLine($"<td>{user.DisplayName ?? "-"}</td>");
            sb.AppendLine($"<td style='font-size:11px'>{user.UserPrincipalName ?? "-"}</td>");
            sb.AppendLine($"<td>{lastInteractive}</td>");
            sb.AppendLine($"<td>{lastNonInteractive}</td>");
            sb.AppendLine($"<td>{mfaMethod}</td>");
            sb.AppendLine($"<td class='{mfaClass}'>{mfaText}</td>");
            sb.AppendLine($"<td class='{enabledClass}'>{enabledText}</td>");
            sb.AppendLine($"</tr>");
        }
        
        sb.AppendLine("</table>");
        sb.AppendLine("</div>");
        return sb.ToString();
    }

    private string GenerateDeletedUsersTable(List<DeletedUserData>? users)
    {
        if (users == null || users.Count == 0)
            return "<div class='section'><h2>Deleted Users in Period</h2><p>No users were deleted during this period.</p></div>";

        var sb = new System.Text.StringBuilder();
        sb.AppendLine("<div class='section'>");
        sb.AppendLine($"<h2>Deleted Users in Period ({users.Count} users)</h2>");
        sb.AppendLine("<table>");
        sb.AppendLine("<tr><th>Display Name</th><th>Email</th><th>Deleted Date</th><th>Job Title</th><th>Department</th></tr>");
        
        foreach (var user in users)
        {
            var deletedDate = user.DeletedDateTime?.ToString("dd MMM yyyy") ?? "-";
            
            sb.AppendLine($"<tr>");
            sb.AppendLine($"<td>{user.DisplayName ?? "-"}</td>");
            sb.AppendLine($"<td style='font-size:11px'>{user.UserPrincipalName ?? user.Mail ?? "-"}</td>");
            sb.AppendLine($"<td class='critical'>{deletedDate}</td>");
            sb.AppendLine($"<td>{user.JobTitle ?? "-"}</td>");
            sb.AppendLine($"<td>{user.Department ?? "-"}</td>");
            sb.AppendLine($"</tr>");
        }
        
        sb.AppendLine("</table>");
        sb.AppendLine("</div>");
        return sb.ToString();
    }

    private string GenerateDomainSecuritySection(DomainSecuritySummary? summary, List<DomainSecurityResult>? results, List<string>? excludedDomains = null)
    {
        if (summary == null) return "";

        var sb = new System.Text.StringBuilder();
        
        // Summary section
        sb.AppendLine("<div class='section'>");
        sb.AppendLine("<h2>Domain Email Security</h2>");
        sb.AppendLine("<table>");
        sb.AppendLine("<tr><th>Metric</th><th>Count</th></tr>");
        sb.AppendLine($"<tr><td>Total Domains Checked</td><td>{summary.TotalDomains}</td></tr>");
        sb.AppendLine($"<tr><td>Domains with MX Records</td><td>{summary.DomainsWithMx}</td></tr>");
        sb.AppendLine($"<tr><td>Domains with SPF</td><td class='good'>{summary.DomainsWithSpf}</td></tr>");
        sb.AppendLine($"<tr><td>Domains with DMARC</td><td class='good'>{summary.DomainsWithDmarc}</td></tr>");
        sb.AppendLine($"<tr><td>Domains with DKIM</td><td class='good'>{summary.DomainsWithDkim}</td></tr>");
        sb.AppendLine($"<tr><td>Domains with MTA-STS</td><td>{summary.DomainsWithMtaSts}</td></tr>");
        sb.AppendLine("</table>");
        
        // DMARC Policy Distribution
        sb.AppendLine("<h3>DMARC Policy Distribution</h3>");
        sb.AppendLine("<table>");
        sb.AppendLine("<tr><th>Policy</th><th>Count</th></tr>");
        sb.AppendLine($"<tr><td>Reject (Full Protection)</td><td class='good'>{summary.DmarcRejectCount}</td></tr>");
        sb.AppendLine($"<tr><td>Quarantine</td><td class='warning'>{summary.DmarcQuarantineCount}</td></tr>");
        sb.AppendLine($"<tr><td>None (Monitor Only)</td><td class='critical'>{summary.DmarcNoneCount}</td></tr>");
        sb.AppendLine("</table>");
        
        // Security Grade Distribution
        sb.AppendLine("<h3>Security Grade Distribution</h3>");
        sb.AppendLine("<table>");
        sb.AppendLine("<tr><th>Grade</th><th>Count</th></tr>");
        sb.AppendLine($"<tr><td>A (90-100)</td><td class='good'>{summary.GradeACount}</td></tr>");
        sb.AppendLine($"<tr><td>B (80-89)</td><td class='good'>{summary.GradeBCount}</td></tr>");
        sb.AppendLine($"<tr><td>C (70-79)</td><td class='warning'>{summary.GradeCCount}</td></tr>");
        sb.AppendLine($"<tr><td>D (60-69)</td><td class='critical'>{summary.GradeDCount}</td></tr>");
        sb.AppendLine($"<tr><td>F (Below 60)</td><td class='critical'>{summary.GradeFCount}</td></tr>");
        sb.AppendLine("</table>");
        
        if (summary.CriticalIssuesCount > 0)
            sb.AppendLine($"<p class='critical'><strong>{summary.CriticalIssuesCount} domains require immediate attention (Grade D or F)</strong></p>");
        
        sb.AppendLine("</div>");
        
        // Domain details table
        if (results?.Any() == true)
        {
            sb.AppendLine("<div class='section'>");
            sb.AppendLine($"<h2>Domain Security Details ({results.Count} domains)</h2>");
            sb.AppendLine("<table>");
            sb.AppendLine("<tr><th>Domain</th><th>MX</th><th>SPF</th><th>DMARC</th><th>DKIM</th><th>MTA-STS</th></tr>");
            
            foreach (var domain in results.Where(d => excludedDomains == null || !excludedDomains.Contains(d.Domain, StringComparer.OrdinalIgnoreCase)).OrderByDescending(d => d.SecurityScore))
            {
                var gradeClass = domain.SecurityGrade switch { "A" or "B" => "good", "C" => "warning", _ => "critical" };
                sb.AppendLine("<tr>");
                sb.AppendLine($"<td>{domain.Domain}</td>");
                sb.AppendLine($"<td>{(domain.HasMx ? "✓" : "✗")}</td>");
                sb.AppendLine($"<td>{(domain.HasSpf ? (domain.SpfPolicy == "-all" ? "✓ Hard" : "~ Soft") : "✗")}</td>");
                sb.AppendLine($"<td>{(domain.HasDmarc ? domain.DmarcPolicy : "✗")}</td>");
                sb.AppendLine($"<td>{(domain.HasDkim ? "✓" : "✗")}</td>");
                sb.AppendLine($"<td>{(domain.HasMtaSts ? "✓" : "✗")}</td>");
                sb.AppendLine("</tr>");
            }
            
            sb.AppendLine("</table>");
            
            // Critical domains
            var criticalDomains = results.Where(d => d.SecurityGrade == "D" || d.SecurityGrade == "F").ToList();
            if (criticalDomains.Any())
            {
                sb.AppendLine("<h3 class='critical'>Domains Requiring Immediate Attention</h3>");
                sb.AppendLine("<ul>");
                foreach (var domain in criticalDomains)
                {
                    var issues = string.Join(", ", domain.Issues ?? new List<string>());
                    sb.AppendLine($"<li><strong>{domain.Domain}</strong>: {issues}</li>");
                }
                sb.AppendLine("</ul>");
            }
            
            sb.AppendLine("</div>");
        }
        
        return sb.ToString();
    }
}

// Data models for the report
public class ExecutiveReportData
{
    public string ReportMonth { get; set; } = string.Empty;
    public DateTime GeneratedAt { get; set; }
    public DateTime StartDate { get; set; }
    public DateTime EndDate { get; set; }
    public SecureScoreData? SecureScore { get; set; }
    public DeviceStatsData? DeviceStats { get; set; }
    public UserStatsData? UserStats { get; set; }
    public DefenderStatsData? DefenderStats { get; set; }
    public MailboxStatsData? MailboxStats { get; set; }
    public SharePointStatsData? SharePointStats { get; set; }
    public AttackSimulationData? AttackSimulation { get; set; }
    public EmailSecurityData? EmailSecurity { get; set; }
    public WindowsUpdateStatsData? WindowsUpdateStats { get; set; }
    public CloudAppDiscoveryData? CloudAppDiscovery { get; set; }
    public int RiskyUsersCount { get; set; }
    public List<string>? HighRiskUsers { get; set; }
    public List<UserSignInDetailData>? UserSignInDetails { get; set; }
    public List<DeletedUserData>? DeletedUsersInPeriod { get; set; }
    public List<MailboxDetailData>? MailboxDetails { get; set; }
    public DeviceDetailsData? DeviceDetails { get; set; }
    public AppCredentialStatusData? AppCredentialStatus { get; set; }
    public List<DomainSecurityResult>? DomainSecurityResults { get; set; }
    public DomainSecuritySummary? DomainSecuritySummary { get; set; }
}

public class SecureScoreData
{
    public double CurrentScore { get; set; }
    public double MaxScore { get; set; }
    public double PercentageScore { get; set; }
}

public class DeviceStatsData
{
    public int TotalDevices { get; set; }
    public int WindowsDevices { get; set; }
    public int MacOsDevices { get; set; }
    public int IosDevices { get; set; }
    public int AndroidDevices { get; set; }
    public int CompliantDevices { get; set; }
    public int NonCompliantDevices { get; set; }
    public double ComplianceRate { get; set; }
}

public class UserStatsData
{
    public int TotalUsers { get; set; }
    public int GuestUsers { get; set; }
    public int DeletedUsers { get; set; }
    public int MfaRegistered { get; set; }
    public int MfaNotRegistered { get; set; }
}

public class DefenderStatsData
{
    public string? ExposureScore { get; set; }
    public double? ExposureScoreNumeric { get; set; }
    public int VulnerabilitiesDetected { get; set; }
    public int CriticalVulnerabilities { get; set; }
    public int HighVulnerabilities { get; set; }
    public int MediumVulnerabilities { get; set; }
    public int LowVulnerabilities { get; set; }
    public int? OnboardedMachines { get; set; }
    public string? Note { get; set; }
}

public class MailboxStatsData
{
    public int TotalMailboxes { get; set; }
    public int ActiveMailboxes { get; set; }
    public double TotalStorageUsedGB { get; set; }
    public double AverageStorageGB { get; set; }
}

public class WindowsDeviceDetailData
{
    public string? DeviceName { get; set; }
    public DateTime? LastCheckIn { get; set; }
    public string? OsVersion { get; set; }
    public string? ComplianceState { get; set; }
    public string? ManagementAgent { get; set; }
    public string? Ownership { get; set; }
    public string? SkuFamily { get; set; }
    public VersionStatus OsVersionStatus { get; set; } = VersionStatus.Unknown;
    public string? OsVersionStatusMessage { get; set; }
    public string? LatestVersion { get; set; }
}

public class IosDeviceDetailData
{
    public string? DeviceName { get; set; }
    public string? ComplianceState { get; set; }
    public string? ManagementAgent { get; set; }
    public string? Ownership { get; set; }
    public string? Os { get; set; }
    public string? OsVersion { get; set; }
    public DateTime? LastCheckIn { get; set; }
    public VersionStatus OsVersionStatus { get; set; } = VersionStatus.Unknown;
    public string? OsVersionStatusMessage { get; set; }
    public string? LatestVersion { get; set; }
}

public class AndroidDeviceDetailData
{
    public string? DeviceName { get; set; }
    public string? ComplianceState { get; set; }
    public string? ManagementAgent { get; set; }
    public string? Os { get; set; }
    public string? OsVersion { get; set; }
    public DateTime? LastCheckIn { get; set; }
    public string? SecurityPatchLevel { get; set; }
    public VersionStatus OsVersionStatus { get; set; } = VersionStatus.Unknown;
    public string? OsVersionStatusMessage { get; set; }
    public string? LatestVersion { get; set; }
}

public class MacDeviceDetailData
{
    public string? DeviceName { get; set; }
    public DateTime? LastCheckIn { get; set; }
    public string? OsVersion { get; set; }
    public string? ComplianceState { get; set; }
    public string? ManagementAgent { get; set; }
    public string? Ownership { get; set; }
    public VersionStatus OsVersionStatus { get; set; } = VersionStatus.Unknown;
    public string? OsVersionStatusMessage { get; set; }
    public string? LatestVersion { get; set; }
}

public class DeviceDetailsData
{
    public List<WindowsDeviceDetailData> WindowsDevices { get; set; } = new();
    public List<IosDeviceDetailData> IosDevices { get; set; } = new();
    public List<AndroidDeviceDetailData> AndroidDevices { get; set; } = new();
    public List<MacDeviceDetailData> MacDevices { get; set; } = new();
}

public class SharePointStatsData
{
    public int TotalSites { get; set; }
    public int ActiveSites { get; set; }
    public double TotalStorageUsedGB { get; set; }
}

public class AttackSimulationData
{
    public int TotalSimulations { get; set; }
    public int CompletedSimulations { get; set; }
    public double AverageCompromiseRate { get; set; }
    public string? Note { get; set; }
}

public class EmailSecurityData
{
    public int TotalMessages { get; set; }
    public int SpamMessages { get; set; }
    public int MalwareMessages { get; set; }
    public int PhishingMessages { get; set; }
    public int BulkMessages { get; set; }
    public string? Note { get; set; }
}

public class WindowsUpdateStatsData
{
    public int TotalWindowsDevices { get; set; }
    public int UpToDate { get; set; }
    public int NeedsUpdate { get; set; }
    public double ComplianceRate { get; set; }
    public string? Note { get; set; }
}

public class CloudAppDiscoveryData
{
    public int DiscoveredApps { get; set; }
    public int SanctionedApps { get; set; }
    public int UnsanctionedApps { get; set; }
    public string? Note { get; set; }
}

public class UserSignInDetailData
{
    public string? DisplayName { get; set; }
    public string? UserPrincipalName { get; set; }
    public DateTime? LastInteractiveSignIn { get; set; }
    public DateTime? LastNonInteractiveSignIn { get; set; }
    public string? DefaultMfaMethod { get; set; }
    public bool IsMfaRegistered { get; set; }
    public bool AccountEnabled { get; set; }
}

public class DeletedUserData
{
    public string? DisplayName { get; set; }
    public string? UserPrincipalName { get; set; }
    public string? Mail { get; set; }
    public DateTime? DeletedDateTime { get; set; }
    public string? JobTitle { get; set; }
    public string? Department { get; set; }
}

public class MailboxDetailData
{
    public string? DisplayName { get; set; }
    public string? UserPrincipalName { get; set; }
    public string? RecipientType { get; set; }
    public long StorageUsedBytes { get; set; }
    public double StorageUsedGB { get; set; }
    public long? QuotaBytes { get; set; }
    public double? QuotaGB { get; set; }
    public double? PercentUsed { get; set; }
    public DateTime? LastActivityDate { get; set; }
    public int? ItemCount { get; set; }
}

public class AppCredentialStatusData
{
    public int TotalApps { get; set; }
    public int AppsWithExpiringSecrets { get; set; }
    public int AppsWithExpiredSecrets { get; set; }
    public int AppsWithExpiringCertificates { get; set; }
    public int AppsWithExpiredCertificates { get; set; }
    public int ThresholdDays { get; set; } = 45;
    public List<AppCredentialDetail> ExpiringSecrets { get; set; } = new();
    public List<AppCredentialDetail> ExpiredSecrets { get; set; } = new();
    public List<AppCredentialDetail> ExpiringCertificates { get; set; } = new();
    public List<AppCredentialDetail> ExpiredCertificates { get; set; } = new();
}

public class AppCredentialDetail
{
    public string? AppName { get; set; }
    public string? AppId { get; set; }
    public string? CredentialType { get; set; }
    public string? Description { get; set; }
    public DateTime? ExpiryDate { get; set; }
    public int DaysUntilExpiry { get; set; }
    public string? Status { get; set; }
}
