using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using M365Dashboard.Api.Models.Dtos;

namespace M365Dashboard.Api.Controllers;

[ApiController]
[Route("api/[controller]")]
[Authorize]
public class SignInsController : ControllerBase
{
    private readonly GraphServiceClient _graphClient;
    private readonly ILogger<SignInsController> _logger;

    public SignInsController(GraphServiceClient graphClient, ILogger<SignInsController> logger)
    {
        _graphClient = graphClient;
        _logger = logger;
    }

    /// <summary>
    /// Get sign-ins with location data for map visualization
    /// </summary>
    [HttpGet("map")]
    public async Task<IActionResult> GetSignInsForMap([FromQuery] int hours = 24)
    {
        try
        {
            _logger.LogInformation("Fetching sign-ins for map visualization, last {Hours} hours", hours);

            var cutoffTime = DateTime.UtcNow.AddHours(-hours);
            var allSignIns = new List<SignIn>();

            var signIns = await _graphClient.AuditLogs.SignIns.GetAsync(config =>
            {
                config.QueryParameters.Filter = $"createdDateTime ge {cutoffTime:yyyy-MM-ddTHH:mm:ssZ}";
                config.QueryParameters.Top = 999;
                config.QueryParameters.Select = new[]
                {
                    "id", "createdDateTime", "userPrincipalName", "userDisplayName",
                    "ipAddress", "location", "status", "clientAppUsed",
                    "deviceDetail", "riskLevelDuringSignIn", "riskState",
                    "riskEventTypes",
                    "conditionalAccessStatus", "isInteractive"
                };
                config.QueryParameters.Orderby = new[] { "createdDateTime desc" };
            });

            if (signIns?.Value != null)
            {
                allSignIns.AddRange(signIns.Value);
            }

            // Page through results (limit to reasonable amount for performance)
            var maxPages = 5;
            var currentPage = 1;
            while (signIns?.OdataNextLink != null && currentPage < maxPages)
            {
                signIns = await _graphClient.AuditLogs.SignIns
                    .WithUrl(signIns.OdataNextLink)
                    .GetAsync();
                if (signIns?.Value != null)
                {
                    allSignIns.AddRange(signIns.Value);
                }
                currentPage++;
            }

            _logger.LogInformation("Retrieved {Count} sign-ins", allSignIns.Count);

            // Convert to DTOs with location data
            var signInDetails = allSignIns
                .Where(s => s.Location?.GeoCoordinates?.Latitude != null && 
                            s.Location?.GeoCoordinates?.Longitude != null)
                .Select(s => new SignInDetailDto(
                    Id: s.Id ?? string.Empty,
                    UserPrincipalName: s.UserPrincipalName ?? "Unknown",
                    DisplayName: s.UserDisplayName,
                    CreatedDateTime: s.CreatedDateTime?.DateTime,
                    IpAddress: s.IpAddress,
                    City: s.Location?.City,
                    State: s.Location?.State,
                    CountryOrRegion: s.Location?.CountryOrRegion,
                    Latitude: s.Location?.GeoCoordinates?.Latitude,
                    Longitude: s.Location?.GeoCoordinates?.Longitude,
                    IsSuccess: s.Status?.ErrorCode == 0,
                    ErrorCode: s.Status?.ErrorCode,
                    FailureReason: s.Status?.FailureReason,
                    ClientAppUsed: s.ClientAppUsed,
                    Browser: s.DeviceDetail?.Browser,
                    OperatingSystem: s.DeviceDetail?.OperatingSystem,
                    DeviceDisplayName: s.DeviceDetail?.DisplayName,
                    IsCompliant: s.DeviceDetail?.IsCompliant,
                    IsManaged: s.DeviceDetail?.IsManaged,
                    RiskLevel: s.RiskLevelDuringSignIn?.ToString(),
                    RiskState: s.RiskState?.ToString(),
                    MfaRequired: null,
                    ConditionalAccessStatus: s.ConditionalAccessStatus?.ToString(),
                    RiskEventTypes: s.RiskEventTypesV2?.Count > 0 ? s.RiskEventTypesV2 : s.RiskEventTypes?.Select(r => r.ToString()).ToList()
                ))
                .ToList();

            // Group by location (rounded to avoid too many unique points)
            var locations = signInDetails
                .Where(s => s.Latitude.HasValue && s.Longitude.HasValue)
                .GroupBy(s => new
                {
                    // Round to 2 decimal places to cluster nearby sign-ins
                    Lat = Math.Round(s.Latitude!.Value, 2),
                    Lon = Math.Round(s.Longitude!.Value, 2),
                    City = s.City ?? "Unknown",
                    Country = s.CountryOrRegion ?? "Unknown"
                })
                .Select(g => new SignInLocationDto(
                    Latitude: (double)g.Key.Lat,
                    Longitude: (double)g.Key.Lon,
                    City: g.Key.City,
                    State: g.First().State,
                    CountryOrRegion: g.Key.Country,
                    SignInCount: g.Count(),
                    SuccessCount: g.Count(s => s.IsSuccess),
                    FailureCount: g.Count(s => !s.IsSuccess),
                    SignIns: g.Take(10).ToList() // Limit details per location
                ))
                .OrderByDescending(l => l.SignInCount)
                .ToList();

            var result = new SignInsMapDataDto(
                Locations: locations,
                TotalSignIns: allSignIns.Count,
                SuccessfulSignIns: allSignIns.Count(s => s.Status?.ErrorCode == 0),
                FailedSignIns: allSignIns.Count(s => s.Status?.ErrorCode != 0),
                UniqueUsers: allSignIns.Select(s => s.UserPrincipalName).Distinct().Count(),
                UniqueLocations: locations.Count,
                StartDate: cutoffTime,
                EndDate: DateTime.UtcNow,
                LastUpdated: DateTime.UtcNow
            );

            return Ok(result);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching sign-ins for map");
            return StatusCode(500, new { error = "Failed to fetch sign-in data", message = ex.Message });
        }
    }

    /// <summary>
    /// Get detailed sign-in list with pagination
    /// </summary>
    [HttpGet]
    public async Task<IActionResult> GetSignIns(
        [FromQuery] int hours = 24,
        [FromQuery] int take = 100,
        [FromQuery] string? status = null)
    {
        try
        {
            _logger.LogInformation("Fetching sign-ins list, last {Hours} hours", hours);

            var cutoffTime = DateTime.UtcNow.AddHours(-hours);

            var filterParts = new List<string>
            {
                $"createdDateTime ge {cutoffTime:yyyy-MM-ddTHH:mm:ssZ}"
            };

            if (status?.ToLower() == "success")
            {
                filterParts.Add("status/errorCode eq 0");
            }
            else if (status?.ToLower() == "failure")
            {
                filterParts.Add("status/errorCode ne 0");
            }

            var signIns = await _graphClient.AuditLogs.SignIns.GetAsync(config =>
            {
                config.QueryParameters.Filter = string.Join(" and ", filterParts);
                config.QueryParameters.Top = take;
                config.QueryParameters.Orderby = new[] { "createdDateTime desc" };
            });

            var signInList = signIns?.Value ?? new List<SignIn>();

            var result = signInList.Select(s => new SignInDetailDto(
                Id: s.Id ?? string.Empty,
                UserPrincipalName: s.UserPrincipalName ?? "Unknown",
                DisplayName: s.UserDisplayName,
                CreatedDateTime: s.CreatedDateTime?.DateTime,
                IpAddress: s.IpAddress,
                City: s.Location?.City,
                State: s.Location?.State,
                CountryOrRegion: s.Location?.CountryOrRegion,
                Latitude: s.Location?.GeoCoordinates?.Latitude,
                Longitude: s.Location?.GeoCoordinates?.Longitude,
                IsSuccess: s.Status?.ErrorCode == 0,
                ErrorCode: s.Status?.ErrorCode,
                FailureReason: s.Status?.FailureReason,
                ClientAppUsed: s.ClientAppUsed,
                Browser: s.DeviceDetail?.Browser,
                OperatingSystem: s.DeviceDetail?.OperatingSystem,
                DeviceDisplayName: s.DeviceDetail?.DisplayName,
                IsCompliant: s.DeviceDetail?.IsCompliant,
                IsManaged: s.DeviceDetail?.IsManaged,
                RiskLevel: s.RiskLevelDuringSignIn?.ToString(),
                RiskState: s.RiskState?.ToString(),
                MfaRequired: null,
                ConditionalAccessStatus: s.ConditionalAccessStatus?.ToString(),
                RiskEventTypes: s.RiskEventTypesV2?.Count > 0 ? s.RiskEventTypesV2 : s.RiskEventTypes?.Select(r => r.ToString()).ToList()
            )).ToList();

            return Ok(new
            {
                signIns = result,
                totalCount = result.Count,
                startDate = cutoffTime,
                endDate = DateTime.UtcNow
            });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching sign-ins list");
            return StatusCode(500, new { error = "Failed to fetch sign-in data", message = ex.Message });
        }
    }

    /// <summary>
    /// Get sign-ins flagged with anonymized IP (VPN/Proxy/Tor) risk event
    /// </summary>
    [HttpGet("vpn-proxy")]
    public async Task<IActionResult> GetVpnProxySignIns([FromQuery] int hours = 24, [FromQuery] int take = 20)
    {
        try
        {
            _logger.LogInformation("Fetching VPN/proxy sign-ins, last {Hours} hours", hours);

            var cutoffTime = DateTime.UtcNow.AddHours(-hours);

            var signIns = await _graphClient.AuditLogs.SignIns.GetAsync(config =>
            {
                config.QueryParameters.Filter =
                    $"createdDateTime ge {cutoffTime:yyyy-MM-ddTHH:mm:ssZ} and riskEventTypes_v2/any(r: r eq 'anonymizedIPAddress')";
                config.QueryParameters.Top = take;
                config.QueryParameters.Select = new[]
                {
                    "id", "createdDateTime", "userPrincipalName", "userDisplayName",
                    "ipAddress", "location", "status", "clientAppUsed",
                    "riskLevelDuringSignIn", "riskState", "riskEventTypes"
                };
                config.QueryParameters.Orderby = new[] { "createdDateTime desc" };
            });

            var result = (signIns?.Value ?? new List<SignIn>()).Select(s => new
            {
                id                 = s.Id,
                userPrincipalName  = s.UserPrincipalName ?? "Unknown",
                displayName        = s.UserDisplayName,
                createdDateTime    = s.CreatedDateTime?.DateTime,
                ipAddress          = s.IpAddress,
                city               = s.Location?.City,
                countryOrRegion    = s.Location?.CountryOrRegion,
                isSuccess          = s.Status?.ErrorCode == 0,
                failureReason      = s.Status?.FailureReason,
                clientAppUsed      = s.ClientAppUsed,
                riskLevel          = s.RiskLevelDuringSignIn?.ToString(),
                riskEventTypes     = s.RiskEventTypesV2?.Count > 0
                                     ? s.RiskEventTypesV2
                                     : s.RiskEventTypes?.Select(r => r.ToString()).ToList(),
            }).ToList();

            return Ok(new { signIns = result, totalCount = result.Count, hours, startDate = cutoffTime });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching VPN/proxy sign-ins");
            return StatusCode(500, new { error = "Failed to fetch VPN/proxy sign-in data", message = ex.Message });
        }
    }

    /// <summary>
    /// Get sign-in statistics summary
    /// </summary>
    [HttpGet("stats")]
    public async Task<IActionResult> GetSignInStats([FromQuery] int hours = 24)
    {
        try
        {
            var cutoffTime = DateTime.UtcNow.AddHours(-hours);

            var signIns = await _graphClient.AuditLogs.SignIns.GetAsync(config =>
            {
                config.QueryParameters.Filter = $"createdDateTime ge {cutoffTime:yyyy-MM-ddTHH:mm:ssZ}";
                config.QueryParameters.Top = 999;
                config.QueryParameters.Select = new[] { "status", "userPrincipalName", "location", "riskLevelDuringSignIn" };
            });

            var allSignIns = signIns?.Value ?? new List<SignIn>();

            // Page through for accurate stats
            while (signIns?.OdataNextLink != null)
            {
                signIns = await _graphClient.AuditLogs.SignIns
                    .WithUrl(signIns.OdataNextLink)
                    .GetAsync();
                if (signIns?.Value != null)
                {
                    allSignIns.AddRange(signIns.Value);
                }
            }

            var successful = allSignIns.Count(s => s.Status?.ErrorCode == 0);
            var failed = allSignIns.Count(s => s.Status?.ErrorCode != 0);
            var risky = allSignIns.Count(s =>
                s.RiskLevelDuringSignIn == RiskLevel.High ||
                s.RiskLevelDuringSignIn == RiskLevel.Medium);

            var topCountries = allSignIns
                .Where(s => s.Location?.CountryOrRegion != null)
                .GroupBy(s => s.Location!.CountryOrRegion)
                .OrderByDescending(g => g.Count())
                .Take(10)
                .Select(g => new { country = g.Key, count = g.Count() })
                .ToList();

            var topCities = allSignIns
                .Where(s => s.Location?.City != null)
                .GroupBy(s => $"{s.Location!.City}, {s.Location.CountryOrRegion}")
                .OrderByDescending(g => g.Count())
                .Take(10)
                .Select(g => new { location = g.Key, count = g.Count() })
                .ToList();

            return Ok(new
            {
                totalSignIns = allSignIns.Count,
                successfulSignIns = successful,
                failedSignIns = failed,
                riskySignIns = risky,
                successRate = allSignIns.Count > 0 ? Math.Round((double)successful / allSignIns.Count * 100, 1) : 0,
                uniqueUsers = allSignIns.Select(s => s.UserPrincipalName).Distinct().Count(),
                topCountries = topCountries,
                topCities = topCities,
                startDate = cutoffTime,
                endDate = DateTime.UtcNow
            });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching sign-in stats");
            return StatusCode(500, new { error = "Failed to fetch sign-in statistics", message = ex.Message });
        }
    }
}
