using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using M365Dashboard.Api.Models.Dtos;
using M365Dashboard.Api.Services;

namespace M365Dashboard.Api.Controllers;

[ApiController]
[Route("api/[controller]")]
[Authorize]
public class LicensesController : ControllerBase
{
    private readonly GraphServiceClient _graphClient;
    private readonly ISkuMappingService _skuMappingService;
    private readonly ILogger<LicensesController> _logger;

    public LicensesController(
        GraphServiceClient graphClient, 
        ISkuMappingService skuMappingService,
        ILogger<LicensesController> logger)
    {
        _graphClient = graphClient;
        _skuMappingService = skuMappingService;
        _logger = logger;
    }

    /// <summary>
    /// Get license overview with statistics
    /// </summary>
    [HttpGet("overview")]
    public async Task<IActionResult> GetOverview([FromQuery] bool excludeFreeTrial = true)
    {
        try
        {
            _logger.LogInformation("Fetching license overview (excludeFreeTrial: {ExcludeFreeTrial})", excludeFreeTrial);

            var allLicenses = await GetAllLicensesAsync();
            
            // Filter based on parameter
            var licenses = excludeFreeTrial 
                ? allLicenses.Where(l => !l.IsTrial).ToList() 
                : allLicenses;
            
            var totalLicenses = licenses.Sum(l => (long)l.TotalUnits);
            var assignedLicenses = licenses.Sum(l => (long)l.ConsumedUnits);
            var availableLicenses = licenses.Sum(l => (long)l.AvailableUnits);
            var warningLicenses = licenses.Count(l => l.WarningUnits > 0);
            
            var trialLicenses = allLicenses.Where(l => l.IsTrial).ToList();

            var overview = new
            {
                stats = new
                {
                    totalSubscriptions = licenses.Count(),
                    totalLicenses,
                    assignedLicenses,
                    availableLicenses,
                    utilizationPercentage = totalLicenses > 0 
                        ? Math.Round((double)assignedLicenses / totalLicenses * 100, 1) 
                        : 0,
                    subscriptionsWithWarnings = warningLicenses,
                    trialSubscriptions = trialLicenses.Count,
                    excludingFreeTrial = excludeFreeTrial
                },
                topUtilized = licenses
                    .Where(l => l.TotalUnits > 0)
                    .OrderByDescending(l => l.UtilizationPercentage)
                    .Take(5)
                    .ToList(),
                lowUtilization = licenses
                    .Where(l => l.TotalUnits > 0 && l.UtilizationPercentage < 50)
                    .OrderBy(l => l.UtilizationPercentage)
                    .Take(5)
                    .ToList(),
                recentlyAdded = licenses
                    .OrderByDescending(l => l.SkuId)
                    .Take(5)
                    .ToList(),
                lastUpdated = DateTime.UtcNow
            };

            return Ok(overview);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching license overview");
            return StatusCode(500, new { error = "Failed to fetch license overview", message = ex.Message });
        }
    }

    /// <summary>
    /// Get all subscribed SKUs (licenses)
    /// </summary>
    [HttpGet]
    public async Task<IActionResult> GetLicenses()
    {
        try
        {
            _logger.LogInformation("Fetching all licenses");

            var licenses = await GetAllLicensesAsync();

            return Ok(new
            {
                licenses,
                totalCount = licenses.Count,
                lastUpdated = DateTime.UtcNow
            });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching licenses");
            return StatusCode(500, new { error = "Failed to fetch licenses", message = ex.Message });
        }
    }

    /// <summary>
    /// Get license details including assigned users
    /// </summary>
    [HttpGet("{skuId}")]
    public async Task<IActionResult> GetLicenseDetails(string skuId)
    {
        try
        {
            _logger.LogInformation("Fetching license details for SKU {SkuId}", skuId);

            // Get the SKU
            var skus = await _graphClient.SubscribedSkus.GetAsync();
            var sku = skus?.Value?.FirstOrDefault(s => s.SkuId.ToString() == skuId);

            if (sku == null)
            {
                return NotFound(new { error = "License not found" });
            }

            // Get users with this license
            var users = await _graphClient.Users.GetAsync(config =>
            {
                config.QueryParameters.Select = new[] { "id", "displayName", "userPrincipalName", "assignedLicenses" };
                config.QueryParameters.Top = 999;
            });

            var assignedUsers = new List<object>();
            if (users?.Value != null)
            {
                foreach (var user in users.Value)
                {
                    if (user.AssignedLicenses?.Any(l => l.SkuId.ToString() == skuId) == true)
                    {
                        assignedUsers.Add(new
                        {
                            id = user.Id,
                            displayName = user.DisplayName,
                            userPrincipalName = user.UserPrincipalName
                        });
                    }
                }
            }

            var license = MapToLicenseDto(sku);

            return Ok(new
            {
                license,
                assignedUsers,
                assignedUsersCount = assignedUsers.Count,
                servicePlans = sku.ServicePlans?.Select(sp => new
                {
                    servicePlanId = sp.ServicePlanId,
                    servicePlanName = sp.ServicePlanName,
                    provisioningStatus = sp.ProvisioningStatus,
                    appliesTo = sp.AppliesTo
                }).ToList()
            });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching license details for SKU {SkuId}", skuId);
            return StatusCode(500, new { error = "Failed to fetch license details", message = ex.Message });
        }
    }

    /// <summary>
    /// Get license assignment summary by department or usage location
    /// </summary>
    [HttpGet("summary")]
    public async Task<IActionResult> GetLicenseSummary()
    {
        try
        {
            _logger.LogInformation("Fetching license summary");

            var users = await _graphClient.Users.GetAsync(config =>
            {
                config.QueryParameters.Select = new[] { "id", "displayName", "department", "usageLocation", "assignedLicenses" };
                config.QueryParameters.Top = 999;
            });

            var allUsers = new List<User>();
            if (users?.Value != null)
            {
                allUsers.AddRange(users.Value);

                while (users?.OdataNextLink != null)
                {
                    users = await _graphClient.Users.WithUrl(users.OdataNextLink).GetAsync();
                    if (users?.Value != null)
                    {
                        allUsers.AddRange(users.Value);
                    }
                }
            }

            // Group by department
            var byDepartment = allUsers
                .Where(u => !string.IsNullOrEmpty(u.Department) && u.AssignedLicenses?.Any() == true)
                .GroupBy(u => u.Department)
                .Select(g => new
                {
                    department = g.Key,
                    userCount = g.Count(),
                    totalLicenses = g.Sum(u => u.AssignedLicenses?.Count ?? 0)
                })
                .OrderByDescending(x => x.userCount)
                .Take(10)
                .ToList();

            // Group by usage location
            var byLocation = allUsers
                .Where(u => !string.IsNullOrEmpty(u.UsageLocation) && u.AssignedLicenses?.Any() == true)
                .GroupBy(u => u.UsageLocation)
                .Select(g => new
                {
                    location = g.Key,
                    userCount = g.Count(),
                    totalLicenses = g.Sum(u => u.AssignedLicenses?.Count ?? 0)
                })
                .OrderByDescending(x => x.userCount)
                .ToList();

            // Users without licenses
            var usersWithoutLicenses = allUsers.Count(u => u.AssignedLicenses == null || !u.AssignedLicenses.Any());

            return Ok(new
            {
                byDepartment,
                byLocation,
                totalUsers = allUsers.Count,
                usersWithLicenses = allUsers.Count - usersWithoutLicenses,
                usersWithoutLicenses,
                lastUpdated = DateTime.UtcNow
            });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching license summary");
            return StatusCode(500, new { error = "Failed to fetch license summary", message = ex.Message });
        }
    }

    /// <summary>
    /// Force refresh SKU mappings from Microsoft
    /// </summary>
    [HttpPost("refresh-mappings")]
    public async Task<IActionResult> RefreshMappings()
    {
        try
        {
            _logger.LogInformation("Manually refreshing SKU mappings");
            await _skuMappingService.RefreshMappingsAsync();
            return Ok(new { message = "SKU mappings refreshed successfully", status = _skuMappingService.GetStatus() });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error refreshing SKU mappings");
            return StatusCode(500, new { error = "Failed to refresh SKU mappings", message = ex.Message });
        }
    }

    /// <summary>
    /// Get SKU mapping service status
    /// </summary>
    [HttpGet("mapping-status")]
    public IActionResult GetMappingStatus()
    {
        return Ok(_skuMappingService.GetStatus());
    }

    #region Private Methods

    private async Task<List<LicenseDto>> GetAllLicensesAsync()
    {
        var skus = await _graphClient.SubscribedSkus.GetAsync();
        
        if (skus?.Value == null)
        {
            return new List<LicenseDto>();
        }

        return skus.Value.Select(MapToLicenseDto).ToList();
    }

    private LicenseDto MapToLicenseDto(SubscribedSku sku)
    {
        var totalUnits = sku.PrepaidUnits?.Enabled ?? 0;
        var consumedUnits = sku.ConsumedUnits ?? 0;
        var warningUnits = sku.PrepaidUnits?.Warning ?? 0;
        var suspendedUnits = sku.PrepaidUnits?.Suspended ?? 0;
        var availableUnits = Math.Max(0, totalUnits - consumedUnits);
        
        var utilizationPercentage = totalUnits > 0 
            ? Math.Round((double)consumedUnits / totalUnits * 100, 1) 
            : 0;

        var skuPartNumber = sku.SkuPartNumber ?? "";
        
        // Use the service to check if it's a free/trial license
        var isTrial = _skuMappingService.IsFreeTrial(skuPartNumber) ||
                      sku.CapabilityStatus == "Warning" ||
                      totalUnits >= 10000; // Licenses with 10,000+ units are typically free/viral

        return new LicenseDto(
            SkuId: sku.SkuId.ToString()!,
            SkuPartNumber: skuPartNumber,
            DisplayName: _skuMappingService.GetFriendlyName(skuPartNumber),
            TotalUnits: (int)totalUnits,
            ConsumedUnits: consumedUnits,
            AvailableUnits: (int)availableUnits,
            WarningUnits: (int)warningUnits,
            SuspendedUnits: (int)suspendedUnits,
            UtilizationPercentage: utilizationPercentage,
            Status: sku.CapabilityStatus ?? "Unknown",
            AppliesTo: sku.AppliesTo ?? "Unknown",
            IsTrial: isTrial,
            ServicePlanCount: sku.ServicePlans?.Count ?? 0
        );
    }

    #endregion
}
