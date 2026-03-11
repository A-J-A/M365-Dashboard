using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using M365Dashboard.Api.Models.Dtos;

namespace M365Dashboard.Api.Controllers;

[ApiController]
[Route("api/[controller]")]
[Authorize]
public class SharePointController : ControllerBase
{
    private readonly GraphServiceClient _graphClient;
    private readonly ILogger<SharePointController> _logger;

    public SharePointController(GraphServiceClient graphClient, ILogger<SharePointController> logger)
    {
        _graphClient = graphClient;
        _logger = logger;
    }

    /// <summary>
    /// Get SharePoint overview with statistics and key insights
    /// </summary>
    [HttpGet("overview")]
    public async Task<IActionResult> GetOverview()
    {
        try
        {
            _logger.LogInformation("Fetching SharePoint overview");

            var allSites = await GetAllSitesAsync();
            var stats = CalculateStats(allSites);

            var largestSites = allSites
                .OrderByDescending(s => s.StorageUsedBytes)
                .Take(10)
                .ToList();

            var recentlyCreated = allSites
                .Where(s => s.CreatedDateTime.HasValue)
                .OrderByDescending(s => s.CreatedDateTime)
                .Take(10)
                .ToList();

            var sitesNearLimit = allSites
                .Where(s => s.StorageUsedPercentage >= 80)
                .OrderByDescending(s => s.StorageUsedPercentage)
                .Take(10)
                .ToList();

            return Ok(new SharePointOverviewDto(
                Stats: stats,
                LargestSites: largestSites,
                RecentlyCreatedSites: recentlyCreated,
                SitesNearStorageLimit: sitesNearLimit,
                LastUpdated: DateTime.UtcNow
            ));
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching SharePoint overview");
            return StatusCode(500, new { error = "Failed to fetch SharePoint overview", message = ex.Message });
        }
    }

    /// <summary>
    /// Get SharePoint statistics
    /// </summary>
    [HttpGet("stats")]
    public async Task<IActionResult> GetStats()
    {
        try
        {
            _logger.LogInformation("Fetching SharePoint statistics");

            var allSites = await GetAllSitesAsync();
            var stats = CalculateStats(allSites);

            return Ok(stats);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching SharePoint stats");
            return StatusCode(500, new { error = "Failed to fetch SharePoint statistics", message = ex.Message });
        }
    }

    /// <summary>
    /// Get list of SharePoint sites with pagination and filtering
    /// </summary>
    [HttpGet("sites")]
    public async Task<IActionResult> GetSites(
        [FromQuery] string? search = null,
        [FromQuery] string? siteType = null,
        [FromQuery] string? orderBy = "name",
        [FromQuery] bool ascending = true,
        [FromQuery] int take = 50)
    {
        try
        {
            _logger.LogInformation("Fetching SharePoint sites list");

            var allSites = await GetAllSitesAsync();

            // Apply search filter
            if (!string.IsNullOrEmpty(search))
            {
                var searchLower = search.ToLower();
                allSites = allSites.Where(s =>
                    s.Name.ToLower().Contains(searchLower) ||
                    s.DisplayName.ToLower().Contains(searchLower) ||
                    (s.WebUrl?.ToLower().Contains(searchLower) ?? false)
                ).ToList();
            }

            // Apply site type filter
            if (!string.IsNullOrEmpty(siteType))
            {
                allSites = siteType.ToLower() switch
                {
                    "team" => allSites.Where(s => s.SiteTemplate == "GROUP#0").ToList(),
                    "communication" => allSites.Where(s => s.SiteTemplate == "SITEPAGEPUBLISHING#0").ToList(),
                    "personal" => allSites.Where(s => s.IsPersonalSite).ToList(),
                    _ => allSites
                };
            }

            var totalCount = allSites.Count;

            // Apply sorting
            allSites = orderBy?.ToLower() switch
            {
                "storage" => ascending 
                    ? allSites.OrderBy(s => s.StorageUsedBytes).ToList()
                    : allSites.OrderByDescending(s => s.StorageUsedBytes).ToList(),
                "created" => ascending
                    ? allSites.OrderBy(s => s.CreatedDateTime).ToList()
                    : allSites.OrderByDescending(s => s.CreatedDateTime).ToList(),
                "modified" => ascending
                    ? allSites.OrderBy(s => s.LastModifiedDateTime).ToList()
                    : allSites.OrderByDescending(s => s.LastModifiedDateTime).ToList(),
                _ => ascending
                    ? allSites.OrderBy(s => s.DisplayName).ToList()
                    : allSites.OrderByDescending(s => s.DisplayName).ToList()
            };

            // Apply pagination
            var pagedSites = allSites.Take(take).ToList();

            return Ok(new SharePointSiteListResultDto(
                Sites: pagedSites,
                TotalCount: totalCount,
                FilteredCount: pagedSites.Count,
                NextLink: null
            ));
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching SharePoint sites");
            return StatusCode(500, new { error = "Failed to fetch SharePoint sites", message = ex.Message });
        }
    }

    /// <summary>
    /// Get details for a specific site
    /// </summary>
    [HttpGet("sites/{siteId}")]
    public async Task<IActionResult> GetSiteDetails(string siteId)
    {
        try
        {
            _logger.LogInformation("Fetching SharePoint site details for {SiteId}", siteId);

            var site = await _graphClient.Sites[siteId].GetAsync(config =>
            {
                config.QueryParameters.Select = new[]
                {
                    "id", "name", "displayName", "description", "webUrl",
                    "createdDateTime", "lastModifiedDateTime", "siteCollection"
                };
            });

            if (site == null)
            {
                return NotFound(new { error = "Site not found" });
            }

            // Get storage info from admin API or drive
            long storageUsed = 0;
            long storageAllocated = 0;

            try
            {
                var drive = await _graphClient.Sites[siteId].Drive.GetAsync(config =>
                {
                    config.QueryParameters.Select = new[] { "quota" };
                });

                storageUsed = drive?.Quota?.Used ?? 0;
                storageAllocated = drive?.Quota?.Total ?? 0;
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "Could not get storage info for site {SiteId}", siteId);
            }

            var storagePercentage = storageAllocated > 0 
                ? Math.Round((double)storageUsed / storageAllocated * 100, 1) 
                : 0;

            var isPersonalSite = site.WebUrl?.Contains("-my.sharepoint.com") == true ||
                                 site.WebUrl?.Contains("/personal/") == true;

            var siteDto = new SharePointSiteDto(
                Id: site.Id ?? string.Empty,
                Name: site.Name ?? "Unknown",
                DisplayName: site.DisplayName ?? site.Name ?? "Unknown",
                Description: site.Description,
                WebUrl: site.WebUrl ?? string.Empty,
                SiteTemplate: null,
                CreatedDateTime: site.CreatedDateTime?.DateTime,
                LastModifiedDateTime: site.LastModifiedDateTime?.DateTime,
                StorageUsedBytes: storageUsed,
                StorageAllocatedBytes: storageAllocated,
                StorageUsedPercentage: storagePercentage,
                OwnerDisplayName: null,
                OwnerEmail: null,
                IsPersonalSite: isPersonalSite,
                ItemCount: null,
                Status: "Active"
            );

            return Ok(siteDto);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching SharePoint site details for {SiteId}", siteId);
            return StatusCode(500, new { error = "Failed to fetch site details", message = ex.Message });
        }
    }

    /// <summary>
    /// Get storage breakdown by site type
    /// </summary>
    [HttpGet("storage")]
    public async Task<IActionResult> GetStorageBreakdown()
    {
        try
        {
            _logger.LogInformation("Fetching SharePoint storage breakdown");

            var allSites = await GetAllSitesAsync();

            var teamSites = allSites.Where(s => s.SiteTemplate == "GROUP#0").ToList();
            var commSites = allSites.Where(s => s.SiteTemplate == "SITEPAGEPUBLISHING#0").ToList();
            var personalSites = allSites.Where(s => s.IsPersonalSite).ToList();
            var otherSites = allSites.Except(teamSites).Except(commSites).Except(personalSites).ToList();

            var breakdown = new
            {
                teamSites = new
                {
                    count = teamSites.Count,
                    storageUsedBytes = teamSites.Sum(s => s.StorageUsedBytes),
                    storageAllocatedBytes = teamSites.Sum(s => s.StorageAllocatedBytes)
                },
                communicationSites = new
                {
                    count = commSites.Count,
                    storageUsedBytes = commSites.Sum(s => s.StorageUsedBytes),
                    storageAllocatedBytes = commSites.Sum(s => s.StorageAllocatedBytes)
                },
                personalSites = new
                {
                    count = personalSites.Count,
                    storageUsedBytes = personalSites.Sum(s => s.StorageUsedBytes),
                    storageAllocatedBytes = personalSites.Sum(s => s.StorageAllocatedBytes)
                },
                otherSites = new
                {
                    count = otherSites.Count,
                    storageUsedBytes = otherSites.Sum(s => s.StorageUsedBytes),
                    storageAllocatedBytes = otherSites.Sum(s => s.StorageAllocatedBytes)
                },
                total = new
                {
                    count = allSites.Count,
                    storageUsedBytes = allSites.Sum(s => s.StorageUsedBytes),
                    storageAllocatedBytes = allSites.Sum(s => s.StorageAllocatedBytes)
                },
                lastUpdated = DateTime.UtcNow
            };

            return Ok(breakdown);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error fetching SharePoint storage breakdown");
            return StatusCode(500, new { error = "Failed to fetch storage breakdown", message = ex.Message });
        }
    }

    #region Private Methods

    private async Task<List<SharePointSiteDto>> GetAllSitesAsync()
    {
        var allSites = new List<Site>();
        var siteIds = new HashSet<string>();
        var groupSiteIds = new HashSet<string>(); // Track which sites came from groups

        // Method 1: Try to enumerate all sites using the sites endpoint without search
        // This requires Sites.Read.All permission
        try
        {
            _logger.LogInformation("Fetching all sites using Graph API");
            
            // First get root site
            var rootSite = await _graphClient.Sites["root"].GetAsync(config =>
            {
                config.QueryParameters.Select = new[]
                {
                    "id", "name", "displayName", "description", "webUrl",
                    "createdDateTime", "lastModifiedDateTime", "siteCollection"
                };
            });
            
            if (rootSite != null && !string.IsNullOrEmpty(rootSite.Id))
            {
                allSites.Add(rootSite);
                siteIds.Add(rootSite.Id);
                _logger.LogInformation("Added root site: {Name} - {Url}", rootSite.DisplayName, rootSite.WebUrl);
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Could not fetch root site");
        }

        // Method 2: Search with wildcard to find indexed sites
        try
        {
            _logger.LogInformation("Searching for all sites using search API");
            
            var searchResults = await _graphClient.Sites.GetAsync(config =>
            {
                config.QueryParameters.Search = "*";
                config.QueryParameters.Select = new[]
                {
                    "id", "name", "displayName", "description", "webUrl",
                    "createdDateTime", "lastModifiedDateTime", "siteCollection"
                };
                config.QueryParameters.Top = 500;
            });

            if (searchResults?.Value != null)
            {
                foreach (var site in searchResults.Value)
                {
                    if (site != null && !string.IsNullOrEmpty(site.Id) && !siteIds.Contains(site.Id))
                    {
                        allSites.Add(site);
                        siteIds.Add(site.Id);
                        _logger.LogDebug("Search found site: {Name} - {Url}", site.DisplayName, site.WebUrl);
                    }
                }
                
                // Page through results
                while (searchResults?.OdataNextLink != null)
                {
                    searchResults = await _graphClient.Sites.WithUrl(searchResults.OdataNextLink).GetAsync();
                    if (searchResults?.Value != null)
                    {
                        foreach (var site in searchResults.Value)
                        {
                            if (site != null && !string.IsNullOrEmpty(site.Id) && !siteIds.Contains(site.Id))
                            {
                                allSites.Add(site);
                                siteIds.Add(site.Id);
                            }
                        }
                    }
                }
            }
            
            _logger.LogInformation("Total sites after search: {Count}", allSites.Count);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Search API failed: {Message}", ex.Message);
        }

        // Method 3: Get sites from Microsoft 365 Groups (these are Team Sites)
        try
        {
            _logger.LogInformation("Fetching sites from Microsoft 365 Groups");
            
            var groups = await _graphClient.Groups.GetAsync(config =>
            {
                config.QueryParameters.Filter = "groupTypes/any(c:c eq 'Unified')";
                config.QueryParameters.Select = new[] { "id", "displayName" };
                config.QueryParameters.Top = 999;
            });

            var groupList = new List<Microsoft.Graph.Models.Group>();
            if (groups?.Value != null)
            {
                groupList.AddRange(groups.Value);
                
                while (groups?.OdataNextLink != null)
                {
                    groups = await _graphClient.Groups.WithUrl(groups.OdataNextLink).GetAsync();
                    if (groups?.Value != null)
                    {
                        groupList.AddRange(groups.Value);
                    }
                }
            }

            _logger.LogInformation("Found {Count} Microsoft 365 Groups", groupList.Count);

            foreach (var group in groupList)
            {
                try
                {
                    var site = await _graphClient.Groups[group.Id].Sites["root"].GetAsync(config =>
                    {
                        config.QueryParameters.Select = new[]
                        {
                            "id", "name", "displayName", "description", "webUrl",
                            "createdDateTime", "lastModifiedDateTime", "siteCollection"
                        };
                    });
                    
                    if (site != null && !string.IsNullOrEmpty(site.Id))
                    {
                        // Mark this site as coming from a group (Team Site)
                        groupSiteIds.Add(site.Id);
                        
                        if (!siteIds.Contains(site.Id))
                        {
                            allSites.Add(site);
                            siteIds.Add(site.Id);
                        }
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogDebug(ex, "Could not get site for group {GroupId}", group.Id);
                }
            }
            
            _logger.LogInformation("Total group sites identified: {Count}", groupSiteIds.Count);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Could not fetch sites from groups");
        }

        _logger.LogInformation("Total sites retrieved: {Count} (Group/Team sites: {GroupCount})", allSites.Count, groupSiteIds.Count);

        // Convert to DTOs with storage info
        var siteDtos = new List<SharePointSiteDto>();
        
        foreach (var site in allSites)
        {
            long storageUsed = 0;

            try
            {
                var drive = await _graphClient.Sites[site.Id].Drive.GetAsync(config =>
                {
                    config.QueryParameters.Select = new[] { "quota" };
                });

                storageUsed = drive?.Quota?.Used ?? 0;
            }
            catch (Exception ex)
            {
                _logger.LogDebug(ex, "Could not get storage info for site {SiteId}", site.Id);
            }

            var isPersonalSite = site.WebUrl?.Contains("-my.sharepoint.com") == true ||
                                 site.WebUrl?.Contains("/personal/") == true;

            // Determine site type
            // 1. If it came from a Group -> Team Site
            // 2. If it's a personal site -> Personal Site  
            // 3. If it has /sites/ or is root -> Communication Site
            // 4. Otherwise -> Other
            bool isTeamSite = groupSiteIds.Contains(site.Id ?? "");
            
            string? siteTemplate;
            if (isPersonalSite)
            {
                siteTemplate = "SPSPERS#10";
            }
            else if (isTeamSite)
            {
                siteTemplate = "GROUP#0"; // Team site connected to M365 Group
            }
            else if (site.WebUrl?.Contains("/sites/") == true || 
                     (site.WebUrl != null && !site.WebUrl.Contains("/teams/") && !site.WebUrl.Contains("/personal/")))
            {
                // Communication site - either /sites/ path OR root site (no special path)
                siteTemplate = "SITEPAGEPUBLISHING#0";
            }
            else
            {
                siteTemplate = "OTHER";
            }

            _logger.LogDebug("Site {Name} ({Url}) - IsTeamSite: {IsTeam}, Template: {Template}", 
                site.DisplayName, site.WebUrl, isTeamSite, siteTemplate);

            siteDtos.Add(new SharePointSiteDto(
                Id: site.Id ?? string.Empty,
                Name: site.Name ?? site.DisplayName ?? "Unknown",
                DisplayName: site.DisplayName ?? site.Name ?? "Unknown",
                Description: site.Description,
                WebUrl: site.WebUrl ?? string.Empty,
                SiteTemplate: siteTemplate,
                CreatedDateTime: site.CreatedDateTime?.DateTime,
                LastModifiedDateTime: site.LastModifiedDateTime?.DateTime,
                StorageUsedBytes: storageUsed,
                StorageAllocatedBytes: 0,
                StorageUsedPercentage: 0,
                OwnerDisplayName: null,
                OwnerEmail: null,
                IsPersonalSite: isPersonalSite,
                ItemCount: null,
                Status: "Active"
            ));
        }

        return siteDtos;
    }

    private static SharePointStatsDto CalculateStats(List<SharePointSiteDto> sites)
    {
        // Team sites: GROUP#0 template (sites from M365 Groups)
        var teamSites = sites.Count(s => s.SiteTemplate == "GROUP#0");
        
        // Communication sites: SITEPAGEPUBLISHING#0 template (sites not from groups)
        var commSites = sites.Count(s => s.SiteTemplate == "SITEPAGEPUBLISHING#0");
        
        // Personal sites (OneDrive)
        var personalSites = sites.Count(s => s.IsPersonalSite);
        
        // Other sites (root site, etc.)
        var otherSites = sites.Count - teamSites - commSites - personalSites;

        // Sum only the storage used (allocated is tenant-wide, not useful)
        var totalStorageUsed = sites.Sum(s => s.StorageUsedBytes);
        
        // We can't get accurate tenant-wide allocation from Graph API
        // Set to 0 to indicate we don't have this info
        long totalStorageAllocated = 0;
        var overallPercentage = totalStorageAllocated > 0
            ? Math.Round((double)totalStorageUsed / totalStorageAllocated * 100, 1)
            : 0;

        var sitesNearQuota = sites.Count(s => s.StorageUsedPercentage >= 80);

        var thirtyDaysAgo = DateTime.UtcNow.AddDays(-30);
        var activeSites = sites.Count(s => s.LastModifiedDateTime > thirtyDaysAgo);
        var inactiveSites = sites.Count(s => s.LastModifiedDateTime <= thirtyDaysAgo || !s.LastModifiedDateTime.HasValue);

        return new SharePointStatsDto(
            TotalSites: sites.Count,
            TeamSites: teamSites,
            CommunicationSites: commSites,
            PersonalSites: personalSites,
            OtherSites: otherSites,
            TotalStorageUsedBytes: totalStorageUsed,
            TotalStorageAllocatedBytes: totalStorageAllocated,
            OverallStorageUsedPercentage: overallPercentage,
            SitesNearQuota: sitesNearQuota,
            ActiveSitesLast30Days: activeSites,
            InactiveSitesLast30Days: inactiveSites,
            LastUpdated: DateTime.UtcNow
        );
    }

    #endregion
}
