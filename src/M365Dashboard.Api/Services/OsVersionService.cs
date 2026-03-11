using System.Text.Json;

namespace M365Dashboard.Api.Services;

public interface IOsVersionService
{
    Task<OsVersionInfo> GetLatestVersionsAsync();
    OsVersionStatus CheckiOSVersion(string? version);
    OsVersionStatus CheckAndroidVersion(string? version, string? securityPatchLevel);
    OsVersionStatus CheckMacOSVersion(string? version);
    OsVersionStatus CheckWindowsVersion(string? version);
}

public class OsVersionService : IOsVersionService
{
    private readonly ILogger<OsVersionService> _logger;
    private readonly HttpClient _httpClient;
    private OsVersionInfo? _cachedVersions;
    private DateTime _cacheExpiry = DateTime.MinValue;
    private static readonly TimeSpan CacheDuration = TimeSpan.FromHours(24);

    // Fallback versions if API calls fail (update these periodically - March 2026)
    private static readonly string FallbackLatestIOS = "19.3";      // iOS 19 released Sep 2025
    private static readonly string FallbackLatestMacOS = "16.3";    // macOS 16 released Sep 2025
    private static readonly int FallbackLatestAndroidVersion = 16;   // Android 16 released 2025
    private static readonly string FallbackLatestAndroidSecurityPatch = "2026-02-01";
    private static readonly string FallbackLatestWindows11Build = "10.0.26100"; // 24H2
    private static readonly string FallbackLatestWindows10Build = "10.0.19045"; // 22H2

    public OsVersionService(ILogger<OsVersionService> logger, IHttpClientFactory httpClientFactory)
    {
        _logger = logger;
        _httpClient = httpClientFactory.CreateClient();
        _httpClient.Timeout = TimeSpan.FromSeconds(10);
    }

    public async Task<OsVersionInfo> GetLatestVersionsAsync()
    {
        if (_cachedVersions != null && DateTime.UtcNow < _cacheExpiry)
        {
            return _cachedVersions;
        }

        var versions = new OsVersionInfo();

        // Fetch iOS versions from endoflife.date
        try
        {
            using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(5));
            var response = await _httpClient.GetStringAsync(
                "https://endoflife.date/api/ios.json", cts.Token);
            var releases = JsonDocument.Parse(response).RootElement;
            
            int latestMajor = 0;
            string? latestVersion = null;
            
            foreach (var release in releases.EnumerateArray())
            {
                var cycle = release.TryGetProperty("cycle", out var c) ? c.GetString() : null;
                if (string.IsNullOrEmpty(cycle)) continue;
                
                // Parse major version
                var majorStr = cycle.Split('.')[0];
                if (!int.TryParse(majorStr, out var majorVersion)) continue;
                
                // Track latest
                if (majorVersion > latestMajor)
                {
                    latestMajor = majorVersion;
                    latestVersion = release.TryGetProperty("latest", out var l) ? l.GetString() : cycle;
                }
                
                // Parse EOL status
                var isEol = false;
                DateTime? eolDate = null;
                if (release.TryGetProperty("eol", out var eol))
                {
                    if (eol.ValueKind == JsonValueKind.True) isEol = true;
                    else if (eol.ValueKind == JsonValueKind.String && DateTime.TryParse(eol.GetString(), out var d))
                    {
                        eolDate = d;
                        isEol = d < DateTime.UtcNow;
                    }
                }
                
                if (!versions.iOSVersions.ContainsKey(majorVersion))
                {
                    versions.iOSVersions[majorVersion] = new iOSVersionInfo
                    {
                        Cycle = cycle,
                        IsEol = isEol,
                        IsMaintained = !isEol,
                        Latest = release.TryGetProperty("latest", out var lat) ? lat.GetString() : null,
                        EolDate = eolDate,
                        ReleaseDate = release.TryGetProperty("releaseDate", out var rd) && 
                            DateTime.TryParse(rd.GetString(), out var relDate) ? relDate : null
                    };
                }
            }
            
            if (!string.IsNullOrEmpty(latestVersion))
                versions.LatestIOS = latestVersion;
            
            _logger.LogInformation("Fetched iOS versions from endoflife.date: latest={Latest}, tracked {Count} versions",
                versions.LatestIOS, versions.iOSVersions.Count);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Failed to fetch iOS versions from endoflife.date");
            versions.LatestIOS = FallbackLatestIOS;
        }
        
        // Fetch macOS versions from endoflife.date
        try
        {
            using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(5));
            var response = await _httpClient.GetStringAsync(
                "https://endoflife.date/api/macos.json", cts.Token);
            var releases = JsonDocument.Parse(response).RootElement;
            
            int latestMajor = 0;
            string? latestVersion = null;
            
            foreach (var release in releases.EnumerateArray())
            {
                var cycle = release.TryGetProperty("cycle", out var c) ? c.GetString() : null;
                if (string.IsNullOrEmpty(cycle)) continue;
                
                // Parse major version (handle "10.15" format)
                var majorStr = cycle.Split('.')[0];
                if (!int.TryParse(majorStr, out var majorVersion)) continue;
                
                // For macOS 10.x, treat them all as version 10
                var versionKey = cycle;
                
                // Track latest (by major version number)
                if (majorVersion > latestMajor || (majorVersion == latestMajor && majorVersion >= 11))
                {
                    latestMajor = majorVersion;
                    latestVersion = release.TryGetProperty("latest", out var l) ? l.GetString() : cycle;
                }
                
                // Parse EOL status
                var isEol = false;
                DateTime? eolDate = null;
                if (release.TryGetProperty("eol", out var eol))
                {
                    if (eol.ValueKind == JsonValueKind.True) isEol = true;
                    else if (eol.ValueKind == JsonValueKind.String && DateTime.TryParse(eol.GetString(), out var d))
                    {
                        eolDate = d;
                        isEol = d < DateTime.UtcNow;
                    }
                }
                
                versions.MacOSVersions[versionKey] = new MacOSVersionInfo
                {
                    Cycle = cycle,
                    Codename = release.TryGetProperty("codename", out var cn) ? cn.GetString() ?? "" : "",
                    IsEol = isEol,
                    Latest = release.TryGetProperty("latest", out var lat) ? lat.GetString() : null,
                    EolDate = eolDate,
                    ReleaseDate = release.TryGetProperty("releaseDate", out var rd) && 
                        DateTime.TryParse(rd.GetString(), out var relDate) ? relDate : null
                };
            }
            
            if (!string.IsNullOrEmpty(latestVersion))
                versions.LatestMacOS = latestVersion;
            
            _logger.LogInformation("Fetched macOS versions from endoflife.date: latest={Latest}, tracked {Count} versions",
                versions.LatestMacOS, versions.MacOSVersions.Count);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Failed to fetch macOS versions from endoflife.date");
            versions.LatestMacOS = FallbackLatestMacOS;
        }

        // Fetch Android versions from endoflife.date API
        try
        {
            using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(5));
            var response = await _httpClient.GetStringAsync(
                "https://endoflife.date/api/v1/products/android/", cts.Token);
            var data = JsonDocument.Parse(response);
            
            if (data.RootElement.TryGetProperty("result", out var result) &&
                result.TryGetProperty("releases", out var releases))
            {
                int latestVersion = FallbackLatestAndroidVersion;
                
                foreach (var release in releases.EnumerateArray())
                {
                    var name = release.TryGetProperty("name", out var n) ? n.GetString() : null;
                    if (string.IsNullOrEmpty(name)) continue;
                    
                    // Parse major version (handle "12.1" -> 12)
                    var versionStr = name.Split('.')[0];
                    if (!int.TryParse(versionStr, out var majorVersion)) continue;
                    
                    // Track the latest version
                    if (majorVersion > latestVersion)
                        latestVersion = majorVersion;
                    
                    // Store version info
                    var versionInfo = new AndroidVersionInfo
                    {
                        Name = name,
                        Codename = release.TryGetProperty("codename", out var c) ? c.GetString() ?? "" : "",
                        IsEol = release.TryGetProperty("isEol", out var eol) && eol.GetBoolean(),
                        IsMaintained = release.TryGetProperty("isMaintained", out var maint) && maint.GetBoolean()
                    };
                    
                    if (release.TryGetProperty("eolFrom", out var eolFrom) && eolFrom.ValueKind == JsonValueKind.String)
                    {
                        if (DateTime.TryParse(eolFrom.GetString(), out var eolDate))
                            versionInfo.EolDate = eolDate;
                    }
                    
                    if (release.TryGetProperty("releaseDate", out var relDate) && relDate.ValueKind == JsonValueKind.String)
                    {
                        if (DateTime.TryParse(relDate.GetString(), out var releaseDate))
                            versionInfo.ReleaseDate = releaseDate;
                    }
                    
                    // Only store if we don't have this major version yet (prefer the main version, not .1 variants)
                    if (!versions.AndroidVersions.ContainsKey(majorVersion))
                        versions.AndroidVersions[majorVersion] = versionInfo;
                }
                
                versions.LatestAndroidVersion = latestVersion;
                _logger.LogInformation("Fetched Android versions from endoflife.date: latest={Latest}, tracked {Count} versions",
                    latestVersion, versions.AndroidVersions.Count);
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Failed to fetch Android versions from endoflife.date, using fallback");
            versions.LatestAndroidVersion = FallbackLatestAndroidVersion;
        }
        
        versions.LatestAndroidSecurityPatch = FallbackLatestAndroidSecurityPatch;
        
        // Fetch Windows versions from endoflife.date API
        try
        {
            using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(5));
            var response = await _httpClient.GetStringAsync(
                "https://endoflife.date/api/windows.json", cts.Token);
            var releases = JsonDocument.Parse(response).RootElement;
            
            foreach (var release in releases.EnumerateArray())
            {
                var cycle = release.TryGetProperty("cycle", out var c) ? c.GetString() : null;
                if (string.IsNullOrEmpty(cycle)) continue;
                
                // Get the build number from latest field (e.g., "10.0.26100")
                var buildNumber = release.TryGetProperty("latest", out var lat) ? lat.GetString() : null;
                
                // Parse EOL status
                var isEol = false;
                DateTime? eolDate = null;
                if (release.TryGetProperty("eol", out var eol))
                {
                    if (eol.ValueKind == JsonValueKind.True) isEol = true;
                    else if (eol.ValueKind == JsonValueKind.String && DateTime.TryParse(eol.GetString(), out var d))
                    {
                        eolDate = d;
                        isEol = d < DateTime.UtcNow;
                    }
                }
                
                var isLts = release.TryGetProperty("lts", out var lts) && lts.ValueKind == JsonValueKind.True;
                
                versions.WindowsVersions[cycle] = new WindowsVersionInfo
                {
                    Name = cycle,
                    Label = cycle,
                    BuildNumber = buildNumber,
                    IsEol = isEol,
                    IsMaintained = !isEol,
                    IsLts = isLts,
                    EolDate = eolDate,
                    ReleaseDate = release.TryGetProperty("releaseDate", out var rd) && 
                        DateTime.TryParse(rd.GetString(), out var relDate) ? relDate : null
                };
                
                // Track latest builds for Windows 10 and 11
                if (!string.IsNullOrEmpty(buildNumber))
                {
                    if (cycle.StartsWith("11-") && !isEol)
                    {
                        var parts = buildNumber.Split('.');
                        if (parts.Length >= 3 && int.TryParse(parts[2], out var build))
                        {
                            var currentParts = versions.LatestWindows11Build.Split('.');
                            if (currentParts.Length >= 3 && int.TryParse(currentParts[2], out var currentBuild))
                            {
                                if (build > currentBuild)
                                    versions.LatestWindows11Build = buildNumber;
                            }
                        }
                    }
                    else if (cycle.StartsWith("10-") && !cycle.Contains("lts") && !isEol)
                    {
                        var parts = buildNumber.Split('.');
                        if (parts.Length >= 3 && int.TryParse(parts[2], out var build))
                        {
                            var currentParts = versions.LatestWindows10Build.Split('.');
                            if (currentParts.Length >= 3 && int.TryParse(currentParts[2], out var currentBuild))
                            {
                                if (build > currentBuild)
                                    versions.LatestWindows10Build = buildNumber;
                            }
                        }
                    }
                }
            }
            
            _logger.LogInformation("Fetched Windows versions from endoflife.date: Win11={Win11}, Win10={Win10}, tracked {Count} versions",
                versions.LatestWindows11Build, versions.LatestWindows10Build, versions.WindowsVersions.Count);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Failed to fetch Windows versions from endoflife.date");
            versions.LatestWindows11Build = FallbackLatestWindows11Build;
            versions.LatestWindows10Build = FallbackLatestWindows10Build;
        }

        _cachedVersions = versions;
        _cacheExpiry = DateTime.UtcNow.Add(CacheDuration);
        
        return versions;
    }

    public OsVersionStatus CheckiOSVersion(string? version)
    {
        if (string.IsNullOrEmpty(version))
            return new OsVersionStatus { Status = VersionStatus.Unknown, Message = "Version unknown" };

        var latest = _cachedVersions?.LatestIOS ?? FallbackLatestIOS;
        var currentParts = ParseVersion(version);
        var iOSVersions = _cachedVersions?.iOSVersions ?? new Dictionary<int, iOSVersionInfo>();

        if (currentParts.Length == 0)
            return new OsVersionStatus { Status = VersionStatus.Unknown, Message = "Invalid version" };

        var majorVersion = currentParts[0];
        
        // Check if we have EOL data for this version from endoflife.date
        if (iOSVersions.TryGetValue(majorVersion, out var versionInfo))
        {
            if (versionInfo.IsEol)
            {
                return new OsVersionStatus 
                { 
                    Status = VersionStatus.Critical, 
                    Message = $"iOS {majorVersion} (EOL)",
                    LatestVersion = latest,
                    VersionsBehind = ParseVersion(latest)[0] - majorVersion
                };
            }
            
            if (versionInfo.IsMaintained)
            {
                var latestMajor = ParseVersion(latest)[0];
                if (majorVersion >= latestMajor - 1)
                {
                    return new OsVersionStatus 
                    { 
                        Status = VersionStatus.Current, 
                        Message = $"iOS {version}",
                        LatestVersion = latest
                    };
                }
                else
                {
                    return new OsVersionStatus 
                    { 
                        Status = VersionStatus.Warning, 
                        Message = $"iOS {majorVersion} (aging)",
                        LatestVersion = latest,
                        VersionsBehind = latestMajor - majorVersion
                    };
                }
            }
        }
        
        // Fallback to version comparison if no EOL data
        var latestParts = ParseVersion(latest);
        var majorDiff = latestParts.Length > 0 ? latestParts[0] - majorVersion : 0;

        if (majorDiff > 2)
            return new OsVersionStatus 
            { 
                Status = VersionStatus.Critical, 
                Message = $"iOS {majorVersion} (likely EOL)",
                LatestVersion = latest,
                VersionsBehind = majorDiff
            };
        
        if (majorDiff > 0)
            return new OsVersionStatus 
            { 
                Status = VersionStatus.Warning, 
                Message = $"Update available ({latest})",
                LatestVersion = latest,
                VersionsBehind = majorDiff
            };
        
        return new OsVersionStatus 
        { 
            Status = VersionStatus.Current, 
            Message = "Up to date",
            LatestVersion = latest
        };
    }

    public OsVersionStatus CheckAndroidVersion(string? version, string? securityPatchLevel)
    {
        // For Android, security patch date is more important than OS version
        if (!string.IsNullOrEmpty(securityPatchLevel))
        {
            if (DateTime.TryParse(securityPatchLevel, out var patchDate))
            {
                var daysSincePatch = (DateTime.UtcNow - patchDate).TotalDays;
                
                if (daysSincePatch > 180)
                    return new OsVersionStatus 
                    { 
                        Status = VersionStatus.Critical, 
                        Message = $"Patch {(int)daysSincePatch} days old",
                        DaysSinceUpdate = (int)daysSincePatch
                    };
                
                if (daysSincePatch > 90)
                    return new OsVersionStatus 
                    { 
                        Status = VersionStatus.Warning, 
                        Message = $"Patch {(int)daysSincePatch} days old",
                        DaysSinceUpdate = (int)daysSincePatch
                    };
                
                return new OsVersionStatus 
                { 
                    Status = VersionStatus.Current, 
                    Message = $"Patch: {patchDate:MMM yyyy}",
                    DaysSinceUpdate = (int)daysSincePatch
                };
            }
        }

        // Fallback to OS version check using endoflife.date data
        if (string.IsNullOrEmpty(version))
            return new OsVersionStatus { Status = VersionStatus.Unknown, Message = "Version unknown" };

        // Try to parse the major version number
        var versionStr = version.Split('.')[0];
        if (int.TryParse(versionStr, out var majorVersion))
        {
            var latestVersion = _cachedVersions?.LatestAndroidVersion ?? FallbackLatestAndroidVersion;
            var androidVersions = _cachedVersions?.AndroidVersions ?? new Dictionary<int, AndroidVersionInfo>();
            
            // Check if we have EOL data for this version
            if (androidVersions.TryGetValue(majorVersion, out var versionInfo))
            {
                var codename = !string.IsNullOrEmpty(versionInfo.Codename) 
                    ? $" '{versionInfo.Codename}'" : "";
                
                if (versionInfo.IsEol)
                {
                    return new OsVersionStatus 
                    { 
                        Status = VersionStatus.Critical, 
                        Message = $"Android {majorVersion}{codename} (EOL)",
                        LatestVersion = latestVersion.ToString(),
                        VersionsBehind = latestVersion - majorVersion
                    };
                }
                
                if (versionInfo.IsMaintained)
                {
                    // Check if it's the latest or close to it
                    if (majorVersion >= latestVersion - 1)
                    {
                        return new OsVersionStatus 
                        { 
                            Status = VersionStatus.Current, 
                            Message = $"Android {majorVersion}{codename}",
                            LatestVersion = latestVersion.ToString()
                        };
                    }
                    else
                    {
                        // Maintained but getting old
                        return new OsVersionStatus 
                        { 
                            Status = VersionStatus.Warning, 
                            Message = $"Android {majorVersion}{codename} (aging)",
                            LatestVersion = latestVersion.ToString(),
                            VersionsBehind = latestVersion - majorVersion
                        };
                    }
                }
            }
            
            // Fallback if no EOL data available
            if (majorVersion >= latestVersion)
                return new OsVersionStatus 
                { 
                    Status = VersionStatus.Current, 
                    Message = $"Android {majorVersion}",
                    LatestVersion = latestVersion.ToString()
                };
            
            if (majorVersion >= latestVersion - 2)
                return new OsVersionStatus 
                { 
                    Status = VersionStatus.Current, 
                    Message = $"Android {majorVersion}",
                    LatestVersion = latestVersion.ToString()
                };
            
            // Older versions without EOL data - assume EOL
            return new OsVersionStatus 
            { 
                Status = VersionStatus.Critical, 
                Message = $"Android {majorVersion} (likely EOL)",
                LatestVersion = latestVersion.ToString(),
                VersionsBehind = latestVersion - majorVersion
            };
        }

        return new OsVersionStatus { Status = VersionStatus.Unknown, Message = version };
    }

    public OsVersionStatus CheckMacOSVersion(string? version)
    {
        if (string.IsNullOrEmpty(version))
            return new OsVersionStatus { Status = VersionStatus.Unknown, Message = "Version unknown" };

        var latest = _cachedVersions?.LatestMacOS ?? FallbackLatestMacOS;
        var currentParts = ParseVersion(version);
        var macOSVersions = _cachedVersions?.MacOSVersions ?? new Dictionary<string, MacOSVersionInfo>();

        if (currentParts.Length == 0)
            return new OsVersionStatus { Status = VersionStatus.Unknown, Message = "Invalid version" };

        var majorVersion = currentParts[0];
        
        // Build version key - for macOS 10.x use full "10.xx", for 11+ use just major
        var versionKey = majorVersion == 10 && currentParts.Length > 1 
            ? $"{majorVersion}.{currentParts[1]}" 
            : majorVersion.ToString();
        
        // Check if we have EOL data for this version from endoflife.date
        if (macOSVersions.TryGetValue(versionKey, out var versionInfo))
        {
            var codename = !string.IsNullOrEmpty(versionInfo.Codename) 
                ? $" '{versionInfo.Codename}'" : "";
            
            if (versionInfo.IsEol)
            {
                return new OsVersionStatus 
                { 
                    Status = VersionStatus.Critical, 
                    Message = $"macOS {versionKey}{codename} (EOL)",
                    LatestVersion = latest,
                    VersionsBehind = ParseVersion(latest)[0] - majorVersion
                };
            }
            
            // macOS is maintained
            var latestMajor = ParseVersion(latest)[0];
            if (majorVersion >= latestMajor - 1)
            {
                return new OsVersionStatus 
                { 
                    Status = VersionStatus.Current, 
                    Message = $"macOS {versionKey}{codename}",
                    LatestVersion = latest
                };
            }
            else
            {
                return new OsVersionStatus 
                { 
                    Status = VersionStatus.Warning, 
                    Message = $"macOS {versionKey}{codename} (aging)",
                    LatestVersion = latest,
                    VersionsBehind = latestMajor - majorVersion
                };
            }
        }
        
        // Fallback to version comparison if no EOL data
        var latestParts = ParseVersion(latest);
        var majorDiff = latestParts.Length > 0 ? latestParts[0] - majorVersion : 0;

        if (majorDiff > 2)
            return new OsVersionStatus 
            { 
                Status = VersionStatus.Critical, 
                Message = $"macOS {version} (likely EOL)",
                LatestVersion = latest,
                VersionsBehind = majorDiff
            };
        
        if (majorDiff > 0)
            return new OsVersionStatus 
            { 
                Status = VersionStatus.Warning, 
                Message = $"Update available ({latest})",
                LatestVersion = latest,
                VersionsBehind = majorDiff
            };
        
        return new OsVersionStatus 
        { 
            Status = VersionStatus.Current, 
            Message = "Up to date",
            LatestVersion = latest
        };
    }

    public OsVersionStatus CheckWindowsVersion(string? version)
    {
        if (string.IsNullOrEmpty(version))
            return new OsVersionStatus { Status = VersionStatus.Unknown, Message = "Version unknown" };

        // Windows version format: 10.0.XXXXX.YYYY
        var parts = version.Split('.');
        if (parts.Length < 3)
            return new OsVersionStatus { Status = VersionStatus.Unknown, Message = version };

        if (!int.TryParse(parts[2], out var buildNumber))
            return new OsVersionStatus { Status = VersionStatus.Unknown, Message = version };

        // Windows 11 builds start at 22000
        var isWindows11 = buildNumber >= 22000;
        var windowsVersions = _cachedVersions?.WindowsVersions ?? new Dictionary<string, WindowsVersionInfo>();
        
        // Try to find the matching Windows version from endoflife.date data
        WindowsVersionInfo? matchedVersion = null;
        string? matchedCycle = null;
        
        foreach (var kvp in windowsVersions)
        {
            if (string.IsNullOrEmpty(kvp.Value.BuildNumber)) continue;
            
            var vBuildParts = kvp.Value.BuildNumber.Split('.');
            if (vBuildParts.Length >= 3 && int.TryParse(vBuildParts[2], out var vBuild))
            {
                // Match by build number (allowing for minor version differences)
                if (vBuild == buildNumber || Math.Abs(vBuild - buildNumber) < 100)
                {
                    matchedVersion = kvp.Value;
                    matchedCycle = kvp.Key;
                    break;
                }
            }
        }
        
        // If we found a match in endoflife.date data
        if (matchedVersion != null && matchedCycle != null)
        {
            var osName = isWindows11 ? "Windows 11" : "Windows 10";
            var edition = matchedCycle.Contains("-w") ? "Home/Pro" : 
                          matchedCycle.Contains("-e") ? "Enterprise" : "";
            var displayName = string.IsNullOrEmpty(edition) ? osName : $"{osName} {edition}";
            
            // Extract version like "24H2" from cycle like "11-24h2-w"
            var versionMatch = System.Text.RegularExpressions.Regex.Match(matchedCycle, @"(\d{2}[hH]\d)");
            var versionLabel = versionMatch.Success ? versionMatch.Groups[1].Value.ToUpper() : "";
            
            if (matchedVersion.IsEol)
            {
                return new OsVersionStatus 
                { 
                    Status = VersionStatus.Critical, 
                    Message = $"{displayName} {versionLabel} (EOL)",
                    LatestVersion = isWindows11 
                        ? (_cachedVersions?.LatestWindows11Build ?? FallbackLatestWindows11Build)
                        : (_cachedVersions?.LatestWindows10Build ?? FallbackLatestWindows10Build)
                };
            }
            
            if (matchedVersion.IsMaintained)
            {
                return new OsVersionStatus 
                { 
                    Status = VersionStatus.Current, 
                    Message = $"{displayName} {versionLabel}",
                    LatestVersion = matchedVersion.BuildNumber
                };
            }
        }
        
        // Fallback to build number comparison
        var latestBuild = isWindows11 
            ? (_cachedVersions?.LatestWindows11Build ?? FallbackLatestWindows11Build)
            : (_cachedVersions?.LatestWindows10Build ?? FallbackLatestWindows10Build);

        var latestBuildParts = latestBuild.Split('.');
        if (latestBuildParts.Length >= 3 && int.TryParse(latestBuildParts[2], out var latestBuildNumber))
        {
            var buildsBehind = latestBuildNumber - buildNumber;
            
            // Each Windows feature update typically increments the build number by ~1000
            if (buildsBehind > 2000)
                return new OsVersionStatus 
                { 
                    Status = VersionStatus.Critical, 
                    Message = $"Outdated (build {buildNumber})",
                    LatestVersion = latestBuild,
                    VersionsBehind = buildsBehind / 1000
                };
            
            if (buildsBehind > 1000)
                return new OsVersionStatus 
                { 
                    Status = VersionStatus.Warning, 
                    Message = "Update available",
                    LatestVersion = latestBuild,
                    VersionsBehind = 1
                };
        }

        // Windows 10 end of support check (Oct 14, 2025) - check from endoflife.date or fallback
        if (!isWindows11)
        {
            // Check if Windows 10 22H2 is EOL according to endoflife.date
            if (windowsVersions.TryGetValue("10-22h2", out var win10Info) && win10Info.IsEol)
            {
                return new OsVersionStatus 
                { 
                    Status = VersionStatus.Critical, 
                    Message = "Windows 10 (EOL)"
                };
            }
            // Fallback date check
            else if (DateTime.UtcNow > new DateTime(2025, 10, 14))
            {
                return new OsVersionStatus 
                { 
                    Status = VersionStatus.Critical, 
                    Message = "Windows 10 (end of support)"
                };
            }
        }

        return new OsVersionStatus { Status = VersionStatus.Current, Message = "Up to date" };
    }

    private static int[] ParseVersion(string version)
    {
        try
        {
            return version.Split('.')
                .Select(p => int.TryParse(p, out var v) ? v : 0)
                .ToArray();
        }
        catch
        {
            return Array.Empty<int>();
        }
    }
}

public class OsVersionInfo
{
    public string LatestIOS { get; set; } = "19.3";
    public string LatestMacOS { get; set; } = "16.3";
    public int LatestAndroidVersion { get; set; } = 16;
    public string LatestAndroidSecurityPatch { get; set; } = "2026-02-01";
    public string LatestWindows11Build { get; set; } = "10.0.26100";
    public string LatestWindows10Build { get; set; } = "10.0.19045";
    public DateTime LastUpdated { get; set; } = DateTime.UtcNow;
    
    // EOL data from endoflife.date
    public Dictionary<int, AndroidVersionInfo> AndroidVersions { get; set; } = new();
    public Dictionary<int, iOSVersionInfo> iOSVersions { get; set; } = new();
    public Dictionary<string, MacOSVersionInfo> MacOSVersions { get; set; } = new();
    public Dictionary<string, WindowsVersionInfo> WindowsVersions { get; set; } = new();
}

public class AndroidVersionInfo
{
    public string Name { get; set; } = "";
    public string Codename { get; set; } = "";
    public bool IsEol { get; set; }
    public bool IsMaintained { get; set; }
    public DateTime? EolDate { get; set; }
    public DateTime? ReleaseDate { get; set; }
}

public class iOSVersionInfo
{
    public string Cycle { get; set; } = "";
    public bool IsEol { get; set; }
    public bool IsMaintained { get; set; }
    public string? Latest { get; set; }
    public DateTime? EolDate { get; set; }
    public DateTime? ReleaseDate { get; set; }
}

public class MacOSVersionInfo
{
    public string Cycle { get; set; } = "";
    public string Codename { get; set; } = "";
    public bool IsEol { get; set; }
    public string? Latest { get; set; }
    public DateTime? EolDate { get; set; }
    public DateTime? ReleaseDate { get; set; }
}

public class WindowsVersionInfo
{
    public string Name { get; set; } = "";
    public string Label { get; set; } = "";
    public string? BuildNumber { get; set; }
    public bool IsEol { get; set; }
    public bool IsMaintained { get; set; }
    public bool IsLts { get; set; }
    public DateTime? EolDate { get; set; }
    public DateTime? ReleaseDate { get; set; }
}

public class OsVersionStatus
{
    public VersionStatus Status { get; set; }
    public string Message { get; set; } = "";
    public string? LatestVersion { get; set; }
    public int? VersionsBehind { get; set; }
    public int? DaysSinceUpdate { get; set; }
}

public enum VersionStatus
{
    Current,    // Up to date or minor update available
    Warning,    // 1 major version behind or 30-90 days old patch
    Critical,   // 2+ major versions behind or 90+ days old patch
    Unknown     // Cannot determine
}
