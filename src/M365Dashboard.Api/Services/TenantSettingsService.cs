using System.Text.Json;
using Microsoft.EntityFrameworkCore;
using M365Dashboard.Api.Data;
using M365Dashboard.Api.Models;
using M365Dashboard.Api.Models.Dtos;

namespace M365Dashboard.Api.Services;

public interface ITenantSettingsService
{
    Task<BreakGlassSettingsDto> GetBreakGlassSettingsAsync(string tenantId);
    Task<BreakGlassSettingsDto> UpdateBreakGlassSettingsAsync(string tenantId, List<string> userPrincipalNames, string modifiedBy);
}

public class TenantSettingsService : ITenantSettingsService
{
    private const string BreakGlassSettingKey = "BreakGlassAccounts";
    
    private readonly ApplicationDbContext _dbContext;
    private readonly IGraphService _graphService;
    private readonly ILogger<TenantSettingsService> _logger;

    public TenantSettingsService(
        ApplicationDbContext dbContext,
        IGraphService graphService,
        ILogger<TenantSettingsService> logger)
    {
        _dbContext = dbContext;
        _graphService = graphService;
        _logger = logger;
    }

    public async Task<BreakGlassSettingsDto> GetBreakGlassSettingsAsync(string tenantId)
    {
        var setting = await _dbContext.TenantSettings
            .FirstOrDefaultAsync(s => s.TenantId == tenantId && s.SettingKey == BreakGlassSettingKey);

        if (setting == null)
        {
            return new BreakGlassSettingsDto(
                Accounts: new List<BreakGlassAccountDto>(),
                LastUpdated: null,
                LastModifiedBy: null
            );
        }

        var upns = JsonSerializer.Deserialize<List<string>>(setting.SettingValue) ?? new List<string>();
        
        // Resolve accounts to get display names and object IDs
        var accounts = new List<BreakGlassAccountDto>();
        foreach (var upn in upns)
        {
            var account = await _graphService.ResolveUserAsync(upn);
            accounts.Add(account ?? new BreakGlassAccountDto(upn, null, null, false));
        }

        return new BreakGlassSettingsDto(
            Accounts: accounts,
            LastUpdated: setting.UpdatedAt,
            LastModifiedBy: setting.LastModifiedBy
        );
    }

    public async Task<BreakGlassSettingsDto> UpdateBreakGlassSettingsAsync(
        string tenantId, 
        List<string> userPrincipalNames, 
        string modifiedBy)
    {
        _logger.LogInformation("Updating break glass accounts for tenant {TenantId}", tenantId);

        // Clean and validate the UPNs
        var cleanedUpns = userPrincipalNames
            .Where(u => !string.IsNullOrWhiteSpace(u))
            .Select(u => u.Trim().ToLowerInvariant())
            .Distinct()
            .ToList();

        var setting = await _dbContext.TenantSettings
            .FirstOrDefaultAsync(s => s.TenantId == tenantId && s.SettingKey == BreakGlassSettingKey);

        if (setting == null)
        {
            setting = new TenantSettings
            {
                TenantId = tenantId,
                SettingKey = BreakGlassSettingKey,
                SettingValue = JsonSerializer.Serialize(cleanedUpns),
                Description = "Break glass (emergency access) account UPNs for Conditional Access policy exclusion checks",
                LastModifiedBy = modifiedBy,
                CreatedAt = DateTime.UtcNow,
                UpdatedAt = DateTime.UtcNow
            };
            _dbContext.TenantSettings.Add(setting);
        }
        else
        {
            setting.SettingValue = JsonSerializer.Serialize(cleanedUpns);
            setting.LastModifiedBy = modifiedBy;
            setting.UpdatedAt = DateTime.UtcNow;
        }

        await _dbContext.SaveChangesAsync();

        _logger.LogInformation("Updated break glass accounts: {Count} accounts configured", cleanedUpns.Count);

        // Return the resolved accounts
        var accounts = new List<BreakGlassAccountDto>();
        foreach (var upn in cleanedUpns)
        {
            var account = await _graphService.ResolveUserAsync(upn);
            accounts.Add(account ?? new BreakGlassAccountDto(upn, null, null, false));
        }

        return new BreakGlassSettingsDto(
            Accounts: accounts,
            LastUpdated: setting.UpdatedAt,
            LastModifiedBy: setting.LastModifiedBy
        );
    }
}
