using Microsoft.Extensions.Caching.Memory;
using System.Collections.Concurrent;
using System.Text.Json;

namespace M365Dashboard.Api.Services;

public interface ICacheService
{
    Task<T?> GetAsync<T>(string key);
    Task SetAsync<T>(string key, T value, TimeSpan expiration);
    Task RemoveAsync(string key);
    Task RemoveByPrefixAsync(string prefix);
    Task ClearAsync();
}

public class CacheService : ICacheService
{
    private readonly IMemoryCache _cache;
    private readonly ILogger<CacheService> _logger;
    private readonly ConcurrentDictionary<string, byte> _cacheKeys = new();

    public CacheService(IMemoryCache cache, ILogger<CacheService> logger)
    {
        _cache = cache;
        _logger = logger;
    }

    public Task<T?> GetAsync<T>(string key)
    {
        if (_cache.TryGetValue(key, out string? json) && json != null)
        {
            try
            {
                var value = JsonSerializer.Deserialize<T>(json);
                return Task.FromResult(value);
            }
            catch (JsonException ex)
            {
                _logger.LogWarning(ex, "Failed to deserialize cache value for key {Key}", key);
                _cache.Remove(key);
                _cacheKeys.TryRemove(key, out _);
            }
        }

        return Task.FromResult<T?>(default);
    }

    public Task SetAsync<T>(string key, T value, TimeSpan expiration)
    {
        var json = JsonSerializer.Serialize(value);
        
        var options = new MemoryCacheEntryOptions
        {
            AbsoluteExpirationRelativeToNow = expiration,
            SlidingExpiration = expiration > TimeSpan.FromMinutes(5) 
                ? TimeSpan.FromMinutes(5) 
                : null
        };

        options.RegisterPostEvictionCallback((k, v, r, s) =>
        {
            _cacheKeys.TryRemove(k.ToString()!, out _);
            _logger.LogDebug("Cache entry evicted: {Key}, Reason: {Reason}", k, r);
        });

        _cache.Set(key, json, options);
        _cacheKeys.TryAdd(key, 0);

        _logger.LogDebug("Cache entry set: {Key}, Expires in: {Expiration}", key, expiration);

        return Task.CompletedTask;
    }

    public Task RemoveAsync(string key)
    {
        _cache.Remove(key);
        _cacheKeys.TryRemove(key, out _);
        _logger.LogDebug("Cache entry removed: {Key}", key);
        return Task.CompletedTask;
    }

    public Task RemoveByPrefixAsync(string prefix)
    {
        var keysToRemove = _cacheKeys.Keys
            .Where(k => k.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
            .ToList();

        foreach (var key in keysToRemove)
        {
            _cache.Remove(key);
            _cacheKeys.TryRemove(key, out _);
        }

        _logger.LogDebug("Removed {Count} cache entries with prefix: {Prefix}", keysToRemove.Count, prefix);

        return Task.CompletedTask;
    }

    public Task ClearAsync()
    {
        foreach (var key in _cacheKeys.Keys.ToList())
        {
            _cache.Remove(key);
        }
        
        _cacheKeys.Clear();
        _logger.LogInformation("Cache cleared");

        return Task.CompletedTask;
    }
}
