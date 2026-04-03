using System.Security.Claims;

namespace M365Dashboard.Api.Middleware;

public class RequestContextMiddleware
{
    private readonly RequestDelegate _next;
    private readonly ILogger<RequestContextMiddleware> _logger;

    public RequestContextMiddleware(RequestDelegate next, ILogger<RequestContextMiddleware> logger)
    {
        _next = next;
        _logger = logger;
    }

    public async Task InvokeAsync(HttpContext context)
    {
        // Add correlation ID
        var correlationId = context.Request.Headers["X-Correlation-ID"].FirstOrDefault() 
            ?? Guid.NewGuid().ToString();
        
        context.Items["CorrelationId"] = correlationId;
        context.Response.Headers["X-Correlation-ID"] = correlationId;

        // Log request context
        var userId = context.User?.FindFirst(ClaimTypes.NameIdentifier)?.Value 
            ?? context.User?.FindFirst("oid")?.Value 
            ?? "anonymous";

        using (_logger.BeginScope(new Dictionary<string, object>
        {
            ["CorrelationId"] = correlationId,
            ["UserId"] = userId,
            ["Path"] = context.Request.Path.ToString()
        }))
        {
            await _next(context);
        }
    }
}
