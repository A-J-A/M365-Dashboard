using Microsoft.AspNetCore.Authentication.JwtBearer;
using Microsoft.IdentityModel.Tokens;
using Microsoft.Identity.Web;
using Microsoft.EntityFrameworkCore;
using Microsoft.Graph;
using Azure.Identity;
using M365Dashboard.Api.Data;
using M365Dashboard.Api.Services;
using M365Dashboard.Api.Configuration;
using M365Dashboard.Api.Middleware;
using Serilog;
using QuestPDF.Infrastructure;

var builder = WebApplication.CreateBuilder(args);

// Configure Serilog
Log.Logger = new LoggerConfiguration()
    .ReadFrom.Configuration(builder.Configuration)
    .Enrich.FromLogContext()
    .WriteTo.Console()
    .CreateLogger();

builder.Host.UseSerilog();

// Add Azure Key Vault configuration
// Loads when KeyVault:Uri is set (always in production via env var, optionally in dev)
// Managed Identity is used in production; DefaultAzureCredential falls back to
// Visual Studio / Azure CLI credentials for local development.
var keyVaultUri = builder.Configuration["KeyVault:Uri"];
if (!string.IsNullOrEmpty(keyVaultUri))
{
    builder.Configuration.AddAzureKeyVault(
        new Uri(keyVaultUri),
        new DefaultAzureCredential());
    Log.Information("Azure Key Vault configuration loaded from {Uri}", keyVaultUri);
}

// Prevent ASP.NET from remapping JWT claim names (e.g. 'oid' → long URI form)
// so User.FindFirst("oid") works as expected in controllers.
Microsoft.IdentityModel.JsonWebTokens.JsonWebTokenHandler.DefaultInboundClaimTypeMap.Clear();
System.IdentityModel.Tokens.Jwt.JwtSecurityTokenHandler.DefaultInboundClaimTypeMap.Clear();

// Configure Entra ID Authentication for API (validates incoming tokens)
builder.Services.AddAuthentication(JwtBearerDefaults.AuthenticationScheme)
    .AddMicrosoftIdentityWebApi(builder.Configuration.GetSection("AzureAd"))
    .EnableTokenAcquisitionToCallDownstreamApi()
    .AddInMemoryTokenCaches();

// Add JWT Bearer events for debugging
builder.Services.Configure<JwtBearerOptions>(JwtBearerDefaults.AuthenticationScheme, options =>
{
    options.Events = new JwtBearerEvents
    {
        OnAuthenticationFailed = context =>
        {
            // Log at debug to avoid leaking stack traces in production logs
            Log.Debug(context.Exception, "Authentication failed: {Error}", context.Exception.Message);
            return Task.CompletedTask;
        },
        OnTokenValidated = context =>
        {
            Log.Information("Token validated for: {User}", context.Principal?.Identity?.Name);
            return Task.CompletedTask;
        },
        OnChallenge = context =>
        {
            Log.Warning("Authentication challenge: {Error} - {ErrorDescription}", context.Error, context.ErrorDescription);
            return Task.CompletedTask;
        }
    };
});

// Configure Authorization Policies based on App Roles
builder.Services.AddAuthorizationBuilder()
    .AddPolicy("RequireAdminRole", policy =>
        policy.RequireRole("Dashboard.Admin"))
    .AddPolicy("RequireReaderRole", policy =>
        policy.RequireRole("Dashboard.Admin", "Dashboard.Reader"));

// Configure Microsoft Graph Client with Client Credentials (Application Permissions)
builder.Services.AddSingleton<GraphServiceClient>(sp =>
{
    var config = builder.Configuration.GetSection("AzureAd");
    var tenantId = config["TenantId"];
    var clientId = config["ClientId"];
    var clientSecret = config["ClientSecret"];

    // Use ClientSecretCredential for application permissions
    var credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
    
    return new GraphServiceClient(credential, new[] { "https://graph.microsoft.com/.default" });
});

// Configure Entity Framework with SQL Server
builder.Services.AddDbContext<ApplicationDbContext>(options =>
{
    var connectionString = builder.Configuration.GetConnectionString("DefaultConnection");
    
    options.UseSqlServer(connectionString, sqlOptions =>
    {
        sqlOptions.EnableRetryOnFailure(
            maxRetryCount: 5,
            maxRetryDelay: TimeSpan.FromSeconds(30),
            errorNumbersToAdd: null);
    });
});

// Register Application Services
builder.Services.AddHttpClient(); // For SKU mapping service
builder.Services.AddSingleton<ISkuMappingService, SkuMappingService>();
builder.Services.AddHostedService(sp => (SkuMappingService)sp.GetRequiredService<ISkuMappingService>());
builder.Services.AddScoped<IGraphService, GraphService>();
builder.Services.AddScoped<IUserSettingsService, UserSettingsService>();
builder.Services.AddScoped<IWidgetDataService, WidgetDataService>();
builder.Services.AddScoped<ICacheService, CacheService>();
builder.Services.AddScoped<IExecutiveReportService, ExecutiveReportService>();
builder.Services.AddScoped<IReportService, ReportService>();
builder.Services.AddHostedService<M365Dashboard.Api.Background.ReportSchedulerService>();
builder.Services.AddScoped<IEmailService, GraphEmailService>();
builder.Services.AddScoped<ITenantSettingsService, TenantSettingsService>();
builder.Services.AddSingleton<IDomainSecurityService, DomainSecurityService>();
builder.Services.AddScoped<IExchangeOnlineService, ExchangeOnlineService>();
builder.Services.AddScoped<ICisBenchmarkService, CisBenchmarkService>();
builder.Services.AddScoped<ISecurityAssessmentService, SecurityAssessmentService>();
builder.Services.AddScoped<IDefenderForOfficeService, DefenderForOfficeService>();
builder.Services.AddSingleton<IOsVersionService, OsVersionService>();
builder.Services.AddScoped<WordReportGenerator>();

// Register PDF generator - QuestPDF works on Windows, Linux, and macOS
QuestPDF.Settings.License = QuestPDF.Infrastructure.LicenseType.Community;
builder.Services.AddScoped<PdfReportGenerator>();
Log.Information("PDF report generation enabled (QuestPDF)");

// Configure caching options
builder.Services.Configure<CacheOptions>(builder.Configuration.GetSection("Cache"));

// Add memory cache for hybrid caching strategy
builder.Services.AddMemoryCache();
builder.Services.AddDistributedMemoryCache();

// Configure CORS for development
builder.Services.AddCors(options =>
{
    options.AddPolicy("Development", policy =>
    {
        policy.SetIsOriginAllowed(origin => 
            origin.StartsWith("http://localhost") || 
            origin.StartsWith("https://localhost"))
        .AllowAnyHeader()
        .AllowAnyMethod()
        .AllowCredentials();
    });
});

builder.Services.AddControllers()
    .AddJsonOptions(options =>
    {
        options.JsonSerializerOptions.Converters.Add(new System.Text.Json.Serialization.JsonStringEnumConverter());
    });
builder.Services.AddEndpointsApiExplorer();
// Swagger only in development - never expose API schema in production
if (builder.Environment.IsDevelopment())
{
    builder.Services.AddSwaggerGen(c =>
    {
        c.SwaggerDoc("v1", new() { Title = "M365 Dashboard API", Version = "v1" });
    });
}

var app = builder.Build();

// Configure the HTTP request pipeline
if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();
    app.UseCors("Development");
}
else
{
    app.UseExceptionHandler("/Error");
    app.UseHsts();
}

app.UseHttpsRedirection();

// Security headers
app.Use(async (context, next) =>
{
    context.Response.Headers["X-Content-Type-Options"]    = "nosniff";
    context.Response.Headers["X-Frame-Options"]           = "DENY";
    context.Response.Headers["X-XSS-Protection"]         = "1; mode=block";
    context.Response.Headers["Referrer-Policy"]          = "strict-origin-when-cross-origin";
    context.Response.Headers["Permissions-Policy"]       = "geolocation=(), camera=(), microphone=()";
    context.Response.Headers["Strict-Transport-Security"] = "max-age=31536000; includeSubDomains";
    var csp = app.Configuration["Security:ContentSecurityPolicy"];
    if (!string.IsNullOrEmpty(csp))
        context.Response.Headers["Content-Security-Policy"] = csp;
    await next();
});

app.UseSerilogRequestLogging();

app.UseRouting();

app.UseAuthentication();
app.UseAuthorization();

// Custom middleware for request context
app.UseMiddleware<RequestContextMiddleware>();

app.MapControllers();

// Apply database migrations on startup
using (var scope = app.Services.CreateScope())
{
    var db = scope.ServiceProvider.GetRequiredService<ApplicationDbContext>();
    await db.Database.MigrateAsync();
    Log.Information("Database migrations applied successfully");
}

Log.Information("M365 Dashboard API started");

// Serve React frontend static files
app.UseStaticFiles();
app.MapFallbackToFile("index.html");

app.Run();