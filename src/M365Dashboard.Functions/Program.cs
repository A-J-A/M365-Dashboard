using Microsoft.Azure.Functions.Worker;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Graph;
using Azure.Identity;
using M365Dashboard.Api.Data;
using M365Dashboard.Api.Services;

var host = new HostBuilder()
    .ConfigureFunctionsWorkerDefaults()
    .ConfigureServices((context, services) =>
    {
        var configuration = context.Configuration;

        // Configure Application Insights
        services.AddApplicationInsightsTelemetryWorkerService();
        services.ConfigureFunctionsApplicationInsights();

        // Configure Entity Framework with SQL Server
        services.AddDbContext<ApplicationDbContext>(options =>
        {
            var connectionString = configuration.GetConnectionString("DefaultConnection");
            options.UseSqlServer(connectionString, sqlOptions =>
            {
                sqlOptions.EnableRetryOnFailure(
                    maxRetryCount: 5,
                    maxRetryDelay: TimeSpan.FromSeconds(30),
                    errorNumbersToAdd: null);
            });
        });

        // Configure Microsoft Graph Client with Client Credentials
        services.AddSingleton<GraphServiceClient>(sp =>
        {
            var tenantId = configuration["AzureAd:TenantId"];
            var clientId = configuration["AzureAd:ClientId"];
            var clientSecret = configuration["AzureAd:ClientSecret"];

            var credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
            return new GraphServiceClient(credential, new[] { "https://graph.microsoft.com/.default" });
        });

        // Register Services
        services.AddScoped<IGraphService, GraphService>();
        services.AddScoped<IReportService, ReportService>();
        services.AddScoped<IEmailService, GraphEmailService>();
        services.AddScoped<ICacheService, CacheService>();
        
        // Add memory cache
        services.AddMemoryCache();
        services.AddDistributedMemoryCache();
    })
    .Build();

host.Run();
