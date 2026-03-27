using M365Dashboard.Api.Models;
using M365Dashboard.Api.Models.Dtos;
using M365Dashboard.Api.Services;

namespace M365Dashboard.Api.Background;

/// <summary>
/// Hosted background service that polls for due scheduled reports every minute,
/// generates the appropriate output and emails it to recipients.
/// </summary>
public class ReportSchedulerService : BackgroundService
{
    private readonly IServiceScopeFactory _scopeFactory;
    private readonly ILogger<ReportSchedulerService> _logger;
    private readonly TimeSpan _pollInterval = TimeSpan.FromMinutes(1);

    public ReportSchedulerService(IServiceScopeFactory scopeFactory, ILogger<ReportSchedulerService> logger)
    {
        _scopeFactory = scopeFactory;
        _logger = logger;
    }

    protected override async Task ExecuteAsync(CancellationToken stoppingToken)
    {
        _logger.LogInformation("Report Scheduler Service started");

        // Stagger startup by 30 seconds so the app fully initialises before first poll
        await Task.Delay(TimeSpan.FromSeconds(30), stoppingToken);

        while (!stoppingToken.IsCancellationRequested)
        {
            try
            {
                await ProcessDueReportsAsync(stoppingToken);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Unhandled error in Report Scheduler poll");
            }

            await Task.Delay(_pollInterval, stoppingToken);
        }

        _logger.LogInformation("Report Scheduler Service stopped");
    }

    private async Task ProcessDueReportsAsync(CancellationToken ct)
    {
        using var scope = _scopeFactory.CreateScope();
        var reportService          = scope.ServiceProvider.GetRequiredService<IReportService>();
        var emailService           = scope.ServiceProvider.GetRequiredService<IEmailService>();
        var executiveReportService = scope.ServiceProvider.GetRequiredService<IExecutiveReportService>();
        var tenantSettingsService  = scope.ServiceProvider.GetRequiredService<ITenantSettingsService>();
        var configuration          = scope.ServiceProvider.GetRequiredService<IConfiguration>();

        var due = await reportService.GetDueScheduledReportsAsync();

        if (due.Count == 0) return;

        _logger.LogInformation("Processing {Count} due scheduled report(s)", due.Count);

        // Load branding settings once for this batch
        var tenantId = configuration["AzureAd:TenantId"] ?? "default";
        ReportSettings reportSettings;
        try
        {
            reportSettings = await tenantSettingsService.GetReportSettingsAsync(tenantId);
        }
        catch
        {
            reportSettings = new ReportSettings();
        }

        var senderEmail = reportSettings.SenderEmail
            ?? configuration["ReportSettings:SenderEmail"]
            ?? "noreply@example.com";

        foreach (var schedule in due)
        {
            if (ct.IsCancellationRequested) break;
            await RunScheduleAsync(schedule, reportService, emailService, executiveReportService, reportSettings, senderEmail);
        }
    }

    private async Task RunScheduleAsync(
        ScheduledReport schedule,
        IReportService reportService,
        IEmailService emailService,
        IExecutiveReportService executiveReportService,
        ReportSettings reportSettings,
        string senderEmail)
    {
        _logger.LogInformation("Running scheduled report {Id} ({Type}) for user {User}",
            schedule.Id, schedule.ReportType, schedule.UserId);

        var recipients = schedule.Recipients
            .Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .ToList();

        if (recipients.Count == 0)
        {
            _logger.LogWarning("Scheduled report {Id} has no recipients — skipping", schedule.Id);
            await reportService.UpdateScheduledReportAfterRunAsync(schedule.Id, false, "No recipients configured");
            return;
        }

        try
        {
            if (schedule.ReportType == "executive-summary-pdf")
            {
                await RunExecutivePdfAsync(schedule, recipients, reportService, emailService, executiveReportService, reportSettings, senderEmail);
            }
            else
            {
                await RunStandardReportAsync(schedule, recipients, reportService, emailService, reportSettings, senderEmail);
            }

            await reportService.UpdateScheduledReportAfterRunAsync(schedule.Id, true);
            _logger.LogInformation("Scheduled report {Id} completed successfully", schedule.Id);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Scheduled report {Id} failed", schedule.Id);
            await reportService.UpdateScheduledReportAfterRunAsync(schedule.Id, false, ex.Message);
        }
    }

    // ── Executive Summary PDF ─────────────────────────────────────────────────

    private async Task RunExecutivePdfAsync(
        ScheduledReport schedule,
        List<string> recipients,
        IReportService reportService,
        IEmailService emailService,
        IExecutiveReportService executiveReportService,
        ReportSettings reportSettings,
        string senderEmail)
    {
        var pdfBytes = await executiveReportService.GeneratePdfAsync();
        var fileName = $"{reportSettings.CompanyName.Replace(" ", "_")}_Executive_Summary_{DateTime.UtcNow:yyyy-MM-dd}.pdf";

        var subject = $"{reportSettings.CompanyName} – Executive Summary Report – {DateTime.UtcNow:MMMM yyyy}";
        var body    = BuildEmailBody(reportSettings, schedule.DisplayName, "Please find attached your scheduled Executive Summary Report.");

        await emailService.SendReportEmailAsync(senderEmail, recipients, subject, body, fileName, pdfBytes);
    }

    // ── Standard (CSV / HTML) Reports ────────────────────────────────────────

    private async Task RunStandardReportAsync(
        ScheduledReport schedule,
        List<string> recipients,
        IReportService reportService,
        IEmailService emailService,
        ReportSettings reportSettings,
        string senderEmail)
    {
        var request = new GenerateReportRequest(
            schedule.ReportType,
            schedule.DateRange ?? "last30days",
            "html"
        );

        // Generate HTML body for email
        string htmlContent;
        try
        {
            htmlContent = await reportService.ExportReportToHtmlAsync(request);
        }
        catch (NotImplementedException)
        {
            // Fall back to CSV attachment if HTML not implemented
            var csvContent = await reportService.ExportReportToCsvAsync(request with { Format = "csv", DateRange = schedule.DateRange ?? "last30days" });
            var csvBytes   = System.Text.Encoding.UTF8.GetBytes(csvContent);
            var fileName   = $"{schedule.ReportType}_{DateTime.UtcNow:yyyy-MM-dd}.csv";
            var subject    = $"{reportSettings.CompanyName} – {schedule.DisplayName} – {DateTime.UtcNow:dd MMM yyyy}";
            var body       = BuildEmailBody(reportSettings, schedule.DisplayName, "Please find the attached CSV report.");

            await emailService.SendReportEmailAsync(senderEmail, recipients, subject, body, fileName, csvBytes);
            return;
        }

        // Inline HTML report — embed it in a branded wrapper and send directly
        var emailSubject = $"{reportSettings.CompanyName} – {schedule.DisplayName} – {DateTime.UtcNow:dd MMM yyyy}";

        await emailService.SendReportEmailAsync(senderEmail, recipients, emailSubject, htmlContent);
    }

    // ── Email body helper ─────────────────────────────────────────────────────

    private static string BuildEmailBody(ReportSettings settings, string reportName, string intro)
    {
        var primary = settings.PrimaryColor ?? "#1E3A5F";
        var accent  = settings.AccentColor  ?? "#E07C3A";
        var company = System.Net.WebUtility.HtmlEncode(settings.CompanyName);
        var name    = System.Net.WebUtility.HtmlEncode(reportName);
        var introEncoded = System.Net.WebUtility.HtmlEncode(intro);
        var footer  = System.Net.WebUtility.HtmlEncode(
            !string.IsNullOrEmpty(settings.FooterText) ? settings.FooterText
            : $"This report was automatically generated by {settings.CompanyName}.");
        var date = DateTime.UtcNow.ToString("dd MMMM yyyy");

        return
            "<!DOCTYPE html><html><head><meta charset='utf-8'/><style>" +
            "body{font-family:'Segoe UI',Arial,sans-serif;margin:0;padding:0;background:#f4f4f4}" +
            ".wrapper{max-width:640px;margin:30px auto;background:#fff;border-radius:8px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,.1)}" +
            $".header{{background:{primary};padding:28px 36px}}" +
            ".header h1{color:#fff;margin:0;font-size:22px;font-weight:600}" +
            ".header p{color:rgba(255,255,255,.8);margin:4px 0 0;font-size:13px}" +
            ".body{padding:28px 36px;color:#374151;font-size:15px;line-height:1.6}" +
            $".accent{{color:{accent};font-weight:600}}" +
            ".footer{padding:16px 36px;background:#f9fafb;border-top:1px solid #e5e7eb;color:#6b7280;font-size:12px}" +
            "</style></head><body>" +
            "<div class='wrapper'>" +
            $"<div class='header'><h1>{company}</h1><p><span class='accent'>{name}</span> &mdash; {date}</p></div>" +
            $"<div class='body'><p>{introEncoded}</p></div>" +
            $"<div class='footer'>{footer}</div>" +
            "</div></body></html>";
    }
}
