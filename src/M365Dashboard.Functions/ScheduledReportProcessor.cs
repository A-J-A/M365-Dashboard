using System.Text;
using System.Text.Json;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;
using M365Dashboard.Api.Models.Dtos;
using M365Dashboard.Api.Services;

namespace M365Dashboard.Functions;

public class ScheduledReportProcessor
{
    private readonly IReportService _reportService;
    private readonly IEmailService _emailService;
    private readonly ILogger<ScheduledReportProcessor> _logger;

    public ScheduledReportProcessor(
        IReportService reportService,
        IEmailService emailService,
        ILogger<ScheduledReportProcessor> logger)
    {
        _reportService = reportService;
        _emailService = emailService;
        _logger = logger;
    }

    /// <summary>
    /// Runs every 15 minutes to check for due scheduled reports
    /// CRON: "0 */15 * * * *" = Every 15 minutes
    /// </summary>
    [Function("ProcessScheduledReports")]
    public async Task ProcessScheduledReports([TimerTrigger("0 */15 * * * *")] TimerInfo timerInfo)
    {
        _logger.LogInformation("Processing scheduled reports at {Time}", DateTime.UtcNow);

        try
        {
            var dueReports = await _reportService.GetDueScheduledReportsAsync();
            
            _logger.LogInformation("Found {Count} due reports to process", dueReports.Count);

            foreach (var schedule in dueReports)
            {
                try
                {
                    _logger.LogInformation("Processing scheduled report {Id}: {ReportType} for user {UserId}", 
                        schedule.Id, schedule.ReportType, schedule.UserId);

                    // Generate the report
                    var request = new GenerateReportRequest(
                        schedule.ReportType,
                        schedule.DateRange,
                        "csv"
                    );

                    var report = await _reportService.GenerateReportAsync(
                        request, 
                        schedule.UserId, 
                        isScheduled: true, 
                        scheduledReportId: schedule.Id
                    );

                    // Convert report to CSV for email attachment
                    var csvContent = await _reportService.ExportReportToCsvAsync(request);
                    var csvBytes = Encoding.UTF8.GetBytes(csvContent);
                    var fileName = $"{schedule.ReportType}_{DateTime.UtcNow:yyyyMMdd}.csv";

                    // Build email content
                    var recipients = schedule.Recipients
                        .Split(',', StringSplitOptions.RemoveEmptyEntries)
                        .Select(r => r.Trim())
                        .ToList();

                    var subject = $"[M365 Dashboard] Scheduled Report: {schedule.DisplayName}";
                    var body = BuildEmailBody(schedule.DisplayName, report, schedule.DateRange);

                    // Send the email
                    var senderEmail = schedule.UserEmail ?? recipients.FirstOrDefault() ?? "noreply@localhost";
                    await _emailService.SendReportEmailAsync(
                        senderEmail,
                        recipients,
                        subject,
                        body,
                        fileName,
                        csvBytes
                    );

                    // Update schedule status
                    await _reportService.UpdateScheduledReportAfterRunAsync(schedule.Id, true);
                    
                    _logger.LogInformation("Successfully processed scheduled report {Id}", schedule.Id);
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "Failed to process scheduled report {Id}: {Error}", schedule.Id, ex.Message);
                    await _reportService.UpdateScheduledReportAfterRunAsync(schedule.Id, false, ex.Message);
                }
            }
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error in scheduled report processor: {Error}", ex.Message);
            throw;
        }

        _logger.LogInformation("Scheduled report processing completed at {Time}", DateTime.UtcNow);
    }

    private static string BuildEmailBody(string reportName, ReportResultDto report, string? dateRange)
    {
        var summaryHtml = "";
        if (report.Summary?.Highlights != null)
        {
            var highlights = report.Summary.Highlights
                .Select(kvp => $"<li><strong>{FormatKeyName(kvp.Key)}:</strong> {kvp.Value}</li>");
            summaryHtml = $@"
                <h3>Summary</h3>
                <ul>
                    <li><strong>Total Records:</strong> {report.Summary.TotalRecords}</li>
                    {string.Join("", highlights)}
                </ul>";
        }

        return $@"
<!DOCTYPE html>
<html>
<head>
    <style>
        body {{ font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 0; padding: 20px; background-color: #f5f5f5; }}
        .container {{ max-width: 600px; margin: 0 auto; background-color: #ffffff; border-radius: 8px; overflow: hidden; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }}
        .header {{ background-color: #0078d4; color: white; padding: 20px; text-align: center; }}
        .header h1 {{ margin: 0; font-size: 24px; }}
        .content {{ padding: 30px; }}
        .content h2 {{ color: #323130; margin-top: 0; }}
        .content h3 {{ color: #605e5c; margin-top: 20px; }}
        .content ul {{ padding-left: 20px; }}
        .content li {{ margin: 8px 0; color: #323130; }}
        .meta {{ background-color: #f3f2f1; padding: 15px; margin-top: 20px; border-radius: 4px; }}
        .meta p {{ margin: 5px 0; color: #605e5c; font-size: 14px; }}
        .footer {{ background-color: #faf9f8; padding: 20px; text-align: center; color: #605e5c; font-size: 12px; }}
        .attachment-note {{ background-color: #fff4ce; border-left: 4px solid #ffb900; padding: 10px 15px; margin-top: 20px; }}
    </style>
</head>
<body>
    <div class='container'>
        <div class='header'>
            <h1>M365 Dashboard Report</h1>
        </div>
        <div class='content'>
            <h2>{reportName}</h2>
            <p>Your scheduled report has been generated and is attached to this email.</p>
            
            {summaryHtml}
            
            <div class='meta'>
                <p><strong>Report Type:</strong> {report.ReportType}</p>
                <p><strong>Date Range:</strong> {dateRange ?? "N/A"}</p>
                <p><strong>Generated:</strong> {report.GeneratedAt:yyyy-MM-dd HH:mm:ss} UTC</p>
            </div>
            
            <div class='attachment-note'>
                <strong>📎 Attachment:</strong> The full report data is attached as a CSV file.
            </div>
        </div>
        <div class='footer'>
            <p>This is an automated report from M365 Dashboard.</p>
            <p>To manage your scheduled reports, visit the Reports page in the dashboard.</p>
        </div>
    </div>
</body>
</html>";
    }

    private static string FormatKeyName(string key)
    {
        // Convert camelCase to Title Case with spaces
        var result = new StringBuilder();
        foreach (var c in key)
        {
            if (char.IsUpper(c) && result.Length > 0)
            {
                result.Append(' ');
            }
            result.Append(result.Length == 0 ? char.ToUpper(c) : c);
        }
        return result.ToString();
    }
}
