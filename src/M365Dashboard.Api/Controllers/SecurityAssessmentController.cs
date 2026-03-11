using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using M365Dashboard.Api.Models;
using M365Dashboard.Api.Services;
using System.Text;
using System.Text.Json;

namespace M365Dashboard.Api.Controllers;

[ApiController]
[Route("api/[controller]")]
[Authorize]
public class SecurityAssessmentController : ControllerBase
{
    private readonly ISecurityAssessmentService _assessmentService;
    private readonly ILogger<SecurityAssessmentController> _logger;
    private readonly IWebHostEnvironment _environment;

    public SecurityAssessmentController(
        ISecurityAssessmentService assessmentService,
        ILogger<SecurityAssessmentController> logger,
        IWebHostEnvironment environment)
    {
        _assessmentService = assessmentService;
        _logger = logger;
        _environment = environment;
    }

    private ReportSettings LoadReportSettings()
    {
        try
        {
            var filePath = Path.Combine(_environment.ContentRootPath, "App_Data", "report-settings.json");
            if (System.IO.File.Exists(filePath))
            {
                var json = System.IO.File.ReadAllText(filePath);
                return JsonSerializer.Deserialize<ReportSettings>(json, new JsonSerializerOptions
                {
                    PropertyNameCaseInsensitive = true
                }) ?? new ReportSettings();
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Could not load report settings, using defaults");
        }
        return new ReportSettings();
    }

    /// <summary>
    /// Run a comprehensive security assessment
    /// </summary>
    [HttpPost("run")]
    [HttpGet("run")]
    public async Task<IActionResult> RunAssessment()
    {
        try
        {
            _logger.LogInformation("Starting Security Assessment");
            var result = await _assessmentService.RunAssessmentAsync();
            _logger.LogInformation("Security Assessment completed: {Compliant}/{Total} checks compliant ({Percentage}%)",
                result.CompliantChecks, result.TotalChecks, result.OverallCompliancePercentage);
            return Ok(result);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error running Security Assessment");
            return StatusCode(500, new { error = "Failed to run security assessment", message = ex.Message });
        }
    }

    /// <summary>
    /// Get a summary of the assessment
    /// </summary>
    [HttpGet("summary")]
    public async Task<IActionResult> GetSummary()
    {
        try
        {
            var result = await _assessmentService.RunAssessmentAsync();
            return Ok(new
            {
                result.GeneratedAt,
                result.TenantName,
                result.TotalChecks,
                result.CompliantChecks,
                result.NonCompliantChecks,
                result.OverallCompliancePercentage,
                UserStats = new
                {
                    result.UserStats.TotalUsers,
                    result.UserStats.LicensedUsers,
                    result.UserStats.BlockedUsers,
                    result.UserStats.GuestUsers
                },
                Sections = new[]
                {
                    new { Name = result.EntraIdCompliance.SectionName, Compliant = result.EntraIdCompliance.CompliantChecks, Total = result.EntraIdCompliance.TotalChecks },
                    new { Name = result.ExchangeCompliance.SectionName, Compliant = result.ExchangeCompliance.CompliantChecks, Total = result.ExchangeCompliance.TotalChecks },
                    new { Name = result.SharePointCompliance.SectionName, Compliant = result.SharePointCompliance.CompliantChecks, Total = result.SharePointCompliance.TotalChecks },
                    new { Name = result.TeamsCompliance.SectionName, Compliant = result.TeamsCompliance.CompliantChecks, Total = result.TeamsCompliance.TotalChecks },
                    new { Name = result.IntuneCompliance.SectionName, Compliant = result.IntuneCompliance.CompliantChecks, Total = result.IntuneCompliance.TotalChecks },
                    new { Name = result.DefenderCompliance.SectionName, Compliant = result.DefenderCompliance.CompliantChecks, Total = result.DefenderCompliance.TotalChecks }
                }
            });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error getting assessment summary");
            return StatusCode(500, new { error = "Failed to get assessment summary", message = ex.Message });
        }
    }

    /// <summary>
    /// View the assessment as HTML report
    /// </summary>
    [HttpGet("html")]
    public async Task<IActionResult> ViewHtmlReport()
    {
        try
        {
            var result = await _assessmentService.RunAssessmentAsync();
            var html = GenerateHtmlReport(result);
            return Content(html, "text/html");
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error generating HTML report");
            return StatusCode(500, new { error = "Failed to generate HTML report", message = ex.Message });
        }
    }

    /// <summary>
    /// Download the assessment as a printable HTML (for PDF)
    /// </summary>
    [HttpGet("download")]
    public async Task<IActionResult> DownloadReport()
    {
        try
        {
            var result = await _assessmentService.RunAssessmentAsync();
            var html = GeneratePrintableHtmlReport(result);
            
            var fileName = $"M365_Security_Assessment_{DateTime.UtcNow:yyyy-MM-dd}.html";
            var bytes = Encoding.UTF8.GetBytes(html);
            return File(bytes, "text/html", fileName);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error generating report");
            return StatusCode(500, new { error = "Failed to generate report", message = ex.Message });
        }
    }

    private string GenerateHtmlReport(SecurityAssessmentResult data)
    {
        return GeneratePrintableHtmlReport(data);
    }

    private string GeneratePrintableHtmlReport(SecurityAssessmentResult data)
    {
        var settings = LoadReportSettings();
        var sb = new StringBuilder();
        
        sb.AppendLine(@"<!DOCTYPE html>
<html>
<head>
    <meta charset='UTF-8'>
    <title>Microsoft 365 Security Assessment</title>
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Segoe+UI:wght@300;400;600;700&display=swap');
        
        * { margin: 0; padding: 0; box-sizing: border-box; }
        
        body { 
            font-family: 'Segoe UI', Arial, sans-serif; 
            background: #f5f5f5; 
            color: #323130;
            line-height: 1.5;
        }
        
        .page { 
            background: white; 
            max-width: 210mm; 
            margin: 20px auto; 
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
        }
        
        @media print {
            .page { 
                margin: 0; 
                box-shadow: none; 
                page-break-after: always;
            }
            .no-print { display: none; }
        }
        
        /* Cover Page */
        .cover-page {
            height: 297mm;
            background: linear-gradient(135deg, #1a365d 0%, #2d3748 100%);
            color: white;
            display: flex;
            flex-direction: column;
            justify-content: space-between;
            padding: 60px;
            position: relative;
            overflow: hidden;
        }
        
        .cover-page::before {
            content: '';
            position: absolute;
            top: 0;
            right: 0;
            width: 60%;
            height: 100%;
            background: url('data:image/svg+xml,<svg xmlns=""http://www.w3.org/2000/svg"" viewBox=""0 0 100 100""><text y="".9em"" font-size=""90"" fill=""rgba(255,255,255,0.03)"">🔒</text></svg>') repeat;
            opacity: 0.5;
        }
        
        .cover-title {
            font-size: 3em;
            font-weight: 300;
            letter-spacing: 2px;
            text-transform: uppercase;
            margin-bottom: 10px;
        }
        
        .cover-subtitle {
            font-size: 1.8em;
            font-weight: 300;
            color: #e07c3a;
            text-transform: uppercase;
            letter-spacing: 4px;
        }
        
        .cover-tenant {
            font-size: 2em;
            font-weight: 600;
            margin: 40px 0;
        }
        
        .cover-date {
            font-size: 1.2em;
            opacity: 0.8;
        }
        
        .cover-logo {
            text-align: right;
            font-size: 1.5em;
            font-weight: 700;
        }
        
        .cover-logo {
            color: #e07c3a;
            font-weight: 600;
        }
        
        /* Content Pages */
        .content-page {
            padding: 40px 50px;
            min-height: 297mm;
        }
        
        .page-header {
            border-bottom: 3px solid #0078d4;
            padding-bottom: 15px;
            margin-bottom: 30px;
        }
        
        .page-title {
            font-size: 1.8em;
            color: #0078d4;
            font-weight: 600;
        }
        
        .page-subtitle {
            color: #666;
            font-size: 0.95em;
            margin-top: 5px;
        }
        
        /* Executive Summary */
        .summary-grid {
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            gap: 20px;
            margin: 30px 0;
        }
        
        .summary-card {
            background: #f8f9fa;
            border-radius: 8px;
            padding: 25px;
            text-align: center;
            border-top: 4px solid #0078d4;
        }
        
        .summary-card.licensed { border-top-color: #107c10; }
        .summary-card.blocked { border-top-color: #d13438; }
        
        .summary-value {
            font-size: 3em;
            font-weight: 700;
            color: #323130;
        }
        
        .summary-label {
            color: #666;
            font-size: 0.9em;
            margin-top: 5px;
        }
        
        .summary-sub {
            color: #999;
            font-size: 0.8em;
        }
        
        /* Tables */
        .data-table {
            width: 100%;
            border-collapse: collapse;
            margin: 20px 0;
            font-size: 0.9em;
        }
        
        .data-table th {
            background: #f0f0f0;
            padding: 12px 15px;
            text-align: left;
            font-weight: 600;
            border-bottom: 2px solid #ddd;
        }
        
        .data-table td {
            padding: 12px 15px;
            border-bottom: 1px solid #eee;
        }
        
        .data-table tr:hover {
            background: #f9f9f9;
        }
        
        /* Compliance Table */
        .compliance-table {
            width: 100%;
            border-collapse: collapse;
            margin: 20px 0;
            font-size: 0.85em;
        }
        
        .compliance-table th {
            background: #1a365d;
            color: white;
            padding: 12px 15px;
            text-align: left;
            font-weight: 500;
        }
        
        .compliance-table td {
            padding: 12px 15px;
            border-bottom: 1px solid #eee;
            vertical-align: top;
        }
        
        .compliance-table tr:nth-child(even) {
            background: #f9f9f9;
        }
        
        .status-compliant {
            color: #107c10;
            font-weight: 600;
        }
        
        .status-noncompliant {
            color: #d13438;
            font-weight: 600;
        }
        
        .status-warning {
            color: #ffb900;
            font-weight: 600;
        }
        
        /* Section Header */
        .section-header {
            background: #1a365d;
            color: white;
            padding: 20px 30px;
            margin: -40px -50px 30px -50px;
        }
        
        .section-header h2 {
            font-size: 1.5em;
            font-weight: 600;
        }
        
        .section-intro {
            color: #666;
            font-size: 0.95em;
            margin-bottom: 25px;
            line-height: 1.6;
        }
        
        /* Role Distribution */
        .role-grid {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 20px;
            margin: 20px 0;
        }
        
        .role-item {
            display: flex;
            justify-content: space-between;
            padding: 10px 0;
            border-bottom: 1px solid #eee;
        }
        
        /* Infographic */
        .infographic-page {
            height: 297mm;
            background: linear-gradient(135deg, #1a1a2e 0%, #16213e 100%);
            display: flex;
            flex-direction: column;
            justify-content: center;
            align-items: center;
            text-align: center;
            color: white;
            padding: 60px;
        }
        
        .infographic-stat {
            font-size: 8em;
            font-weight: 700;
            color: #e07c3a;
            line-height: 1;
        }
        
        .infographic-text {
            font-size: 1.8em;
            font-weight: 300;
            max-width: 600px;
            margin-top: 20px;
        }
        
        .infographic-source {
            margin-top: 40px;
            font-size: 0.9em;
            opacity: 0.7;
        }
        
        /* Footer */
        .page-footer {
            margin-top: auto;
            padding-top: 30px;
            border-top: 1px solid #eee;
            display: flex;
            justify-content: space-between;
            align-items: center;
            font-size: 0.8em;
            color: #999;
        }
        
        .footer-logo {
            font-weight: 700;
            color: #333;
        }
        

    </style>
</head>
<body>");

        // Cover Page - with settings
        var logoHtml = !string.IsNullOrEmpty(settings.LogoBase64) 
            ? $"<img src='data:{settings.LogoContentType};base64,{settings.LogoBase64}' style='max-height:60px;max-width:200px;' alt='Logo' />"
            : $"<span style='color:{settings.AccentColor};font-weight:600;font-size:1.5em;'>{settings.CompanyName}</span>";

        sb.AppendLine($@"
    <div class='page cover-page' style='background: linear-gradient(135deg, {settings.PrimaryColor} 0%, #2d3748 100%);'>
        <div>
            <div class='cover-title'>MICROSOFT 365</div>
            <div class='cover-subtitle' style='color:{settings.AccentColor};'>{settings.ReportTitle.ToUpper().Replace("MICROSOFT 365 ", "")}</div>
        </div>
        <div>
            <div class='cover-tenant'>{data.TenantName}</div>
            <div class='cover-date'>Generated on {data.GeneratedAt:dd MMMM yyyy}</div>
        </div>
        <div class='cover-logo' style='text-align:right;'>
            {logoHtml}
        </div>
    </div>");

        // Executive Summary Page
        sb.AppendLine($@"
    <div class='page content-page'>
        <div class='page-header'>
            <div class='page-title'>Executive Summary</div>
        </div>
        
        <div class='summary-grid'>
            <div class='summary-card'>
                <div class='summary-value'>{data.UserStats.TotalUsers}</div>
                <div class='summary-label'>Total Users</div>
                <div class='summary-sub'>Including {data.UserStats.GuestUsers} guest users</div>
            </div>
            <div class='summary-card licensed'>
                <div class='summary-value'>{data.UserStats.LicensedUsers}</div>
                <div class='summary-label'>Licensed Users</div>
                <div class='summary-sub'>{data.UserStats.UnlicensedUsers} unlicensed users</div>
            </div>
            <div class='summary-card blocked'>
                <div class='summary-value'>{data.UserStats.BlockedUsers}</div>
                <div class='summary-label'>Blocked Users</div>
                <div class='summary-sub'>{data.UserStats.BlockedUsersWithLicenses} blocked & licensed</div>
            </div>
        </div>
        
        <p style='color:#666;line-height:1.8;margin:20px 0;'>
            This report was prepared for <strong>{data.TenantName}</strong> in {data.GeneratedAt:MMMM yyyy}. 
            This Microsoft 365 Security Report provides a comprehensive analysis of over {data.TotalChecks} critical security checks 
            across key Microsoft 365 services, including Entra ID (Azure AD), Exchange Online, Intune, SharePoint, and Teams. 
            The report identifies and evaluates your organization's configuration against best practices outlined by the 
            Center for Internet Security (CIS) and the National Cyber Security Centre (NCSC).
        </p>
        
        <h3 style='margin-top:30px;color:#323130;'>User and Role Distribution</h3>
        <div class='role-grid'>
            <div>
                <table class='data-table'>
                    <tr><th>Description</th><th style='text-align:right'>Count</th></tr>
                    <tr><td>Member Users</td><td style='text-align:right'>{data.UserStats.MemberUsers}</td></tr>
                    <tr><td>Guest Users</td><td style='text-align:right'>{data.UserStats.GuestUsers}</td></tr>
                    <tr><td>Total of All Users</td><td style='text-align:right'>{data.UserStats.TotalUsers}</td></tr>
                    <tr><td>Licensed Users</td><td style='text-align:right'>{data.UserStats.LicensedUsers}</td></tr>
                    <tr><td>Unlicensed Users</td><td style='text-align:right'>{data.UserStats.UnlicensedUsers}</td></tr>
                    <tr><td>Blocked Users</td><td style='text-align:right'>{data.UserStats.BlockedUsers}</td></tr>
                    <tr><td>Blocked Users with Licenses</td><td style='text-align:right'>{data.UserStats.BlockedUsersWithLicenses}</td></tr>
                </table>
            </div>
            <div>
                <table class='data-table'>
                    <tr><th>Role</th><th style='text-align:right'>Count</th></tr>
                    {string.Join("", data.RoleDistribution.Take(8).Select(r => $"<tr><td>{r.RoleName}</td><td style='text-align:right'>{r.MemberCount}</td></tr>"))}
                </table>
            </div>
        </div>
        
        <div class='page-footer'>
            <div class='footer-logo'>{settings.CompanyName}</div>
            {(string.IsNullOrEmpty(settings.FooterText) ? "" : $"<div>{settings.FooterText}</div>")}
        </div>
    </div>");

        // Infographic Page (conditional)
        if (settings.ShowInfoGraphics)
        {
            sb.AppendLine($@"
    <div class='page infographic-page'>
        <div class='infographic-stat' style='color:{settings.AccentColor};'>3%</div>
        <div class='infographic-text'>
            Only <strong>3%</strong> of businesses have implemented <strong>all recommended</strong> security steps...
        </div>
        <div class='infographic-text' style='margin-top:60px;color:{settings.AccentColor};'>
            ...We make <strong>best practices</strong> your <strong>standard practice</strong>
        </div>
        <div class='infographic-source'>
            Source: Cyber security breaches survey 2024<br/>Gov.uk
        </div>
    </div>");
        }

        // Entra ID Compliance Section
        sb.AppendLine(GenerateComplianceSectionHtml(data.EntraIdCompliance, settings));

        // Infographic - Phishing (conditional)
        if (settings.ShowInfoGraphics)
        {
            sb.AppendLine($@"
    <div class='page infographic-page'>
        <div class='infographic-stat' style='color:{settings.AccentColor};'>84%</div>
        <div class='infographic-text'>
            of businesses fell victim to <strong style='color:{settings.AccentColor};'>phishing attacks</strong> in 2024
        </div>
        <div class='infographic-source'>
            Source: Cyber security breaches survey 2024<br/>Gov.uk
        </div>
    </div>");
        }

        // Exchange Compliance Section
        sb.AppendLine(GenerateComplianceSectionHtml(data.ExchangeCompliance, settings));

        // Infographic - 60 seconds (conditional)
        if (settings.ShowInfoGraphics)
        {
            sb.AppendLine($@"
    <div class='page infographic-page'>
        <div class='infographic-stat' style='color:{settings.AccentColor};'>60</div>
        <div class='infographic-text'>
            <strong>Seconds</strong><br/><br/>
            is all it takes for an employee to fall for a <strong style='color:{settings.AccentColor};'>phishing attack</strong>
        </div>
        <div class='infographic-source'>
            Source: verizon.com
        </div>
    </div>");
        }

        // SharePoint Compliance Section
        sb.AppendLine(GenerateComplianceSectionHtml(data.SharePointCompliance, settings));

        // Infographic - Licenses (conditional)
        if (settings.ShowInfoGraphics)
        {
            sb.AppendLine($@"
    <div class='page infographic-page'>
        <div class='infographic-stat' style='color:{settings.AccentColor};'>18%</div>
        <div class='infographic-text'>
            of Microsoft 365 licenses are left <strong style='color:{settings.AccentColor};'>unassigned</strong>.
        </div>
        <div class='infographic-source'>
            Source: How to close the 365 license management gap.<br/>Quest.com
        </div>
    </div>");
        }

        // Teams Compliance Section
        sb.AppendLine(GenerateComplianceSectionHtml(data.TeamsCompliance, settings));

        // Infographic - Stolen Credentials (conditional)
        if (settings.ShowInfoGraphics)
        {
            sb.AppendLine($@"
    <div class='page infographic-page'>
        <div class='infographic-stat' style='color:{settings.AccentColor};'>31%</div>
        <div class='infographic-text'>
            of all breaches over the past <strong>10 years</strong> have involved the use of <strong style='color:{settings.AccentColor};'>stolen credentials</strong>.
        </div>
        <div class='infographic-source'>
            Source: verizon.com
        </div>
    </div>");
        }

        // Intune Compliance Section
        sb.AppendLine(GenerateComplianceSectionHtml(data.IntuneCompliance, settings));

        // Defender Compliance Section
        sb.AppendLine(GenerateComplianceSectionHtml(data.DefenderCompliance, settings));

        sb.AppendLine(@"
</body>
</html>");

        return sb.ToString();
    }

    private string GenerateComplianceSectionHtml(ComplianceSection section, ReportSettings settings)
    {
        var checksHtml = new StringBuilder();
        
        foreach (var check in section.Checks)
        {
            var statusClass = check.Status switch
            {
                SecurityCheckStatus.Compliant => "status-compliant",
                SecurityCheckStatus.NonCompliant => "status-noncompliant",
                _ => "status-warning"
            };
            
            var statusText = check.Status switch
            {
                SecurityCheckStatus.Compliant => "Compliant",
                SecurityCheckStatus.NonCompliant => "Non-Compliant",
                SecurityCheckStatus.Warning => "Warning",
                SecurityCheckStatus.NotApplicable => "N/A",
                _ => "Unknown"
            };
            
            checksHtml.AppendLine($@"
                <tr>
                    <td style='width:30%'><strong>{check.Name}</strong>{(check.IsBeta ? " (Beta)" : "")}</td>
                    <td style='width:45%'>{check.Description}</td>
                    <td style='width:12%;text-align:center'>{check.CheckedAt:dd MMM yyyy}</td>
                    <td style='width:13%' class='{statusClass}'>{statusText}</td>
                </tr>");
        }

        var footerContent = string.IsNullOrEmpty(settings.FooterText) 
            ? "" 
            : $"<div style='color:#666;'>{settings.FooterText}</div>";

        return $@"
    <div class='page content-page'>
        <div class='section-header' style='background:{settings.PrimaryColor};'>
            <h2>{section.SectionName} Compliance</h2>
        </div>
        
        <p class='section-intro'>{section.SectionDescription}</p>
        
        <table class='compliance-table'>
            <thead>
                <tr>
                    <th style='background:{settings.PrimaryColor};'>Name</th>
                    <th style='background:{settings.PrimaryColor};'>Description</th>
                    <th style='background:{settings.PrimaryColor};text-align:center'>Date</th>
                    <th style='background:{settings.PrimaryColor};'>Status</th>
                </tr>
            </thead>
            <tbody>
                {checksHtml}
            </tbody>
        </table>
        
        <div class='page-footer'>
            <div class='footer-logo'>{settings.CompanyName}</div>
            {footerContent}
        </div>
    </div>";
    }
}
