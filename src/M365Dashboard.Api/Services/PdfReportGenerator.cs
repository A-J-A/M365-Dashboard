using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;
using M365Dashboard.Api.Models;
using M365Dashboard.Api.Controllers;

namespace M365Dashboard.Api.Services;

/// <summary>
/// Professional PDF report generator matching the Cloud1st PDF design
/// </summary>
public class PdfReportGenerator : IDocument
{
    private ExecutiveReportData _data = null!;
    private ReportSettings _settings = null!;
    private byte[]? _logoBytes;
    private List<ReportQuote> _selectedQuotes = new();
    
    // Branding colors (matching the PDF example)
    private string _primaryColor = "#1E3A5F";  // Dark navy blue
    private string _accentColor = "#E07C3A";   // Orange
    private const string CompliantColor = "#107C6C";  // Teal
    private const string NonCompliantColor = "#6B7280"; // Gray
    private const string WarningColor = "#F59E0B";  // Amber
    private const string CriticalColor = "#DC2626"; // Red
    private const string TextColor = "#374151";
    private const string LightGray = "#F9FAFB";
    private const string BorderColor = "#E5E7EB";

    public byte[] GenerateReport(ExecutiveReportData data, ReportSettings settings)
    {
        // Configure QuestPDF license (Community license)
        QuestPDF.Settings.License = LicenseType.Community;
        
        _data = data;
        _settings = settings;
        
        // Apply custom colors from settings
        if (!string.IsNullOrEmpty(settings.PrimaryColor))
            _primaryColor = settings.PrimaryColor.StartsWith("#") ? settings.PrimaryColor : $"#{settings.PrimaryColor}";
        if (!string.IsNullOrEmpty(settings.AccentColor))
            _accentColor = settings.AccentColor.StartsWith("#") ? settings.AccentColor : $"#{settings.AccentColor}";
        
        // Load logo if available
        if (!string.IsNullOrEmpty(settings.LogoBase64))
        {
            try { _logoBytes = Convert.FromBase64String(settings.LogoBase64); }
            catch { _logoBytes = null; }
        }

        // Pick 3 random quotes from the settings pool
        _selectedQuotes = PickRandomQuotes(settings, 3);
        
        return Document.Create(Compose).GeneratePdf();
    }

    // QuestPDF font helper - falls back gracefully if Segoe UI isn't installed
    private static TextStyle SafeFont(string preferred = "Segoe UI") =>
        TextStyle.Default.FontFamily(preferred, "DejaVu Sans", "Arial", "Liberation Sans");

    public void Compose(IDocumentContainer container)
    {
        container.Page(page =>
        {
            page.Size(PageSizes.A4);
            page.Margin(0);
            page.DefaultTextStyle(x => x.FontFamily("Segoe UI", "DejaVu Sans", "Liberation Sans", "Arial").FontColor(TextColor));
            
            page.Content().Column(column =>
            {
                // === COVER PAGE ===
                ComposeCoverPage(column);
                
                // === EXECUTIVE SUMMARY ===
                column.Item().PageBreak();
                ComposeExecutiveSummaryPage(column);
                
                // === INFOGRAPHIC 1 (if enabled) ===
                if (_settings.ShowInfoGraphics && _settings.ShowQuotes && _selectedQuotes.Count > 0)
                {
                    column.Item().PageBreak();
                    ComposeInfoGraphicPage(column,
                        _selectedQuotes[0].BigNumber,
                        _selectedQuotes[0].Line1,
                        _selectedQuotes[0].Line2,
                        _selectedQuotes[0].Source);
                }
                
                // === SECURITY METRICS ===
                column.Item().PageBreak();
                ComposeSecurityMetricsPage(column);
                
                // === INFOGRAPHIC 2 (if enabled) ===
                if (_settings.ShowInfoGraphics && _settings.ShowQuotes && _selectedQuotes.Count > 1)
                {
                    column.Item().PageBreak();
                    ComposeInfoGraphicPage(column,
                        _selectedQuotes[1].BigNumber,
                        _selectedQuotes[1].Line1,
                        _selectedQuotes[1].Line2,
                        _selectedQuotes[1].Source);
                }
                
                // === DEVICE DETAILS ===
                if (_data.DeviceDetails != null && HasDevices())
                {
                    column.Item().PageBreak();
                    ComposeDeviceDetailsPage(column);
                }
                
                // === INFOGRAPHIC 3 (if enabled) ===
                if (_settings.ShowInfoGraphics && _settings.ShowQuotes && _selectedQuotes.Count > 2)
                {
                    column.Item().PageBreak();
                    ComposeInfoGraphicPage(column,
                        _selectedQuotes[2].BigNumber,
                        _selectedQuotes[2].Line1,
                        _selectedQuotes[2].Line2,
                        _selectedQuotes[2].Source);
                }
                
                // === USER DETAILS ===
                if (_data.UserSignInDetails?.Any() == true)
                {
                    column.Item().PageBreak();
                    ComposeUserDetailsPage(column);
                }
                
                // === DOMAIN SECURITY ===
                if (_data.DomainSecuritySummary != null)
                {
                    column.Item().PageBreak();
                    ComposeDomainSecurityPage(column);
                }
            });
        });
    }

    #region Cover Page
    
    private void ComposeCoverPage(ColumnDescriptor column)
    {
        // Full-page dark blue background with title
        column.Item().Height(600).Background(_primaryColor).Padding(50).Column(coverCol =>
        {
            coverCol.Item().Height(120); // Top spacing
            
            // Parse title
            var titleParts = ParseReportTitle(_settings.ReportTitle);
            
            // "MICROSOFT 365" in orange
            coverCol.Item().Text(titleParts.line1)
                .FontSize(28)
                .FontColor(_accentColor)
                .LetterSpacing(0.15f)
                .FontFamily("Segoe UI");
            
            coverCol.Item().Height(5);
            
            // Main title in white
            coverCol.Item().Text(titleParts.line2)
                .FontSize(48)
                .FontColor(Colors.White)
                .FontFamily("Segoe UI Light")
                .LetterSpacing(0.05f);
        });
        
        // White section with company info
        column.Item().ExtendVertical().Background(Colors.White).Padding(50).AlignCenter().Column(infoCol =>
        {
            infoCol.Item().Height(40);
            
            // Company name
            infoCol.Item().AlignCenter().Text(_settings.CompanyName)
                .FontSize(24)
                .Bold()
                .FontColor(_primaryColor);
            
            infoCol.Item().Height(15);
            
            // Date
            infoCol.Item().AlignCenter().Text($"Generated on {_data.GeneratedAt:d MMMM yyyy}")
                .FontSize(12)
                .FontColor(Colors.Grey.Medium);
            
            infoCol.Item().Height(40);
            
            // Logo
            if (_logoBytes != null)
            {
                infoCol.Item().AlignCenter().Height(60).Image(_logoBytes).FitHeight();
            }
        });
    }
    
    #endregion

    #region Executive Summary Page
    
    private void ComposeExecutiveSummaryPage(ColumnDescriptor column)
    {
        column.Item().Padding(40).Column(content =>
        {
            // Page title
            content.Item().Text("Executive Summary")
                .FontSize(36)
                .FontColor(_primaryColor)
                .FontFamily("Segoe UI Light");
            
            content.Item().Height(25);
            
            // KPI Cards Row
            content.Item().Row(row =>
            {
                // Total Users
                row.RelativeItem().Border(1).BorderColor(BorderColor).Padding(20).Column(kpi =>
                {
                    kpi.Item().Row(r =>
                    {
                        r.AutoItem().Text("👤").FontSize(24);
                        r.RelativeItem().AlignRight().Text($"{_data.UserStats?.TotalUsers ?? 0}")
                            .FontSize(32)
                            .FontColor(_primaryColor)
                            .FontFamily("Segoe UI Light");
                    });
                    kpi.Item().Height(5);
                    kpi.Item().Text("Total Users").FontSize(12).Bold();
                    kpi.Item().Text($"Including {_data.UserStats?.GuestUsers ?? 0} guest users")
                        .FontSize(10).FontColor(Colors.Grey.Medium);
                });
                
                row.ConstantItem(15);
                
                // Licensed Users
                row.RelativeItem().Border(1).BorderColor(BorderColor).Padding(20).Column(kpi =>
                {
                    kpi.Item().Row(r =>
                    {
                        r.AutoItem().Text("📋").FontSize(24);
                        r.RelativeItem().AlignRight().Text($"{(_data.UserStats?.TotalUsers ?? 0) - (_data.UserStats?.GuestUsers ?? 0)}")
                            .FontSize(32)
                            .FontColor(_primaryColor)
                            .FontFamily("Segoe UI Light");
                    });
                    kpi.Item().Height(5);
                    kpi.Item().Text("Licensed Users").FontSize(12).Bold();
                    kpi.Item().Text($"{_data.UserStats?.GuestUsers ?? 0} unlicensed users")
                        .FontSize(10).FontColor(Colors.Grey.Medium);
                });
                
                row.ConstantItem(15);
                
                // MFA Registered
                row.RelativeItem().Border(1).BorderColor(BorderColor).Padding(20).Column(kpi =>
                {
                    kpi.Item().Row(r =>
                    {
                        r.AutoItem().Text("🔐").FontSize(24);
                        r.RelativeItem().AlignRight().Text($"{_data.UserStats?.MfaRegistered ?? 0}")
                            .FontSize(32)
                            .FontColor(_primaryColor)
                            .FontFamily("Segoe UI Light");
                    });
                    kpi.Item().Height(5);
                    kpi.Item().Text("MFA Registered").FontSize(12).Bold();
                    kpi.Item().Text($"{_data.UserStats?.MfaNotRegistered ?? 0} not registered")
                        .FontSize(10).FontColor(Colors.Grey.Medium);
                });
            });
            
            content.Item().Height(25);
            
            // Introduction paragraphs
            content.Item().Text(text =>
            {
                text.Span($"This report was prepared for {_settings.CompanyName} in {_data.GeneratedAt:MMMM yyyy}. ")
                    .FontSize(11);
                text.Span($"This {_settings.ReportTitle} provides a comprehensive analysis of your organization's security configuration across key Microsoft 365 services, including Entra ID (Azure AD), Exchange Online, Intune, SharePoint, and Teams.")
                    .FontSize(11);
            });
            
            content.Item().Height(15);
            
            content.Item().Text("The aim of this review is to provide a clear and actionable understanding of your current security posture within Microsoft 365, helping to mitigate potential risks, safeguard sensitive data, and ensure compliance with leading security benchmarks.")
                .FontSize(11);
            
            content.Item().Height(30);
            
            // User and Role Distribution Table
            content.Item().Text("User and Role Distribution")
                .FontSize(20)
                .FontColor(_primaryColor)
                .FontFamily("Segoe UI Light");
            
            content.Item().Height(15);
            
            content.Item().Row(row =>
            {
                // Left table - User counts
                row.RelativeItem().Table(table =>
                {
                    table.ColumnsDefinition(cols =>
                    {
                        cols.RelativeColumn(3);
                        cols.RelativeColumn(1);
                    });
                    
                    // Header
                    table.Cell().Border(1).BorderColor(BorderColor).Background(LightGray).Padding(8)
                        .Text("Description").FontSize(10).Bold();
                    table.Cell().Border(1).BorderColor(BorderColor).Background(LightGray).Padding(8)
                        .AlignRight().Text("Count").FontSize(10).Bold();
                    
                    // Data rows
                    AddTableRow(table, "Member Users", $"{(_data.UserStats?.TotalUsers ?? 0) - (_data.UserStats?.GuestUsers ?? 0)}");
                    AddTableRow(table, "Guest Users", $"{_data.UserStats?.GuestUsers ?? 0}");
                    AddTableRow(table, "Total of All Users", $"{_data.UserStats?.TotalUsers ?? 0}");
                    AddTableRow(table, "Licensed Users", $"{(_data.UserStats?.TotalUsers ?? 0) - (_data.UserStats?.GuestUsers ?? 0)}");
                    AddTableRow(table, "MFA Registered", $"{_data.UserStats?.MfaRegistered ?? 0}");
                    AddTableRow(table, "MFA Not Registered", $"{_data.UserStats?.MfaNotRegistered ?? 0}");
                });
                
                row.ConstantItem(20);
                
                // Right table - Role counts (placeholder)
                row.RelativeItem().Table(table =>
                {
                    table.ColumnsDefinition(cols =>
                    {
                        cols.RelativeColumn(3);
                        cols.RelativeColumn(1);
                    });
                    
                    // Header
                    table.Cell().Border(1).BorderColor(BorderColor).Background(LightGray).Padding(8)
                        .Text("Risky Users").FontSize(10).Bold();
                    table.Cell().Border(1).BorderColor(BorderColor).Background(LightGray).Padding(8)
                        .AlignRight().Text("Count").FontSize(10).Bold();
                    
                    // Data rows
                    AddTableRow(table, "Total Risky Users", $"{_data.RiskyUsersCount}");
                    if (_data.HighRiskUsers?.Any() == true)
                    {
                        foreach (var user in _data.HighRiskUsers.Take(5))
                        {
                            AddTableRow(table, user, "High", CriticalColor);
                        }
                    }
                });
            });
        });
        
        // Footer
        ComposePageFooter(column);
    }
    
    #endregion

    #region Infographic Page
    
    private void ComposeInfoGraphicPage(ColumnDescriptor column, string bigNumber, string line1, string line2, string source)
    {
        column.Item().Height(841).Background(_primaryColor).AlignCenter().AlignMiddle().Column(content =>
        {
            // Big statistic number
            content.Item().AlignCenter().Text(bigNumber)
                .FontSize(140)
                .FontColor(Colors.White)
                .FontFamily("Segoe UI Light");
            
            content.Item().Height(20);
            
            // First line
            content.Item().AlignCenter().Text(line1)
                .FontSize(28)
                .FontColor(Colors.White);
            
            content.Item().Height(8);
            
            // Second line (bold, accent color)
            content.Item().AlignCenter().Text(line2)
                .FontSize(28)
                .FontColor(_accentColor)
                .Bold();
            
            content.Item().Height(100);
            
            // Source
            content.Item().AlignCenter().Text(source)
                .FontSize(10)
                .FontColor(Colors.Grey.Lighten2)
                .Italic();
        });
    }
    
    #endregion

    #region Security Metrics Page
    
    private void ComposeSecurityMetricsPage(ColumnDescriptor column)
    {
        column.Item().Padding(40).Column(content =>
        {
            // Microsoft Secure Score
            if (_data.SecureScore != null)
            {
                content.Item().Text("Microsoft Secure Score")
                    .FontSize(20).FontColor(_primaryColor).FontFamily("Segoe UI Light");
                
                content.Item().Height(8);
                content.Item().Text("Microsoft Secure Score is a measurement of an organization's security posture, with a higher number indicating more improvement actions taken.")
                    .FontSize(10).FontColor(Colors.Grey.Darken1);
                
                content.Item().Height(15);
                
                content.Item().Table(table =>
                {
                    table.ColumnsDefinition(cols =>
                    {
                        cols.RelativeColumn(3);
                        cols.RelativeColumn(1);
                    });
                    
                    AddTableRow(table, "Current Score", $"{_data.SecureScore.CurrentScore:N0}");
                    AddTableRow(table, "Maximum Score", $"{_data.SecureScore.MaxScore:N0}");
                    AddTableRow(table, "Percentage", $"{_data.SecureScore.PercentageScore:F1}%", 
                        _data.SecureScore.PercentageScore >= 70 ? CompliantColor : 
                        _data.SecureScore.PercentageScore >= 50 ? WarningColor : CriticalColor);
                });
                
                content.Item().Height(25);
            }
            
            // Intune Managed Devices
            if (_data.DeviceStats != null)
            {
                content.Item().Text("Intune Managed Devices")
                    .FontSize(20).FontColor(_primaryColor).FontFamily("Segoe UI Light");
                
                content.Item().Height(8);
                content.Item().Text("Overview of all devices managed through Microsoft Intune, including compliance status across different platforms.")
                    .FontSize(10).FontColor(Colors.Grey.Darken1);
                
                content.Item().Height(15);
                
                content.Item().Table(table =>
                {
                    table.ColumnsDefinition(cols =>
                    {
                        cols.RelativeColumn(3);
                        cols.RelativeColumn(1);
                    });
                    
                    AddTableRow(table, "Total Devices", $"{_data.DeviceStats.TotalDevices}");
                    AddTableRow(table, "Windows", $"{_data.DeviceStats.WindowsDevices}");
                    AddTableRow(table, "macOS", $"{_data.DeviceStats.MacOsDevices}");
                    AddTableRow(table, "iOS/iPadOS", $"{_data.DeviceStats.IosDevices}");
                    AddTableRow(table, "Android", $"{_data.DeviceStats.AndroidDevices}");
                    AddTableRow(table, "Compliant", $"{_data.DeviceStats.CompliantDevices}", CompliantColor);
                    AddTableRow(table, "Non-Compliant", $"{_data.DeviceStats.NonCompliantDevices}", 
                        _data.DeviceStats.NonCompliantDevices > 0 ? CriticalColor : null);
                    AddTableRow(table, "Compliance Rate", $"{_data.DeviceStats.ComplianceRate}%",
                        _data.DeviceStats.ComplianceRate >= 90 ? CompliantColor : WarningColor);
                });
                
                content.Item().Height(25);
            }
            
            // Defender Stats
            if (_data.DefenderStats != null)
            {
                content.Item().Text("Microsoft Defender for Endpoint")
                    .FontSize(20).FontColor(_primaryColor).FontFamily("Segoe UI Light");
                
                content.Item().Height(15);
                
                content.Item().Table(table =>
                {
                    table.ColumnsDefinition(cols =>
                    {
                        cols.RelativeColumn(3);
                        cols.RelativeColumn(1);
                    });
                    
                    AddTableRow(table, "Exposure Score", _data.DefenderStats.ExposureScore ?? "N/A");
                    AddTableRow(table, "Onboarded Machines", $"{_data.DefenderStats.OnboardedMachines ?? 0}");
                    AddTableRow(table, "Total Vulnerabilities", $"{_data.DefenderStats.VulnerabilitiesDetected}");
                    AddTableRow(table, "Critical", $"{_data.DefenderStats.CriticalVulnerabilities}",
                        _data.DefenderStats.CriticalVulnerabilities > 0 ? CriticalColor : null);
                    AddTableRow(table, "High", $"{_data.DefenderStats.HighVulnerabilities}",
                        _data.DefenderStats.HighVulnerabilities > 0 ? CriticalColor : null);
                    AddTableRow(table, "Medium", $"{_data.DefenderStats.MediumVulnerabilities}",
                        _data.DefenderStats.MediumVulnerabilities > 0 ? WarningColor : null);
                    AddTableRow(table, "Low", $"{_data.DefenderStats.LowVulnerabilities}");
                });
            }
        });
        
        ComposePageFooter(column);
    }
    
    #endregion

    #region Device Details Page
    
    private void ComposeDeviceDetailsPage(ColumnDescriptor column)
    {
        column.Item().Padding(40).Column(content =>
        {
            // Windows Devices
            if (_data.DeviceDetails?.WindowsDevices?.Any() == true)
            {
                content.Item().Text($"Windows Devices ({_data.DeviceDetails.WindowsDevices.Count})")
                    .FontSize(20).FontColor(_primaryColor).FontFamily("Segoe UI Light");
                
                content.Item().Height(10);
                
                ComposeDeviceTable(content, _data.DeviceDetails.WindowsDevices.Take(15).Select(d => new DeviceRow
                {
                    Name = d.DeviceName ?? "-",
                    OsVersion = d.OsVersion ?? "-",
                    UpdateStatus = d.OsVersionStatusMessage ?? d.OsVersionStatus.ToString(),
                    UpdateStatusColor = GetStatusColor(d.OsVersionStatus.ToString()),
                    Compliance = d.ComplianceState ?? "-",
                    LastCheckIn = d.LastCheckIn?.ToString("dd MMM yyyy") ?? "Never"
                }).ToList());
                
                content.Item().Height(20);
            }
            
            // macOS Devices
            if (_data.DeviceDetails?.MacDevices?.Any() == true)
            {
                content.Item().Text($"macOS Devices ({_data.DeviceDetails.MacDevices.Count})")
                    .FontSize(20).FontColor(_primaryColor).FontFamily("Segoe UI Light");
                
                content.Item().Height(10);
                
                ComposeDeviceTable(content, _data.DeviceDetails.MacDevices.Take(10).Select(d => new DeviceRow
                {
                    Name = d.DeviceName ?? "-",
                    OsVersion = d.OsVersion ?? "-",
                    UpdateStatus = d.OsVersionStatusMessage ?? d.OsVersionStatus.ToString(),
                    UpdateStatusColor = GetStatusColor(d.OsVersionStatus.ToString()),
                    Compliance = d.ComplianceState ?? "-",
                    LastCheckIn = d.LastCheckIn?.ToString("dd MMM yyyy") ?? "Never"
                }).ToList());
                
                content.Item().Height(20);
            }
            
            // iOS Devices
            if (_data.DeviceDetails?.IosDevices?.Any() == true)
            {
                content.Item().Text($"iOS/iPadOS Devices ({_data.DeviceDetails.IosDevices.Count})")
                    .FontSize(20).FontColor(_primaryColor).FontFamily("Segoe UI Light");
                
                content.Item().Height(10);
                
                ComposeDeviceTable(content, _data.DeviceDetails.IosDevices.Take(10).Select(d => new DeviceRow
                {
                    Name = d.DeviceName ?? "-",
                    OsVersion = d.OsVersion ?? "-",
                    UpdateStatus = d.OsVersionStatusMessage ?? d.OsVersionStatus.ToString(),
                    UpdateStatusColor = GetStatusColor(d.OsVersionStatus.ToString()),
                    Compliance = d.ComplianceState ?? "-",
                    LastCheckIn = d.LastCheckIn?.ToString("dd MMM yyyy") ?? "Never"
                }).ToList());
            }
        });
        
        ComposePageFooter(column);
    }
    
    private void ComposeDeviceTable(ColumnDescriptor content, List<DeviceRow> devices)
    {
        content.Item().Table(table =>
        {
            table.ColumnsDefinition(cols =>
            {
                cols.RelativeColumn(2.5f);  // Device Name
                cols.RelativeColumn(2);     // OS Version
                cols.RelativeColumn(2);     // Update Status
                cols.RelativeColumn(1.5f);  // Compliance
                cols.RelativeColumn(1.5f);  // Last Check-in
            });
            
            // Header
            table.Cell().Border(1).BorderColor(BorderColor).Background(LightGray).Padding(6)
                .Text("Device Name").FontSize(9).Bold();
            table.Cell().Border(1).BorderColor(BorderColor).Background(LightGray).Padding(6)
                .Text("OS Version").FontSize(9).Bold();
            table.Cell().Border(1).BorderColor(BorderColor).Background(LightGray).Padding(6)
                .Text("Update Status").FontSize(9).Bold();
            table.Cell().Border(1).BorderColor(BorderColor).Background(LightGray).Padding(6)
                .Text("Compliance").FontSize(9).Bold();
            table.Cell().Border(1).BorderColor(BorderColor).Background(LightGray).Padding(6)
                .Text("Last Check-in").FontSize(9).Bold();
            
            foreach (var device in devices)
            {
                table.Cell().Border(1).BorderColor(BorderColor).Padding(6)
                    .Text(device.Name).FontSize(9);
                table.Cell().Border(1).BorderColor(BorderColor).Padding(6)
                    .Text(device.OsVersion).FontSize(8);
                table.Cell().Border(1).BorderColor(BorderColor).Padding(6)
                    .Text(device.UpdateStatus).FontSize(9).FontColor(device.UpdateStatusColor);
                table.Cell().Border(1).BorderColor(BorderColor).Padding(6)
                    .Text(device.Compliance).FontSize(9)
                    .FontColor(device.Compliance.ToLower() == "compliant" ? CompliantColor : NonCompliantColor);
                table.Cell().Border(1).BorderColor(BorderColor).Padding(6)
                    .Text(device.LastCheckIn).FontSize(9);
            }
        });
    }
    
    private class DeviceRow
    {
        public string Name { get; set; } = "";
        public string OsVersion { get; set; } = "";
        public string UpdateStatus { get; set; } = "";
        public string UpdateStatusColor { get; set; } = TextColor;
        public string Compliance { get; set; } = "";
        public string LastCheckIn { get; set; } = "";
    }
    
    #endregion

    #region User Details Page
    
    private void ComposeUserDetailsPage(ColumnDescriptor column)
    {
        column.Item().Padding(40).Column(content =>
        {
            content.Item().Text($"User Sign-in & MFA Details ({_data.UserSignInDetails!.Count} users)")
                .FontSize(20).FontColor(_primaryColor).FontFamily("Segoe UI Light");
            
            content.Item().Height(8);
            content.Item().Text("Detailed view of user sign-in activity and multi-factor authentication registration status.")
                .FontSize(10).FontColor(Colors.Grey.Darken1);
            
            content.Item().Height(15);
            
            content.Item().Table(table =>
            {
                table.ColumnsDefinition(cols =>
                {
                    cols.RelativeColumn(2);    // Display Name
                    cols.RelativeColumn(2.5f); // Email
                    cols.RelativeColumn(1.5f); // Last Sign-in
                    cols.RelativeColumn(2);    // MFA Method
                    cols.RelativeColumn(0.8f); // MFA
                    cols.RelativeColumn(0.8f); // Enabled
                });
                
                // Header
                table.Cell().Border(1).BorderColor(BorderColor).Background(LightGray).Padding(6)
                    .Text("Display Name").FontSize(9).Bold();
                table.Cell().Border(1).BorderColor(BorderColor).Background(LightGray).Padding(6)
                    .Text("Email").FontSize(9).Bold();
                table.Cell().Border(1).BorderColor(BorderColor).Background(LightGray).Padding(6)
                    .Text("Last Sign-in").FontSize(9).Bold();
                table.Cell().Border(1).BorderColor(BorderColor).Background(LightGray).Padding(6)
                    .Text("Default MFA Method").FontSize(9).Bold();
                table.Cell().Border(1).BorderColor(BorderColor).Background(LightGray).Padding(6)
                    .Text("MFA").FontSize(9).Bold();
                table.Cell().Border(1).BorderColor(BorderColor).Background(LightGray).Padding(6)
                    .Text("Enabled").FontSize(9).Bold();
                
                foreach (var user in _data.UserSignInDetails.Take(30))
                {
                    table.Cell().Border(1).BorderColor(BorderColor).Padding(6)
                        .Text(user.DisplayName ?? "-").FontSize(9);
                    table.Cell().Border(1).BorderColor(BorderColor).Padding(6)
                        .Text(user.UserPrincipalName ?? "-").FontSize(8);
                    table.Cell().Border(1).BorderColor(BorderColor).Padding(6)
                        .Text(user.LastInteractiveSignIn?.ToString("dd MMM yyyy") ?? "Never").FontSize(9);
                    table.Cell().Border(1).BorderColor(BorderColor).Padding(6)
                        .Text(user.DefaultMfaMethod ?? "None").FontSize(9);
                    table.Cell().Border(1).BorderColor(BorderColor).Padding(6)
                        .Text(user.IsMfaRegistered ? "Yes" : "No").FontSize(9)
                        .FontColor(user.IsMfaRegistered ? CompliantColor : CriticalColor);
                    table.Cell().Border(1).BorderColor(BorderColor).Padding(6)
                        .Text(user.AccountEnabled ? "Yes" : "No").FontSize(9)
                        .FontColor(user.AccountEnabled ? CompliantColor : NonCompliantColor);
                }
            });
        });
        
        ComposePageFooter(column);
    }
    
    #endregion

    #region Domain Security Page
    
    private void ComposeDomainSecurityPage(ColumnDescriptor column)
    {
        column.Item().Padding(40).Column(content =>
        {
            content.Item().Text("Domain Email Security")
                .FontSize(20).FontColor(_primaryColor).FontFamily("Segoe UI Light");
            
            content.Item().Height(8);
            content.Item().Text("Email authentication protocols (SPF, DMARC, DKIM) help protect your organization from email spoofing and phishing attacks.")
                .FontSize(10).FontColor(Colors.Grey.Darken1);
            
            content.Item().Height(15);
            
            // Summary table
            content.Item().Table(table =>
            {
                table.ColumnsDefinition(cols =>
                {
                    cols.RelativeColumn(3);
                    cols.RelativeColumn(1);
                });
                
                AddTableRow(table, "Total Domains Checked", $"{_data.DomainSecuritySummary!.TotalDomains}");
                AddTableRow(table, "Domains with MX Records", $"{_data.DomainSecuritySummary.DomainsWithMx}");
                AddTableRow(table, "Domains with SPF", $"{_data.DomainSecuritySummary.DomainsWithSpf}", CompliantColor);
                AddTableRow(table, "Domains with DMARC", $"{_data.DomainSecuritySummary.DomainsWithDmarc}", CompliantColor);
                AddTableRow(table, "Domains with DKIM", $"{_data.DomainSecuritySummary.DomainsWithDkim}", CompliantColor);
            });
            
            if (_data.DomainSecurityResults?.Any() == true)
            {
                content.Item().Height(25);
                
                content.Item().Text("Domain Security Details")
                    .FontSize(20).FontColor(_primaryColor).FontFamily("Segoe UI Light");
                
                content.Item().Height(15);
                
                content.Item().Table(table =>
                {
                    table.ColumnsDefinition(cols =>
                    {
                        cols.RelativeColumn(3);   // Domain
                        cols.RelativeColumn(0.8f); // MX
                        cols.RelativeColumn(0.8f); // SPF
                        cols.RelativeColumn(1);    // DMARC
                        cols.RelativeColumn(0.8f); // DKIM
                        cols.RelativeColumn(0.8f); // Score
                        cols.RelativeColumn(0.8f); // Grade
                    });
                    
                    // Header
                    table.Cell().Border(1).BorderColor(BorderColor).Background(LightGray).Padding(6)
                        .Text("Domain").FontSize(9).Bold();
                    table.Cell().Border(1).BorderColor(BorderColor).Background(LightGray).Padding(6)
                        .Text("MX").FontSize(9).Bold();
                    table.Cell().Border(1).BorderColor(BorderColor).Background(LightGray).Padding(6)
                        .Text("SPF").FontSize(9).Bold();
                    table.Cell().Border(1).BorderColor(BorderColor).Background(LightGray).Padding(6)
                        .Text("DMARC").FontSize(9).Bold();
                    table.Cell().Border(1).BorderColor(BorderColor).Background(LightGray).Padding(6)
                        .Text("DKIM").FontSize(9).Bold();
                    table.Cell().Border(1).BorderColor(BorderColor).Background(LightGray).Padding(6)
                        .Text("Score").FontSize(9).Bold();
                    table.Cell().Border(1).BorderColor(BorderColor).Background(LightGray).Padding(6)
                        .Text("Grade").FontSize(9).Bold();
                    
                    foreach (var domain in _data.DomainSecurityResults.OrderByDescending(d => d.SecurityScore).Take(15))
                    {
                        table.Cell().Border(1).BorderColor(BorderColor).Padding(6)
                            .Text(domain.Domain).FontSize(9);
                        table.Cell().Border(1).BorderColor(BorderColor).Padding(6)
                            .Text(domain.HasMx ? "✓" : "✗").FontSize(9)
                            .FontColor(domain.HasMx ? CompliantColor : CriticalColor);
                        table.Cell().Border(1).BorderColor(BorderColor).Padding(6)
                            .Text(domain.HasSpf ? "✓" : "✗").FontSize(9)
                            .FontColor(domain.HasSpf ? CompliantColor : CriticalColor);
                        table.Cell().Border(1).BorderColor(BorderColor).Padding(6)
                            .Text(domain.HasDmarc ? domain.DmarcPolicy ?? "✓" : "✗").FontSize(9)
                            .FontColor(domain.HasDmarc ? CompliantColor : CriticalColor);
                        table.Cell().Border(1).BorderColor(BorderColor).Padding(6)
                            .Text(domain.HasDkim ? "✓" : "✗").FontSize(9)
                            .FontColor(domain.HasDkim ? CompliantColor : CriticalColor);
                        table.Cell().Border(1).BorderColor(BorderColor).Padding(6)
                            .Text($"{domain.SecurityScore}").FontSize(9);
                        table.Cell().Border(1).BorderColor(BorderColor).Padding(6)
                            .Text(domain.SecurityGrade).FontSize(9).Bold()
                            .FontColor(domain.SecurityGrade switch
                            {
                                "A" or "B" => CompliantColor,
                                "C" => WarningColor,
                                _ => CriticalColor
                            });
                    }
                });
            }
        });
        
        ComposePageFooter(column);
    }
    
    #endregion

    #region Footer
    
    private void ComposePageFooter(ColumnDescriptor column)
    {
        column.Item().ExtendVertical().AlignBottom().Padding(40).Column(footer =>
        {
            // Horizontal line
            footer.Item().Height(1).Background(_primaryColor);
            
            footer.Item().Height(15);
            
            footer.Item().Row(row =>
            {
                // Footer text
                if (!string.IsNullOrWhiteSpace(_settings.FooterText))
                {
                    row.RelativeItem().Text(_settings.FooterText)
                        .FontSize(8).FontColor(Colors.Grey.Darken1).Italic();
                }
                else
                {
                    row.RelativeItem().Text("Generated by M365 Dashboard")
                        .FontSize(8).FontColor(Colors.Grey.Darken1).Italic();
                }
                
                // Logo on right
                if (_logoBytes != null)
                {
                    row.ConstantItem(80).AlignRight().Height(25).Image(_logoBytes).FitHeight();
                }
            });
        });
    }
    
    #endregion

    #region Helpers
    
    private void AddTableRow(TableDescriptor table, string label, string value, string? valueColor = null)
    {
        table.Cell().Border(1).BorderColor(BorderColor).Padding(8)
            .Text(label).FontSize(10);
        
        var valueCell = table.Cell().Border(1).BorderColor(BorderColor).Padding(8).AlignRight();
        if (valueColor != null)
            valueCell.Text(value).FontSize(10).FontColor(valueColor);
        else
            valueCell.Text(value).FontSize(10);
    }
    
    private (string line1, string line2) ParseReportTitle(string title)
    {
        if (string.IsNullOrWhiteSpace(title))
            return ("MICROSOFT 365", "SECURITY ASSESSMENT");
        
        var lower = title.ToLower();
        if (lower.StartsWith("microsoft 365"))
            return ("MICROSOFT 365", title.Substring(13).Trim().ToUpper());
        if (lower.StartsWith("m365"))
            return ("MICROSOFT 365", title.Substring(4).Trim().ToUpper());
        
        return ("MICROSOFT 365", title.ToUpper());
    }
    
    private static List<ReportQuote> PickRandomQuotes(ReportSettings settings, int count)
    {
        var pool = settings.Quotes
            .Where(q => q.Enabled &&
                        !string.IsNullOrWhiteSpace(q.BigNumber) &&
                        !string.IsNullOrWhiteSpace(q.Line1))
            .ToList();
        if (pool.Count == 0) pool = ReportSettings.DefaultQuotes();
        if (pool.Count <= count) return pool;
        var rng = new Random();
        for (int i = pool.Count - 1; i > 0; i--)
        {
            int j = rng.Next(i + 1);
            (pool[i], pool[j]) = (pool[j], pool[i]);
        }
        return pool.Take(count).ToList();
    }

    private bool HasDevices()
    {
        return _data.DeviceDetails?.WindowsDevices?.Any() == true ||
               _data.DeviceDetails?.MacDevices?.Any() == true ||
               _data.DeviceDetails?.IosDevices?.Any() == true ||
               _data.DeviceDetails?.AndroidDevices?.Any() == true;
    }
    
    private string GetStatusColor(string? status)
    {
        return status?.ToLower() switch
        {
            "current" => CompliantColor,
            "warning" => WarningColor,
            "critical" => CriticalColor,
            _ => TextColor
        };
    }
    
    #endregion
}
