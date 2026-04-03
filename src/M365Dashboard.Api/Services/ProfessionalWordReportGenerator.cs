using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using M365Dashboard.Api.Models;
using M365Dashboard.Api.Controllers;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace M365Dashboard.Api.Services;

/// <summary>
/// Professional Word document generator for Executive Reports
/// Styled to match the Cloud1st PDF example
/// </summary>
public class ProfessionalWordReportGenerator
{
    // Branding colors (matching PDF example)
    private string _primaryColor = "1E3A5F";  // Dark blue
    private string _accentColor = "E07C3A";   // Orange
    private string _compliantColor = "107C6C"; // Teal
    private string _nonCompliantColor = "DC2626"; // Red
    private string _warningColor = "F59E0B"; // Amber
    private string _headerBgColor = "F9FAFB"; // Light gray for table headers
    private string _altRowColor = "F3F4F6";   // Alternating row color
    private string _borderColor = "E5E7EB";   // Table border color
    private string _textColor = "374151";     // Body text
    
    private ReportSettings _settings = new();
    private uint _imageCounter = 1;
    
    // Page dimensions (A4)
    private const int PAGE_WIDTH = 11906;  // A4 width in twentieths of a point
    private const int PAGE_HEIGHT = 16838; // A4 height
    private const int MARGIN = 720;        // 0.5 inch margins
    private const int CONTENT_WIDTH = 10466; // PAGE_WIDTH - 2*MARGIN

    public byte[] GenerateReport(ExecutiveReportData data, ReportSettings settings)
    {
        _settings = settings;
        if (!string.IsNullOrEmpty(settings.PrimaryColor))
            _primaryColor = settings.PrimaryColor.TrimStart('#');
        if (!string.IsNullOrEmpty(settings.AccentColor))
            _accentColor = settings.AccentColor.TrimStart('#');
        
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document();
            var body = mainPart.Document.AppendChild(new Body());
            
            // Add styles
            AddStyles(mainPart);
            
            // === COVER PAGE ===
            AddCoverPage(body, mainPart, data);
            body.Append(new Paragraph(new Run(new Break() { Type = BreakValues.Page })));
            
            // === EXECUTIVE SUMMARY ===
            AddExecutiveSummary(body, data);
            body.Append(new Paragraph(new Run(new Break() { Type = BreakValues.Page })));
            
            // === SECURITY METRICS ===
            AddSecurityMetricsPage(body, data);
            
            // === DEVICE DETAILS ===
            if (data.DeviceDetails != null)
            {
                body.Append(new Paragraph(new Run(new Break() { Type = BreakValues.Page })));
                AddDeviceDetailsPage(body, data);
            }
            
            // === USER DETAILS ===
            if (data.UserSignInDetails?.Any() == true)
            {
                body.Append(new Paragraph(new Run(new Break() { Type = BreakValues.Page })));
                AddUserDetailsPage(body, data);
            }
            
            // === DOMAIN SECURITY ===
            if (data.DomainSecuritySummary != null)
            {
                body.Append(new Paragraph(new Run(new Break() { Type = BreakValues.Page })));
                AddDomainSecurityPage(body, data);
            }
            
            // === FOOTER ===
            AddFooter(body);
            
            // Set section properties (margins, etc.)
            body.Append(new SectionProperties(
                new PageMargin() { Top = MARGIN, Bottom = MARGIN, Left = (uint)MARGIN, Right = (uint)MARGIN }
            ));
            
            mainPart.Document.Save();
        }
        
        return stream.ToArray();
    }

    #region Styles
    
    private void AddStyles(MainDocumentPart mainPart)
    {
        var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
        var styles = new Styles();
        
        // Normal style
        var normal = new Style() { Type = StyleValues.Paragraph, StyleId = "Normal", Default = true };
        normal.Append(new StyleName() { Val = "Normal" });
        normal.Append(new StyleRunProperties(
            new RunFonts() { Ascii = "Segoe UI", HighAnsi = "Segoe UI" },
            new FontSize() { Val = "22" },
            new Color() { Val = _textColor }
        ));
        styles.Append(normal);
        
        // Heading 1
        var h1 = new Style() { Type = StyleValues.Paragraph, StyleId = "Heading1" };
        h1.Append(new StyleName() { Val = "Heading 1" });
        h1.Append(new BasedOn() { Val = "Normal" });
        h1.Append(new StyleRunProperties(
            new RunFonts() { Ascii = "Segoe UI Light", HighAnsi = "Segoe UI Light" },
            new FontSize() { Val = "52" },
            new Color() { Val = _primaryColor }
        ));
        h1.Append(new StyleParagraphProperties(
            new SpacingBetweenLines() { Before = "400", After = "200" }
        ));
        styles.Append(h1);
        
        // Heading 2
        var h2 = new Style() { Type = StyleValues.Paragraph, StyleId = "Heading2" };
        h2.Append(new StyleName() { Val = "Heading 2" });
        h2.Append(new BasedOn() { Val = "Normal" });
        h2.Append(new StyleRunProperties(
            new RunFonts() { Ascii = "Segoe UI", HighAnsi = "Segoe UI" },
            new Bold(),
            new FontSize() { Val = "28" },
            new Color() { Val = _primaryColor }
        ));
        h2.Append(new StyleParagraphProperties(
            new SpacingBetweenLines() { Before = "300", After = "120" }
        ));
        styles.Append(h2);
        
        stylesPart.Styles = styles;
    }
    
    #endregion

    #region Cover Page
    
    private void AddCoverPage(Body body, MainDocumentPart mainPart, ExecutiveReportData data)
    {
        // Dark blue header section using a table for full-width background
        var coverTable = new Table();
        coverTable.Append(new TableProperties(
            new TableWidth() { Width = "5000", Type = TableWidthUnitValues.Pct },
            new TableBorders(
                new TopBorder() { Val = BorderValues.None },
                new BottomBorder() { Val = BorderValues.None },
                new LeftBorder() { Val = BorderValues.None },
                new RightBorder() { Val = BorderValues.None }
            )
        ));
        
        var coverRow = new TableRow();
        coverRow.Append(new TableRowProperties(new TableRowHeight() { Val = 7000, HeightType = HeightRuleValues.Exact }));
        
        var coverCell = new TableCell();
        coverCell.Append(new TableCellProperties(
            new TableCellWidth() { Width = "5000", Type = TableWidthUnitValues.Pct },
            new Shading() { Fill = _primaryColor, Val = ShadingPatternValues.Clear },
            new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Bottom },
            new TableCellMargin(
                new LeftMargin() { Width = "400", Type = TableWidthUnitValues.Dxa },
                new BottomMargin() { Width = "400", Type = TableWidthUnitValues.Dxa }
            )
        ));
        
        // Parse report title
        var titleParts = ParseReportTitle(_settings.ReportTitle);
        
        // "MICROSOFT 365" in orange
        var subtitlePara = new Paragraph();
        subtitlePara.Append(new ParagraphProperties(
            new SpacingBetweenLines() { Before = "0", After = "100" }
        ));
        var subtitleRun = new Run();
        subtitleRun.Append(new RunProperties(
            new RunFonts() { Ascii = "Segoe UI", HighAnsi = "Segoe UI" },
            new FontSize() { Val = "32" },
            new Color() { Val = _accentColor },
            new Caps()
        ));
        subtitleRun.Append(new Text(titleParts.line1));
        subtitlePara.Append(subtitleRun);
        coverCell.Append(subtitlePara);
        
        // Main title in white
        var titlePara = new Paragraph();
        titlePara.Append(new ParagraphProperties(
            new SpacingBetweenLines() { Before = "0", After = "0" }
        ));
        var titleRun = new Run();
        titleRun.Append(new RunProperties(
            new RunFonts() { Ascii = "Segoe UI Light", HighAnsi = "Segoe UI Light" },
            new FontSize() { Val = "72" },
            new Color() { Val = "FFFFFF" }
        ));
        titleRun.Append(new Text(titleParts.line2.ToUpper()));
        titlePara.Append(titleRun);
        coverCell.Append(titlePara);
        
        coverRow.Append(coverCell);
        coverTable.Append(coverRow);
        body.Append(coverTable);
        
        // White section with company info
        AddEmptyParagraphs(body, 3);
        
        // Company name
        var companyPara = new Paragraph();
        companyPara.Append(new ParagraphProperties(
            new Justification() { Val = JustificationValues.Center },
            new SpacingBetweenLines() { Before = "400", After = "100" }
        ));
        var companyRun = new Run();
        companyRun.Append(new RunProperties(
            new RunFonts() { Ascii = "Segoe UI", HighAnsi = "Segoe UI" },
            new Bold(),
            new FontSize() { Val = "36" },
            new Color() { Val = _primaryColor }
        ));
        companyRun.Append(new Text(_settings.CompanyName));
        companyPara.Append(companyRun);
        body.Append(companyPara);
        
        // Date
        var datePara = new Paragraph();
        datePara.Append(new ParagraphProperties(
            new Justification() { Val = JustificationValues.Center },
            new SpacingBetweenLines() { Before = "0", After = "200" }
        ));
        var dateRun = new Run();
        dateRun.Append(new RunProperties(
            new RunFonts() { Ascii = "Segoe UI", HighAnsi = "Segoe UI" },
            new FontSize() { Val = "22" },
            new Color() { Val = "666666" }
        ));
        dateRun.Append(new Text($"Generated on {data.GeneratedAt:d MMMM yyyy}"));
        datePara.Append(dateRun);
        body.Append(datePara);
        
        // Logo
        if (!string.IsNullOrEmpty(_settings.LogoBase64))
        {
            AddCenteredLogo(body, mainPart);
        }
    }
    
    private void AddCenteredLogo(Body body, MainDocumentPart mainPart)
    {
        try
        {
            var logoBytes = Convert.FromBase64String(_settings.LogoBase64!);
            var imagePart = mainPart.AddImagePart(
                _settings.LogoContentType?.Contains("png") == true ? ImagePartType.Png : ImagePartType.Jpeg);
            using var imgStream = new MemoryStream(logoBytes);
            imagePart.FeedData(imgStream);
            
            var relationshipId = mainPart.GetIdOfPart(imagePart);
            const long cx = 1800000L; // ~2 inches
            const long cy = 600000L;
            
            var logoPara = new Paragraph();
            logoPara.Append(new ParagraphProperties(
                new Justification() { Val = JustificationValues.Center },
                new SpacingBetweenLines() { Before = "200" }
            ));
            
            logoPara.Append(new Run(new Drawing(
                new DW.Inline(
                    new DW.Extent() { Cx = cx, Cy = cy },
                    new DW.EffectExtent(),
                    new DW.DocProperties() { Id = _imageCounter++, Name = "Logo" },
                    new DW.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks() { NoChangeAspect = true }),
                    new A.Graphic(new A.GraphicData(
                        new PIC.Picture(
                            new PIC.NonVisualPictureProperties(
                                new PIC.NonVisualDrawingProperties() { Id = 0U, Name = "logo" },
                                new PIC.NonVisualPictureDrawingProperties()),
                            new PIC.BlipFill(
                                new A.Blip() { Embed = relationshipId },
                                new A.Stretch(new A.FillRectangle())),
                            new PIC.ShapeProperties(
                                new A.Transform2D(
                                    new A.Offset() { X = 0, Y = 0 },
                                    new A.Extents() { Cx = cx, Cy = cy }),
                                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle })))
                    { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })))));
            
            body.Append(logoPara);
        }
        catch { }
    }
    
    #endregion

    #region Executive Summary
    
    private void AddExecutiveSummary(Body body, ExecutiveReportData data)
    {
        // Section heading
        AddHeading1(body, "Executive Summary");
        
        // KPI Cards using a table
        var kpiTable = new Table();
        kpiTable.Append(new TableProperties(
            new TableWidth() { Width = "5000", Type = TableWidthUnitValues.Pct },
            new TableBorders(
                new TopBorder() { Val = BorderValues.None },
                new BottomBorder() { Val = BorderValues.None },
                new LeftBorder() { Val = BorderValues.None },
                new RightBorder() { Val = BorderValues.None },
                new InsideHorizontalBorder() { Val = BorderValues.None },
                new InsideVerticalBorder() { Val = BorderValues.None }
            ),
            new TableCellSpacing() { Width = "100", Type = TableWidthUnitValues.Dxa }
        ));
        
        var kpiRow = new TableRow();
        kpiRow.Append(new TableRowProperties(new TableRowHeight() { Val = 1400 }));
        
        // KPI 1: Total Users
        kpiRow.Append(CreateKpiCell(
            $"{data.UserStats?.TotalUsers ?? 0}",
            "Total Users",
            $"Including {data.UserStats?.GuestUsers ?? 0} guest users"
        ));
        
        // KPI 2: Licensed Users
        kpiRow.Append(CreateKpiCell(
            $"{(data.UserStats?.TotalUsers ?? 0) - (data.UserStats?.GuestUsers ?? 0)}",
            "Licensed Users",
            $"{data.UserStats?.GuestUsers ?? 0} unlicensed users"
        ));
        
        // KPI 3: MFA Registered
        kpiRow.Append(CreateKpiCell(
            $"{data.UserStats?.MfaRegistered ?? 0}",
            "MFA Registered",
            $"{data.UserStats?.MfaNotRegistered ?? 0} not registered"
        ));
        
        kpiTable.Append(kpiRow);
        body.Append(kpiTable);
        
        AddEmptyParagraphs(body, 1);
        
        // Introduction text
        AddBodyText(body, $"This report was prepared for {_settings.CompanyName} in {data.GeneratedAt:MMMM yyyy}. This {_settings.ReportTitle} provides a comprehensive analysis of your organization's security configuration across key Microsoft 365 services, including Entra ID (Azure AD), Exchange Online, Intune, SharePoint, and Teams.");
        
        AddBodyText(body, "The aim of this review is to provide a clear and actionable understanding of your current security posture within Microsoft 365, helping to mitigate potential risks, safeguard sensitive data, and ensure compliance with leading security benchmarks.");
    }
    
    private TableCell CreateKpiCell(string value, string label, string sublabel)
    {
        var cell = new TableCell();
        cell.Append(new TableCellProperties(
            new TableCellWidth() { Width = "0", Type = TableWidthUnitValues.Auto },
            new Shading() { Fill = _altRowColor, Val = ShadingPatternValues.Clear },
            new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center },
            new TableCellMargin(
                new LeftMargin() { Width = "200", Type = TableWidthUnitValues.Dxa },
                new RightMargin() { Width = "200", Type = TableWidthUnitValues.Dxa },
                new TopMargin() { Width = "200", Type = TableWidthUnitValues.Dxa },
                new BottomMargin() { Width = "200", Type = TableWidthUnitValues.Dxa }
            )
        ));
        
        // Value
        var valuePara = new Paragraph();
        valuePara.Append(new ParagraphProperties(
            new Justification() { Val = JustificationValues.Center },
            new SpacingBetweenLines() { Before = "0", After = "40" }
        ));
        var valueRun = new Run();
        valueRun.Append(new RunProperties(
            new RunFonts() { Ascii = "Segoe UI Light", HighAnsi = "Segoe UI Light" },
            new FontSize() { Val = "56" },
            new Color() { Val = _primaryColor }
        ));
        valueRun.Append(new Text(value));
        valuePara.Append(valueRun);
        cell.Append(valuePara);
        
        // Label
        var labelPara = new Paragraph();
        labelPara.Append(new ParagraphProperties(
            new Justification() { Val = JustificationValues.Center },
            new SpacingBetweenLines() { Before = "0", After = "20" }
        ));
        var labelRun = new Run();
        labelRun.Append(new RunProperties(
            new RunFonts() { Ascii = "Segoe UI", HighAnsi = "Segoe UI" },
            new Bold(),
            new FontSize() { Val = "22" },
            new Color() { Val = _textColor }
        ));
        labelRun.Append(new Text(label));
        labelPara.Append(labelRun);
        cell.Append(labelPara);
        
        // Sublabel
        var sublabelPara = new Paragraph();
        sublabelPara.Append(new ParagraphProperties(
            new Justification() { Val = JustificationValues.Center },
            new SpacingBetweenLines() { Before = "0", After = "0" }
        ));
        var sublabelRun = new Run();
        sublabelRun.Append(new RunProperties(
            new RunFonts() { Ascii = "Segoe UI", HighAnsi = "Segoe UI" },
            new FontSize() { Val = "18" },
            new Color() { Val = "666666" }
        ));
        sublabelRun.Append(new Text(sublabel));
        sublabelPara.Append(sublabelRun);
        cell.Append(sublabelPara);
        
        return cell;
    }
    
    #endregion

    #region Security Metrics
    
    private void AddSecurityMetricsPage(Body body, ExecutiveReportData data)
    {
        // Secure Score
        if (data.SecureScore != null)
        {
            AddHeading2(body, "Microsoft Secure Score");
            AddIntroText(body, "Microsoft Secure Score is a measurement of an organization's security posture, with a higher number indicating more improvement actions taken.");
            
            AddDataTable(body, new[] { "Metric", "Value" }, new List<string[]>
            {
                new[] { "Current Score", $"{data.SecureScore.CurrentScore:N0}" },
                new[] { "Maximum Score", $"{data.SecureScore.MaxScore:N0}" },
                new[] { "Percentage", $"{data.SecureScore.PercentageScore:F1}%" }
            });
        }
        
        // Device Stats
        if (data.DeviceStats != null)
        {
            AddHeading2(body, "Intune Managed Devices");
            AddIntroText(body, "Overview of all devices managed through Microsoft Intune, including compliance status across different platforms.");
            
            AddDataTable(body, new[] { "Platform", "Count" }, new List<string[]>
            {
                new[] { "Total Devices", $"{data.DeviceStats.TotalDevices}" },
                new[] { "Windows", $"{data.DeviceStats.WindowsDevices}" },
                new[] { "macOS", $"{data.DeviceStats.MacOsDevices}" },
                new[] { "iOS/iPadOS", $"{data.DeviceStats.IosDevices}" },
                new[] { "Android", $"{data.DeviceStats.AndroidDevices}" },
                new[] { "Compliant", $"{data.DeviceStats.CompliantDevices}" },
                new[] { "Non-Compliant", $"{data.DeviceStats.NonCompliantDevices}" },
                new[] { "Compliance Rate", $"{data.DeviceStats.ComplianceRate}%" }
            }, valueColorColumn: 1, goodValues: new[] { "Compliant" }, badValues: new[] { "Non-Compliant" });
        }
        
        // Defender Stats
        if (data.DefenderStats != null)
        {
            AddHeading2(body, "Microsoft Defender for Endpoint");
            
            AddDataTable(body, new[] { "Metric", "Value" }, new List<string[]>
            {
                new[] { "Exposure Score", data.DefenderStats.ExposureScore ?? "N/A" },
                new[] { "Onboarded Machines", $"{data.DefenderStats.OnboardedMachines ?? 0}" },
                new[] { "Total Vulnerabilities", $"{data.DefenderStats.VulnerabilitiesDetected}" },
                new[] { "Critical", $"{data.DefenderStats.CriticalVulnerabilities}" },
                new[] { "High", $"{data.DefenderStats.HighVulnerabilities}" },
                new[] { "Medium", $"{data.DefenderStats.MediumVulnerabilities}" },
                new[] { "Low", $"{data.DefenderStats.LowVulnerabilities}" }
            });
        }
        
        // User Stats
        if (data.UserStats != null)
        {
            AddHeading2(body, "User Accounts");
            AddIntroText(body, "Summary of user accounts including MFA registration status and high-risk users identified by Microsoft Entra ID Protection.");
            
            AddDataTable(body, new[] { "Type", "Count" }, new List<string[]>
            {
                new[] { "Total Users", $"{data.UserStats.TotalUsers}" },
                new[] { "Guest Users", $"{data.UserStats.GuestUsers}" },
                new[] { "Deleted Users (Soft)", $"{data.UserStats.DeletedUsers}" },
                new[] { "MFA Registered", $"{data.UserStats.MfaRegistered}" },
                new[] { "MFA Not Registered", $"{data.UserStats.MfaNotRegistered}" },
                new[] { "Risky Users", $"{data.RiskyUsersCount}" }
            });
            
            if (data.HighRiskUsers?.Any() == true)
            {
                AddWarningText(body, $"⚠ High Risk Users: {string.Join(", ", data.HighRiskUsers)}");
            }
        }
    }
    
    #endregion

    #region Device Details
    
    private void AddDeviceDetailsPage(Body body, ExecutiveReportData data)
    {
        // Windows Devices
        if (data.DeviceDetails?.WindowsDevices?.Any() == true)
        {
            AddHeading2(body, $"Windows Devices ({data.DeviceDetails.WindowsDevices.Count})");
            
            var headers = new[] { "Device Name", "OS Version", "Update Status", "Compliance", "Last Check-in" };
            var rows = data.DeviceDetails.WindowsDevices.Take(30).Select(d => new[]
            {
                d.DeviceName ?? "-",
                d.OsVersion ?? "-",
                FormatVersionStatus(d.OsVersionStatus, d.OsVersionStatusMessage),
                d.ComplianceState ?? "-",
                d.LastCheckIn?.ToString("dd MMM yyyy") ?? "Never"
            }).ToList();
            
            AddDeviceTable(body, headers, rows);
        }
        
        // macOS Devices
        if (data.DeviceDetails?.MacDevices?.Any() == true)
        {
            AddHeading2(body, $"macOS Devices ({data.DeviceDetails.MacDevices.Count})");
            
            var headers = new[] { "Device Name", "OS Version", "Update Status", "Compliance", "Last Check-in" };
            var rows = data.DeviceDetails.MacDevices.Take(30).Select(d => new[]
            {
                d.DeviceName ?? "-",
                d.OsVersion ?? "-",
                FormatVersionStatus(d.OsVersionStatus, d.OsVersionStatusMessage),
                d.ComplianceState ?? "-",
                d.LastCheckIn?.ToString("dd MMM yyyy") ?? "Never"
            }).ToList();
            
            AddDeviceTable(body, headers, rows);
        }
        
        // iOS Devices
        if (data.DeviceDetails?.IosDevices?.Any() == true)
        {
            AddHeading2(body, $"iOS/iPadOS Devices ({data.DeviceDetails.IosDevices.Count})");
            
            var headers = new[] { "Device Name", "OS Version", "Update Status", "Compliance", "Last Check-in" };
            var rows = data.DeviceDetails.IosDevices.Take(30).Select(d => new[]
            {
                d.DeviceName ?? "-",
                d.OsVersion ?? "-",
                FormatVersionStatus(d.OsVersionStatus, d.OsVersionStatusMessage),
                d.ComplianceState ?? "-",
                d.LastCheckIn?.ToString("dd MMM yyyy") ?? "Never"
            }).ToList();
            
            AddDeviceTable(body, headers, rows);
        }
        
        // Android Devices
        if (data.DeviceDetails?.AndroidDevices?.Any() == true)
        {
            AddHeading2(body, $"Android Devices ({data.DeviceDetails.AndroidDevices.Count})");
            
            var headers = new[] { "Device Name", "OS Version", "Update Status", "Compliance", "Last Check-in" };
            var rows = data.DeviceDetails.AndroidDevices.Take(30).Select(d => new[]
            {
                d.DeviceName ?? "-",
                d.OsVersion ?? "-",
                FormatVersionStatus(d.OsVersionStatus, d.OsVersionStatusMessage),
                d.ComplianceState ?? "-",
                d.LastCheckIn?.ToString("dd MMM yyyy") ?? "Never"
            }).ToList();
            
            AddDeviceTable(body, headers, rows);
        }
    }
    
    #endregion

    #region User Details
    
    private void AddUserDetailsPage(Body body, ExecutiveReportData data)
    {
        AddHeading2(body, $"User Sign-in & MFA Details ({data.UserSignInDetails!.Count} users)");
        AddIntroText(body, "Detailed view of user sign-in activity and MFA registration status.");
        
        var headers = new[] { "Display Name", "Email", "Last Sign-in", "MFA Method", "MFA", "Enabled" };
        var rows = data.UserSignInDetails.Take(50).Select(u => new[]
        {
            u.DisplayName ?? "-",
            u.UserPrincipalName ?? "-",
            u.LastInteractiveSignIn?.ToString("dd MMM yyyy") ?? "Never",
            u.DefaultMfaMethod ?? "None",
            u.IsMfaRegistered ? "Yes" : "No",
            u.AccountEnabled ? "Yes" : "No"
        }).ToList();
        
        AddUserTable(body, headers, rows);
    }
    
    #endregion

    #region Domain Security
    
    private void AddDomainSecurityPage(Body body, ExecutiveReportData data)
    {
        AddHeading2(body, "Domain Email Security");
        AddIntroText(body, "Email authentication protocols (SPF, DMARC, DKIM) help protect your organization from email spoofing and phishing attacks.");
        
        AddDataTable(body, new[] { "Metric", "Count" }, new List<string[]>
        {
            new[] { "Total Domains Checked", $"{data.DomainSecuritySummary!.TotalDomains}" },
            new[] { "Domains with MX Records", $"{data.DomainSecuritySummary.DomainsWithMx}" },
            new[] { "Domains with SPF", $"{data.DomainSecuritySummary.DomainsWithSpf}" },
            new[] { "Domains with DMARC", $"{data.DomainSecuritySummary.DomainsWithDmarc}" },
            new[] { "Domains with DKIM", $"{data.DomainSecuritySummary.DomainsWithDkim}" }
        });
        
        if (data.DomainSecurityResults?.Any() == true)
        {
            AddHeading2(body, "Domain Security Details");
            
            var headers = new[] { "Domain", "MX", "SPF", "DMARC", "DKIM", "Score", "Grade" };
            var rows = data.DomainSecurityResults.OrderByDescending(d => d.SecurityScore).Select(d => new[]
            {
                d.Domain,
                d.HasMx ? "✓" : "✗",
                d.HasSpf ? "✓" : "✗",
                d.HasDmarc ? d.DmarcPolicy ?? "✓" : "✗",
                d.HasDkim ? "✓" : "✗",
                $"{d.SecurityScore}",
                d.SecurityGrade
            }).ToList();
            
            AddDomainTable(body, headers, rows);
        }
    }
    
    #endregion

    #region Footer
    
    private void AddFooter(Body body)
    {
        AddEmptyParagraphs(body, 2);
        
        // Horizontal rule
        var rulePara = new Paragraph();
        rulePara.Append(new ParagraphProperties(
            new ParagraphBorders(new BottomBorder() { Val = BorderValues.Single, Size = 6, Color = _primaryColor }),
            new SpacingBetweenLines() { Before = "400", After = "200" }
        ));
        rulePara.Append(new Run(new Text(" ") { Space = SpaceProcessingModeValues.Preserve }));
        body.Append(rulePara);
        
        if (!string.IsNullOrWhiteSpace(_settings.FooterText))
        {
            AddNoteText(body, _settings.FooterText);
        }
        
        AddNoteText(body, "This report was automatically generated by M365 Dashboard.");
        AddNoteText(body, "Some metrics may require additional licensing or API permissions.");
    }
    
    #endregion

    #region Table Helpers
    
    private void AddDataTable(Body body, string[] headers, List<string[]> rows, int valueColorColumn = -1, string[]? goodValues = null, string[]? badValues = null)
    {
        var table = CreateTable();
        
        // Header row
        var headerRow = new TableRow();
        foreach (var header in headers)
        {
            headerRow.Append(CreateCell(header, isHeader: true));
        }
        table.Append(headerRow);
        
        // Data rows
        for (int i = 0; i < rows.Count; i++)
        {
            var row = rows[i];
            var tableRow = new TableRow();
            
            for (int j = 0; j < row.Length; j++)
            {
                string? textColor = null;
                if (j == valueColorColumn && goodValues != null && goodValues.Any(g => row[0].Contains(g)))
                    textColor = _compliantColor;
                else if (j == valueColorColumn && badValues != null && badValues.Any(b => row[0].Contains(b)))
                    textColor = _nonCompliantColor;
                    
                tableRow.Append(CreateCell(row[j], isHeader: false, isAlternate: i % 2 == 1, textColor: textColor));
            }
            table.Append(tableRow);
        }
        
        body.Append(table);
        AddEmptyParagraphs(body, 1);
    }
    
    private void AddDeviceTable(Body body, string[] headers, List<string[]> rows)
    {
        var table = CreateTable();
        
        // Header row
        var headerRow = new TableRow();
        foreach (var header in headers)
        {
            headerRow.Append(CreateCell(header, isHeader: true, fontSize: "18"));
        }
        table.Append(headerRow);
        
        // Data rows
        for (int i = 0; i < rows.Count; i++)
        {
            var row = rows[i];
            var tableRow = new TableRow();
            
            for (int j = 0; j < row.Length; j++)
            {
                string? textColor = null;
                
                // Color the Update Status column (index 2)
                if (j == 2)
                {
                    if (row[j].StartsWith("✓")) textColor = _compliantColor;
                    else if (row[j].StartsWith("⚠")) textColor = _warningColor;
                    else if (row[j].StartsWith("❌")) textColor = _nonCompliantColor;
                }
                // Color the Compliance column (index 3)
                else if (j == 3)
                {
                    textColor = row[j].ToLower() == "compliant" ? _compliantColor : _nonCompliantColor;
                }
                
                tableRow.Append(CreateCell(row[j], isHeader: false, isAlternate: i % 2 == 1, textColor: textColor, fontSize: "18"));
            }
            table.Append(tableRow);
        }
        
        body.Append(table);
        AddEmptyParagraphs(body, 1);
    }
    
    private void AddUserTable(Body body, string[] headers, List<string[]> rows)
    {
        var table = CreateTable();
        
        // Header row
        var headerRow = new TableRow();
        foreach (var header in headers)
        {
            headerRow.Append(CreateCell(header, isHeader: true, fontSize: "18"));
        }
        table.Append(headerRow);
        
        // Data rows
        for (int i = 0; i < rows.Count; i++)
        {
            var row = rows[i];
            var tableRow = new TableRow();
            
            for (int j = 0; j < row.Length; j++)
            {
                string? textColor = null;
                
                // Color MFA column (index 4) and Enabled column (index 5)
                if (j == 4 || j == 5)
                {
                    textColor = row[j] == "Yes" ? _compliantColor : _nonCompliantColor;
                }
                
                tableRow.Append(CreateCell(row[j], isHeader: false, isAlternate: i % 2 == 1, textColor: textColor, fontSize: "18"));
            }
            table.Append(tableRow);
        }
        
        body.Append(table);
        AddEmptyParagraphs(body, 1);
    }
    
    private void AddDomainTable(Body body, string[] headers, List<string[]> rows)
    {
        var table = CreateTable();
        
        // Header row
        var headerRow = new TableRow();
        foreach (var header in headers)
        {
            headerRow.Append(CreateCell(header, isHeader: true, fontSize: "18"));
        }
        table.Append(headerRow);
        
        // Data rows
        for (int i = 0; i < rows.Count; i++)
        {
            var row = rows[i];
            var tableRow = new TableRow();
            
            for (int j = 0; j < row.Length; j++)
            {
                string? textColor = null;
                
                // Color check marks
                if (j >= 1 && j <= 4)
                {
                    textColor = row[j].Contains("✓") ? _compliantColor : _nonCompliantColor;
                }
                // Color grade
                else if (j == 6)
                {
                    textColor = row[j] switch
                    {
                        "A" or "B" => _compliantColor,
                        "C" => _warningColor,
                        _ => _nonCompliantColor
                    };
                }
                
                tableRow.Append(CreateCell(row[j], isHeader: false, isAlternate: i % 2 == 1, textColor: textColor, fontSize: "18"));
            }
            table.Append(tableRow);
        }
        
        body.Append(table);
        AddEmptyParagraphs(body, 1);
    }
    
    private Table CreateTable()
    {
        var table = new Table();
        table.Append(new TableProperties(
            new TableWidth() { Width = "5000", Type = TableWidthUnitValues.Pct },
            new TableBorders(
                new TopBorder() { Val = BorderValues.Single, Size = 4, Color = _borderColor },
                new BottomBorder() { Val = BorderValues.Single, Size = 4, Color = _borderColor },
                new LeftBorder() { Val = BorderValues.Single, Size = 4, Color = _borderColor },
                new RightBorder() { Val = BorderValues.Single, Size = 4, Color = _borderColor },
                new InsideHorizontalBorder() { Val = BorderValues.Single, Size = 4, Color = _borderColor },
                new InsideVerticalBorder() { Val = BorderValues.Single, Size = 4, Color = _borderColor }
            )
        ));
        return table;
    }
    
    private TableCell CreateCell(string text, bool isHeader = false, bool isAlternate = false, string? textColor = null, string fontSize = "20")
    {
        var cell = new TableCell();
        
        var cellProps = new TableCellProperties(
            new TableCellMargin(
                new LeftMargin() { Width = "80", Type = TableWidthUnitValues.Dxa },
                new RightMargin() { Width = "80", Type = TableWidthUnitValues.Dxa },
                new TopMargin() { Width = "40", Type = TableWidthUnitValues.Dxa },
                new BottomMargin() { Width = "40", Type = TableWidthUnitValues.Dxa }
            )
        );
        
        if (isHeader)
            cellProps.Append(new Shading() { Fill = _headerBgColor, Val = ShadingPatternValues.Clear });
        else if (isAlternate)
            cellProps.Append(new Shading() { Fill = "FFFFFF", Val = ShadingPatternValues.Clear });
        
        cell.Append(cellProps);
        
        var para = new Paragraph();
        var run = new Run();
        
        var runProps = new RunProperties(
            new RunFonts() { Ascii = "Segoe UI", HighAnsi = "Segoe UI" },
            new FontSize() { Val = fontSize }
        );
        
        if (isHeader)
        {
            runProps.Append(new Bold());
            runProps.Append(new Color() { Val = _textColor });
        }
        else if (textColor != null)
        {
            runProps.Append(new Color() { Val = textColor });
        }
        else
        {
            runProps.Append(new Color() { Val = _textColor });
        }
        
        run.Append(runProps);
        run.Append(new Text(text));
        para.Append(run);
        cell.Append(para);
        
        return cell;
    }
    
    #endregion

    #region Text Helpers
    
    private void AddHeading1(Body body, string text)
    {
        var para = new Paragraph();
        para.Append(new ParagraphProperties(new ParagraphStyleId() { Val = "Heading1" }));
        var run = new Run();
        run.Append(new RunProperties(
            new RunFonts() { Ascii = "Segoe UI Light", HighAnsi = "Segoe UI Light" },
            new FontSize() { Val = "52" },
            new Color() { Val = _primaryColor }
        ));
        run.Append(new Text(text));
        para.Append(run);
        body.Append(para);
    }
    
    private void AddHeading2(Body body, string text)
    {
        var para = new Paragraph();
        para.Append(new ParagraphProperties(
            new SpacingBetweenLines() { Before = "300", After = "120" }
        ));
        var run = new Run();
        run.Append(new RunProperties(
            new RunFonts() { Ascii = "Segoe UI", HighAnsi = "Segoe UI" },
            new Bold(),
            new FontSize() { Val = "28" },
            new Color() { Val = _primaryColor }
        ));
        run.Append(new Text(text));
        para.Append(run);
        body.Append(para);
    }
    
    private void AddIntroText(Body body, string text)
    {
        var para = new Paragraph();
        para.Append(new ParagraphProperties(
            new SpacingBetweenLines() { Before = "60", After = "120", Line = "276", LineRule = LineSpacingRuleValues.Auto }
        ));
        var run = new Run();
        run.Append(new RunProperties(
            new RunFonts() { Ascii = "Segoe UI", HighAnsi = "Segoe UI" },
            new FontSize() { Val = "20" },
            new Color() { Val = "666666" }
        ));
        run.Append(new Text(text));
        para.Append(run);
        body.Append(para);
    }
    
    private void AddBodyText(Body body, string text)
    {
        var para = new Paragraph();
        para.Append(new ParagraphProperties(
            new SpacingBetweenLines() { Before = "100", After = "150", Line = "276", LineRule = LineSpacingRuleValues.Auto }
        ));
        var run = new Run();
        run.Append(new RunProperties(
            new RunFonts() { Ascii = "Segoe UI", HighAnsi = "Segoe UI" },
            new FontSize() { Val = "22" },
            new Color() { Val = _textColor }
        ));
        run.Append(new Text(text));
        para.Append(run);
        body.Append(para);
    }
    
    private void AddWarningText(Body body, string text)
    {
        var para = new Paragraph();
        para.Append(new ParagraphProperties(
            new SpacingBetweenLines() { Before = "100", After = "100" }
        ));
        var run = new Run();
        run.Append(new RunProperties(
            new RunFonts() { Ascii = "Segoe UI", HighAnsi = "Segoe UI" },
            new Bold(),
            new FontSize() { Val = "22" },
            new Color() { Val = _nonCompliantColor }
        ));
        run.Append(new Text(text));
        para.Append(run);
        body.Append(para);
    }
    
    private void AddNoteText(Body body, string text)
    {
        var para = new Paragraph();
        para.Append(new ParagraphProperties(
            new SpacingBetweenLines() { Before = "40", After = "40" }
        ));
        var run = new Run();
        run.Append(new RunProperties(
            new RunFonts() { Ascii = "Segoe UI", HighAnsi = "Segoe UI" },
            new Italic(),
            new FontSize() { Val = "18" },
            new Color() { Val = "666666" }
        ));
        run.Append(new Text(text));
        para.Append(run);
        body.Append(para);
    }
    
    private void AddEmptyParagraphs(Body body, int count)
    {
        for (int i = 0; i < count; i++)
        {
            body.Append(new Paragraph(new Run(new Text(" ") { Space = SpaceProcessingModeValues.Preserve })));
        }
    }
    
    #endregion

    #region Helpers
    
    private (string line1, string line2) ParseReportTitle(string title)
    {
        if (string.IsNullOrWhiteSpace(title))
            return ("MICROSOFT 365", "SECURITY ASSESSMENT");
        
        var lower = title.ToLower();
        if (lower.StartsWith("microsoft 365"))
            return ("MICROSOFT 365", title.Substring(13).Trim());
        if (lower.StartsWith("m365"))
            return ("MICROSOFT 365", title.Substring(4).Trim());
        
        return ("MICROSOFT 365", title);
    }
    
    private string FormatVersionStatus(VersionStatus status, string? message)
    {
        var prefix = status switch
        {
            VersionStatus.Current => "✓",
            VersionStatus.Warning => "⚠",
            VersionStatus.Critical => "❌",
            _ => "?"
        };
        return $"{prefix} {message ?? status.ToString()}";
    }
    
    #endregion
}
