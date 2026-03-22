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
/// Styled to match Cloud1st branding standards
/// </summary>
public class WordReportGenerator
{
    // Branding colors (matching PDF example)
    private string _primaryColor = "1E3A5F";  // Dark blue (cover page background)
    private string _accentColor = "E07C3A";   // Orange (accent text)
    private string _compliantColor = "107C6C"; // Teal (compliant status)
    private string _nonCompliantColor = "6B7280"; // Gray (non-compliant status)
    private string _exceptionColor = "F59E0B"; // Amber (exception status)
    private string _criticalColor = "DC2626"; // Red (critical issues)
    private string _headerTextColor = "1E3A5F"; // Dark blue for headers
    private string _bodyTextColor = "374151"; // Dark gray for body text
    private string _lightGray = "F3F4F6"; // Light gray for table backgrounds
    private string _borderColor = "E5E7EB"; // Table border color
    
    private ReportSettings _settings = new();
    
    public void SetSettings(ReportSettings settings)
    {
        _settings = settings;
        if (!string.IsNullOrEmpty(settings.PrimaryColor))
            _primaryColor = settings.PrimaryColor.TrimStart('#');
        if (!string.IsNullOrEmpty(settings.AccentColor))
            _accentColor = settings.AccentColor.TrimStart('#');
    }

    public byte[] GenerateReport(ExecutiveReportData data, ReportSettings settings)
    {
        SetSettings(settings);
        
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document();
            var body = mainPart.Document.AppendChild(new Body());
            
            // Set up page margins
            var sectionProps = new SectionProperties(
                new PageMargin() 
                { 
                    Top = 720, Bottom = 720, Left = 720, Right = 720, // 0.5 inch margins
                    Header = 360, Footer = 360 
                }
            );
            
            // Add styles
            AddDocumentStyles(mainPart);
            
            // === COVER PAGE ===
            AddCoverPage(body, mainPart, data);
            AddPageBreak(body);
            
            // === EXECUTIVE SUMMARY ===
            AddExecutiveSummary(body, data);
            AddPageBreak(body);
            
            // === INFOGRAPHIC: Pick 3 random quotes from settings pool ===
            var selectedQuotes = PickRandomQuotes(_settings, 3);

            if (_settings.ShowInfoGraphics && _settings.ShowQuotes && selectedQuotes.Count > 0)
            {
                AddInfoGraphicPage(body, selectedQuotes[0].BigNumber, selectedQuotes[0].Line1, selectedQuotes[0].Line2, selectedQuotes[0].Source);
                AddPageBreak(body);
            }
            
            // === SECURITY SCORE ===
            if (data.SecureScore != null)
            {
                AddSecurityScoreSection(body, data);
            }
            
            // === DEVICE STATISTICS ===
            if (data.DeviceStats != null)
            {
                AddDeviceStatisticsSection(body, data);
            }
            
            // === WINDOWS PATCH STATUS ===
            if (data.WindowsUpdateStats != null)
            {
                AddWindowsPatchSection(body, data);
            }
            
            // === MICROSOFT DEFENDER ===
            if (data.DefenderStats != null)
            {
                AddDefenderSection(body, data);
            }
            
            // === USER ACCOUNTS ===
            if (data.UserStats != null)
            {
                AddUserAccountsSection(body, data);
            }
            
            // === INFOGRAPHIC: Second quote (if enabled) ===
            if (_settings.ShowInfoGraphics && _settings.ShowQuotes && selectedQuotes.Count > 1)
            {
                AddPageBreak(body);
                AddInfoGraphicPage(body, selectedQuotes[1].BigNumber, selectedQuotes[1].Line1, selectedQuotes[1].Line2, selectedQuotes[1].Source);
            }
            
            // === DEVICE DETAILS ===
            if (data.DeviceDetails != null)
            {
                AddDeviceDetailsSection(body, data);
            }
            
            // === USER SIGN-IN DETAILS ===
            if (data.UserSignInDetails?.Any() == true)
            {
                AddPageBreak(body);
                AddUserSignInSection(body, data);
            }
            
            // === DELETED USERS ===
            if (data.DeletedUsersInPeriod?.Any() == true)
            {
                AddDeletedUsersSection(body, data);
            }
            
            // === MAILBOX DETAILS ===
            if (data.MailboxDetails?.Any() == true)
            {
                AddPageBreak(body);
                AddMailboxDetailsSection(body, data);
            }
            
            // === INFOGRAPHIC: Third quote (if enabled) ===
            if (_settings.ShowInfoGraphics && _settings.ShowQuotes && selectedQuotes.Count > 2)
            {
                AddPageBreak(body);
                AddInfoGraphicPage(body, selectedQuotes[2].BigNumber, selectedQuotes[2].Line1, selectedQuotes[2].Line2, selectedQuotes[2].Source);
            }
            
            // === DOMAIN SECURITY ===
            if (data.DomainSecuritySummary != null)
            {
                AddPageBreak(body);
                AddDomainSecuritySection(body, data);
            }
            
            // === APP CREDENTIALS ===
            if (data.AppCredentialStatus != null)
            {
                AddAppCredentialsSection(body, data);
            }
            
            // Footer note
            AddFooterNote(body);
            
            // Add section properties
            body.Append(sectionProps);
            
            mainPart.Document.Save();
        }
        
        return stream.ToArray();
    }

    #region Document Styles
    
    private void AddDocumentStyles(MainDocumentPart mainPart)
    {
        var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
        var styles = new Styles();
        
        // Default document font - Segoe UI like the PDF
        var defaultStyle = new Style()
        {
            Type = StyleValues.Paragraph,
            StyleId = "Normal",
            Default = true
        };
        defaultStyle.Append(new StyleName() { Val = "Normal" });
        defaultStyle.Append(new StyleRunProperties(
            new RunFonts() { Ascii = "Segoe UI", HighAnsi = "Segoe UI" },
            new FontSize() { Val = "22" }, // 11pt
            new Color() { Val = _bodyTextColor }
        ));
        defaultStyle.Append(new StyleParagraphProperties(
            new SpacingBetweenLines() { After = "120", Line = "276", LineRule = LineSpacingRuleValues.Auto }
        ));
        styles.Append(defaultStyle);
        
        // Heading 1 - Main section titles (like "Executive Summary")
        var h1 = new Style() { Type = StyleValues.Paragraph, StyleId = "Heading1" };
        h1.Append(new StyleName() { Val = "Heading 1" });
        h1.Append(new BasedOn() { Val = "Normal" });
        h1.Append(new StyleRunProperties(
            new RunFonts() { Ascii = "Segoe UI Light", HighAnsi = "Segoe UI Light" },
            new FontSize() { Val = "56" }, // 28pt
            new Color() { Val = _headerTextColor }
        ));
        h1.Append(new StyleParagraphProperties(
            new SpacingBetweenLines() { Before = "400", After = "240" }
        ));
        styles.Append(h1);
        
        // Heading 2 - Sub-section titles
        var h2 = new Style() { Type = StyleValues.Paragraph, StyleId = "Heading2" };
        h2.Append(new StyleName() { Val = "Heading 2" });
        h2.Append(new BasedOn() { Val = "Normal" });
        h2.Append(new StyleRunProperties(
            new RunFonts() { Ascii = "Segoe UI", HighAnsi = "Segoe UI" },
            new Bold(),
            new FontSize() { Val = "32" }, // 16pt
            new Color() { Val = _headerTextColor }
        ));
        h2.Append(new StyleParagraphProperties(
            new SpacingBetweenLines() { Before = "300", After = "160" }
        ));
        styles.Append(h2);
        
        stylesPart.Styles = styles;
        stylesPart.Styles.Save();
    }
    
    #endregion

    #region Cover Page & Infographics
    
    /// <summary>
    /// Add a full-page infographic with a large statistic (like the PDF example)
    /// </summary>
    private void AddInfoGraphicPage(Body body, string bigNumber, string line1, string line2, string source)
    {
        // Create a centered statistic callout using a table for the colored box
        // This approach works better in Word than trying to shade entire paragraphs
        
        // Top spacing
        AddEmptyParagraph(body, 3);
        
        // Create a single-cell table with dark background as a "card"
        var table = new Table();
        var tableProps = new TableProperties(
            new TableWidth() { Width = "5000", Type = TableWidthUnitValues.Pct },
            new TableBorders(
                new TopBorder() { Val = BorderValues.None },
                new BottomBorder() { Val = BorderValues.None },
                new LeftBorder() { Val = BorderValues.None },
                new RightBorder() { Val = BorderValues.None }
            ),
            new TableJustification() { Val = TableRowAlignmentValues.Center }
        );
        table.Append(tableProps);
        
        var row = new TableRow();
        var cell = new TableCell();
        cell.Append(new TableCellProperties(
            new TableCellWidth() { Width = "0", Type = TableWidthUnitValues.Auto },
            new Shading() { Fill = _primaryColor, Val = ShadingPatternValues.Clear },
            new TableCellMargin(
                new LeftMargin() { Width = "400", Type = TableWidthUnitValues.Dxa },
                new RightMargin() { Width = "400", Type = TableWidthUnitValues.Dxa },
                new TopMargin() { Width = "600", Type = TableWidthUnitValues.Dxa },
                new BottomMargin() { Width = "600", Type = TableWidthUnitValues.Dxa }
            )
        ));
        
        // Big number
        var numberPara = new Paragraph();
        numberPara.Append(new ParagraphProperties(
            new Justification() { Val = JustificationValues.Center },
            new SpacingBetweenLines() { Before = "0", After = "100" }
        ));
        var numberRun = new Run();
        numberRun.Append(new RunProperties(
            new RunFonts() { Ascii = "Segoe UI Light", HighAnsi = "Segoe UI Light" },
            new FontSize() { Val = "144" }, // 72pt
            new Color() { Val = "FFFFFF" }
        ));
        numberRun.Append(new Text(bigNumber));
        numberPara.Append(numberRun);
        cell.Append(numberPara);
        
        // First line
        var line1Para = new Paragraph();
        line1Para.Append(new ParagraphProperties(
            new Justification() { Val = JustificationValues.Center },
            new SpacingBetweenLines() { Before = "0", After = "40" }
        ));
        var line1Run = new Run();
        line1Run.Append(new RunProperties(
            new RunFonts() { Ascii = "Segoe UI", HighAnsi = "Segoe UI" },
            new FontSize() { Val = "28" },
            new Color() { Val = "FFFFFF" }
        ));
        line1Run.Append(new Text(line1));
        line1Para.Append(line1Run);
        cell.Append(line1Para);
        
        // Second line (accent color)
        var line2Para = new Paragraph();
        line2Para.Append(new ParagraphProperties(
            new Justification() { Val = JustificationValues.Center },
            new SpacingBetweenLines() { Before = "0", After = "0" }
        ));
        var line2Run = new Run();
        line2Run.Append(new RunProperties(
            new RunFonts() { Ascii = "Segoe UI", HighAnsi = "Segoe UI" },
            new FontSize() { Val = "28" },
            new Bold(),
            new Color() { Val = _accentColor }
        ));
        line2Run.Append(new Text(line2));
        line2Para.Append(line2Run);
        cell.Append(line2Para);
        
        row.Append(cell);
        table.Append(row);
        body.Append(table);
        
        // Source citation below the card
        var sourcePara = new Paragraph();
        sourcePara.Append(new ParagraphProperties(
            new Justification() { Val = JustificationValues.Center },
            new SpacingBetweenLines() { Before = "200", After = "0" }
        ));
        var sourceRun = new Run();
        sourceRun.Append(new RunProperties(
            new RunFonts() { Ascii = "Segoe UI", HighAnsi = "Segoe UI" },
            new FontSize() { Val = "18" },
            new Italic(),
            new Color() { Val = "888888" }
        ));
        sourceRun.Append(new Text(source));
        sourcePara.Append(sourceRun);
        body.Append(sourcePara);
    }
    
    private void AddCoverPage(Body body, MainDocumentPart mainPart, ExecutiveReportData data)
    {
        // Create a full-page dark blue header section
        // Multiple paragraphs with dark background to simulate the cover page
        
        // Top spacer with dark background
        for (int i = 0; i < 8; i++)
        {
            var spacer = new Paragraph();
            spacer.Append(new ParagraphProperties(
                new Shading() { Fill = _primaryColor, Val = ShadingPatternValues.Clear },
                new SpacingBetweenLines() { Before = "0", After = "0", Line = "480", LineRule = LineSpacingRuleValues.Exact }
            ));
            spacer.Append(new Run(new Text(" ") { Space = SpaceProcessingModeValues.Preserve }));
            body.Append(spacer);
        }
        
        // Parse report title into two lines (e.g., "Microsoft 365" and "Security Assessment")
        var titleParts = ParseReportTitle(_settings.ReportTitle);
        
        // First line - small caps in orange/accent (e.g., "MICROSOFT 365")
        var m365Para = new Paragraph();
        m365Para.Append(new ParagraphProperties(
            new Shading() { Fill = _primaryColor, Val = ShadingPatternValues.Clear },
            new SpacingBetweenLines() { Before = "0", After = "120" }
        ));
        var m365Run = new Run();
        m365Run.Append(new RunProperties(
            new RunFonts() { Ascii = "Segoe UI", HighAnsi = "Segoe UI" },
            new FontSize() { Val = "36" }, // 18pt
            new Color() { Val = _accentColor },
            new Caps()
        ));
        m365Run.Append(new Text(titleParts.line1));
        m365Para.Append(m365Run);
        body.Append(m365Para);
        
        // Second line - large white text (e.g., "SECURITY ASSESSMENT")
        var titlePara = new Paragraph();
        titlePara.Append(new ParagraphProperties(
            new Shading() { Fill = _primaryColor, Val = ShadingPatternValues.Clear },
            new SpacingBetweenLines() { Before = "0", After = "120" }
        ));
        var titleRun = new Run();
        titleRun.Append(new RunProperties(
            new RunFonts() { Ascii = "Segoe UI Light", HighAnsi = "Segoe UI Light" },
            new FontSize() { Val = "64" }, // 32pt
            new Color() { Val = "FFFFFF" },
            new Caps()
        ));
        titleRun.Append(new Text(titleParts.line2));
        titlePara.Append(titleRun);
        body.Append(titlePara);
        
        // More dark background spacers
        for (int i = 0; i < 20; i++)
        {
            var spacer = new Paragraph();
            spacer.Append(new ParagraphProperties(
                new Shading() { Fill = _primaryColor, Val = ShadingPatternValues.Clear },
                new SpacingBetweenLines() { Before = "0", After = "0", Line = "480", LineRule = LineSpacingRuleValues.Exact }
            ));
            spacer.Append(new Run(new Text(" ") { Space = SpaceProcessingModeValues.Preserve }));
            body.Append(spacer);
        }
        
        // White section for company name and date
        AddEmptyParagraph(body, 2);
        
        // Company name
        var companyPara = new Paragraph();
        companyPara.Append(new ParagraphProperties(
            new Justification() { Val = JustificationValues.Center },
            new SpacingBetweenLines() { Before = "400", After = "120" }
        ));
        var companyRun = new Run();
        companyRun.Append(new RunProperties(
            new RunFonts() { Ascii = "Segoe UI", HighAnsi = "Segoe UI" },
            new Bold(),
            new FontSize() { Val = "36" }, // 18pt
            new Color() { Val = _headerTextColor }
        ));
        companyRun.Append(new Text(_settings.CompanyName));
        companyPara.Append(companyRun);
        body.Append(companyPara);
        
        // Generated date
        var datePara = new Paragraph();
        datePara.Append(new ParagraphProperties(
            new Justification() { Val = JustificationValues.Center },
            new SpacingBetweenLines() { Before = "60", After = "200" }
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
        
        // Logo if available
        if (!string.IsNullOrEmpty(_settings.LogoBase64))
        {
            try
            {
                AddCenteredLogo(body, mainPart);
            }
            catch { /* Skip if logo fails */ }
        }
    }
    
    private void AddCenteredLogo(Body body, MainDocumentPart mainPart)
    {
        var logoPara = new Paragraph();
        logoPara.Append(new ParagraphProperties(
            new Justification() { Val = JustificationValues.Center },
            new SpacingBetweenLines() { Before = "400", After = "200" }
        ));
        
        try
        {
            var logoBytes = Convert.FromBase64String(_settings.LogoBase64!);
            var imagePart = mainPart.AddImagePart(
                _settings.LogoContentType?.Contains("png") == true 
                    ? ImagePartType.Png 
                    : ImagePartType.Jpeg);
            using var imgStream = new MemoryStream(logoBytes);
            imagePart.FeedData(imgStream);
            
            var relationshipId = mainPart.GetIdOfPart(imagePart);
            const long cx = 1800000L; // ~2 inches
            const long cy = 600000L;  // ~0.65 inch
            
            var logoRun = new Run(new Drawing(
                new DW.Inline(
                    new DW.Extent() { Cx = cx, Cy = cy },
                    new DW.EffectExtent() { LeftEdge = 0, TopEdge = 0, RightEdge = 0, BottomEdge = 0 },
                    new DW.DocProperties() { Id = 1U, Name = "Logo" },
                    new DW.NonVisualGraphicFrameDrawingProperties(
                        new A.GraphicFrameLocks() { NoChangeAspect = true }),
                    new A.Graphic(
                        new A.GraphicData(
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
                                    new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }
                                )))
                        { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                )
                { DistanceFromTop = 0U, DistanceFromBottom = 0U, DistanceFromLeft = 0U, DistanceFromRight = 0U }));
            
            logoPara.Append(logoRun);
        }
        catch { return; }
        
        body.Append(logoPara);
    }
    
    #endregion

    #region Executive Summary
    
    private void AddExecutiveSummary(Body body, ExecutiveReportData data)
    {
        AddSectionHeading(body, "Executive Summary");
        
        // KPI Cards Row - matching the PDF layout with 3 cards
        var kpiTable = CreateKpiCardsTable(new[]
        {
            (
                Icon: "👤",
                Value: $"{data.UserStats?.TotalUsers ?? 0}",
                Title: "Total Users",
                Subtitle: $"Including {data.UserStats?.GuestUsers ?? 0} guest users"
            ),
            (
                Icon: "📋",
                Value: $"{data.UserStats?.TotalUsers - (data.UserStats?.GuestUsers ?? 0) - (data.UserStats?.DeletedUsers ?? 0) ?? 0}",
                Title: "Licensed Users",
                Subtitle: $"{data.UserStats?.GuestUsers ?? 0} unlicensed users"
            ),
            (
                Icon: "🔒",
                Value: $"{data.UserStats?.MfaRegistered ?? 0}",
                Title: "MFA Registered",
                Subtitle: $"{data.UserStats?.MfaNotRegistered ?? 0} not registered"
            )
        });
        body.Append(kpiTable);
        
        AddEmptyParagraph(body, 1);
        
        // Introduction paragraph
        var introPara = new Paragraph();
        introPara.Append(new ParagraphProperties(
            new SpacingBetweenLines() { Before = "200", After = "200", Line = "276", LineRule = LineSpacingRuleValues.Auto }
        ));
        var introRun = new Run();
        introRun.Append(new RunProperties(
            new RunFonts() { Ascii = "Segoe UI", HighAnsi = "Segoe UI" },
            new FontSize() { Val = "22" },
            new Color() { Val = _bodyTextColor }
        ));
        introRun.Append(new Text($"This report was prepared for {_settings.CompanyName} in {data.GeneratedAt:MMMM yyyy}. This {_settings.ReportTitle} provides a comprehensive analysis of your organization's security configuration across key Microsoft 365 services, including Entra ID (Azure AD), Exchange Online, Intune, SharePoint, and Teams."));
        introPara.Append(introRun);
        body.Append(introPara);
        
        var intro2Para = new Paragraph();
        intro2Para.Append(new ParagraphProperties(
            new SpacingBetweenLines() { Before = "120", After = "200", Line = "276", LineRule = LineSpacingRuleValues.Auto }
        ));
        var intro2Run = new Run();
        intro2Run.Append(new RunProperties(
            new RunFonts() { Ascii = "Segoe UI", HighAnsi = "Segoe UI" },
            new FontSize() { Val = "22" },
            new Color() { Val = _bodyTextColor }
        ));
        intro2Run.Append(new Text("The aim of this review is to provide a clear and actionable understanding of your current security posture within Microsoft 365, helping to mitigate potential risks, safeguard sensitive data, and ensure compliance with leading security benchmarks."));
        intro2Para.Append(intro2Run);
        body.Append(intro2Para);
    }
    
    private Table CreateKpiCardsTable((string Icon, string Value, string Title, string Subtitle)[] kpis)
    {
        var table = new Table();
        
        var tableProps = new TableProperties(
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
        );
        table.Append(tableProps);
        
        var row = new TableRow();
        row.Append(new TableRowProperties(new TableRowHeight() { Val = 1800 }));
        
        foreach (var kpi in kpis)
        {
            var cell = new TableCell();
            cell.Append(new TableCellProperties(
                new TableCellWidth() { Width = "0", Type = TableWidthUnitValues.Auto },
                new Shading() { Fill = _lightGray, Val = ShadingPatternValues.Clear },
                new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center },
                new TableCellMargin(
                    new LeftMargin() { Width = "200", Type = TableWidthUnitValues.Dxa },
                    new RightMargin() { Width = "200", Type = TableWidthUnitValues.Dxa },
                    new TopMargin() { Width = "200", Type = TableWidthUnitValues.Dxa },
                    new BottomMargin() { Width = "200", Type = TableWidthUnitValues.Dxa }
                )
            ));
            
            // Large number/value
            var valuePara = new Paragraph();
            valuePara.Append(new ParagraphProperties(
                new Justification() { Val = JustificationValues.Center },
                new SpacingBetweenLines() { Before = "0", After = "60" }
            ));
            var valueRun = new Run();
            valueRun.Append(new RunProperties(
                new RunFonts() { Ascii = "Segoe UI Light", HighAnsi = "Segoe UI Light" },
                new FontSize() { Val = "72" }, // 36pt - large number
                new Color() { Val = _headerTextColor }
            ));
            valueRun.Append(new Text(kpi.Value));
            valuePara.Append(valueRun);
            cell.Append(valuePara);
            
            // Title
            var titlePara = new Paragraph();
            titlePara.Append(new ParagraphProperties(
                new Justification() { Val = JustificationValues.Center },
                new SpacingBetweenLines() { Before = "0", After = "40" }
            ));
            var titleRun = new Run();
            titleRun.Append(new RunProperties(
                new RunFonts() { Ascii = "Segoe UI", HighAnsi = "Segoe UI" },
                new FontSize() { Val = "24" }, // 12pt
                new Bold(),
                new Color() { Val = _bodyTextColor }
            ));
            titleRun.Append(new Text(kpi.Title));
            titlePara.Append(titleRun);
            cell.Append(titlePara);
            
            // Subtitle
            var subtitlePara = new Paragraph();
            subtitlePara.Append(new ParagraphProperties(
                new Justification() { Val = JustificationValues.Center },
                new SpacingBetweenLines() { Before = "0", After = "0" }
            ));
            var subtitleRun = new Run();
            subtitleRun.Append(new RunProperties(
                new RunFonts() { Ascii = "Segoe UI", HighAnsi = "Segoe UI" },
                new FontSize() { Val = "20" }, // 10pt
                new Color() { Val = "6B7280" }
            ));
            subtitleRun.Append(new Text(kpi.Subtitle));
            subtitlePara.Append(subtitleRun);
            cell.Append(subtitlePara);
            
            row.Append(cell);
        }
        
        table.Append(row);
        return table;
    }
    
    #endregion

    #region Report Sections
    
    private void AddSecurityScoreSection(Body body, ExecutiveReportData data)
    {
        AddSubSectionHeading(body, "Microsoft Secure Score");
        AddIntroText(body, "Microsoft Secure Score is a measurement of an organization's security posture, with a higher number indicating more improvement actions taken.");
        
        var rows = new List<string[]>
        {
            new[] { "Current Score", $"{data.SecureScore?.CurrentScore ?? 0}" },
            new[] { "Maximum Score", $"{data.SecureScore?.MaxScore ?? 0}" },
            new[] { "Percentage", $"{data.SecureScore?.PercentageScore ?? 0}%" }
        };
        
        AddSimpleTable(body, new[] { "Metric", "Value" }, rows);
    }
    
    private void AddDeviceStatisticsSection(Body body, ExecutiveReportData data)
    {
        AddSubSectionHeading(body, "Intune Managed Devices");
        AddIntroText(body, "Overview of all devices managed through Microsoft Intune, including compliance status across different platforms.");
        
        var rows = new List<string[]>
        {
            new[] { "Total Devices", $"{data.DeviceStats?.TotalDevices ?? 0}" },
            new[] { "Windows", $"{data.DeviceStats?.WindowsDevices ?? 0}" },
            new[] { "macOS", $"{data.DeviceStats?.MacOsDevices ?? 0}" },
            new[] { "iOS/iPadOS", $"{data.DeviceStats?.IosDevices ?? 0}" },
            new[] { "Android", $"{data.DeviceStats?.AndroidDevices ?? 0}" },
            new[] { "Compliant", $"{data.DeviceStats?.CompliantDevices ?? 0}" },
            new[] { "Non-Compliant", $"{data.DeviceStats?.NonCompliantDevices ?? 0}" },
            new[] { "Compliance Rate", $"{data.DeviceStats?.ComplianceRate ?? 0}%" }
        };
        
        AddSimpleTable(body, new[] { "Platform", "Count" }, rows);
    }
    
    private void AddWindowsPatchSection(Body body, ExecutiveReportData data)
    {
        AddSubSectionHeading(body, "Windows Patch Status");
        
        var rows = new List<string[]>
        {
            new[] { "Total Windows Devices", $"{data.WindowsUpdateStats?.TotalWindowsDevices ?? 0}" },
            new[] { "Up to Date", $"{data.WindowsUpdateStats?.UpToDate ?? 0}" },
            new[] { "Needs Update", $"{data.WindowsUpdateStats?.NeedsUpdate ?? 0}" },
            new[] { "Compliance Rate", $"{data.WindowsUpdateStats?.ComplianceRate ?? 0}%" }
        };
        
        AddSimpleTable(body, new[] { "Status", "Count" }, rows);
        
        if (!string.IsNullOrEmpty(data.WindowsUpdateStats?.Note))
        {
            AddNoteText(body, data.WindowsUpdateStats.Note);
        }
    }
    
    private void AddDefenderSection(Body body, ExecutiveReportData data)
    {
        AddSubSectionHeading(body, "Microsoft Defender for Endpoint");
        
        var rows = new List<string[]>
        {
            new[] { "Exposure Score", data.DefenderStats?.ExposureScore ?? "N/A" }
        };
        
        if (data.DefenderStats?.OnboardedMachines.HasValue == true)
            rows.Add(new[] { "Onboarded Machines", $"{data.DefenderStats.OnboardedMachines}" });
            
        rows.Add(new[] { "Total Vulnerabilities", $"{data.DefenderStats?.VulnerabilitiesDetected ?? 0}" });
        rows.Add(new[] { "Critical", $"{data.DefenderStats?.CriticalVulnerabilities ?? 0}" });
        rows.Add(new[] { "High", $"{data.DefenderStats?.HighVulnerabilities ?? 0}" });
        rows.Add(new[] { "Medium", $"{data.DefenderStats?.MediumVulnerabilities ?? 0}" });
        rows.Add(new[] { "Low", $"{data.DefenderStats?.LowVulnerabilities ?? 0}" });
        
        AddSimpleTable(body, new[] { "Metric", "Value" }, rows);
    }
    
    private void AddUserAccountsSection(Body body, ExecutiveReportData data)
    {
        AddSubSectionHeading(body, "User Accounts");
        AddIntroText(body, "Summary of user accounts including MFA registration status and high-risk users identified by Microsoft Entra ID Protection.");
        
        var rows = new List<string[]>
        {
            new[] { "Total Users", $"{data.UserStats?.TotalUsers ?? 0}" },
            new[] { "Guest Users", $"{data.UserStats?.GuestUsers ?? 0}" },
            new[] { "Deleted Users (Soft)", $"{data.UserStats?.DeletedUsers ?? 0}" },
            new[] { "MFA Registered", $"{data.UserStats?.MfaRegistered ?? 0}" },
            new[] { "MFA Not Registered", $"{data.UserStats?.MfaNotRegistered ?? 0}" },
            new[] { "Risky Users", $"{data.RiskyUsersCount}" }
        };
        
        AddSimpleTable(body, new[] { "Type", "Count" }, rows);
        
        if (data.HighRiskUsers?.Any() == true)
        {
            AddWarningText(body, $"High Risk Users: {string.Join(", ", data.HighRiskUsers)}");
        }
    }
    
    private void AddDeviceDetailsSection(Body body, ExecutiveReportData data)
    {
        // Windows Devices
        if (data.DeviceDetails?.WindowsDevices?.Any() == true)
        {
            AddPageBreak(body);
            AddSubSectionHeading(body, $"Windows Devices ({data.DeviceDetails.WindowsDevices.Count})");
            
            var headers = new[] { "Device Name", "OS Version", "Update Status", "Compliance", "Last Check-in" };
            var rows = data.DeviceDetails.WindowsDevices.Take(50).Select(d => new[]
            {
                d.DeviceName ?? "-",
                d.OsVersion ?? "-",
                GetVersionStatusText(d.OsVersionStatus, d.OsVersionStatusMessage),
                d.ComplianceState ?? "-",
                d.LastCheckIn?.ToString("dd MMM yyyy") ?? "Never"
            }).ToList();
            
            AddDetailTable(body, headers, rows, statusColumn: 2, complianceColumn: 3);
            
            if (data.DeviceDetails.WindowsDevices.Count > 50)
                AddNoteText(body, $"Showing first 50 of {data.DeviceDetails.WindowsDevices.Count} devices.");
        }
        
        // macOS Devices
        if (data.DeviceDetails?.MacDevices?.Any() == true)
        {
            AddSubSectionHeading(body, $"macOS Devices ({data.DeviceDetails.MacDevices.Count})");
            
            var headers = new[] { "Device Name", "OS Version", "Update Status", "Compliance", "Last Check-in" };
            var rows = data.DeviceDetails.MacDevices.Take(50).Select(d => new[]
            {
                d.DeviceName ?? "-",
                d.OsVersion ?? "-",
                GetVersionStatusText(d.OsVersionStatus, d.OsVersionStatusMessage),
                d.ComplianceState ?? "-",
                d.LastCheckIn?.ToString("dd MMM yyyy") ?? "Never"
            }).ToList();
            
            AddDetailTable(body, headers, rows, statusColumn: 2, complianceColumn: 3);
        }
        
        // iOS Devices
        if (data.DeviceDetails?.IosDevices?.Any() == true)
        {
            AddSubSectionHeading(body, $"iOS/iPadOS Devices ({data.DeviceDetails.IosDevices.Count})");
            
            var headers = new[] { "Device Name", "OS Version", "Update Status", "Compliance", "Last Check-in" };
            var rows = data.DeviceDetails.IosDevices.Take(50).Select(d => new[]
            {
                d.DeviceName ?? "-",
                d.OsVersion ?? "-",
                GetVersionStatusText(d.OsVersionStatus, d.OsVersionStatusMessage),
                d.ComplianceState ?? "-",
                d.LastCheckIn?.ToString("dd MMM yyyy") ?? "Never"
            }).ToList();
            
            AddDetailTable(body, headers, rows, statusColumn: 2, complianceColumn: 3);
        }
        
        // Android Devices
        if (data.DeviceDetails?.AndroidDevices?.Any() == true)
        {
            AddSubSectionHeading(body, $"Android Devices ({data.DeviceDetails.AndroidDevices.Count})");
            
            var headers = new[] { "Device Name", "OS Version", "Update Status", "Compliance", "Last Check-in" };
            var rows = data.DeviceDetails.AndroidDevices.Take(50).Select(d => new[]
            {
                d.DeviceName ?? "-",
                d.OsVersion ?? "-",
                GetVersionStatusText(d.OsVersionStatus, d.OsVersionStatusMessage),
                d.ComplianceState ?? "-",
                d.LastCheckIn?.ToString("dd MMM yyyy") ?? "Never"
            }).ToList();
            
            AddDetailTable(body, headers, rows, statusColumn: 2, complianceColumn: 3);
        }
    }
    
    private void AddUserSignInSection(Body body, ExecutiveReportData data)
    {
        AddSubSectionHeading(body, $"User Sign-in & MFA Details ({data.UserSignInDetails!.Count} users)");
        AddIntroText(body, "Detailed view of user sign-in activity and MFA registration status.");
        
        var headers = new[] { "Display Name", "Email", "Last Sign-in", "Default MFA", "MFA", "Enabled" };
        var rows = data.UserSignInDetails.Take(100).Select(u => new[]
        {
            u.DisplayName ?? "-",
            u.UserPrincipalName ?? "-",
            u.LastInteractiveSignIn?.ToString("dd MMM yyyy") ?? "Never",
            u.DefaultMfaMethod ?? "None",
            u.IsMfaRegistered ? "Yes" : "No",
            u.AccountEnabled ? "Yes" : "No"
        }).ToList();
        
        AddDetailTable(body, headers, rows, mfaColumn: 4, enabledColumn: 5);
    }
    
    private void AddDeletedUsersSection(Body body, ExecutiveReportData data)
    {
        AddSubSectionHeading(body, $"Deleted Users in Period ({data.DeletedUsersInPeriod!.Count} users)");
        
        var headers = new[] { "Display Name", "Email", "Deleted Date", "Job Title", "Department" };
        var rows = data.DeletedUsersInPeriod.Select(u => new[]
        {
            u.DisplayName ?? "-",
            u.UserPrincipalName ?? u.Mail ?? "-",
            u.DeletedDateTime?.ToString("dd MMM yyyy") ?? "-",
            u.JobTitle ?? "-",
            u.Department ?? "-"
        }).ToList();
        
        AddSimpleTable(body, headers, rows);
    }
    
    private void AddMailboxDetailsSection(Body body, ExecutiveReportData data)
    {
        AddSubSectionHeading(body, $"Mailbox Storage Details ({data.MailboxDetails!.Count} mailboxes)");
        
        var headers = new[] { "Display Name", "Email", "Size (GB)", "Quota (GB)", "% Used", "Items" };
        var rows = data.MailboxDetails.Take(100).Select(m => new[]
        {
            m.DisplayName ?? "-",
            m.UserPrincipalName ?? "-",
            $"{m.StorageUsedGB:F2}",
            m.QuotaGB.HasValue ? $"{m.QuotaGB:F0}" : "-",
            m.PercentUsed.HasValue ? $"{m.PercentUsed:F1}%" : "-",
            m.ItemCount?.ToString("N0") ?? "-"
        }).ToList();
        
        AddSimpleTable(body, headers, rows);
    }
    
    private void AddDomainSecuritySection(Body body, ExecutiveReportData data)
    {
        AddSubSectionHeading(body, "Domain Email Security");
        AddIntroText(body, "Email authentication protocols (SPF, DMARC, DKIM) help protect your organization from email spoofing and phishing attacks.");
        
        // Summary
        var summaryRows = new List<string[]>
        {
            new[] { "Total Domains Checked", $"{data.DomainSecuritySummary!.TotalDomains}" },
            new[] { "Domains with MX Records", $"{data.DomainSecuritySummary.DomainsWithMx}" },
            new[] { "Domains with SPF", $"{data.DomainSecuritySummary.DomainsWithSpf}" },
            new[] { "Domains with DMARC", $"{data.DomainSecuritySummary.DomainsWithDmarc}" },
            new[] { "Domains with DKIM", $"{data.DomainSecuritySummary.DomainsWithDkim}" }
        };
        
        AddSimpleTable(body, new[] { "Metric", "Count" }, summaryRows);
        
        // Domain details
        if (data.DomainSecurityResults?.Any() == true)
        {
            AddEmptyParagraph(body, 1);
            AddBodyText(body, "Domain Security Details:");
            
            var detailHeaders = new[] { "Domain", "MX", "SPF", "DMARC", "DKIM", "Score", "Grade" };
            var detailRows = data.DomainSecurityResults
                .OrderByDescending(d => d.SecurityScore)
                .Select(d => new[]
                {
                    d.Domain,
                    d.HasMx ? "✓" : "✗",
                    d.HasSpf ? "✓" : "✗",
                    d.HasDmarc ? d.DmarcPolicy ?? "✓" : "✗",
                    d.HasDkim ? "✓" : "✗",
                    $"{d.SecurityScore}",
                    d.SecurityGrade
                }).ToList();
            
            AddDetailTable(body, detailHeaders, detailRows, gradeColumn: 6);
        }
    }
    
    private void AddAppCredentialsSection(Body body, ExecutiveReportData data)
    {
        AddSubSectionHeading(body, "App Registration Credentials");
        AddIntroText(body, "Status of application secrets and certificates. Expired or expiring credentials can cause service disruptions.");
        
        var rows = new List<string[]>
        {
            new[] { "Total App Registrations", $"{data.AppCredentialStatus!.TotalApps}" },
            new[] { "Apps with Expiring Secrets", $"{data.AppCredentialStatus.AppsWithExpiringSecrets}" },
            new[] { "Apps with Expired Secrets", $"{data.AppCredentialStatus.AppsWithExpiredSecrets}" },
            new[] { "Apps with Expiring Certificates", $"{data.AppCredentialStatus.AppsWithExpiringCertificates}" },
            new[] { "Apps with Expired Certificates", $"{data.AppCredentialStatus.AppsWithExpiredCertificates}" }
        };
        
        AddSimpleTable(body, new[] { "Metric", "Count" }, rows);
    }
    
    private void AddFooterNote(Body body)
    {
        AddEmptyParagraph(body, 2);
        
        // Horizontal rule
        var hr = new Paragraph();
        hr.Append(new ParagraphProperties(
            new ParagraphBorders(
                new BottomBorder() { Val = BorderValues.Single, Size = 6, Color = _primaryColor }
            ),
            new SpacingBetweenLines() { Before = "400", After = "200" }
        ));
        hr.Append(new Run(new Text(" ") { Space = SpaceProcessingModeValues.Preserve }));
        body.Append(hr);
        
        // Custom footer text from settings (e.g., confidentiality notice)
        if (!string.IsNullOrWhiteSpace(_settings.FooterText))
        {
            AddNoteText(body, _settings.FooterText);
        }
        
        AddNoteText(body, "This report was automatically generated by M365 Dashboard.");
        AddNoteText(body, "Some metrics may require additional licensing or API permissions.");
    }
    
    #endregion

    #region Table Helpers
    
    private void AddSimpleTable(Body body, string[] headers, List<string[]> rows)
    {
        var table = new Table();
        
        var tableProps = new TableProperties(
            new TableWidth() { Width = "5000", Type = TableWidthUnitValues.Pct },
            new TableBorders(
                new TopBorder() { Val = BorderValues.Single, Size = 4, Color = _borderColor },
                new BottomBorder() { Val = BorderValues.Single, Size = 4, Color = _borderColor },
                new LeftBorder() { Val = BorderValues.Single, Size = 4, Color = _borderColor },
                new RightBorder() { Val = BorderValues.Single, Size = 4, Color = _borderColor },
                new InsideHorizontalBorder() { Val = BorderValues.Single, Size = 4, Color = _borderColor },
                new InsideVerticalBorder() { Val = BorderValues.Single, Size = 4, Color = _borderColor }
            )
        );
        table.Append(tableProps);
        
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
            var tableRow = new TableRow();
            foreach (var cellValue in rows[i])
            {
                tableRow.Append(CreateCell(cellValue, isHeader: false, isAlternate: i % 2 == 1));
            }
            table.Append(tableRow);
        }
        
        body.Append(table);
        AddEmptyParagraph(body, 1);
    }
    
    private void AddDetailTable(Body body, string[] headers, List<string[]> rows, 
        int statusColumn = -1, int complianceColumn = -1, int gradeColumn = -1, 
        int mfaColumn = -1, int enabledColumn = -1)
    {
        var table = new Table();
        
        var tableProps = new TableProperties(
            new TableWidth() { Width = "5000", Type = TableWidthUnitValues.Pct },
            new TableBorders(
                new TopBorder() { Val = BorderValues.Single, Size = 4, Color = _borderColor },
                new BottomBorder() { Val = BorderValues.Single, Size = 4, Color = _borderColor },
                new LeftBorder() { Val = BorderValues.Single, Size = 4, Color = _borderColor },
                new RightBorder() { Val = BorderValues.Single, Size = 4, Color = _borderColor },
                new InsideHorizontalBorder() { Val = BorderValues.Single, Size = 4, Color = _borderColor },
                new InsideVerticalBorder() { Val = BorderValues.Single, Size = 4, Color = _borderColor }
            )
        );
        table.Append(tableProps);
        
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
                
                if (j == complianceColumn)
                    textColor = GetComplianceColor(row[j]);
                else if (j == statusColumn)
                    textColor = GetStatusColor(row[j]);
                else if (j == gradeColumn)
                    textColor = GetGradeColor(row[j]);
                else if (j == mfaColumn || j == enabledColumn)
                    textColor = row[j] == "Yes" ? _compliantColor : _nonCompliantColor;
                
                tableRow.Append(CreateCell(row[j], isHeader: false, isAlternate: i % 2 == 1, textColor: textColor));
            }
            
            table.Append(tableRow);
        }
        
        body.Append(table);
        AddEmptyParagraph(body, 1);
    }
    
    private TableCell CreateCell(string text, bool isHeader = false, bool isAlternate = false, string? textColor = null)
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
        {
            cellProps.Append(new Shading() { Fill = _lightGray, Val = ShadingPatternValues.Clear });
        }
        
        cell.Append(cellProps);
        
        var para = new Paragraph();
        var run = new Run();
        
        var runProps = new RunProperties(
            new RunFonts() { Ascii = "Segoe UI", HighAnsi = "Segoe UI" },
            new FontSize() { Val = isHeader ? "20" : "18" }
        );
        
        if (isHeader)
        {
            runProps.Append(new Bold());
            runProps.Append(new Color() { Val = _bodyTextColor });
        }
        else if (textColor != null)
        {
            runProps.Append(new Color() { Val = textColor });
        }
        else
        {
            runProps.Append(new Color() { Val = _bodyTextColor });
        }
        
        run.Append(runProps);
        run.Append(new Text(text));
        para.Append(run);
        cell.Append(para);
        
        return cell;
    }
    
    #endregion

    #region Text Helpers
    
    private void AddSectionHeading(Body body, string text)
    {
        var para = new Paragraph();
        para.Append(new ParagraphProperties(
            new SpacingBetweenLines() { Before = "400", After = "240" }
        ));
        var run = new Run();
        run.Append(new RunProperties(
            new RunFonts() { Ascii = "Segoe UI Light", HighAnsi = "Segoe UI Light" },
            new FontSize() { Val = "56" },
            new Color() { Val = _headerTextColor }
        ));
        run.Append(new Text(text));
        para.Append(run);
        body.Append(para);
    }
    
    private void AddSubSectionHeading(Body body, string text)
    {
        var para = new Paragraph();
        para.Append(new ParagraphProperties(
            new SpacingBetweenLines() { Before = "300", After = "160" }
        ));
        var run = new Run();
        run.Append(new RunProperties(
            new RunFonts() { Ascii = "Segoe UI", HighAnsi = "Segoe UI" },
            new Bold(),
            new FontSize() { Val = "28" },
            new Color() { Val = _headerTextColor }
        ));
        run.Append(new Text(text));
        para.Append(run);
        body.Append(para);
    }
    
    private void AddIntroText(Body body, string text)
    {
        var para = new Paragraph();
        para.Append(new ParagraphProperties(
            new SpacingBetweenLines() { Before = "60", After = "160", Line = "276", LineRule = LineSpacingRuleValues.Auto }
        ));
        var run = new Run();
        run.Append(new RunProperties(
            new RunFonts() { Ascii = "Segoe UI", HighAnsi = "Segoe UI" },
            new FontSize() { Val = "22" },
            new Color() { Val = _bodyTextColor }
        ));
        run.Append(new Text(text));
        para.Append(run);
        body.Append(para);
    }
    
    private void AddBodyText(Body body, string text)
    {
        var para = new Paragraph();
        para.Append(new ParagraphProperties(
            new SpacingBetweenLines() { Before = "60", After = "120" }
        ));
        var run = new Run();
        run.Append(new RunProperties(
            new RunFonts() { Ascii = "Segoe UI", HighAnsi = "Segoe UI" },
            new Bold(),
            new FontSize() { Val = "22" },
            new Color() { Val = _bodyTextColor }
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
            new FontSize() { Val = "20" },
            new Color() { Val = "6B7280" }
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
            new Color() { Val = _criticalColor }
        ));
        run.Append(new Text(text));
        para.Append(run);
        body.Append(para);
    }
    
    private void AddEmptyParagraph(Body body, int count = 1)
    {
        for (int i = 0; i < count; i++)
        {
            var para = new Paragraph();
            para.Append(new Run(new Text(" ") { Space = SpaceProcessingModeValues.Preserve }));
            body.Append(para);
        }
    }
    
    private void AddPageBreak(Body body)
    {
        body.Append(new Paragraph(new Run(new Break() { Type = BreakValues.Page })));
    }
    
    #endregion

    #region Quote Helpers

    /// <summary>
    /// Selects <paramref name="count"/> quotes at random from the enabled quotes in settings.
    /// Falls back to the built-in defaults if the settings pool is empty.
    /// </summary>
    private static List<ReportQuote> PickRandomQuotes(ReportSettings settings, int count)
    {
        var pool = settings.Quotes
            .Where(q => q.Enabled &&
                        !string.IsNullOrWhiteSpace(q.BigNumber) &&
                        !string.IsNullOrWhiteSpace(q.Line1))
            .ToList();

        if (pool.Count == 0)
            pool = ReportSettings.DefaultQuotes();

        if (pool.Count <= count)
            return pool;

        // Fisher-Yates shuffle, take first <count>
        var rng = new Random();
        for (int i = pool.Count - 1; i > 0; i--)
        {
            int j = rng.Next(i + 1);
            (pool[i], pool[j]) = (pool[j], pool[i]);
        }

        return pool.Take(count).ToList();
    }

    #endregion

    #region Color Helpers
    
    private string GetComplianceColor(string value)
    {
        return value.ToLower() switch
        {
            "compliant" => _compliantColor,
            "non-compliant" => _criticalColor,
            "noncompliant" => _criticalColor,
            _ => _nonCompliantColor
        };
    }
    
    private string GetStatusColor(string value)
    {
        if (value.StartsWith("✓")) return _compliantColor;
        if (value.StartsWith("⚠")) return _exceptionColor;
        if (value.StartsWith("❌")) return _criticalColor;
        if (value.Contains("EOL") || value.Contains("Critical")) return _criticalColor;
        if (value.Contains("aging") || value.Contains("Warning")) return _exceptionColor;
        return _compliantColor;
    }
    
    private string GetGradeColor(string grade)
    {
        return grade switch
        {
            "A" => _compliantColor,
            "B" => _compliantColor,
            "C" => _exceptionColor,
            "D" => _criticalColor,
            "F" => _criticalColor,
            _ => _nonCompliantColor
        };
    }
    
    private string GetVersionStatusText(VersionStatus status, string? message)
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
    
    /// <summary>
    /// Parse the report title into two lines for the cover page
    /// e.g., "Microsoft 365 Security Assessment" -> ("Microsoft 365", "Security Assessment")
    /// </summary>
    private (string line1, string line2) ParseReportTitle(string title)
    {
        if (string.IsNullOrWhiteSpace(title))
            return ("MICROSOFT 365", "SECURITY ASSESSMENT");
        
        // Try common splits
        var lowerTitle = title.ToLower();
        
        // Check for "Microsoft 365" prefix
        if (lowerTitle.StartsWith("microsoft 365"))
        {
            var rest = title.Substring(13).Trim();
            return ("MICROSOFT 365", string.IsNullOrEmpty(rest) ? "SECURITY ASSESSMENT" : rest);
        }
        
        // Check for "M365" prefix
        if (lowerTitle.StartsWith("m365"))
        {
            var rest = title.Substring(4).Trim();
            return ("MICROSOFT 365", string.IsNullOrEmpty(rest) ? "SECURITY ASSESSMENT" : rest);
        }
        
        // Try to split at common words
        var splitWords = new[] { " Security ", " Report ", " Assessment ", " Analysis " };
        foreach (var splitWord in splitWords)
        {
            var idx = title.IndexOf(splitWord, StringComparison.OrdinalIgnoreCase);
            if (idx > 0)
            {
                return (title.Substring(0, idx).Trim(), title.Substring(idx).Trim());
            }
        }
        
        // If title is long enough, try to split in middle at a space
        if (title.Length > 20)
        {
            var midPoint = title.Length / 2;
            var spaceIdx = title.IndexOf(' ', midPoint);
            if (spaceIdx > 0)
            {
                return (title.Substring(0, spaceIdx).Trim(), title.Substring(spaceIdx).Trim());
            }
        }
        
        // Fallback: use whole title as line 2
        return ("MICROSOFT 365", title);
    }
    
    #endregion
}
