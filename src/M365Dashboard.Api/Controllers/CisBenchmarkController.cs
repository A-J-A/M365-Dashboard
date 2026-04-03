using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using M365Dashboard.Api.Models;
using M365Dashboard.Api.Services;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace M365Dashboard.Api.Controllers;

[ApiController]
[Route("api/[controller]")]
[Authorize]
public class CisBenchmarkController : ControllerBase
{
    private readonly ICisBenchmarkService _benchmarkService;
    private readonly ILogger<CisBenchmarkController> _logger;

    public CisBenchmarkController(
        ICisBenchmarkService benchmarkService,
        ILogger<CisBenchmarkController> logger)
    {
        _benchmarkService = benchmarkService;
        _logger = logger;
    }

    /// <summary>
    /// Run the CIS Microsoft 365 Foundations Benchmark assessment
    /// </summary>
    [HttpPost("run")]
    [HttpGet("run")]
    public async Task<IActionResult> RunBenchmark([FromBody] CisBenchmarkRequest? request = null)
    {
        try
        {
            _logger.LogInformation("Starting CIS Benchmark assessment");
            var result = await _benchmarkService.RunBenchmarkAsync(request);
            _logger.LogInformation("CIS Benchmark assessment completed: {Passed}/{Total} controls passed ({Percentage}%)",
                result.PassedControls, result.TotalControls, result.CompliancePercentage);
            return Ok(result);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error running CIS Benchmark assessment");
            return StatusCode(500, new { error = "Failed to run benchmark assessment", message = ex.Message });
        }
    }

    /// <summary>
    /// Get a summary of the last benchmark run
    /// </summary>
    [HttpGet("summary")]
    public async Task<IActionResult> GetSummary()
    {
        try
        {
            var result = await _benchmarkService.RunBenchmarkAsync();
            var summary = new CisBenchmarkSummary
            {
                LastScanDate = result.GeneratedAt,
                TotalControls = result.TotalControls,
                PassedControls = result.PassedControls,
                FailedControls = result.FailedControls,
                CompliancePercentage = result.CompliancePercentage,
                CriticalFailures = result.Controls
                    .Where(c => c.Status == CisControlStatus.Fail && c.Level == CisLevel.L1)
                    .OrderBy(c => c.ControlId)
                    .Take(10)
                    .ToList()
            };
            return Ok(summary);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error getting benchmark summary");
            return StatusCode(500, new { error = "Failed to get benchmark summary", message = ex.Message });
        }
    }

    /// <summary>
    /// Check a specific control
    /// </summary>
    [HttpGet("control/{controlId}")]
    public async Task<IActionResult> CheckControl(string controlId)
    {
        try
        {
            var result = await _benchmarkService.CheckControlAsync(controlId);
            return Ok(result);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error checking control {ControlId}", controlId);
            return StatusCode(500, new { error = $"Failed to check control {controlId}", message = ex.Message });
        }
    }

    /// <summary>
    /// Download the benchmark report as a Word document
    /// </summary>
    [HttpGet("download")]
    public async Task<IActionResult> DownloadReport([FromQuery] bool includeLevel2 = true, [FromQuery] bool includeE5 = true)
    {
        try
        {
            var request = new CisBenchmarkRequest
            {
                IncludeLevel2 = includeLevel2,
                IncludeE5Only = includeE5
            };
            
            var result = await _benchmarkService.RunBenchmarkAsync(request);
            var documentBytes = GenerateWordDocument(result);
            
            var fileName = $"CIS_M365_Benchmark_Report_{DateTime.UtcNow:yyyy-MM-dd}.docx";
            return File(documentBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", fileName);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error generating benchmark report");
            return StatusCode(500, new { error = "Failed to generate benchmark report", message = ex.Message });
        }
    }

    /// <summary>
    /// View the benchmark report as HTML
    /// </summary>
    [HttpGet("html")]
    public async Task<IActionResult> ViewHtmlReport([FromQuery] bool includeLevel2 = true, [FromQuery] bool includeE5 = true)
    {
        try
        {
            var request = new CisBenchmarkRequest
            {
                IncludeLevel2 = includeLevel2,
                IncludeE5Only = includeE5
            };
            
            var result = await _benchmarkService.RunBenchmarkAsync(request);
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
    /// Get available categories
    /// </summary>
    [HttpGet("categories")]
    public IActionResult GetCategories()
    {
        var categories = new[]
        {
            new { Id = "1", Name = "Microsoft 365 admin center" },
            new { Id = "2", Name = "Microsoft 365 Defender" },
            new { Id = "3", Name = "Microsoft Purview" },
            new { Id = "5", Name = "Microsoft Entra admin center" },
            new { Id = "6", Name = "Exchange Online" },
            new { Id = "7", Name = "SharePoint & OneDrive" },
            new { Id = "8", Name = "Microsoft Teams" },
            new { Id = "9", Name = "Microsoft Fabric (Power BI)" }
        };
        return Ok(categories);
    }

    private byte[] GenerateWordDocument(CisBenchmarkResult data)
    {
        using var stream = new MemoryStream();
        
        using (var wordDocument = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true))
        {
            var mainPart = wordDocument.AddMainDocumentPart();
            mainPart.Document = new Document();
            var body = mainPart.Document.AppendChild(new Body());

            AddStyles(mainPart);

            // Title
            AddHeading(body, "CIS Microsoft 365 Foundations Benchmark Report", 1, true);
            AddParagraph(body, $"Benchmark Version: {data.BenchmarkVersion}");
            AddParagraph(body, $"Generated: {data.GeneratedAt:dd MMMM yyyy HH:mm} UTC");
            AddParagraph(body, $"Tenant: {data.TenantName}");
            AddParagraph(body, "");

            // Executive Summary
            AddHeading(body, "Executive Summary", 2);
            
            var summaryTable = CreateTable(body, new[] { "Metric", "Value" });
            AddTableRow(summaryTable, new[] { "Total Controls Assessed", data.TotalControls.ToString() });
            AddTableRow(summaryTable, new[] { "Passed", data.PassedControls.ToString() });
            AddTableRow(summaryTable, new[] { "Failed", data.FailedControls.ToString() });
            AddTableRow(summaryTable, new[] { "Manual Review Required", data.ManualControls.ToString() });
            AddTableRow(summaryTable, new[] { "Compliance Score", $"{data.CompliancePercentage}%" });

            AddParagraph(body, "");
            
            // Level Breakdown
            AddHeading(body, "Compliance by Level", 2);
            var levelTable = CreateTable(body, new[] { "Level", "Passed", "Total", "Percentage" });
            var l1Pct = data.Level1Total > 0 ? Math.Round((double)data.Level1Passed / data.Level1Total * 100, 1) : 0;
            var l2Pct = data.Level2Total > 0 ? Math.Round((double)data.Level2Passed / data.Level2Total * 100, 1) : 0;
            AddTableRow(levelTable, new[] { "Level 1 (Essential)", data.Level1Passed.ToString(), data.Level1Total.ToString(), $"{l1Pct}%" });
            AddTableRow(levelTable, new[] { "Level 2 (Defense in Depth)", data.Level2Passed.ToString(), data.Level2Total.ToString(), $"{l2Pct}%" });

            AddParagraph(body, "");

            // Category Breakdown
            AddHeading(body, "Compliance by Category", 2);
            var catTable = CreateTable(body, new[] { "Category", "Passed", "Failed", "Manual", "Score" });
            foreach (var cat in data.Categories)
            {
                AddTableRow(catTable, new[] { 
                    cat.CategoryName, 
                    cat.PassedControls.ToString(), 
                    cat.FailedControls.ToString(), 
                    cat.ManualControls.ToString(),
                    $"{cat.CompliancePercentage}%" 
                });
            }

            AddParagraph(body, "");

            // Failed Controls
            var failedControls = data.Controls.Where(c => c.Status == CisControlStatus.Fail).OrderBy(c => c.ControlId).ToList();
            if (failedControls.Any())
            {
                AddHeading(body, $"Failed Controls ({failedControls.Count})", 2);
                foreach (var control in failedControls)
                {
                    AddControlSection(body, control);
                }
            }

            // Manual Review Controls
            var manualControls = data.Controls.Where(c => c.Status == CisControlStatus.Manual).OrderBy(c => c.ControlId).ToList();
            if (manualControls.Any())
            {
                AddHeading(body, $"Controls Requiring Manual Review ({manualControls.Count})", 2);
                foreach (var control in manualControls)
                {
                    AddControlSection(body, control);
                }
            }

            // Passed Controls
            var passedControls = data.Controls.Where(c => c.Status == CisControlStatus.Pass).OrderBy(c => c.ControlId).ToList();
            if (passedControls.Any())
            {
                AddHeading(body, $"Passed Controls ({passedControls.Count})", 2);
                foreach (var control in passedControls)
                {
                    AddControlSectionCompact(body, control);
                }
            }

            // Footer
            AddParagraph(body, "");
            AddParagraph(body, "This report was automatically generated by M365 Dashboard using the CIS Microsoft 365 Foundations Benchmark v6.0.0.", true);

            mainPart.Document.Save();
        }

        return stream.ToArray();
    }

    private void AddControlSection(Body body, CisControlResult control)
    {
        AddParagraph(body, $"{control.ControlId} - {control.Title}", false, control.Status == CisControlStatus.Fail);
        AddParagraph(body, $"Level: {control.Level} | Profile: {control.Profile} | Status: {control.Status}", true);
        AddParagraph(body, control.Description, true);
        AddParagraph(body, $"Current Value: {control.CurrentValue}");
        AddParagraph(body, $"Expected Value: {control.ExpectedValue}");
        if (control.Status == CisControlStatus.Fail || control.Status == CisControlStatus.Manual)
        {
            AddParagraph(body, $"Remediation: {control.Remediation}");
        }
        if (control.AffectedItems?.Any() == true)
        {
            AddParagraph(body, $"Affected Items: {string.Join(", ", control.AffectedItems.Take(5))}{(control.AffectedItems.Count > 5 ? $" (+{control.AffectedItems.Count - 5} more)" : "")}");
        }
        AddParagraph(body, "");
    }

    private void AddControlSectionCompact(Body body, CisControlResult control)
    {
        AddParagraph(body, $"✓ {control.ControlId} - {control.Title}");
        AddParagraph(body, $"   {control.CurrentValue}", true);
    }

    private void AddStyles(MainDocumentPart mainPart)
    {
        var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
        var styles = new Styles();
        
        var heading1 = new Style() { Type = StyleValues.Paragraph, StyleId = "Heading1" };
        heading1.Append(new StyleName() { Val = "Heading 1" });
        heading1.Append(new StyleRunProperties(new Bold(), new FontSize() { Val = "48" }, new Color() { Val = "0078D4" }));
        styles.Append(heading1);

        var heading2 = new Style() { Type = StyleValues.Paragraph, StyleId = "Heading2" };
        heading2.Append(new StyleName() { Val = "Heading 2" });
        heading2.Append(new StyleRunProperties(new Bold(), new FontSize() { Val = "28" }, new Color() { Val = "323130" }));
        styles.Append(heading2);

        stylesPart.Styles = styles;
        stylesPart.Styles.Save();
    }

    private void AddHeading(Body body, string text, int level, bool isTitle = false)
    {
        var para = new Paragraph();
        var run = new Run();
        var runProps = new RunProperties();
        
        runProps.Append(new Bold());
        runProps.Append(new FontSize() { Val = level == 1 ? (isTitle ? "56" : "48") : "28" });
        runProps.Append(new Color() { Val = level == 1 ? "0078D4" : "323130" });
        
        run.Append(runProps);
        run.Append(new Text(text));
        para.Append(run);
        body.Append(para);
    }

    private void AddParagraph(Body body, string text, bool isItalic = false, bool isRed = false)
    {
        var para = new Paragraph();
        var run = new Run();
        var runProps = new RunProperties();
        
        if (isItalic) runProps.Append(new Italic());
        if (isRed) runProps.Append(new Color() { Val = "D13438" });
        runProps.Append(new FontSize() { Val = "22" });
        
        run.Append(runProps);
        run.Append(new Text(text));
        para.Append(run);
        body.Append(para);
    }

    private Table CreateTable(Body body, string[] headers)
    {
        var table = new Table();
        var tableProps = new TableProperties(
            new TableBorders(
                new TopBorder() { Val = BorderValues.Single, Size = 4, Color = "DDDDDD" },
                new BottomBorder() { Val = BorderValues.Single, Size = 4, Color = "DDDDDD" },
                new LeftBorder() { Val = BorderValues.Single, Size = 4, Color = "DDDDDD" },
                new RightBorder() { Val = BorderValues.Single, Size = 4, Color = "DDDDDD" },
                new InsideHorizontalBorder() { Val = BorderValues.Single, Size = 4, Color = "DDDDDD" },
                new InsideVerticalBorder() { Val = BorderValues.Single, Size = 4, Color = "DDDDDD" }
            ),
            new TableWidth() { Type = TableWidthUnitValues.Pct, Width = "5000" }
        );
        table.Append(tableProps);

        var headerRow = new TableRow();
        foreach (var header in headers)
        {
            var cell = new TableCell();
            cell.Append(new TableCellProperties(new Shading() { Fill = "0078D4" }));
            var para = new Paragraph();
            var run = new Run();
            run.Append(new RunProperties(new Bold(), new Color() { Val = "FFFFFF" }, new FontSize() { Val = "20" }));
            run.Append(new Text(header));
            para.Append(run);
            cell.Append(para);
            headerRow.Append(cell);
        }
        table.Append(headerRow);
        body.Append(table);
        body.Append(new Paragraph());
        return table;
    }

    private void AddTableRow(Table table, string[] values)
    {
        var row = new TableRow();
        foreach (var value in values)
        {
            var cell = new TableCell();
            var para = new Paragraph();
            var run = new Run();
            run.Append(new RunProperties(new FontSize() { Val = "20" }));
            run.Append(new Text(value));
            para.Append(run);
            cell.Append(para);
            row.Append(cell);
        }
        var lastRow = table.Elements<TableRow>().LastOrDefault();
        if (lastRow != null) lastRow.InsertAfterSelf(row);
        else table.Append(row);
    }

    private string GenerateHtmlReport(CisBenchmarkResult data)
    {
        var failedControls = data.Controls.Where(c => c.Status == CisControlStatus.Fail).OrderBy(c => c.ControlId).ToList();
        var manualControls = data.Controls.Where(c => c.Status == CisControlStatus.Manual).OrderBy(c => c.ControlId).ToList();
        var passedControls = data.Controls.Where(c => c.Status == CisControlStatus.Pass).OrderBy(c => c.ControlId).ToList();

        return $@"<!DOCTYPE html>
<html>
<head>
    <title>CIS Microsoft 365 Benchmark Report</title>
    <style>
        body {{ font-family: 'Segoe UI', Arial, sans-serif; margin: 40px; background: #f5f5f5; }}
        .container {{ max-width: 1200px; margin: 0 auto; background: white; padding: 40px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }}
        h1 {{ color: #0078D4; border-bottom: 3px solid #0078D4; padding-bottom: 10px; }}
        h2 {{ color: #323130; margin-top: 30px; }}
        .summary-grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(150px, 1fr)); gap: 15px; margin: 20px 0; }}
        .summary-card {{ background: #f8f9fa; border-radius: 8px; padding: 20px; text-align: center; }}
        .summary-card.pass {{ border-left: 4px solid #107C10; }}
        .summary-card.fail {{ border-left: 4px solid #D13438; }}
        .summary-card.manual {{ border-left: 4px solid #FFB900; }}
        .summary-value {{ font-size: 2em; font-weight: bold; color: #323130; }}
        .summary-label {{ color: #666; font-size: 0.9em; }}
        table {{ width: 100%; border-collapse: collapse; margin: 20px 0; }}
        th {{ background: #0078D4; color: white; padding: 12px; text-align: left; }}
        td {{ padding: 10px; border-bottom: 1px solid #ddd; }}
        tr:hover {{ background: #f5f5f5; }}
        .control {{ background: #fff; border: 1px solid #ddd; border-radius: 8px; padding: 20px; margin: 15px 0; }}
        .control.fail {{ border-left: 4px solid #D13438; }}
        .control.manual {{ border-left: 4px solid #FFB900; }}
        .control.pass {{ border-left: 4px solid #107C10; }}
        .control-header {{ display: flex; justify-content: space-between; align-items: center; }}
        .control-id {{ font-weight: bold; color: #0078D4; }}
        .control-title {{ font-size: 1.1em; font-weight: 600; margin: 10px 0; }}
        .badge {{ display: inline-block; padding: 4px 12px; border-radius: 4px; font-size: 0.8em; font-weight: 500; }}
        .badge-pass {{ background: #DFF6DD; color: #107C10; }}
        .badge-fail {{ background: #FDE7E9; color: #D13438; }}
        .badge-manual {{ background: #FFF4CE; color: #8A6914; }}
        .badge-l1 {{ background: #E8F4FC; color: #0078D4; }}
        .badge-l2 {{ background: #F3E5F5; color: #7B1FA2; }}
        .control-details {{ margin-top: 15px; font-size: 0.95em; }}
        .control-details dt {{ font-weight: 600; color: #323130; margin-top: 10px; }}
        .control-details dd {{ margin-left: 0; color: #666; }}
        .affected-items {{ background: #FDE7E9; padding: 10px; border-radius: 4px; margin-top: 10px; }}
        .score-ring {{ width: 150px; height: 150px; margin: 20px auto; }}
        .category-row {{ transition: background 0.2s; }}
        .footer {{ margin-top: 40px; padding-top: 20px; border-top: 1px solid #ddd; color: #666; font-size: 0.9em; }}
    </style>
</head>
<body>
    <div class='container'>
        <h1>CIS Microsoft 365 Foundations Benchmark Report</h1>
        <p><strong>Benchmark Version:</strong> {data.BenchmarkVersion}</p>
        <p><strong>Generated:</strong> {data.GeneratedAt:dd MMMM yyyy HH:mm} UTC</p>
        <p><strong>Tenant:</strong> {data.TenantName}</p>

        <h2>Executive Summary</h2>
        <div class='summary-grid'>
            <div class='summary-card'>
                <div class='summary-value'>{data.TotalControls}</div>
                <div class='summary-label'>Total Controls</div>
            </div>
            <div class='summary-card pass'>
                <div class='summary-value' style='color:#107C10'>{data.PassedControls}</div>
                <div class='summary-label'>Passed</div>
            </div>
            <div class='summary-card fail'>
                <div class='summary-value' style='color:#D13438'>{data.FailedControls}</div>
                <div class='summary-label'>Failed</div>
            </div>
            <div class='summary-card manual'>
                <div class='summary-value' style='color:#8A6914'>{data.ManualControls}</div>
                <div class='summary-label'>Manual Review</div>
            </div>
            <div class='summary-card'>
                <div class='summary-value' style='color:#0078D4'>{data.CompliancePercentage}%</div>
                <div class='summary-label'>Compliance Score</div>
            </div>
        </div>

        <h2>Compliance by Category</h2>
        <table>
            <tr><th>Category</th><th>Passed</th><th>Failed</th><th>Manual</th><th>Score</th></tr>
            {string.Join("", data.Categories.Select(c => $@"
            <tr class='category-row'>
                <td>{c.CategoryName}</td>
                <td style='color:#107C10'>{c.PassedControls}</td>
                <td style='color:#D13438'>{c.FailedControls}</td>
                <td style='color:#8A6914'>{c.ManualControls}</td>
                <td><strong>{c.CompliancePercentage}%</strong></td>
            </tr>"))}
        </table>

        {(failedControls.Any() ? $@"
        <h2>Failed Controls ({failedControls.Count})</h2>
        {string.Join("", failedControls.Select(c => GenerateControlHtml(c)))}" : "")}

        {(manualControls.Any() ? $@"
        <h2>Controls Requiring Manual Review ({manualControls.Count})</h2>
        {string.Join("", manualControls.Select(c => GenerateControlHtml(c)))}" : "")}

        {(passedControls.Any() ? $@"
        <h2>Passed Controls ({passedControls.Count})</h2>
        {string.Join("", passedControls.Select(c => GenerateControlHtmlCompact(c)))}" : "")}

        <div class='footer'>
            <p>This report was automatically generated by M365 Dashboard using the CIS Microsoft 365 Foundations Benchmark v6.0.0.</p>
            <p>For the complete benchmark documentation, visit <a href='https://www.cisecurity.org/benchmark/microsoft_365'>cisecurity.org</a></p>
        </div>
    </div>
</body>
</html>";
    }

    private string GenerateControlHtml(CisControlResult c)
    {
        var statusClass = c.Status.ToString().ToLower();
        var statusBadge = c.Status switch
        {
            CisControlStatus.Pass => "badge-pass",
            CisControlStatus.Fail => "badge-fail",
            _ => "badge-manual"
        };
        var levelBadge = c.Level == CisLevel.L1 ? "badge-l1" : "badge-l2";

        return $@"
        <div class='control {statusClass}'>
            <div class='control-header'>
                <span class='control-id'>{c.ControlId}</span>
                <div>
                    <span class='badge {levelBadge}'>{c.Level}</span>
                    <span class='badge {statusBadge}'>{c.Status}</span>
                </div>
            </div>
            <div class='control-title'>{c.Title}</div>
            <p style='color:#666'>{c.Description}</p>
            <dl class='control-details'>
                <dt>Current Value</dt>
                <dd>{c.CurrentValue}</dd>
                <dt>Expected Value</dt>
                <dd>{c.ExpectedValue}</dd>
                {(c.Status != CisControlStatus.Pass ? $@"
                <dt>Remediation</dt>
                <dd>{c.Remediation}</dd>" : "")}
            </dl>
            {(c.AffectedItems?.Any() == true ? $@"
            <div class='affected-items'>
                <strong>Affected Items:</strong> {string.Join(", ", c.AffectedItems.Take(5))}{(c.AffectedItems.Count > 5 ? $" (+{c.AffectedItems.Count - 5} more)" : "")}
            </div>" : "")}
        </div>";
    }

    private string GenerateControlHtmlCompact(CisControlResult c)
    {
        return $@"
        <div class='control pass' style='padding:12px'>
            <span class='control-id'>{c.ControlId}</span> - {c.Title}
            <div style='color:#666;font-size:0.9em;margin-top:5px'>{c.CurrentValue}</div>
        </div>";
    }
}
