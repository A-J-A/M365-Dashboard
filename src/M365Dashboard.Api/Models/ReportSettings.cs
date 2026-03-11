namespace M365Dashboard.Api.Models;

/// <summary>
/// Report branding settings
/// </summary>
public class ReportSettings
{
    public string CompanyName { get; set; } = "M365 Dashboard";
    public string ReportTitle { get; set; } = "Microsoft 365 Security Assessment";
    public string? LogoBase64 { get; set; }
    public string? LogoContentType { get; set; }
    public string PrimaryColor { get; set; } = "#0078d4";
    public string AccentColor { get; set; } = "#e07c3a";
    public bool ShowInfoGraphics { get; set; } = true;
    public string? FooterText { get; set; }
    public DateTime UpdatedAt { get; set; } = DateTime.UtcNow;
}
