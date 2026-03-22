namespace M365Dashboard.Api.Models;

/// <summary>
/// A single quote/statistic displayed as a full-page infographic in the executive report.
/// </summary>
public class ReportQuote
{
    public string BigNumber { get; set; } = "";
    public string Line1     { get; set; } = "";
    public string Line2     { get; set; } = "";
    public string Source    { get; set; } = "";
    public bool   Enabled   { get; set; } = true;
}

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
    public bool ShowQuotes { get; set; } = true;
    public string? FooterText { get; set; }
    public DateTime UpdatedAt { get; set; } = DateTime.UtcNow;

    /// <summary>
    /// Pool of quotes to choose from. 3 are selected at random each time a report is generated.
    /// </summary>
    public List<ReportQuote> Quotes { get; set; } = DefaultQuotes();

    public static List<ReportQuote> DefaultQuotes() => new()
    {
        new() { BigNumber = "99%",  Line1 = "of breaches could be mitigated",          Line2 = "with strong passwords and MFA",                Source = "Source: Microsoft Security Report",          Enabled = true },
        new() { BigNumber = "84%",  Line1 = "of businesses fell victim",                Line2 = "to phishing attacks in 2024",                    Source = "Source: Cyber Security Breaches Survey",      Enabled = true },
        new() { BigNumber = "31%",  Line1 = "of all breaches over the past",            Line2 = "10 years involved stolen credentials",            Source = "Source: Verizon DBIR",                       Enabled = true },
        new() { BigNumber = "300%", Line1 = "increase in reported cyber incidents",     Line2 = "since the start of remote working",               Source = "Source: NCSC Annual Review",                 Enabled = true },
        new() { BigNumber = "4.5M", Line1 = "average cost of a data breach in 2023",   Line2 = "a record high for the 13th consecutive year",      Source = "Source: IBM Cost of a Data Breach Report",   Enabled = true },
        new() { BigNumber = "11s",  Line1 = "a business falls victim to ransomware",   Line2 = "every 11 seconds globally",                        Source = "Source: Cybersecurity Ventures",             Enabled = true },
        new() { BigNumber = "74%",  Line1 = "of all breaches include",                  Line2 = "a human element",                                  Source = "Source: Verizon DBIR 2023",                  Enabled = true },
        new() { BigNumber = "85%",  Line1 = "of organisations have experienced",        Line2 = "at least one cloud data breach",                    Source = "Source: Thales Cloud Security Study",        Enabled = true },
        new() { BigNumber = "50%",  Line1 = "of SMBs have suffered a cyberattack",     Line2 = "and 60% close within 6 months",                    Source = "Source: SCORE.org",                          Enabled = true },
        new() { BigNumber = "98%",  Line1 = "of cyberattacks can be prevented",         Line2 = "by implementing basic cyber hygiene",               Source = "Source: Microsoft Digital Defence Report",   Enabled = true },
    };
}
