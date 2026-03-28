using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;
using M365Dashboard.Api.Models;
using M365Dashboard.Api.Controllers;

namespace M365Dashboard.Api.Services;

/// <summary>
/// PDF report generator.
/// Each section is a separate container.Page() call to avoid QuestPDF conflicting size-constraint errors.
/// </summary>
public class PdfReportGenerator : IDocument
{
    private ExecutiveReportData _data = null!;
    private ReportSettings _settings = null!;
    private byte[]? _logoBytes;
    private List<ReportQuote> _selectedQuotes = new();

    private string _primary = "#1E3A5F";
    private string _accent  = "#E07C3A";
    private const string Compliant    = "#107C6C";
    private const string NonCompliant = "#6B7280";
    private const string Warn         = "#F59E0B";
    private const string Crit         = "#DC2626";
    private const string BodyText     = "#374151";
    private const string LightGray    = "#F9FAFB";
    private const string BorderCol    = "#E5E7EB";

    // Fonts available on Debian ASP.NET 8 image
    private static readonly string[] Fonts = { "Liberation Sans", "DejaVu Sans", "Arial" };

    public byte[] GenerateReport(ExecutiveReportData data, ReportSettings settings)
    {
        QuestPDF.Settings.License = LicenseType.Community;
        _data     = data;
        _settings = settings;

        if (!string.IsNullOrEmpty(settings.PrimaryColor))
            _primary = settings.PrimaryColor.StartsWith("#") ? settings.PrimaryColor : $"#{settings.PrimaryColor}";
        if (!string.IsNullOrEmpty(settings.AccentColor))
            _accent = settings.AccentColor.StartsWith("#") ? settings.AccentColor : $"#{settings.AccentColor}";

        if (!string.IsNullOrEmpty(settings.LogoBase64))
        {
            try { _logoBytes = Convert.FromBase64String(settings.LogoBase64); }
            catch { _logoBytes = null; }
        }

        _selectedQuotes = PickRandomQuotes(settings, 3);
        return Document.Create(Compose).GeneratePdf();
    }

    public void Compose(IDocumentContainer c)
    {
        void ContentPage(IDocumentContainer container, Action<ColumnDescriptor> body) =>
            container.Page(p =>
            {
                p.Size(PageSizes.A4);
                p.MarginHorizontal(40);
                p.MarginTop(30);
                p.MarginBottom(20);
                p.DefaultTextStyle(s => s.FontFamily(Fonts).FontColor(BodyText));
                p.Content().Column(body);
            });

        void InfoPage(IDocumentContainer container, ReportQuote q) =>
            container.Page(p =>
            {
                p.Size(PageSizes.A4);
                p.Margin(0);
                p.DefaultTextStyle(s => s.FontFamily(Fonts));
                p.Content().Background(_primary).Padding(60).Column(col =>
                {
                    col.Item().Height(80);
                    col.Item().AlignCenter().Text(q.BigNumber).FontSize(110).FontColor(Colors.White);
                    col.Item().Height(20);
                    col.Item().AlignCenter().Text(q.Line1).FontSize(22).FontColor(Colors.White);
                    col.Item().Height(6);
                    col.Item().AlignCenter().Text(q.Line2).FontSize(22).Bold().FontColor(_accent);
                    col.Item().Height(60);
                    col.Item().AlignCenter().Text(q.Source).FontSize(9).Italic().FontColor(Colors.Grey.Lighten2);
                });
            });

        // 1. Cover
        c.Page(p =>
        {
            p.Size(PageSizes.A4);
            p.Margin(0);
            p.DefaultTextStyle(s => s.FontFamily(Fonts));
            p.Content().Column(CoverPage);
        });

        // 2. Executive Summary
        ContentPage(c, ExecutiveSummaryPage);

        // 3. Infographic 1
        if (_settings.ShowInfoGraphics && _settings.ShowQuotes && _selectedQuotes.Count > 0)
            InfoPage(c, _selectedQuotes[0]);

        // 4. Security Metrics
        ContentPage(c, SecurityMetricsPage);

        // 5. Infographic 2
        if (_settings.ShowInfoGraphics && _settings.ShowQuotes && _selectedQuotes.Count > 1)
            InfoPage(c, _selectedQuotes[1]);

        // 6. Device Details
        if (HasDevices())
            ContentPage(c, DeviceDetailsPage);

        // 7. Infographic 3
        if (_settings.ShowInfoGraphics && _settings.ShowQuotes && _selectedQuotes.Count > 2)
            InfoPage(c, _selectedQuotes[2]);

        // 8. User Details
        if (_data.UserSignInDetails?.Any() == true)
            ContentPage(c, UserDetailsPage);

        // 9. Deleted Users
        if (_data.DeletedUsersInPeriod?.Any() == true)
            ContentPage(c, DeletedUsersPage);

        // 10. Mailbox Storage Details
        if (_data.MailboxDetails?.Any() == true)
            ContentPage(c, MailboxDetailsPage);

        // 11. Domain Security
        if (_data.DomainSecuritySummary != null)
            ContentPage(c, DomainSecurityPage);

        // 12. App Secrets & Certificates
        if (_data.AppCredentialStatus != null &&
            (_data.AppCredentialStatus.ExpiringSecrets.Any() || _data.AppCredentialStatus.ExpiredSecrets.Any() ||
             _data.AppCredentialStatus.ExpiringCertificates.Any() || _data.AppCredentialStatus.ExpiredCertificates.Any()))
            ContentPage(c, AppCredentialsPage);

        // 13. Sign-in Locations Map
        if (_data.SignInMapImageBytes != null && _data.SignInLocations?.Any() == true)
            ContentPage(c, SignInLocationsPage);
    }

    private void CoverPage(ColumnDescriptor col)
    {
        var (line1, line2) = ParseTitle(_settings.ReportTitle);

        col.Item().Height(400).Background(_primary).Padding(50).Column(h =>
        {
            h.Item().Height(70);
            h.Item().Text(line1).FontSize(20).FontColor(_accent).Bold();
            h.Item().Height(12);
            h.Item().Text(line2).FontSize(38).FontColor(Colors.White);
        });

        col.Item().Background(Colors.White).Padding(50).Column(info =>
        {
            info.Item().Height(15);
            info.Item().AlignCenter().Text(_settings.CompanyName).FontSize(20).Bold().FontColor(_primary);
            info.Item().Height(8);
            info.Item().AlignCenter().Text($"Generated {_data.GeneratedAt:d MMMM yyyy}")
                .FontSize(11).FontColor(Colors.Grey.Medium);
            if (_logoBytes != null)
            {
                info.Item().Height(15);
                info.Item().AlignCenter().MaxHeight(55).Image(_logoBytes).FitHeight();
            }
        });
    }

    private void ExecutiveSummaryPage(ColumnDescriptor col)
    {
        col.Item().Text("Executive Summary").FontSize(28).FontColor(_primary);
        col.Item().Height(18);

        col.Item().Row(row =>
        {
            KpiCard(row.RelativeItem(), "Total Users",
                $"{_data.UserStats?.TotalUsers ?? 0}",
                $"inc. {_data.UserStats?.GuestUsers ?? 0} guests");
            row.ConstantItem(10);
            KpiCard(row.RelativeItem(), "MFA Registered",
                $"{_data.UserStats?.MfaRegistered ?? 0}",
                $"{_data.UserStats?.MfaNotRegistered ?? 0} not registered");
            row.ConstantItem(10);
            KpiCard(row.RelativeItem(), "Secure Score",
                _data.SecureScore != null ? $"{_data.SecureScore.PercentageScore:F0}%" : "N/A",
                _data.SecureScore != null ? $"{_data.SecureScore.CurrentScore:N0}/{_data.SecureScore.MaxScore:N0}" : "");
        });

        col.Item().Height(18);
        col.Item().Text(
            $"This {_settings.ReportTitle} for {_data.GeneratedAt:MMMM yyyy} provides a comprehensive overview of your " +
            $"Microsoft 365 security posture, covering Entra ID, Exchange Online, Intune, SharePoint, and Teams.")
            .FontSize(10);

        col.Item().Height(22);
        col.Item().Text("User Summary").FontSize(16).FontColor(_primary).Bold();
        col.Item().Height(8);
        col.Item().Table(t =>
        {
            TwoCols(t);
            TH(t, "Metric", "Count");
            TR(t, "Total Users",        $"{_data.UserStats?.TotalUsers ?? 0}");
            TR(t, "Guest Users",        $"{_data.UserStats?.GuestUsers ?? 0}");
            TR(t, "MFA Registered",     $"{_data.UserStats?.MfaRegistered ?? 0}", Compliant);
            TR(t, "MFA Not Registered", $"{_data.UserStats?.MfaNotRegistered ?? 0}",
                (_data.UserStats?.MfaNotRegistered ?? 0) > 0 ? Warn : null);
            TR(t, "Risky Users", $"{_data.RiskyUsersCount}",
                _data.RiskyUsersCount > 0 ? Crit : null);
        });
    }

    private void SecurityMetricsPage(ColumnDescriptor col)
    {
        if (_data.SecureScore != null)
        {
            col.Item().Text("Microsoft Secure Score").FontSize(16).FontColor(_primary).Bold();
            col.Item().Height(4);
            col.Item().Text("Microsoft Secure Score measures your organisation's security posture across Microsoft 365. " +
                "A higher score indicates more security controls are in place. Scores above 70% are considered good, " +
                "with the maximum achievable score depending on your licensed services.").FontSize(8).FontColor(Colors.Grey.Darken1);
            col.Item().Height(8);
            col.Item().Table(t =>
            {
                TwoCols(t);
                TH(t, "Metric", "Value");
                TR(t, "Current Score", $"{_data.SecureScore.CurrentScore:N0} / {_data.SecureScore.MaxScore:N0}");
                TR(t, "Percentage",    $"{_data.SecureScore.PercentageScore:F1}%",
                    _data.SecureScore.PercentageScore >= 70 ? Compliant :
                    _data.SecureScore.PercentageScore >= 50 ? Warn : Crit);
            });
            col.Item().Height(18);
        }

        if (_data.DeviceStats != null)
        {
            col.Item().Text("Intune Managed Devices").FontSize(16).FontColor(_primary).Bold();
            col.Item().Height(4);
            col.Item().Text("Microsoft Intune is the mobile device and application management solution within Microsoft 365. " +
                "Compliant devices meet the security requirements defined in your compliance policies, such as requiring " +
                "encryption, PIN protection, and up-to-date operating systems.").FontSize(8).FontColor(Colors.Grey.Darken1);
            col.Item().Height(8);
            col.Item().Table(t =>
            {
                TwoCols(t);
                TH(t, "Platform", "Count");
                TR(t, "Total",         $"{_data.DeviceStats.TotalDevices}");
                TR(t, "Windows",       $"{_data.DeviceStats.WindowsDevices}");
                TR(t, "macOS",         $"{_data.DeviceStats.MacOsDevices}");
                TR(t, "iOS/iPadOS",    $"{_data.DeviceStats.IosDevices}");
                TR(t, "Android",       $"{_data.DeviceStats.AndroidDevices}");
                TR(t, "Compliant",     $"{_data.DeviceStats.CompliantDevices}", Compliant);
                TR(t, "Non-Compliant", $"{_data.DeviceStats.NonCompliantDevices}",
                    _data.DeviceStats.NonCompliantDevices > 0 ? Crit : null);
                TR(t, "Compliance %",  $"{_data.DeviceStats.ComplianceRate:F1}%",
                    _data.DeviceStats.ComplianceRate >= 90 ? Compliant : Warn);
            });
            col.Item().Height(18);
        }

        if (_data.DefenderStats != null)
        {
            col.Item().Text("Microsoft Defender for Endpoint").FontSize(16).FontColor(_primary).Bold();
            col.Item().Height(4);
            col.Item().Text("Microsoft Defender for Endpoint provides advanced threat protection for devices. " +
                "The Exposure Score reflects how vulnerable your organisation is to threats — a lower score is better. " +
                "Low (0-30) indicates good posture, Medium (31-70) needs attention, and High (71-100) requires immediate action.").FontSize(8).FontColor(Colors.Grey.Darken1);
            col.Item().Height(8);

            // Exposure score gauge (visual bar)
            if (_data.DefenderStats.ExposureScoreNumeric.HasValue)
            {
                var score = _data.DefenderStats.ExposureScoreNumeric.Value;
                var gaugeColor = score <= 30 ? Compliant : score <= 70 ? Warn : Crit;
                var gaugeLabel = score <= 30 ? "Low" : score <= 70 ? "Medium" : "High";
                var fillPct = (float)(score / 100.0);

                col.Item().Column(g =>
                {
                    g.Item().Text($"Exposure Score: {gaugeLabel} ({score:F1})").FontSize(9).FontColor(gaugeColor).Bold();
                    g.Item().Height(4);
                    // Gauge bar using a row: filled portion + empty remainder
                    g.Item().Height(12).Row(bar =>
                    {
                        if (fillPct > 0)
                            bar.RelativeItem((int)Math.Round(fillPct * 100)).Height(12).Background(gaugeColor);
                        if (fillPct < 1)
                            bar.RelativeItem((int)Math.Round((1 - fillPct) * 100)).Height(12).Background("#E5E7EB");
                    });
                });
                col.Item().Height(12);
            }

            // Only show exposure level in table, rest removed
            col.Item().Table(t =>
            {
                TwoCols(t);
                TH(t, "Metric", "Value");
                TR(t, "Exposure Level", _data.DefenderStats.ExposureScore ?? "N/A",
                    _data.DefenderStats.ExposureScore switch { "Low" => Compliant, "High" => Crit, _ => Warn });
            });
        }
    }

    private void DeviceDetailsPage(ColumnDescriptor col)
    {
        col.Item().Text("Intune Managed Devices").FontSize(28).FontColor(_primary);
        col.Item().Height(6);
        col.Item().Text("Detailed view of all Intune-managed devices with OS version status. " +
            "Each device's OS is compared against the latest known release sourced from endoflife.date. " +
            "Keeping devices up to date is critical for patching known vulnerabilities and maintaining compliance.").FontSize(8).FontColor(Colors.Grey.Darken1);
        col.Item().Height(16);

        // Colour-coded legend
        col.Item().Row(row =>
        {
            row.AutoItem().Text("● Current").FontSize(8).FontColor(Compliant);
            row.ConstantItem(16);
            row.AutoItem().Text("● Update available").FontSize(8).FontColor(Warn);
            row.ConstantItem(16);
            row.AutoItem().Text("● Critical / EOL").FontSize(8).FontColor(Crit);
        });
        col.Item().Height(14);

        void DevSection<T>(string title, IEnumerable<T> devices,
            Func<T, string> name, Func<T, string> osVer, Func<T, VersionStatus> status,
            Func<T, string> statusMsg, Func<T, string?> latest,
            Func<T, string> compliance, Func<T, string> checkIn)
        {
            col.Item().Text(title).FontSize(14).FontColor(_primary).Bold();
            col.Item().Height(6);
            col.Item().Table(t =>
            {
                t.ColumnsDefinition(c =>
                {
                    c.RelativeColumn(2.5f); c.RelativeColumn(1.8f); c.RelativeColumn(1.5f);
                    c.RelativeColumn(1.5f); c.RelativeColumn(1.5f); c.RelativeColumn(1.5f);
                });
                foreach (var h in new[] { "Device", "Current OS", "Latest OS", "Update Status", "Compliance", "Last Check-in" })
                    t.Cell().Background(LightGray).Border(1).BorderColor(BorderCol)
                        .Padding(4).Text(h).FontSize(7).Bold();
                foreach (var d in devices)
                {
                    var st = status(d);
                    var stColor = st switch
                    {
                        VersionStatus.Current  => Compliant,
                        VersionStatus.Warning  => Warn,
                        VersionStatus.Critical => Crit,
                        _                      => BodyText
                    };
                    t.Cell().Border(1).BorderColor(BorderCol).Padding(4).Text(name(d)).FontSize(7);
                    t.Cell().Border(1).BorderColor(BorderCol).Padding(4).Text(osVer(d)).FontSize(7);
                    t.Cell().Border(1).BorderColor(BorderCol).Padding(4)
                        .Text(latest(d) ?? "-").FontSize(7).FontColor(Colors.Grey.Darken1);
                    t.Cell().Border(1).BorderColor(BorderCol).Padding(4)
                        .Text(statusMsg(d)).FontSize(7).FontColor(stColor);
                    t.Cell().Border(1).BorderColor(BorderCol).Padding(4)
                        .Text(compliance(d)).FontSize(7)
                        .FontColor(compliance(d).Equals("Compliant", StringComparison.OrdinalIgnoreCase) ? Compliant :
                                   compliance(d).Equals("Non-Compliant", StringComparison.OrdinalIgnoreCase) ? Crit : BodyText);
                    t.Cell().Border(1).BorderColor(BorderCol).Padding(4).Text(checkIn(d)).FontSize(7);
                }
            });
            col.Item().Height(14);
        }

        if (_data.DeviceDetails?.WindowsDevices?.Any() == true)
            DevSection($"Windows ({_data.DeviceDetails.WindowsDevices.Count})",
                _data.DeviceDetails.WindowsDevices.Take(20),
                d => d.DeviceName ?? "-", d => d.OsVersion ?? "-",
                d => d.OsVersionStatus, d => d.OsVersionStatusMessage ?? "-",
                d => d.LatestVersion,
                d => d.ComplianceState ?? "-", d => d.LastCheckIn?.ToString("dd MMM yyyy") ?? "Never");

        if (_data.DeviceDetails?.MacDevices?.Any() == true)
            DevSection($"macOS ({_data.DeviceDetails.MacDevices.Count})",
                _data.DeviceDetails.MacDevices.Take(10),
                d => d.DeviceName ?? "-", d => d.OsVersion ?? "-",
                d => d.OsVersionStatus, d => d.OsVersionStatusMessage ?? "-",
                d => d.LatestVersion,
                d => d.ComplianceState ?? "-", d => d.LastCheckIn?.ToString("dd MMM yyyy") ?? "Never");

        if (_data.DeviceDetails?.IosDevices?.Any() == true)
            DevSection($"iOS/iPadOS ({_data.DeviceDetails.IosDevices.Count})",
                _data.DeviceDetails.IosDevices.Take(10),
                d => d.DeviceName ?? "-", d => d.OsVersion ?? "-",
                d => d.OsVersionStatus, d => d.OsVersionStatusMessage ?? "-",
                d => d.LatestVersion,
                d => d.ComplianceState ?? "-", d => d.LastCheckIn?.ToString("dd MMM yyyy") ?? "Never");

        if (_data.DeviceDetails?.AndroidDevices?.Any() == true)
            DevSection($"Android ({_data.DeviceDetails.AndroidDevices.Count})",
                _data.DeviceDetails.AndroidDevices.Take(10),
                d => d.DeviceName ?? "-", d => d.OsVersion ?? "-",
                d => d.OsVersionStatus, d => d.OsVersionStatusMessage ?? "-",
                d => d.LatestVersion,
                d => d.ComplianceState ?? "-", d => d.LastCheckIn?.ToString("dd MMM yyyy") ?? "Never");
    }

    private void UserDetailsPage(ColumnDescriptor col)
    {
        col.Item().Text($"User Sign-in & MFA ({_data.UserSignInDetails!.Count} users)")
            .FontSize(16).FontColor(_primary).Bold();
        col.Item().Height(4);
        col.Item().Text("Multi-Factor Authentication (MFA) requires users to verify their identity using a second method beyond their password, " +
            "significantly reducing the risk of account compromise. All users should have MFA registered. " +
            "Last sign-in dates help identify inactive accounts that may pose a security risk.").FontSize(8).FontColor(Colors.Grey.Darken1);
        col.Item().Height(8);
        col.Item().Table(t =>
        {
            t.ColumnsDefinition(c =>
            {
                c.RelativeColumn(2); c.RelativeColumn(2.5f); c.RelativeColumn(1.5f);
                c.RelativeColumn(2); c.RelativeColumn(0.8f); c.RelativeColumn(0.8f);
            });
            foreach (var h in new[] { "Name", "Email", "Last Sign-in", "MFA Method", "MFA", "Enabled" })
                t.Cell().Background(LightGray).Border(1).BorderColor(BorderCol)
                    .Padding(5).Text(h).FontSize(8).Bold();
            foreach (var u in _data.UserSignInDetails.Take(35))
            {
                t.Cell().Border(1).BorderColor(BorderCol).Padding(5).Text(u.DisplayName ?? "-").FontSize(8);
                t.Cell().Border(1).BorderColor(BorderCol).Padding(5).Text(u.UserPrincipalName ?? "-").FontSize(7);
                t.Cell().Border(1).BorderColor(BorderCol).Padding(5)
                    .Text(u.LastInteractiveSignIn?.ToString("dd MMM yyyy") ?? "Never").FontSize(8);
                t.Cell().Border(1).BorderColor(BorderCol).Padding(5)
                    .Text(u.DefaultMfaMethod ?? "None").FontSize(8);
                t.Cell().Border(1).BorderColor(BorderCol).Padding(5)
                    .Text(u.IsMfaRegistered ? "Yes" : "No").FontSize(8)
                    .FontColor(u.IsMfaRegistered ? Compliant : Crit);
                t.Cell().Border(1).BorderColor(BorderCol).Padding(5)
                    .Text(u.AccountEnabled ? "Yes" : "No").FontSize(8)
                    .FontColor(u.AccountEnabled ? Compliant : NonCompliant);
            }
        });
    }

    private void DomainSecurityPage(ColumnDescriptor col)
    {
        col.Item().Text("Domain Email Security").FontSize(16).FontColor(_primary).Bold();
        col.Item().Height(4);
        col.Item().Text("Email authentication records protect your domains from spoofing and phishing attacks. " +
            "SPF (Sender Policy Framework) specifies which mail servers are authorised to send email for your domain. " +
            "DKIM (DomainKeys Identified Mail) adds a digital signature to emails to verify they have not been tampered with. " +
            "DMARC (Domain-based Message Authentication) tells receiving mail servers what to do with emails that fail SPF or DKIM checks — " +
            "a policy of 'reject' provides the strongest protection.").FontSize(8).FontColor(Colors.Grey.Darken1);
        col.Item().Height(8);
        col.Item().Table(t =>
        {
            TwoCols(t);
            TH(t, "Metric", "Count");
            TR(t, "Total Domains", $"{_data.DomainSecuritySummary!.TotalDomains}");
            TR(t, "With MX",      $"{_data.DomainSecuritySummary.DomainsWithMx}");
            TR(t, "With SPF",     $"{_data.DomainSecuritySummary.DomainsWithSpf}", Compliant);
            TR(t, "With DMARC",   $"{_data.DomainSecuritySummary.DomainsWithDmarc}", Compliant);
            TR(t, "With DKIM",    $"{_data.DomainSecuritySummary.DomainsWithDkim}", Compliant);
        });

        if (_data.DomainSecurityResults?.Any() == true)
        {
            col.Item().Height(18);
            col.Item().Text("Domain Details").FontSize(16).FontColor(_primary).Bold();
            col.Item().Height(10);
            col.Item().Table(t =>
            {
                t.ColumnsDefinition(c =>
                {
                    c.RelativeColumn(3); c.RelativeColumn(1); c.RelativeColumn(1);
                    c.RelativeColumn(1.5f); c.RelativeColumn(1);
                });
                foreach (var h in new[] { "Domain", "MX", "SPF", "DMARC", "DKIM" })
                    t.Cell().Background(LightGray).Border(1).BorderColor(BorderCol)
                        .Padding(5).Text(h).FontSize(8).Bold();
                foreach (var d in _data.DomainSecurityResults
                    .Where(x => !_settings.ExcludedDomains.Any(e => string.Equals(e, x.Domain, StringComparison.OrdinalIgnoreCase)))
                    .OrderByDescending(x => x.SecurityScore).Take(20))
                {
                    t.Cell().Border(1).BorderColor(BorderCol).Padding(5).Text(d.Domain).FontSize(8);
                    Tick(t, d.HasMx);
                    Tick(t, d.HasSpf);
                    t.Cell().Border(1).BorderColor(BorderCol).Padding(5)
                        .Text(d.HasDmarc ? (d.DmarcPolicy ?? "✓") : "✗").FontSize(8)
                        .FontColor(d.HasDmarc ? Compliant : Crit);
                    Tick(t, d.HasDkim);
                }
            });
        }
    }

    private void DeletedUsersPage(ColumnDescriptor col)
    {
        col.Item().Text($"Deleted Users ({_data.DeletedUsersInPeriod!.Count} in period)")
            .FontSize(16).FontColor(_primary).Bold();
        col.Item().Height(6);
        col.Item().Text("Users deleted during the report period. Deleted accounts are soft-deleted and retained in the directory for 30 days " +
            "before permanent removal. Reviewing deletions helps identify unauthorised account removals or offboarding that may " +
            "require follow-up, such as revoking access to shared resources or reassigning licences.").FontSize(8).FontColor(Colors.Grey.Darken1);
        col.Item().Height(12);
        col.Item().Table(t =>
        {
            t.ColumnsDefinition(c =>
            {
                c.RelativeColumn(2); c.RelativeColumn(2.5f);
                c.RelativeColumn(1.5f); c.RelativeColumn(1.5f); c.RelativeColumn(1.5f);
            });
            foreach (var h in new[] { "Name", "Email", "Deleted", "Job Title", "Department" })
                t.Cell().Background(LightGray).Border(1).BorderColor(BorderCol)
                    .Padding(5).Text(h).FontSize(8).Bold();
            foreach (var u in _data.DeletedUsersInPeriod)
            {
                t.Cell().Border(1).BorderColor(BorderCol).Padding(5).Text(u.DisplayName ?? "-").FontSize(8);
                t.Cell().Border(1).BorderColor(BorderCol).Padding(5).Text(u.UserPrincipalName ?? u.Mail ?? "-").FontSize(7);
                t.Cell().Border(1).BorderColor(BorderCol).Padding(5)
                    .Text(u.DeletedDateTime?.ToString("dd MMM yyyy") ?? "-").FontSize(8).FontColor(Crit);
                t.Cell().Border(1).BorderColor(BorderCol).Padding(5).Text(u.JobTitle ?? "-").FontSize(8);
                t.Cell().Border(1).BorderColor(BorderCol).Padding(5).Text(u.Department ?? "-").FontSize(8);
            }
        });
    }

    private void MailboxDetailsPage(ColumnDescriptor col)
    {
        col.Item().Text($"Mailbox Storage ({_data.MailboxDetails!.Count} mailboxes)")
            .FontSize(16).FontColor(_primary).Bold();
        col.Item().Height(6);
        col.Item().Text("Mailbox storage usage over the last 30 days, sorted by size. Exchange Online mailboxes have a default quota of 50GB (100GB with Microsoft 365 E3/E5). " +
            "Mailboxes approaching their quota will begin to experience send/receive restrictions. " +
            "Inactive mailboxes (no recent activity) may indicate unused licences that could be reclaimed.").FontSize(8).FontColor(Colors.Grey.Darken1);
        col.Item().Height(12);
        col.Item().Table(t =>
        {
            t.ColumnsDefinition(c =>
            {
                c.RelativeColumn(2); c.RelativeColumn(2.5f);
                c.RelativeColumn(1); c.RelativeColumn(1); c.RelativeColumn(1); c.RelativeColumn(1.2f);
            });
            foreach (var h in new[] { "Name", "Email", "Size (GB)", "Quota (GB)", "% Used", "Last Active" })
                t.Cell().Background(LightGray).Border(1).BorderColor(BorderCol)
                    .Padding(5).Text(h).FontSize(8).Bold();
            foreach (var m in _data.MailboxDetails.Take(40))
            {
                var pctColor = m.PercentUsed >= 90 ? Crit : m.PercentUsed >= 75 ? Warn : null;
                t.Cell().Border(1).BorderColor(BorderCol).Padding(5).Text(m.DisplayName ?? "-").FontSize(8);
                t.Cell().Border(1).BorderColor(BorderCol).Padding(5).Text(m.UserPrincipalName ?? "-").FontSize(7);
                t.Cell().Border(1).BorderColor(BorderCol).Padding(5).AlignRight()
                    .Text($"{m.StorageUsedGB:F2}").FontSize(8);
                t.Cell().Border(1).BorderColor(BorderCol).Padding(5).AlignRight()
                    .Text(m.QuotaGB.HasValue ? $"{m.QuotaGB:F0}" : "-").FontSize(8);
                var pctCell = t.Cell().Border(1).BorderColor(BorderCol).Padding(5).AlignRight();
                if (pctColor != null)
                    pctCell.Text(m.PercentUsed.HasValue ? $"{m.PercentUsed:F1}%" : "-").FontSize(8).FontColor(pctColor);
                else
                    pctCell.Text(m.PercentUsed.HasValue ? $"{m.PercentUsed:F1}%" : "-").FontSize(8);
                t.Cell().Border(1).BorderColor(BorderCol).Padding(5)
                    .Text(m.LastActivityDate?.ToString("dd MMM yyyy") ?? "Never").FontSize(8);
            }
        });
    }

    private void AppCredentialsPage(ColumnDescriptor col)
    {
        var creds = _data.AppCredentialStatus!;

        col.Item().Text($"App Secrets & Certificates ({creds.ThresholdDays}-day threshold)")
            .FontSize(28).FontColor(_primary);
        col.Item().Height(6);
        col.Item().Text(
            $"App registrations use secrets and certificates to authenticate to Microsoft 365 and other services. " +
            $"Expired credentials cause application outages and broken integrations. " +
            $"{creds.TotalApps} app registrations were scanned and credentials expiring within {creds.ThresholdDays} days or already expired are listed below. " +
            $"Renew expiring credentials before they expire to avoid service disruption.").FontSize(8).FontColor(Colors.Grey.Darken1);
        col.Item().Height(16);

        void CredTable(string title, List<AppCredentialDetail> items, string rowColor)
        {
            if (!items.Any()) return;
            col.Item().Text(title).FontSize(14).FontColor(_primary).Bold();
            col.Item().Height(6);
            col.Item().Table(t =>
            {
                t.ColumnsDefinition(c =>
                {
                    c.RelativeColumn(2.5f); c.RelativeColumn(1.5f);
                    c.RelativeColumn(1.5f); c.RelativeColumn(2.5f);
                });
                foreach (var h in new[] { "App Name", "Type", "Expiry Date", "Status" })
                    t.Cell().Background(LightGray).Border(1).BorderColor(BorderCol)
                        .Padding(4).Text(h).FontSize(8).Bold();
                foreach (var item in items)
                {
                    t.Cell().Border(1).BorderColor(BorderCol).Padding(4)
                        .Text(item.AppName ?? "-").FontSize(8);
                    t.Cell().Border(1).BorderColor(BorderCol).Padding(4)
                        .Text(item.CredentialType ?? "-").FontSize(8);
                    t.Cell().Border(1).BorderColor(BorderCol).Padding(4)
                        .Text(item.ExpiryDate.HasValue ? item.ExpiryDate.Value.ToString("dd MMM yyyy") : "-").FontSize(8);
                    t.Cell().Border(1).BorderColor(BorderCol).Padding(4)
                        .Text(item.Status ?? "-").FontSize(8).FontColor(rowColor);
                }
            });
            col.Item().Height(14);
        }

        CredTable($"Expired Secrets ({creds.ExpiredSecrets.Count})", creds.ExpiredSecrets, Crit);
        CredTable($"Expiring Secrets ({creds.ExpiringSecrets.Count})", creds.ExpiringSecrets, Warn);
        CredTable($"Expired Certificates ({creds.ExpiredCertificates.Count})", creds.ExpiredCertificates, Crit);
        CredTable($"Expiring Certificates ({creds.ExpiringCertificates.Count})", creds.ExpiringCertificates, Warn);
    }

    // ── Helpers ───────────────────────────────────────────────────────────

    private static void KpiCard(IContainer c, string label, string value, string sub) =>
        c.Border(1).BorderColor(BorderCol).Padding(14).Column(k =>
        {
            k.Item().AlignRight().Text(value).FontSize(26).FontColor("#1E3A5F");
            k.Item().Height(3);
            k.Item().Text(label).FontSize(10).Bold();
            k.Item().Text(sub).FontSize(9).FontColor(Colors.Grey.Medium);
        });

    private static void TwoCols(TableDescriptor t) =>
        t.ColumnsDefinition(c => { c.RelativeColumn(3); c.RelativeColumn(1); });

    private static void TH(TableDescriptor t, string c1, string c2)
    {
        t.Cell().Background(LightGray).Border(1).BorderColor(BorderCol).Padding(6)
            .Text(c1).FontSize(9).Bold();
        t.Cell().Background(LightGray).Border(1).BorderColor(BorderCol).Padding(6).AlignRight()
            .Text(c2).FontSize(9).Bold();
    }

    private void SignInLocationsPage(ColumnDescriptor col)
    {
        col.Item().Text("Sign-in Locations – Last 30 Days")
            .FontSize(16).FontColor(_primary).Bold();
        col.Item().Height(6);
        col.Item().Text("Geographic distribution of authentication activity across your Microsoft 365 tenant.")
            .FontSize(10).FontColor(BodyText).Italic();
        col.Item().Height(16);

        // Map image
        col.Item().Border(1).BorderColor(BorderCol)
            .Image(_data.SignInMapImageBytes!).FitWidth();
        col.Item().Height(4);
        col.Item().AlignRight()
            .Text("Powered by Azure Maps").FontSize(7).FontColor("#9CA3AF").Italic();
        col.Item().Height(20);

        // Summary stats + top countries side by side
        if (_data.SignInLocations?.Any() == true)
        {
            var totalSignIns  = _data.SignInLocations.Sum(l => l.SignInCount);
            var totalCountries = _data.SignInLocations.Count;
            var topCountry    = _data.SignInLocations.First();

            // Stats row
            col.Item().Row(row =>
            {
                void StatBox(string value, string label, string color) =>
                    row.RelativeItem().Border(1).BorderColor(BorderCol)
                        .Background(LightGray).Padding(12).Column(c =>
                        {
                            c.Item().AlignCenter().Text(value).FontSize(22).Bold().FontColor(color);
                            c.Item().Height(4);
                            c.Item().AlignCenter().Text(label).FontSize(9).FontColor(BodyText);
                        });

                StatBox(totalSignIns.ToString("N0"),   "Total Sign-ins",    _primary);
                row.ConstantItem(8);
                StatBox(totalCountries.ToString(),      "Countries / Regions", _accent);
                row.ConstantItem(8);
                StatBox(topCountry.Country ?? "—",     "Top Sign-in Country", Compliant);
            });

            col.Item().Height(20);

            // Top 10 countries table
            col.Item().Text("Top Sign-in Countries").FontSize(12).FontColor(_primary).Bold();
            col.Item().Height(8);
            col.Item().Table(t =>
            {
                t.ColumnsDefinition(cd =>
                {
                    cd.ConstantColumn(20);   // Rank
                    cd.RelativeColumn(4);    // Country
                    cd.RelativeColumn(2);    // Sign-ins
                    cd.RelativeColumn(2);    // % of total
                });

                t.Header(h =>
                {
                    foreach (var hdr in new[] { "#", "Country", "Sign-ins", "% of Total" })
                        h.Cell().Background(_primary).Padding(6)
                            .Text(hdr).FontColor(Colors.White).FontSize(9).Bold();
                });

                var rank = 1;
                foreach (var loc in _data.SignInLocations.Take(10))
                {
                    var pct = totalSignIns > 0
                        ? Math.Round((double)loc.SignInCount / totalSignIns * 100, 1)
                        : 0;
                    var bg = rank % 2 == 0 ? LightGray : "#FFFFFF";

                    t.Cell().Background(bg).Border(1).BorderColor(BorderCol).Padding(6)
                        .AlignCenter().Text(rank.ToString()).FontSize(9);
                    t.Cell().Background(bg).Border(1).BorderColor(BorderCol).Padding(6)
                        .Text(loc.Country ?? "Unknown").FontSize(9);
                    t.Cell().Background(bg).Border(1).BorderColor(BorderCol).Padding(6)
                        .AlignRight().Text(loc.SignInCount.ToString("N0")).FontSize(9);
                    t.Cell().Background(bg).Border(1).BorderColor(BorderCol).Padding(6)
                        .AlignRight().Text($"{pct}%").FontSize(9);
                    rank++;
                }
            });
        }
    }

    private static void TR(TableDescriptor t, string label, string value, string? colour = null)
    {
        t.Cell().Border(1).BorderColor(BorderCol).Padding(6).Text(label).FontSize(9);
        var cell = t.Cell().Border(1).BorderColor(BorderCol).Padding(6).AlignRight();
        if (colour != null) cell.Text(value).FontSize(9).FontColor(colour);
        else cell.Text(value).FontSize(9);
    }

    private static void Tick(TableDescriptor t, bool ok) =>
        t.Cell().Border(1).BorderColor(BorderCol).Padding(5)
            .Text(ok ? "✓" : "✗").FontSize(8).FontColor(ok ? Compliant : Crit);

    private static (string line1, string line2) ParseTitle(string title)
    {
        if (string.IsNullOrWhiteSpace(title)) return ("MICROSOFT 365", "SECURITY ASSESSMENT");
        var l = title.ToLower();
        if (l.StartsWith("microsoft 365")) return ("MICROSOFT 365", title[13..].Trim().ToUpper());
        if (l.StartsWith("m365"))          return ("MICROSOFT 365", title[4..].Trim().ToUpper());
        return ("MICROSOFT 365", title.ToUpper());
    }

    private static List<ReportQuote> PickRandomQuotes(ReportSettings s, int count)
    {
        var pool = s.Quotes.Where(q => q.Enabled && !string.IsNullOrWhiteSpace(q.BigNumber)).ToList();
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

    private bool HasDevices() =>
        _data.DeviceDetails?.WindowsDevices?.Any() == true ||
        _data.DeviceDetails?.MacDevices?.Any()     == true ||
        _data.DeviceDetails?.IosDevices?.Any()     == true ||
        _data.DeviceDetails?.AndroidDevices?.Any() == true;
}
