using DnsClient;
using DnsClient.Protocol;
using System.Text.RegularExpressions;

namespace M365Dashboard.Api.Services;

public interface IDomainSecurityService
{
    Task<List<DomainSecurityResult>> CheckDomainsAsync(IEnumerable<string> domains);
    Task<DomainSecurityResult> CheckDomainAsync(string domain);
    Task<DomainSecuritySummary> GetSecuritySummaryAsync(List<DomainSecurityResult> results);
}

public class DomainSecurityService : IDomainSecurityService
{
    private readonly ILogger<DomainSecurityService> _logger;
    private readonly ILookupClient _lookupClient;
    
    // Common DKIM selectors to check
    private static readonly string[] DkimSelectors = new[]
    {
        "selector1",      // Microsoft 365
        "selector2",      // Microsoft 365
        "google",         // Google Workspace
        "default",        // Common default
        "k1",             // Various ESPs
        "s1",             // Various ESPs
        "dkim",           // Generic
        "mail",           // Generic
        "mcsv",           // Mailchimp
        "mandrill",       // Mailchimp/Mandrill
        "sm"              // Salesforce Marketing Cloud
    };

    public DomainSecurityService(ILogger<DomainSecurityService> logger)
    {
        _logger = logger;
        _lookupClient = new LookupClient(new LookupClientOptions
        {
            UseCache = true,
            Timeout = TimeSpan.FromSeconds(5),
            Retries = 2
        });
    }

    public async Task<List<DomainSecurityResult>> CheckDomainsAsync(IEnumerable<string> domains)
    {
        var results = new List<DomainSecurityResult>();
        
        foreach (var domain in domains)
        {
            try
            {
                var result = await CheckDomainAsync(domain);
                results.Add(result);
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "Error checking domain {Domain}", domain);
                results.Add(new DomainSecurityResult
                {
                    Domain = domain,
                    Error = ex.Message
                });
            }
            
            // Small delay between requests
            await Task.Delay(100);
        }
        
        return results;
    }

    public async Task<DomainSecurityResult> CheckDomainAsync(string domain)
    {
        _logger.LogDebug("Checking domain security for {Domain}", domain);
        
        var result = new DomainSecurityResult
        {
            Domain = domain,
            CheckedAt = DateTime.UtcNow
        };

        // Check MX records
        await CheckMxRecordsAsync(domain, result);
        
        // Check SPF record
        await CheckSpfRecordAsync(domain, result);
        
        // Check DMARC record
        await CheckDmarcRecordAsync(domain, result);
        
        // Check DKIM selectors
        await CheckDkimRecordsAsync(domain, result);
        
        // Check MTA-STS
        await CheckMtaStsAsync(domain, result);
        
        // Calculate security score
        CalculateSecurityScore(result);
        
        return result;
    }

    public Task<DomainSecuritySummary> GetSecuritySummaryAsync(List<DomainSecurityResult> results)
    {
        var summary = new DomainSecuritySummary
        {
            TotalDomains = results.Count,
            DomainsWithMx = results.Count(r => r.HasMx),
            DomainsWithSpf = results.Count(r => r.HasSpf),
            DomainsWithDmarc = results.Count(r => r.HasDmarc),
            DomainsWithDkim = results.Count(r => r.HasDkim),
            DomainsWithMtaSts = results.Count(r => r.HasMtaSts),
            
            DmarcRejectCount = results.Count(r => r.DmarcPolicy == "reject"),
            DmarcQuarantineCount = results.Count(r => r.DmarcPolicy == "quarantine"),
            DmarcNoneCount = results.Count(r => r.DmarcPolicy == "none"),
            
            SpfHardFailCount = results.Count(r => r.SpfPolicy == "-all"),
            SpfSoftFailCount = results.Count(r => r.SpfPolicy == "~all"),
            
            GradeACount = results.Count(r => r.SecurityGrade == "A"),
            GradeBCount = results.Count(r => r.SecurityGrade == "B"),
            GradeCCount = results.Count(r => r.SecurityGrade == "C"),
            GradeDCount = results.Count(r => r.SecurityGrade == "D"),
            GradeFCount = results.Count(r => r.SecurityGrade == "F"),
            
            CriticalIssuesCount = results.Count(r => r.SecurityGrade == "D" || r.SecurityGrade == "F")
        };

        // Identify mail providers
        var providers = results
            .Where(r => !string.IsNullOrEmpty(r.MailProvider) && r.MailProvider != "Unknown" && r.MailProvider != "None")
            .GroupBy(r => r.MailProvider)
            .ToDictionary(g => g.Key!, g => g.Count());
        summary.MailProviders = providers;

        return Task.FromResult(summary);
    }

    private async Task CheckMxRecordsAsync(string domain, DomainSecurityResult result)
    {
        try
        {
            var response = await _lookupClient.QueryAsync(domain, QueryType.MX);
            var mxRecords = response.Answers.MxRecords().ToList();
            
            if (mxRecords.Any())
            {
                result.HasMx = true;
                result.MxRecords = mxRecords
                    .OrderBy(mx => mx.Preference)
                    .Select(mx => $"{mx.Preference} {mx.Exchange.Value.TrimEnd('.')}")
                    .ToList();
                
                // Detect mail provider
                var mxString = string.Join(" ", result.MxRecords).ToLower();
                result.MailProvider = DetectMailProvider(mxString);
            }
            else
            {
                result.HasMx = false;
                result.MailProvider = "None";
            }
        }
        catch (DnsResponseException ex) when (ex.Code == DnsResponseCode.NotExistentDomain)
        {
            result.HasMx = false;
            result.MailProvider = "NXDOMAIN";
            _logger.LogDebug("Domain {Domain} does not exist (NXDOMAIN)", domain);
        }
        catch (Exception ex)
        {
            _logger.LogDebug(ex, "Error checking MX for {Domain}", domain);
            result.HasMx = false;
        }
    }

    private async Task CheckSpfRecordAsync(string domain, DomainSecurityResult result)
    {
        try
        {
            var response = await _lookupClient.QueryAsync(domain, QueryType.TXT);
            var txtRecords = response.Answers.TxtRecords().ToList();
            
            foreach (var txt in txtRecords)
            {
                var recordText = string.Join("", txt.Text);
                if (recordText.StartsWith("v=spf1", StringComparison.OrdinalIgnoreCase))
                {
                    result.HasSpf = true;
                    result.SpfRecord = recordText;
                    
                    // Determine SPF policy
                    if (recordText.EndsWith("-all", StringComparison.OrdinalIgnoreCase))
                        result.SpfPolicy = "-all";
                    else if (recordText.EndsWith("~all", StringComparison.OrdinalIgnoreCase))
                        result.SpfPolicy = "~all";
                    else if (recordText.EndsWith("?all", StringComparison.OrdinalIgnoreCase))
                        result.SpfPolicy = "?all";
                    else if (recordText.EndsWith("+all", StringComparison.OrdinalIgnoreCase))
                        result.SpfPolicy = "+all";
                    else
                        result.SpfPolicy = "unknown";
                    
                    // Count DNS lookups
                    var lookupCount = CountSpfLookups(recordText);
                    result.SpfLookupCount = lookupCount;
                    result.SpfLookupStatus = lookupCount > 10 ? $"EXCEEDS LIMIT ({lookupCount})" :
                                             lookupCount > 7 ? $"Warning ({lookupCount})" :
                                             $"OK ({lookupCount})";
                    
                    break;
                }
            }
        }
        catch (Exception ex)
        {
            _logger.LogDebug(ex, "Error checking SPF for {Domain}", domain);
        }
    }

    private async Task CheckDmarcRecordAsync(string domain, DomainSecurityResult result)
    {
        try
        {
            var response = await _lookupClient.QueryAsync($"_dmarc.{domain}", QueryType.TXT);
            var txtRecords = response.Answers.TxtRecords().ToList();
            
            foreach (var txt in txtRecords)
            {
                var recordText = string.Join("", txt.Text);
                if (recordText.Contains("v=DMARC1", StringComparison.OrdinalIgnoreCase))
                {
                    result.HasDmarc = true;
                    result.DmarcRecord = recordText;
                    
                    // Extract policy
                    var policyMatch = Regex.Match(recordText, @"p=(\w+)", RegexOptions.IgnoreCase);
                    if (policyMatch.Success)
                        result.DmarcPolicy = policyMatch.Groups[1].Value.ToLower();
                    
                    // Extract subdomain policy
                    var spMatch = Regex.Match(recordText, @"sp=(\w+)", RegexOptions.IgnoreCase);
                    result.DmarcSubdomainPolicy = spMatch.Success ? spMatch.Groups[1].Value.ToLower() : "inherit";
                    
                    // Extract percentage
                    var pctMatch = Regex.Match(recordText, @"pct=(\d+)", RegexOptions.IgnoreCase);
                    result.DmarcPercentage = pctMatch.Success ? int.Parse(pctMatch.Groups[1].Value) : 100;
                    
                    // Check reporting
                    result.HasDmarcRua = recordText.Contains("rua=", StringComparison.OrdinalIgnoreCase);
                    result.HasDmarcRuf = recordText.Contains("ruf=", StringComparison.OrdinalIgnoreCase);
                    
                    break;
                }
            }
        }
        catch (Exception ex)
        {
            _logger.LogDebug(ex, "Error checking DMARC for {Domain}", domain);
        }
    }

    private async Task CheckDkimRecordsAsync(string domain, DomainSecurityResult result)
    {
        var foundSelectors = new List<string>();
        
        foreach (var selector in DkimSelectors)
        {
            try
            {
                var response = await _lookupClient.QueryAsync($"{selector}._domainkey.{domain}", QueryType.TXT);
                var txtRecords = response.Answers.TxtRecords().ToList();
                
                if (txtRecords.Any(txt => string.Join("", txt.Text).Contains("v=DKIM1", StringComparison.OrdinalIgnoreCase) ||
                                          string.Join("", txt.Text).Contains("k=rsa", StringComparison.OrdinalIgnoreCase) ||
                                          string.Join("", txt.Text).Contains("p=", StringComparison.OrdinalIgnoreCase)))
                {
                    foundSelectors.Add(selector);
                }
            }
            catch
            {
                // Selector not found, continue
            }
        }
        
        result.HasDkim = foundSelectors.Any();
        result.DkimSelectors = foundSelectors;
    }

    private async Task CheckMtaStsAsync(string domain, DomainSecurityResult result)
    {
        try
        {
            var response = await _lookupClient.QueryAsync($"_mta-sts.{domain}", QueryType.TXT);
            var txtRecords = response.Answers.TxtRecords().ToList();
            
            foreach (var txt in txtRecords)
            {
                var recordText = string.Join("", txt.Text);
                if (recordText.Contains("v=STSv1", StringComparison.OrdinalIgnoreCase))
                {
                    result.HasMtaSts = true;
                    result.MtaStsRecord = recordText;
                    break;
                }
            }
        }
        catch (Exception ex)
        {
            _logger.LogDebug(ex, "Error checking MTA-STS for {Domain}", domain);
        }
    }

    private void CalculateSecurityScore(DomainSecurityResult result)
    {
        int score = 0;
        var issues = new List<string>();
        
        // MX (10 points)
        if (result.HasMx)
            score += 10;
        else
            issues.Add("No MX records");
        
        // SPF (25 points)
        if (result.HasSpf)
        {
            score += 15;
            if (result.SpfPolicy == "-all")
                score += 10;
            else if (result.SpfPolicy == "~all")
            {
                score += 5;
                issues.Add("SPF uses soft fail (~all)");
            }
            else
                issues.Add("SPF policy too permissive");
            
            if (result.SpfLookupCount > 10)
            {
                score -= 5;
                issues.Add("SPF exceeds DNS lookup limit");
            }
        }
        else
            issues.Add("No SPF record");
        
        // DMARC (30 points)
        if (result.HasDmarc)
        {
            score += 15;
            if (result.DmarcPolicy == "reject")
                score += 15;
            else if (result.DmarcPolicy == "quarantine")
            {
                score += 10;
                issues.Add("DMARC not at reject");
            }
            else
            {
                score += 5;
                issues.Add("DMARC in monitor mode only");
            }
            
            if (!result.HasDmarcRua && !result.HasDmarcRuf)
                issues.Add("No DMARC reporting configured");
        }
        else
            issues.Add("No DMARC record");
        
        // DKIM (20 points)
        if (result.HasDkim)
            score += 20;
        else
            issues.Add("No DKIM selectors found");
        
        // MTA-STS (10 points)
        if (result.HasMtaSts)
            score += 10;
        else
            issues.Add("No MTA-STS");
        
        // TLS-RPT (5 points) - not currently checked, placeholder
        
        result.SecurityScore = Math.Max(0, Math.Min(100, score));
        result.SecurityGrade = score switch
        {
            >= 90 => "A",
            >= 80 => "B",
            >= 70 => "C",
            >= 60 => "D",
            _ => "F"
        };
        result.Issues = issues;
    }

    private string DetectMailProvider(string mxString)
    {
        if (mxString.Contains("outlook") || mxString.Contains("microsoft"))
            return "Microsoft 365";
        if (mxString.Contains("google") || mxString.Contains("gmail"))
            return "Google Workspace";
        if (mxString.Contains("mimecast"))
            return "Mimecast";
        if (mxString.Contains("proofpoint") || mxString.Contains("pphosted"))
            return "Proofpoint";
        if (mxString.Contains("barracuda"))
            return "Barracuda";
        if (mxString.Contains("messagelabs") || mxString.Contains("symantec"))
            return "Symantec/Broadcom";
        if (mxString.Contains("forcepoint"))
            return "Forcepoint";
        if (mxString.Contains("mailgun"))
            return "Mailgun";
        if (mxString.Contains("sendgrid"))
            return "SendGrid";
        if (mxString.Contains("zoho"))
            return "Zoho";
        if (mxString.Contains("protonmail"))
            return "ProtonMail";
        if (mxString.Contains("amazonses"))
            return "Amazon SES";
        
        return "Unknown";
    }

    private int CountSpfLookups(string spfRecord)
    {
        // Count mechanisms that require DNS lookups
        var lookupPatterns = new[] { "include:", "a:", "a/", "mx:", "mx/", "ptr:", "exists:", "redirect=" };
        int count = 0;
        
        foreach (var pattern in lookupPatterns)
        {
            count += Regex.Matches(spfRecord, Regex.Escape(pattern), RegexOptions.IgnoreCase).Count;
        }
        
        // Check for bare 'a' and 'mx' mechanisms
        if (Regex.IsMatch(spfRecord, @"\ba\b(?!:)", RegexOptions.IgnoreCase))
            count++;
        if (Regex.IsMatch(spfRecord, @"\bmx\b(?!:)", RegexOptions.IgnoreCase))
            count++;
        
        return count;
    }
}

// Data models
public class DomainSecurityResult
{
    public string Domain { get; set; } = string.Empty;
    public DateTime CheckedAt { get; set; }
    public string? Error { get; set; }
    
    // MX
    public bool HasMx { get; set; }
    public List<string> MxRecords { get; set; } = new();
    public string? MailProvider { get; set; }
    
    // SPF
    public bool HasSpf { get; set; }
    public string? SpfRecord { get; set; }
    public string? SpfPolicy { get; set; }
    public int SpfLookupCount { get; set; }
    public string? SpfLookupStatus { get; set; }
    
    // DMARC
    public bool HasDmarc { get; set; }
    public string? DmarcRecord { get; set; }
    public string? DmarcPolicy { get; set; }
    public string? DmarcSubdomainPolicy { get; set; }
    public int DmarcPercentage { get; set; } = 100;
    public bool HasDmarcRua { get; set; }
    public bool HasDmarcRuf { get; set; }
    
    // DKIM
    public bool HasDkim { get; set; }
    public List<string> DkimSelectors { get; set; } = new();
    
    // MTA-STS
    public bool HasMtaSts { get; set; }
    public string? MtaStsRecord { get; set; }
    
    // Score
    public int SecurityScore { get; set; }
    public string SecurityGrade { get; set; } = "F";
    public List<string> Issues { get; set; } = new();
}

public class DomainSecuritySummary
{
    public int TotalDomains { get; set; }
    public int DomainsWithMx { get; set; }
    public int DomainsWithSpf { get; set; }
    public int DomainsWithDmarc { get; set; }
    public int DomainsWithDkim { get; set; }
    public int DomainsWithMtaSts { get; set; }
    
    public int DmarcRejectCount { get; set; }
    public int DmarcQuarantineCount { get; set; }
    public int DmarcNoneCount { get; set; }
    
    public int SpfHardFailCount { get; set; }
    public int SpfSoftFailCount { get; set; }
    
    public int GradeACount { get; set; }
    public int GradeBCount { get; set; }
    public int GradeCCount { get; set; }
    public int GradeDCount { get; set; }
    public int GradeFCount { get; set; }
    public int CriticalIssuesCount { get; set; }
    
    public Dictionary<string, int> MailProviders { get; set; } = new();
}
