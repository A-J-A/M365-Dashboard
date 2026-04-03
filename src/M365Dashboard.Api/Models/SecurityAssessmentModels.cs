using System.Text.Json.Serialization;

namespace M365Dashboard.Api.Models;

/// <summary>
/// Microsoft 365 Security Assessment Report models
/// Inspired by professional MSP security assessment reports
/// </summary>

public class SecurityAssessmentResult
{
    public string ReportTitle { get; set; } = "Microsoft 365 Security Assessment";
    public DateTime GeneratedAt { get; set; } = DateTime.UtcNow;
    public string TenantId { get; set; } = string.Empty;
    public string TenantName { get; set; } = string.Empty;
    public string TenantDomain { get; set; } = string.Empty;
    
    // Executive Summary Stats
    public UserStatistics UserStats { get; set; } = new();
    public LicenseStatistics LicenseStats { get; set; } = new();
    public List<AdminRoleAssignment> RoleDistribution { get; set; } = new();
    
    // Compliance Sections
    public ComplianceSection EntraIdCompliance { get; set; } = new() { SectionName = "Entra ID", SectionDescription = "Microsoft have stated that 99% of breaches could be mitigated with strong passwords and multi-factor authentication. Enabling and enforcing MFA across your organization is one of the easiest and most effective ways to increase your security posture. This section evaluates your Entra ID (Azure AD) configuration against security best practices." };
    public ComplianceSection ExchangeCompliance { get; set; } = new() { SectionName = "Exchange Online", SectionDescription = "Business Email Compromise is a common attack vector. Microsoft Exchange Online provides a number of security features to help protect your organization from email-based threats. This section evaluates your Exchange Online configuration against security best practices to ensure your email environment is properly secured." };
    public ComplianceSection SharePointCompliance { get; set; } = new() { SectionName = "SharePoint & OneDrive", SectionDescription = "Internal and External sharing of business data is one of the most challenging aspects to manage within Microsoft 365. SharePoint Online provides a number of security features to help protect your organization's data. This section evaluates your SharePoint Online configuration against security best practices to ensure your data sharing environment is properly secured." };
    public ComplianceSection TeamsCompliance { get; set; } = new() { SectionName = "Microsoft Teams", SectionDescription = "Microsoft Teams essentially front-ends a number of services which have already been checked e.g. identity, internal and external document sharing and remote access. This section evaluates your Teams configuration against security best practices to ensure your collaboration environment is properly secured." };
    public ComplianceSection IntuneCompliance { get; set; } = new() { SectionName = "Microsoft Intune", SectionDescription = "Business data is accessed by employees across multiple devices, both organisation owned and personal. This section evaluates your Intune configuration against security best practices to ensure your devices are properly managed and secured." };
    public ComplianceSection DefenderCompliance { get; set; } = new() { SectionName = "Microsoft Defender", SectionDescription = "Microsoft Defender for Office 365 provides protection against advanced threats in email, attachments, and links. This section evaluates your Defender configuration against security best practices." };
    
    // Overall Summary
    public int TotalChecks { get; set; }
    public int CompliantChecks { get; set; }
    public int NonCompliantChecks { get; set; }
    public double OverallCompliancePercentage { get; set; }
}

public class UserStatistics
{
    public int TotalUsers { get; set; }
    public int MemberUsers { get; set; }
    public int GuestUsers { get; set; }
    public int LicensedUsers { get; set; }
    public int UnlicensedUsers { get; set; }
    public int BlockedUsers { get; set; }
    public int BlockedUsersWithLicenses { get; set; }
    public int AdminUsers { get; set; }
    public int MfaRegisteredUsers { get; set; }
    public int MfaNotRegisteredUsers { get; set; }
}

public class LicenseStatistics
{
    public int TotalLicenses { get; set; }
    public int AssignedLicenses { get; set; }
    public int AvailableLicenses { get; set; }
    public List<LicenseSummary> LicenseBreakdown { get; set; } = new();
}

public class LicenseSummary
{
    public string SkuName { get; set; } = string.Empty;
    public string SkuPartNumber { get; set; } = string.Empty;
    public int Total { get; set; }
    public int Assigned { get; set; }
    public int Available { get; set; }
}

public class AdminRoleAssignment
{
    public string RoleName { get; set; } = string.Empty;
    public int MemberCount { get; set; }
    public List<string> Members { get; set; } = new();
}

public class ComplianceSection
{
    public string SectionName { get; set; } = string.Empty;
    public string SectionDescription { get; set; } = string.Empty;
    public int TotalChecks { get; set; }
    public int CompliantChecks { get; set; }
    public int NonCompliantChecks { get; set; }
    public double CompliancePercentage { get; set; }
    public List<SecurityCheck> Checks { get; set; } = new();
}

public class SecurityCheck
{
    public string Name { get; set; } = string.Empty;
    public string Description { get; set; } = string.Empty;
    public DateTime CheckedAt { get; set; } = DateTime.UtcNow;
    public SecurityCheckStatus Status { get; set; } = SecurityCheckStatus.Unknown;
    public string CurrentValue { get; set; } = string.Empty;
    public string ExpectedValue { get; set; } = string.Empty;
    public string Remediation { get; set; } = string.Empty;
    public string Reference { get; set; } = string.Empty;
    public List<string> AffectedItems { get; set; } = new();
    public bool IsBeta { get; set; } = false;
}

[JsonConverter(typeof(JsonStringEnumConverter))]
public enum SecurityCheckStatus
{
    Compliant,
    NonCompliant,
    Warning,
    NotApplicable,
    Error,
    Unknown
}
