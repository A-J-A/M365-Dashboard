using System.Text.Json.Serialization;

namespace M365Dashboard.Api.Models;

/// <summary>
/// CIS Microsoft 365 Foundations Benchmark v6.0.0 data models
/// </summary>

public class CisBenchmarkResult
{
    public string ReportTitle { get; set; } = "CIS Microsoft 365 Foundations Benchmark";
    public string BenchmarkVersion { get; set; } = "6.0.0";
    public DateTime GeneratedAt { get; set; } = DateTime.UtcNow;
    public string TenantId { get; set; } = string.Empty;
    public string TenantName { get; set; } = string.Empty;
    
    // Summary
    public int TotalControls { get; set; }
    public int PassedControls { get; set; }
    public int FailedControls { get; set; }
    public int ManualControls { get; set; }
    public int NotApplicableControls { get; set; }
    public int ErrorControls { get; set; }
    public double CompliancePercentage { get; set; }
    
    // By Level
    public int Level1Total { get; set; }
    public int Level1Passed { get; set; }
    public int Level2Total { get; set; }
    public int Level2Passed { get; set; }
    
    // By Category
    public List<CisCategoryResult> Categories { get; set; } = new();
    
    // All Controls
    public List<CisControlResult> Controls { get; set; } = new();
}

public class CisCategoryResult
{
    public string CategoryId { get; set; } = string.Empty;
    public string CategoryName { get; set; } = string.Empty;
    public int TotalControls { get; set; }
    public int PassedControls { get; set; }
    public int FailedControls { get; set; }
    public int ManualControls { get; set; }
    public double CompliancePercentage { get; set; }
}

public class CisControlResult
{
    public string ControlId { get; set; } = string.Empty;
    public string Title { get; set; } = string.Empty;
    public string Description { get; set; } = string.Empty;
    public string Rationale { get; set; } = string.Empty;
    public string Category { get; set; } = string.Empty;
    public string SubCategory { get; set; } = string.Empty;
    public CisLevel Level { get; set; } = CisLevel.L1;
    public CisLicenseProfile Profile { get; set; } = CisLicenseProfile.E3;
    public CisControlStatus Status { get; set; } = CisControlStatus.Unknown;
    public string StatusReason { get; set; } = string.Empty;
    public string CurrentValue { get; set; } = string.Empty;
    public string ExpectedValue { get; set; } = string.Empty;
    public string Remediation { get; set; } = string.Empty;
    public string Impact { get; set; } = string.Empty;
    public string Reference { get; set; } = string.Empty;
    public bool IsAutomated { get; set; } = true;
    public List<string> AffectedItems { get; set; } = new();
}

[JsonConverter(typeof(JsonStringEnumConverter))]
public enum CisLevel
{
    L1,  // Level 1 - Basic security, minimal impact
    L2   // Level 2 - Defense in depth, may impact functionality
}

[JsonConverter(typeof(JsonStringEnumConverter))]
public enum CisLicenseProfile
{
    E3,
    E5
}

[JsonConverter(typeof(JsonStringEnumConverter))]
public enum CisControlStatus
{
    Pass,
    Fail,
    Manual,
    NotApplicable,
    Error,
    Unknown
}

// Request/Response models
public class CisBenchmarkRequest
{
    public bool IncludeLevel2 { get; set; } = true;
    public bool IncludeE5Only { get; set; } = true;
    public List<string>? Categories { get; set; }
}

public class CisBenchmarkSummary
{
    public DateTime LastScanDate { get; set; }
    public int TotalControls { get; set; }
    public int PassedControls { get; set; }
    public int FailedControls { get; set; }
    public double CompliancePercentage { get; set; }
    public List<CisControlResult> CriticalFailures { get; set; } = new();
}
