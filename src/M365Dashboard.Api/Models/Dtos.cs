namespace M365Dashboard.Api.Models.Dtos;

// User Settings DTOs
public record UserSettingsDto(
    string Theme,
    int RefreshIntervalSeconds,
    string DateRangePreference,
    bool ShowWelcomeMessage,
    bool CompactMode
);

public record UpdateUserSettingsDto(
    string? Theme,
    int? RefreshIntervalSeconds,
    string? DateRangePreference,
    bool? ShowWelcomeMessage,
    bool? CompactMode
);

// Widget DTOs
public record WidgetConfigurationDto(
    int Id,
    string WidgetType,
    bool IsEnabled,
    int DisplayOrder,
    int GridColumn,
    int GridRow,
    int GridWidth,
    int GridHeight,
    Dictionary<string, object>? CustomSettings
);

public record UpdateWidgetConfigurationDto(
    bool? IsEnabled,
    int? DisplayOrder,
    int? GridColumn,
    int? GridRow,
    int? GridWidth,
    int? GridHeight,
    Dictionary<string, object>? CustomSettings
);

public record WidgetDefinitionDto(
    string Type,
    string Name,
    string Description,
    string Category,
    string[] RequiredPermissions,
    int DefaultWidth,
    int DefaultHeight
);

// Dashboard Layout DTOs
public record DashboardLayoutDto(
    int Id,
    string Name,
    bool IsDefault,
    List<WidgetConfigurationDto> Widgets,
    DateTime CreatedAt,
    DateTime UpdatedAt
);

public record CreateDashboardLayoutDto(
    string Name,
    bool IsDefault,
    List<WidgetConfigurationDto> Widgets
);

// Widget Data DTOs
public record ActiveUsersDataDto(
    int DailyActiveUsers,
    int WeeklyActiveUsers,
    int MonthlyActiveUsers,
    List<DailyActiveUsersTrendDto> Trend,
    DateTime LastUpdated
);

public record DailyActiveUsersTrendDto(
    DateTime Date,
    int Count
);

public record SignInAnalyticsDto(
    int TotalSignIns,
    int SuccessfulSignIns,
    int FailedSignIns,
    int RiskySignIns,
    double SuccessRate,
    List<SignInTrendDto> Trend,
    List<TopSignInLocationDto> TopLocations,
    DateTime LastUpdated
);

public record SignInTrendDto(
    DateTime Date,
    int Successful,
    int Failed
);

public record TopSignInLocationDto(
    string Location,
    int Count
);

public record LicenseUsageDto(
    List<LicenseSkuDto> Licenses,
    int TotalConsumed,
    int TotalAvailable,
    double OverallUtilization,
    DateTime LastUpdated
);

public record LicenseSkuDto(
    string SkuId,
    string SkuName,
    int ConsumedUnits,
    int PrepaidUnits,
    double UtilizationPercent
);

public record DeviceComplianceDto(
    int TotalDevices,
    int CompliantDevices,
    int NonCompliantDevices,
    int UnknownDevices,
    double ComplianceRate,
    List<DeviceByPlatformDto> ByPlatform,
    DateTime LastUpdated
);

public record DeviceByPlatformDto(
    string Platform,
    int Total,
    int Compliant,
    int NonCompliant
);

public record MailActivityDto(
    long TotalEmailsSent,
    long TotalEmailsReceived,
    long TotalEmailsRead,
    List<MailActivityTrendDto> Trend,
    DateTime LastUpdated
);

public record MailActivityTrendDto(
    DateTime Date,
    long Sent,
    long Received
);

public record TeamsActivityDto(
    long TotalMessages,
    long TotalCalls,
    long TotalMeetings,
    int ActiveUsers,
    List<TeamsActivityTrendDto> Trend,
    DateTime LastUpdated
);

public record TeamsActivityTrendDto(
    DateTime Date,
    long Messages,
    long Calls,
    long Meetings
);

// User Profile DTO
public record UserProfileDto(
    string Id,
    string DisplayName,
    string Email,
    string? JobTitle,
    string? Department,
    string? ProfilePhoto,
    List<string> Roles
);

// Dashboard Summary
public record DashboardSummaryDto(
    int ActiveUsers,
    double SignInSuccessRate,
    double LicenseUtilization,
    double DeviceComplianceRate,
    DateTime LastUpdated
);

// User List DTOs
public record TenantUserDto(
    string Id,
    string DisplayName,
    string UserPrincipalName,
    string? Mail,
    string UserType,
    bool AccountEnabled,
    DateTime? CreatedDateTime,
    DateTime? LastSignInDateTime,
    DateTime? LastNonInteractiveSignInDateTime,
    string? JobTitle,
    string? Department,
    string? OfficeLocation,
    string? City,
    string? Country,
    string? MobilePhone,
    string? BusinessPhones,
    List<string>? AssignedLicenses,
    bool HasMailbox,
    string? ManagerDisplayName,
    string? ProfilePhoto,
    bool IsMfaRegistered,
    bool IsMfaCapable,
    string? DefaultMfaMethod,
    List<string>? MfaMethods
);

public record UserListResultDto(
    List<TenantUserDto> Users,
    int TotalCount,
    int FilteredCount,
    string? NextLink
);

public record UserStatsDto(
    int TotalUsers,
    int EnabledUsers,
    int DisabledUsers,
    int MemberUsers,
    int GuestUsers,
    int LicensedUsers,
    int UnlicensedUsers,
    int UsersSignedInLast30Days,
    int UsersNeverSignedIn,
    int DeletedUsers,
    int MfaRegistered,
    int MfaNotRegistered,
    DateTime LastUpdated
);

public record UserDetailDto(
    string Id,
    string DisplayName,
    string UserPrincipalName,
    string? Mail,
    string UserType,
    bool AccountEnabled,
    DateTime? CreatedDateTime,
    DateTime? LastSignInDateTime,
    DateTime? LastNonInteractiveSignInDateTime,
    DateTime? LastPasswordChangeDateTime,
    string? JobTitle,
    string? Department,
    string? CompanyName,
    string? OfficeLocation,
    string? StreetAddress,
    string? City,
    string? State,
    string? PostalCode,
    string? Country,
    string? MobilePhone,
    List<string>? BusinessPhones,
    List<LicenseDetailDto>? Licenses,
    List<GroupMembershipDto>? GroupMemberships,
    string? ManagerId,
    string? ManagerDisplayName,
    List<string>? DirectReports,
    string? ProfilePhoto,
    string? OnPremisesSamAccountName,
    bool? OnPremisesSyncEnabled,
    DateTime? OnPremisesLastSyncDateTime
);

public record LicenseDetailDto(
    string SkuId,
    string SkuPartNumber,
    string? DisplayName
);

public record GroupMembershipDto(
    string Id,
    string DisplayName,
    string? Description,
    string GroupType
);

// Groups & Teams DTOs
public record TenantGroupDto(
    string Id,
    string DisplayName,
    string? Description,
    string? Mail,
    string GroupType,
    bool? MailEnabled,
    bool? SecurityEnabled,
    string? Visibility,
    DateTime? CreatedDateTime,
    DateTime? RenewedDateTime,
    int MemberCount,
    int OwnerCount,
    bool IsTeam,
    string? TeamWebUrl,
    List<string>? ResourceProvisioningOptions
);

public record GroupListResultDto(
    List<TenantGroupDto> Groups,
    int TotalCount,
    int FilteredCount,
    string? NextLink
);

public record GroupStatsDto(
    int TotalGroups,
    int Microsoft365Groups,
    int SecurityGroups,
    int DistributionGroups,
    int TeamsEnabled,
    int PublicGroups,
    int PrivateGroups,
    int GroupsWithNoOwner,
    int GroupsWithNoMembers,
    DateTime LastUpdated
);

public record GroupDetailDto(
    string Id,
    string DisplayName,
    string? Description,
    string? Mail,
    string GroupType,
    bool? MailEnabled,
    bool? SecurityEnabled,
    string? Visibility,
    DateTime? CreatedDateTime,
    DateTime? RenewedDateTime,
    DateTime? ExpirationDateTime,
    List<GroupMemberDto>? Members,
    List<GroupMemberDto>? Owners,
    bool IsTeam,
    string? TeamWebUrl,
    bool? IsArchived,
    List<string>? ResourceProvisioningOptions
);

public record GroupMemberDto(
    string Id,
    string DisplayName,
    string? UserPrincipalName,
    string? Mail,
    string MemberType
);

// Intune Device DTOs
public record IntuneDeviceDto(
    string Id,
    string DeviceName,
    string? UserDisplayName,
    string? UserPrincipalName,
    string? ManagedDeviceOwnerType,
    string? OperatingSystem,
    string? OsVersion,
    string? ComplianceState,
    string? ManagementState,
    string? DeviceEnrollmentType,
    DateTime? LastSyncDateTime,
    DateTime? EnrolledDateTime,
    string? Model,
    string? Manufacturer,
    string? SerialNumber,
    string? JailBroken,
    bool? IsEncrypted,
    bool? IsSupervised,
    string? DeviceRegistrationState,
    string? ManagementAgent,
    long? TotalStorageSpaceInBytes,
    long? FreeStorageSpaceInBytes,
    string? WiFiMacAddress,
    string? EthernetMacAddress,
    string? Imei,
    string? PhoneNumber,
    string? AzureAdDeviceId,
    bool? AzureAdRegistered,
    string? DeviceCategoryDisplayName,
    string? SkuFamily,
    string? WindowsEdition
);

public record DeviceListResultDto(
    List<IntuneDeviceDto> Devices,
    int TotalCount,
    int FilteredCount,
    string? NextLink
);

public record DeviceStatsDto(
    int TotalDevices,
    int CompliantDevices,
    int NonCompliantDevices,
    int InGracePeriod,
    int ConfigurationManagerDevices,
    int WindowsDevices,
    int MacOsDevices,
    int IosDevices,
    int AndroidDevices,
    int LinuxDevices,
    int CorporateDevices,
    int PersonalDevices,
    int ManagedDevices,
    int EncryptedDevices,
    DateTime LastUpdated
);

public record DeviceDetailDto(
    string Id,
    string DeviceName,
    string? UserDisplayName,
    string? UserPrincipalName,
    string? UserId,
    string? EmailAddress,
    string? ManagedDeviceOwnerType,
    string? OperatingSystem,
    string? OsVersion,
    string? ComplianceState,
    string? ManagementState,
    string? DeviceEnrollmentType,
    DateTime? LastSyncDateTime,
    DateTime? EnrolledDateTime,
    DateTime? ComplianceGracePeriodExpirationDateTime,
    string? Model,
    string? Manufacturer,
    string? SerialNumber,
    string? JailBroken,
    bool? IsEncrypted,
    bool? IsSupervised,
    string? DeviceRegistrationState,
    string? ManagementAgent,
    long? TotalStorageSpaceInBytes,
    long? FreeStorageSpaceInBytes,
    string? WiFiMacAddress,
    string? EthernetMacAddress,
    string? Imei,
    string? Meid,
    string? PhoneNumber,
    string? SubscriberCarrier,
    string? AzureAdDeviceId,
    bool? AzureAdRegistered,
    string? DeviceCategoryDisplayName,
    List<string>? ConfigurationManagerClientEnabledFeatures,
    string? Notes
);

// Mailflow DTOs
public record MailboxDto(
    string Id,
    string DisplayName,
    string UserPrincipalName,
    string? Mail,
    string? RecipientType,
    string? RecipientTypeDetails,
    DateTime? WhenCreated,
    DateTime? WhenMailboxCreated,
    bool? HiddenFromAddressListsEnabled,
    bool? IsMailboxEnabled,
    long? ProhibitSendQuota,
    long? ProhibitSendReceiveQuota,
    long? IssueWarningQuota,
    long? TotalItemSize,
    int? ItemCount,
    DateTime? LastLogonTime,
    string? PrimarySmtpAddress,
    List<string>? EmailAddresses,
    string? ForwardingAddress,
    string? ForwardingSmtpAddress,
    bool? DeliverToMailboxAndForward,
    string? ArchiveStatus,
    long? ArchiveQuota,
    string? LitigationHoldEnabled,
    string? RetentionPolicy
);

public record MailboxListResultDto(
    List<MailboxDto> Mailboxes,
    int TotalCount,
    int FilteredCount,
    string? NextLink
);

public record MailboxStatsDto(
    int TotalMailboxes,
    int UserMailboxes,
    int SharedMailboxes,
    int RoomMailboxes,
    int EquipmentMailboxes,
    int ActiveMailboxes,
    int InactiveMailboxes,
    long TotalStorageUsedBytes,
    int MailboxesNearQuota,
    int MailboxesWithForwarding,
    int MailboxesOnHold,
    int MailboxesWithArchive,
    DateTime LastUpdated
);

public record MailTrafficReportDto(
    DateTime Date,
    long MessagesSent,
    long MessagesReceived,
    long SpamReceived,
    long MalwareReceived,
    long GoodMail
);

public record MailflowSummaryDto(
    long TotalMessagesSent,
    long TotalMessagesReceived,
    long TotalSpamBlocked,
    long TotalMalwareBlocked,
    double AverageMessagesPerDay,
    List<MailTrafficReportDto> DailyTraffic,
    List<TopSenderDto> TopSenders,
    List<TopRecipientDto> TopRecipients,
    DateTime LastUpdated
);

public record TopSenderDto(
    string UserPrincipalName,
    string? DisplayName,
    long MessageCount
);

public record TopRecipientDto(
    string UserPrincipalName,
    string? DisplayName,
    long MessageCount
);

// Security DTOs
public record SecurityScoreDto(
    double CurrentScore,
    double MaxScore,
    double PercentageScore,
    List<SecurityControlScoreDto> ControlScores,
    DateTime LastUpdated
);

public record SecurityControlScoreDto(
    string ControlName,
    string ControlCategory,
    string? Description,
    double Score,
    double MaxScore,
    string? Implementation,
    string? UserImpact,
    List<string>? Threats
);

public record RiskyUserDto(
    string Id,
    string UserPrincipalName,
    string? DisplayName,
    string RiskLevel,
    string RiskState,
    string? RiskDetail,
    DateTime? RiskLastUpdatedDateTime,
    bool IsDeleted,
    bool IsProcessing
);

public record RiskySignInDto(
    string Id,
    string UserPrincipalName,
    string? DisplayName,
    DateTime? CreatedDateTime,
    string? IpAddress,
    string? Location,
    string RiskLevel,
    string RiskState,
    string? RiskDetail,
    string? ClientAppUsed,
    string? DeviceDetail
);

public record SecurityAlertDto(
    string Id,
    string Title,
    string? Description,
    string Severity,
    string Status,
    string? Category,
    DateTime? CreatedDateTime,
    string? UserPrincipalName,
    List<string>? RecommendedActions
);

public record SecurityStatsDto(
    int TotalRiskyUsers,
    int HighRiskUsers,
    int MediumRiskUsers,
    int LowRiskUsers,
    int UsersAtRisk,
    int RiskySignInsLast24Hours,
    int ActiveAlerts,
    int HighSeverityAlerts,
    int MediumSeverityAlerts,
    int LowSeverityAlerts,
    int MfaRegisteredUsers,
    int MfaNotRegisteredUsers,
    double MfaRegistrationPercentage,
    DateTime LastUpdated
);

public record SecurityOverviewDto(
    SecurityScoreDto? SecureScore,
    SecurityStatsDto Stats,
    List<RiskyUserDto> RiskyUsers,
    List<RiskySignInDto> RiskySignIns,
    List<SecurityAlertDto> RecentAlerts,
    DateTime LastUpdated
);

// MFA Registration DTOs
public record MfaUserDetailDto(
    string Id,
    string UserPrincipalName,
    string? DisplayName,
    bool IsMfaRegistered,
    bool IsMfaCapable,
    string? DefaultMfaMethod,
    List<string>? MethodsRegistered,
    bool IsAdmin,
    DateTime? LastUpdated
);

public record MfaRegistrationListDto(
    List<MfaUserDetailDto> Users,
    int TotalCount,
    int MfaRegisteredCount,
    int MfaNotRegisteredCount,
    double MfaRegistrationPercentage,
    DateTime LastUpdated
);

// Permission Status DTOs
public record PermissionStatusDto(
    string PermissionName,
    string DisplayName,
    string Description,
    bool IsGranted,
    string? ErrorMessage,
    string Category
);

public record PermissionsStatusResponseDto(
    List<PermissionStatusDto> Permissions,
    int TotalPermissions,
    int GrantedPermissions,
    int MissingPermissions,
    bool AllPermissionsGranted,
    DateTime LastChecked
);

// Report DTOs
public record ReportDefinitionDto(
    string ReportType,
    string DisplayName,
    string Description,
    string Category,
    List<string> AvailableFormats,
    bool SupportsDateRange,
    bool SupportsScheduling
);

public record GenerateReportRequest(
    string ReportType,
    string? DateRange = "last30days",
    string? Format = "json",
    Dictionary<string, string>? Parameters = null
);

public record ReportResultDto(
    string ReportType,
    string DisplayName,
    DateTime GeneratedAt,
    string DateRange,
    object Data,
    ReportSummaryDto? Summary
);

public record ReportSummaryDto(
    int TotalRecords,
    Dictionary<string, object>? Highlights
);

public record ScheduledReportDto(
    string Id,
    string ReportType,
    string DisplayName,
    string Schedule,
    string Frequency,
    List<string> Recipients,
    string? DateRange,
    bool IsEnabled,
    DateTime? LastRunAt,
    DateTime? NextRunAt,
    DateTime CreatedAt
);

public record CreateScheduledReportRequest(
    string ReportType,
    string Frequency,
    string? Time,
    int? DayOfWeek,
    int? DayOfMonth,
    List<string> Recipients,
    string? DateRange
);

public record UpdateScheduledReportRequest(
    string? Frequency,
    string? Time,
    int? DayOfWeek,
    int? DayOfMonth,
    List<string>? Recipients,
    string? DateRange,
    bool? IsEnabled
);

public record ReportHistoryDto(
    string Id,
    string ReportType,
    string DisplayName,
    DateTime GeneratedAt,
    string Status,
    string? ErrorMessage,
    int? RecordCount,
    bool WasScheduled
);

// Sign-in DTOs for map visualization
public record SignInDetailDto(
    string Id,
    string UserPrincipalName,
    string? DisplayName,
    DateTime? CreatedDateTime,
    string? IpAddress,
    string? City,
    string? State,
    string? CountryOrRegion,
    double? Latitude,
    double? Longitude,
    bool IsSuccess,
    int? ErrorCode,
    string? FailureReason,
    string? ClientAppUsed,
    string? Browser,
    string? OperatingSystem,
    string? DeviceDisplayName,
    bool? IsCompliant,
    bool? IsManaged,
    string? RiskLevel,
    string? RiskState,
    bool? MfaRequired,
    string? ConditionalAccessStatus
);

public record SignInLocationDto(
    double Latitude,
    double Longitude,
    string City,
    string? State,
    string CountryOrRegion,
    int SignInCount,
    int SuccessCount,
    int FailureCount,
    List<SignInDetailDto> SignIns
);

public record SignInsMapDataDto(
    List<SignInLocationDto> Locations,
    int TotalSignIns,
    int SuccessfulSignIns,
    int FailedSignIns,
    int UniqueUsers,
    int UniqueLocations,
    DateTime StartDate,
    DateTime EndDate,
    DateTime LastUpdated
);

// SharePoint DTOs
public record SharePointSiteDto(
    string Id,
    string Name,
    string DisplayName,
    string? Description,
    string WebUrl,
    string? SiteTemplate,
    DateTime? CreatedDateTime,
    DateTime? LastModifiedDateTime,
    long StorageUsedBytes,
    long StorageAllocatedBytes,
    double StorageUsedPercentage,
    string? OwnerDisplayName,
    string? OwnerEmail,
    bool IsPersonalSite,
    int? ItemCount,
    string? Status
);

public record SharePointSiteListResultDto(
    List<SharePointSiteDto> Sites,
    int TotalCount,
    int FilteredCount,
    string? NextLink
);

public record SharePointStatsDto(
    int TotalSites,
    int TeamSites,
    int CommunicationSites,
    int PersonalSites,
    int OtherSites,
    long TotalStorageUsedBytes,
    long TotalStorageAllocatedBytes,
    double OverallStorageUsedPercentage,
    int SitesNearQuota,
    int ActiveSitesLast30Days,
    int InactiveSitesLast30Days,
    DateTime LastUpdated
);

public record SharePointOverviewDto(
    SharePointStatsDto Stats,
    List<SharePointSiteDto> LargestSites,
    List<SharePointSiteDto> RecentlyCreatedSites,
    List<SharePointSiteDto> SitesNearStorageLimit,
    DateTime LastUpdated
);

// License DTOs
public record LicenseDto(
    string SkuId,
    string SkuPartNumber,
    string DisplayName,
    int TotalUnits,
    int ConsumedUnits,
    int AvailableUnits,
    int WarningUnits,
    int SuspendedUnits,
    double UtilizationPercentage,
    string Status,
    string AppliesTo,
    bool IsTrial,
    int ServicePlanCount
);

// App Registration Credential DTOs
public record AppCredentialDetailDto(
    string AppId,
    string AppDisplayName,
    string CredentialType,
    string? KeyId,
    string? DisplayName,
    DateTime? StartDateTime,
    DateTime EndDateTime,
    int DaysUntilExpiry,
    string Status
);

public record AppCredentialStatusDto(
    int TotalApps,
    int AppsWithExpiringSecrets,
    int AppsWithExpiredSecrets,
    int AppsWithExpiringCertificates,
    int AppsWithExpiredCertificates,
    int ThresholdDays,
    List<AppCredentialDetailDto> ExpiringSecrets,
    List<AppCredentialDetailDto> ExpiredSecrets,
    List<AppCredentialDetailDto> ExpiringCertificates,
    List<AppCredentialDetailDto> ExpiredCertificates,
    DateTime LastUpdated
);

// Public Groups DTOs
public record PublicGroupDto(
    string Id,
    string DisplayName,
    DateTime? CreatedDateTime,
    string GroupType,
    bool IsTeam,
    int OwnerCount,
    int MemberCount,
    string? Description
);

public record PublicGroupsReportDto(
    int TotalPublicGroups,
    int TotalTeams,
    int TotalM365Groups,
    int GroupsWithNoOwner,
    int GroupsWithSingleOwner,
    List<PublicGroupDto> Groups,
    DateTime LastUpdated
);

// Stale Privileged Accounts DTOs
public record StalePrivilegedAccountDto(
    string UserId,
    string DisplayName,
    string UserPrincipalName,
    string Role,
    DateTime? LastSignIn,
    DateTime? LastNonInteractiveSignIn,
    string AccountStatus,
    int DaysSinceLastSignIn
);

public record StalePrivilegedAccountsReportDto(
    int TotalPrivilegedUsers,
    int TotalStaleAccounts,
    int AccountsNeverSignedIn,
    int AccountsDisabled,
    int InactiveDaysThreshold,
    List<string> MonitoredRoles,
    List<StalePrivilegedAccountDto> StaleAccounts,
    List<StalePrivilegedAccountDto> AllPrivilegedUsers,
    DateTime LastUpdated
);

// Conditional Access Break Glass DTOs
public record BreakGlassAccountDto(
    string UserPrincipalName,
    string? DisplayName,
    string? ObjectId,
    bool IsResolved
);

public record BreakGlassSettingsDto(
    List<BreakGlassAccountDto> Accounts,
    DateTime? LastUpdated,
    string? LastModifiedBy
);

public record UpdateBreakGlassSettingsRequest(
    List<string> UserPrincipalNames
);

public record ConditionalAccessPolicyDto(
    string Id,
    string DisplayName,
    string State,
    string DisplayState,
    bool IsBreakGlassExcluded,
    List<string> ExcludedBreakGlassAccounts,
    List<string> MissingBreakGlassAccounts,
    string ExclusionStatus
);

public record CABreakGlassReportDto(
    string TenantName,
    string TenantId,
    int TotalPolicies,
    int PoliciesWithFullExclusion,
    int PoliciesWithPartialExclusion,
    int PoliciesWithNoExclusion,
    int EnabledPolicies,
    int DisabledPolicies,
    int ReportOnlyPolicies,
    List<BreakGlassAccountDto> ConfiguredBreakGlassAccounts,
    List<ConditionalAccessPolicyDto> Policies,
    DateTime LastUpdated
);
