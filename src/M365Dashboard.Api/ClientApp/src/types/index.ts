// User and Authentication Types
export interface UserProfile {
  id: string;
  displayName: string;
  email: string;
  jobTitle?: string;
  department?: string;
  profilePhoto?: string;
  roles: string[];
}

export interface UserRoles {
  roles: string[];
  isAdmin: boolean;
  isReader: boolean;
}

// Settings Types
export interface UserSettings {
  theme: 'light' | 'dark' | 'system';
  refreshIntervalSeconds: number;
  dateRangePreference: DateRangePreset;
  showWelcomeMessage: boolean;
  compactMode: boolean;
}

export type DateRangePreset = 
  | 'last7days' 
  | 'last30days' 
  | 'last90days' 
  | 'thismonth' 
  | 'lastmonth' 
  | 'custom';

// Widget Types
export interface WidgetConfiguration {
  id: number;
  widgetType: WidgetType;
  isEnabled: boolean;
  displayOrder: number;
  gridColumn: number;
  gridRow: number;
  gridWidth: number;
  gridHeight: number;
  customSettings?: Record<string, unknown>;
}

export type WidgetType = 
  | 'active-users'
  | 'sign-in-analytics'
  | 'license-usage'
  | 'device-compliance'
  | 'mail-activity'
  | 'teams-activity';

export interface WidgetDefinition {
  type: WidgetType;
  name: string;
  description: string;
  category: string;
  requiredPermissions: string[];
  defaultWidth: number;
  defaultHeight: number;
}

// Dashboard Data Types
export interface DashboardSummary {
  activeUsers: number;
  signInSuccessRate: number;
  licenseUtilization: number;
  deviceComplianceRate: number;
  lastUpdated: string;
}

export interface ActiveUsersData {
  dailyActiveUsers: number;
  weeklyActiveUsers: number;
  monthlyActiveUsers: number;
  trend: DailyActiveUsersTrend[];
  lastUpdated: string;
}

export interface DailyActiveUsersTrend {
  date: string;
  count: number;
}

export interface SignInAnalyticsData {
  totalSignIns: number;
  successfulSignIns: number;
  failedSignIns: number;
  riskySignIns: number;
  successRate: number;
  trend: SignInTrend[];
  topLocations: SignInLocation[];
  lastUpdated: string;
}

export interface SignInTrend {
  date: string;
  successful: number;
  failed: number;
}

export interface SignInLocation {
  location: string;
  count: number;
}

export interface LicenseUsageData {
  licenses: LicenseSku[];
  totalConsumed: number;
  totalAvailable: number;
  overallUtilization: number;
  lastUpdated: string;
}

export interface LicenseSku {
  skuId: string;
  skuName: string;
  consumedUnits: number;
  prepaidUnits: number;
  utilizationPercent: number;
}

export interface DeviceComplianceData {
  totalDevices: number;
  compliantDevices: number;
  nonCompliantDevices: number;
  unknownDevices: number;
  complianceRate: number;
  byPlatform: DeviceByPlatform[];
  lastUpdated: string;
}

export interface DeviceByPlatform {
  platform: string;
  total: number;
  compliant: number;
  nonCompliant: number;
}

export interface MailActivityData {
  totalEmailsSent: number;
  totalEmailsReceived: number;
  totalEmailsRead: number;
  trend: MailActivityTrend[];
  lastUpdated: string;
}

export interface MailActivityTrend {
  date: string;
  sent: number;
  received: number;
}

export interface TeamsActivityData {
  totalMessages: number;
  totalCalls: number;
  totalMeetings: number;
  activeUsers: number;
  trend: TeamsActivityTrend[];
  lastUpdated: string;
}

export interface TeamsActivityTrend {
  date: string;
  messages: number;
  calls: number;
  meetings: number;
}

// Layout Types
export interface DashboardLayout {
  id: number;
  name: string;
  isDefault: boolean;
  widgets: WidgetConfiguration[];
  createdAt: string;
  updatedAt: string;
}

// API Response Types
export interface ApiError {
  error: string;
  message?: string;
}
