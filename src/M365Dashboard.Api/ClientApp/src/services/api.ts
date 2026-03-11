import { msalInstance, apiRequest } from './authConfig';
import type {
  UserProfile,
  UserSettings,
  WidgetConfiguration,
  DashboardSummary,
  ActiveUsersData,
  SignInAnalyticsData,
  LicenseUsageData,
  DeviceComplianceData,
  MailActivityData,
  TeamsActivityData,
  WidgetDefinition,
  DateRangePreset,
} from '../types';

const API_BASE = '/api';

// Get access token for API calls
async function getAccessToken(): Promise<string> {
  const account = msalInstance.getActiveAccount();
  if (!account) {
    throw new Error('No active account');
  }

  try {
    const response = await msalInstance.acquireTokenSilent({
      ...apiRequest,
      account,
    });
    return response.accessToken;
  } catch (error) {
    // If silent acquisition fails, try interactive
    const response = await msalInstance.acquireTokenPopup(apiRequest);
    return response.accessToken;
  }
}

// Generic fetch wrapper with authentication
async function apiFetch<T>(endpoint: string, options: RequestInit = {}): Promise<T> {
  const token = await getAccessToken();
  
  const response = await fetch(`${API_BASE}${endpoint}`, {
    ...options,
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${token}`,
      ...options.headers,
    },
  });

  if (!response.ok) {
    const error = await response.json().catch(() => ({ error: 'Unknown error' }));
    throw new Error(error.error || `API error: ${response.status}`);
  }

  return response.json();
}

// User API
export const userApi = {
  getProfile: () => apiFetch<UserProfile>('/user/profile'),
  getRoles: () => apiFetch<{ roles: string[]; isAdmin: boolean; isReader: boolean }>('/user/roles'),
};

// Settings API
export const settingsApi = {
  getSettings: () => apiFetch<UserSettings>('/settings'),
  
  updateSettings: (settings: Partial<UserSettings>) =>
    apiFetch<UserSettings>('/settings', {
      method: 'PUT',
      body: JSON.stringify(settings),
    }),

  getWidgets: () => apiFetch<WidgetConfiguration[]>('/settings/widgets'),
  
  updateWidget: (widgetType: string, config: Partial<WidgetConfiguration>) =>
    apiFetch<WidgetConfiguration>(`/settings/widgets/${widgetType}`, {
      method: 'PUT',
      body: JSON.stringify(config),
    }),

  resetWidgets: () =>
    apiFetch<WidgetConfiguration[]>('/settings/widgets/reset', {
      method: 'POST',
    }),
};

// Dashboard API
export const dashboardApi = {
  getSummary: () => apiFetch<DashboardSummary>('/dashboard/summary'),
  
  getWidgetDefinitions: () => apiFetch<WidgetDefinition[]>('/dashboard/widgets/definitions'),

  getActiveUsers: (dateRange: DateRangePreset = 'last30days') =>
    apiFetch<ActiveUsersData>(`/dashboard/widgets/active-users?dateRange=${dateRange}`),

  getSignInAnalytics: (dateRange: DateRangePreset = 'last7days') =>
    apiFetch<SignInAnalyticsData>(`/dashboard/widgets/sign-in-analytics?dateRange=${dateRange}`),

  getLicenseUsage: () =>
    apiFetch<LicenseUsageData>('/dashboard/widgets/license-usage'),

  getDeviceCompliance: () =>
    apiFetch<DeviceComplianceData>('/dashboard/widgets/device-compliance'),

  getMailActivity: (dateRange: DateRangePreset = 'last30days') =>
    apiFetch<MailActivityData>(`/dashboard/widgets/mail-activity?dateRange=${dateRange}`),

  getTeamsActivity: (dateRange: DateRangePreset = 'last30days') =>
    apiFetch<TeamsActivityData>(`/dashboard/widgets/teams-activity?dateRange=${dateRange}`),

  refreshCache: (metricType: string = 'all') =>
    apiFetch<{ message: string }>(`/dashboard/cache/refresh?metricType=${metricType}`, {
      method: 'POST',
    }),
};
