import { useState, useEffect, useCallback } from 'react';
import { dashboardApi } from '../services/api';
import { useSettings } from '../contexts/AppContext';
import type {
  DashboardSummary,
  ActiveUsersData,
  SignInAnalyticsData,
  LicenseUsageData,
  DeviceComplianceData,
  MailActivityData,
  TeamsActivityData,
  DateRangePreset,
} from '../types';

interface UseDataResult<T> {
  data: T | null;
  isLoading: boolean;
  error: string | null;
  refresh: () => Promise<void>;
}

// Generic hook for data fetching with caching
function useApiData<T>(
  fetchFn: () => Promise<T>,
  refreshInterval?: number
): UseDataResult<T> {
  const [data, setData] = useState<T | null>(null);
  const [isLoading, setIsLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  const refresh = useCallback(async () => {
    setIsLoading(true);
    setError(null);
    try {
      const result = await fetchFn();
      setData(result);
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Failed to load data');
    } finally {
      setIsLoading(false);
    }
  }, [fetchFn]);

  useEffect(() => {
    refresh();

    if (refreshInterval && refreshInterval > 0) {
      const interval = setInterval(refresh, refreshInterval * 1000);
      return () => clearInterval(interval);
    }
  }, [refresh, refreshInterval]);

  return { data, isLoading, error, refresh };
}

// Dashboard Summary Hook
export function useDashboardSummary(): UseDataResult<DashboardSummary> {
  const { settings } = useSettings();
  return useApiData(
    dashboardApi.getSummary,
    settings?.refreshIntervalSeconds
  );
}

// Active Users Hook
export function useActiveUsers(dateRange?: DateRangePreset): UseDataResult<ActiveUsersData> {
  const { settings } = useSettings();
  const range = dateRange || settings?.dateRangePreference || 'last30days';
  
  return useApiData(
    useCallback(() => dashboardApi.getActiveUsers(range), [range]),
    settings?.refreshIntervalSeconds
  );
}

// Sign-in Analytics Hook
export function useSignInAnalytics(dateRange?: DateRangePreset): UseDataResult<SignInAnalyticsData> {
  const { settings } = useSettings();
  const range = dateRange || 'last7days';
  
  return useApiData(
    useCallback(() => dashboardApi.getSignInAnalytics(range), [range]),
    settings?.refreshIntervalSeconds
  );
}

// License Usage Hook
export function useLicenseUsage(): UseDataResult<LicenseUsageData> {
  const { settings } = useSettings();
  return useApiData(
    dashboardApi.getLicenseUsage,
    settings?.refreshIntervalSeconds
  );
}

// Device Compliance Hook
export function useDeviceCompliance(): UseDataResult<DeviceComplianceData> {
  const { settings } = useSettings();
  return useApiData(
    dashboardApi.getDeviceCompliance,
    settings?.refreshIntervalSeconds
  );
}

// Mail Activity Hook
export function useMailActivity(dateRange?: DateRangePreset): UseDataResult<MailActivityData> {
  const { settings } = useSettings();
  const range = dateRange || settings?.dateRangePreference || 'last30days';
  
  return useApiData(
    useCallback(() => dashboardApi.getMailActivity(range), [range]),
    settings?.refreshIntervalSeconds
  );
}

// Teams Activity Hook
export function useTeamsActivity(dateRange?: DateRangePreset): UseDataResult<TeamsActivityData> {
  const { settings } = useSettings();
  const range = dateRange || settings?.dateRangePreference || 'last30days';
  
  return useApiData(
    useCallback(() => dashboardApi.getTeamsActivity(range), [range]),
    settings?.refreshIntervalSeconds
  );
}
