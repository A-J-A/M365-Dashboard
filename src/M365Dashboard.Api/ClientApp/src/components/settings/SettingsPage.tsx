import React, { useState, useEffect, useCallback, useMemo } from 'react';
import {
  Switch,
  Dropdown,
  Option,
  Button,
  Spinner,
  Toast,
  Toaster,
  useToastController,
  useId,
  Badge,
  Textarea,
  Tab,
  TabList,
  SelectTabData,
  SelectTabEvent,
} from '@fluentui/react-components';
import {
  Settings24Regular,
  Grid24Regular,
  Color24Regular,
  Clock24Regular,
  ArrowReset24Regular,
  ShieldCheckmark24Regular,
  Checkmark16Regular,
  Dismiss16Regular,
  ArrowSync24Regular,
  Open16Regular,
  PlugConnected24Regular,
  Globe24Regular,
  Info24Regular,
  PersonKey24Regular,
  Shield24Regular,
  Server24Regular,
  Options24Regular,
  ArrowDownload24Regular,
  Delete16Regular,
  DocumentRegular,
  SaveRegular,
  ImageRegular,
  DeleteRegular,
  CheckmarkCircleFilled,
  DismissCircleFilled,
  ColorRegular,
  BuildingRegular,
  ChevronDown20Regular,
  ChevronUp20Regular,
} from '@fluentui/react-icons';
import { useSettings, useTheme, useUser, useAppContext } from '../../contexts/AppContext';
import type { DateRangePreset, WidgetConfiguration } from '../../types';

interface PermissionStatus {
  permissionName: string;
  displayName: string;
  description: string;
  isGranted: boolean;
  errorMessage?: string;
  category: string;
}

interface PermissionsStatusResponse {
  permissions: PermissionStatus[];
  totalPermissions: number;
  grantedPermissions: number;
  missingPermissions: number;
  allPermissionsGranted: boolean;
  lastChecked: string;
}

interface ExternalServiceStatus {
  name: string;
  displayName: string;
  description: string;
  isConfigured: boolean;
  isWorking: boolean;
  errorMessage?: string;
  setupUrl?: string;
  docsUrl?: string;
}

interface SkuMappingStatus {
  lastRefreshed: string | null;
  totalMappings: number;
  freeTrialSkusCount: number;
  isRefreshing: boolean;
  lastError: string | null;
  nextScheduledRefresh: string | null;
}

interface BreakGlassAccount {
  userPrincipalName: string;
  displayName: string | null;
  objectId: string | null;
  isResolved: boolean;
}

interface BreakGlassSettings {
  accounts: BreakGlassAccount[];
  lastUpdated: string | null;
  lastModifiedBy: string | null;
}

interface UpdateStatus {
  currentVersion: string;
  latestVersion: string | null;
  updateAvailable: boolean;
  releaseNotes: string | null;
  releaseUrl: string | null;
  publishedAt: string | null;
  error: string | null;
  updateConfigured: boolean;
}

type SettingsTab = 'general' | 'security' | 'system' | 'reports';

export function SettingsPage() {
  const { settings, widgets, updateSettings, updateWidget, resetWidgets, isLoading } = useSettings();
  const { theme, setTheme } = useTheme();
  const { isAdmin } = useUser();
  const { getAccessToken } = useAppContext();
  const [selectedTab, setSelectedTab] = useState<SettingsTab>('general');
  const [isSaving, setIsSaving] = useState(false);
  const [permissionsStatus, setPermissionsStatus] = useState<PermissionsStatusResponse | null>(null);
  const [isLoadingPermissions, setIsLoadingPermissions] = useState(false);
  const [externalServices, setExternalServices] = useState<ExternalServiceStatus[]>([]);
  const [isLoadingServices, setIsLoadingServices] = useState(false);
  const [skuMappingStatus, setSkuMappingStatus] = useState<SkuMappingStatus | null>(null);
  const [isLoadingSkuStatus, setIsLoadingSkuStatus] = useState(false);
  const [isRefreshingSkuMappings, setIsRefreshingSkuMappings] = useState(false);
  const [updateStatus, setUpdateStatus] = useState<UpdateStatus | null>(null);
  const [isCheckingUpdate, setIsCheckingUpdate] = useState(false);
  const [isApplyingUpdate, setIsApplyingUpdate] = useState(false);
  const [updateMessage, setUpdateMessage] = useState<{ type: 'success' | 'error'; text: string } | null>(null);
  const [breakGlassSettings, setBreakGlassSettings] = useState<BreakGlassSettings | null>(null);
  const [isLoadingBreakGlass, setIsLoadingBreakGlass] = useState(false);
  const [hasLoadedBreakGlass, setHasLoadedBreakGlass] = useState(false);
  const [isSavingBreakGlass, setIsSavingBreakGlass] = useState(false);
  const [notification, setNotification] = useState<{ message: string; type: 'success' | 'error' } | null>(null);
  const [breakGlassInput, setBreakGlassInput] = useState('');

  // Report settings state
  const [reportSettings, setReportSettings] = useState<{
    companyName: string;
    reportTitle: string;
    logoBase64: string | null;
    logoContentType: string | null;
    primaryColor: string;
    accentColor: string;
    showInfoGraphics: boolean;
    footerText: string | null;
    updatedAt: string;
  }>({
    companyName: 'M365 Dashboard',
    reportTitle: 'Microsoft 365 Security Assessment',
    logoBase64: null,
    logoContentType: null,
    primaryColor: '#0078d4',
    accentColor: '#e07c3a',
    showInfoGraphics: true,
    footerText: null,
    updatedAt: new Date().toISOString(),
  });
  const [reportLoading, setReportLoading] = useState(false);
  const [reportSaving, setReportSaving] = useState(false);
  const [reportUploading, setReportUploading] = useState(false);
  const [reportMessage, setReportMessage] = useState<{ type: 'success' | 'error'; text: string } | null>(null);
  const [hasLoadedReport, setHasLoadedReport] = useState(false);
  const fileInputRef = React.useRef<HTMLInputElement>(null);
  const toasterId = useId('toaster');
  const { dispatchToast } = useToastController(toasterId);

  // Simple notification that auto-dismisses
  const showNotification = useCallback((message: string, type: 'success' | 'error' = 'success') => {
    setNotification({ message, type });
    setTimeout(() => setNotification(null), 3000);
  }, []);

  const showToast = useCallback((message: string, intent: 'success' | 'error' = 'success') => {
    dispatchToast(
      <Toast>
        <span>{message}</span>
      </Toast>,
      { intent }
    );
  }, [dispatchToast]);

  const loadPermissionsStatus = async () => {
    setIsLoadingPermissions(true);
    try {
      const token = await getAccessToken();
      const response = await fetch('/api/permissions/status', {
        headers: {
          'Authorization': `Bearer ${token}`,
          'Content-Type': 'application/json',
        },
      });
      if (!response.ok) {
        throw new Error('Failed to fetch permissions status');
      }
      const data: PermissionsStatusResponse = await response.json();
      setPermissionsStatus(data);
    } catch (error) {
      console.error('Failed to load permissions status:', error);
      showToast('Failed to check permissions status', 'error');
    } finally {
      setIsLoadingPermissions(false);
    }
  };

  const loadExternalServices = async () => {
    setIsLoadingServices(true);
    try {
      const token = await getAccessToken();
      const azureMapsStatus = await checkAzureMaps(token);
      setExternalServices([azureMapsStatus]);
    } catch (error) {
      console.error('Failed to load external services status:', error);
    } finally {
      setIsLoadingServices(false);
    }
  };

  const loadSkuMappingStatus = async () => {
    setIsLoadingSkuStatus(true);
    try {
      const token = await getAccessToken();
      const response = await fetch('/api/licenses/mapping-status', {
        headers: {
          'Authorization': `Bearer ${token}`,
          'Content-Type': 'application/json',
        },
      });
      if (!response.ok) {
        throw new Error('Failed to fetch SKU mapping status');
      }
      const data: SkuMappingStatus = await response.json();
      setSkuMappingStatus(data);
    } catch (error) {
      console.error('Failed to load SKU mapping status:', error);
    } finally {
      setIsLoadingSkuStatus(false);
    }
  };

  const refreshSkuMappings = async () => {
    setIsRefreshingSkuMappings(true);
    try {
      const token = await getAccessToken();
      const response = await fetch('/api/licenses/refresh-mappings', {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${token}`,
          'Content-Type': 'application/json',
        },
      });
      if (!response.ok) {
        throw new Error('Failed to refresh SKU mappings');
      }
      const data = await response.json();
      setSkuMappingStatus(data.status);
      showToast('SKU mappings refreshed successfully');
    } catch (error) {
      console.error('Failed to refresh SKU mappings:', error);
      showToast('Failed to refresh SKU mappings', 'error');
    } finally {
      setIsRefreshingSkuMappings(false);
    }
  };

  const checkForUpdates = useCallback(async () => {
    setIsCheckingUpdate(true);
    try {
      const token = await getAccessToken();
      const response = await fetch('/api/update/check', {
        headers: { 'Authorization': `Bearer ${token}` },
      });
      if (response.ok) {
        setUpdateStatus(await response.json());
      }
    } catch (error) {
      console.error('Failed to check for updates:', error);
    } finally {
      setIsCheckingUpdate(false);
    }
  }, [getAccessToken]);

  const applyUpdate = async (version: string) => {
    if (!confirm(`Apply update to ${version}? The dashboard will restart and be unavailable for about 60 seconds.`)) return;
    setIsApplyingUpdate(true);
    setUpdateMessage(null);
    try {
      const token = await getAccessToken();
      const response = await fetch('/api/update/apply', {
        method: 'POST',
        headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' },
        body: JSON.stringify({ version }),
      });
      const data = await response.json();
      if (response.ok) {
        setUpdateMessage({ type: 'success', text: data.message ?? `Update to ${version} initiated. The page will reload in 90 seconds.` });
        // Reload after 90s to pick up the new version
        setTimeout(() => window.location.reload(), 90000);
      } else {
        setUpdateMessage({ type: 'error', text: data.error ?? 'Failed to apply update' });
      }
    } catch (error) {
      setUpdateMessage({ type: 'error', text: 'Failed to apply update' });
    } finally {
      setIsApplyingUpdate(false);
    }
  };

  const loadBreakGlassSettings = async (showSpinner = true) => {
    if (showSpinner && !hasLoadedBreakGlass) {
      setIsLoadingBreakGlass(true);
    }
    try {
      const token = await getAccessToken();
      const response = await fetch('/api/settings/breakglass', {
        headers: {
          'Authorization': `Bearer ${token}`,
          'Content-Type': 'application/json',
        },
      });
      if (!response.ok) {
        throw new Error('Failed to fetch break glass settings');
      }
      const data: BreakGlassSettings = await response.json();
      setBreakGlassSettings(data);
      setBreakGlassInput(data.accounts.map(a => a.userPrincipalName).join('\n'));
      setHasLoadedBreakGlass(true);
    } catch (error) {
      console.error('Failed to load break glass settings:', error);
    } finally {
      setIsLoadingBreakGlass(false);
    }
  };

  const loadReportSettings = useCallback(async () => {
    try {
      setReportLoading(true);
      const token = await getAccessToken();
      const response = await fetch('/api/settings/report', { headers: { 'Authorization': `Bearer ${token}` } });
      if (response.ok) {
        setReportSettings(await response.json());
        setHasLoadedReport(true);
      }
    } catch (err) {
      console.error('Error loading report settings:', err);
    } finally {
      setReportLoading(false);
    }
  }, [getAccessToken]);

  const saveReportSettings = async () => {
    try {
      setReportSaving(true);
      setReportMessage(null);
      const token = await getAccessToken();
      const response = await fetch('/api/settings/report', {
        method: 'POST',
        headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' },
        body: JSON.stringify(reportSettings),
      });
      if (response.ok) {
        setReportSettings(await response.json());
        setReportMessage({ type: 'success', text: 'Settings saved successfully' });
      } else {
        throw new Error('Failed to save');
      }
    } catch {
      setReportMessage({ type: 'error', text: 'Failed to save settings' });
    } finally {
      setReportSaving(false);
    }
  };

  const uploadReportLogo = async (file: File) => {
    try {
      setReportUploading(true);
      setReportMessage(null);
      const token = await getAccessToken();
      const formData = new FormData();
      formData.append('file', file);
      const response = await fetch('/api/settings/report/logo', {
        method: 'POST',
        headers: { 'Authorization': `Bearer ${token}` },
        body: formData,
      });
      if (response.ok) {
        const data = await response.json();
        setReportSettings(prev => ({ ...prev, logoBase64: data.logoBase64, logoContentType: data.contentType }));
        setReportMessage({ type: 'success', text: 'Logo uploaded successfully' });
      } else {
        const err = await response.json();
        throw new Error(err.error || 'Failed to upload');
      }
    } catch (err) {
      setReportMessage({ type: 'error', text: err instanceof Error ? err.message : 'Failed to upload logo' });
    } finally {
      setReportUploading(false);
    }
  };

  const importEntraLogo = async () => {
    try {
      setReportUploading(true);
      setReportMessage(null);
      const token = await getAccessToken();
      const response = await fetch('/api/settings/report/logo/entra', {
        method: 'POST',
        headers: { 'Authorization': `Bearer ${token}` },
      });
      const data = await response.json();
      if (response.ok) {
        setReportSettings(prev => ({ ...prev, logoBase64: data.logoBase64, logoContentType: data.contentType }));
        setReportMessage({ type: 'success', text: 'Entra branding logo imported successfully' });
      } else {
        throw new Error(data.message || data.error || 'Failed to import');
      }
    } catch (err) {
      setReportMessage({ type: 'error', text: err instanceof Error ? err.message : 'Failed to import Entra logo' });
    } finally {
      setReportUploading(false);
    }
  };

  const removeReportLogo = async () => {
    try {
      setReportUploading(true);
      setReportMessage(null);
      const token = await getAccessToken();
      const response = await fetch('/api/settings/report/logo', {
        method: 'DELETE',
        headers: { 'Authorization': `Bearer ${token}` },
      });
      if (response.ok) {
        setReportSettings(prev => ({ ...prev, logoBase64: null, logoContentType: null }));
        setReportMessage({ type: 'success', text: 'Logo removed' });
      }
    } catch {
      setReportMessage({ type: 'error', text: 'Failed to remove logo' });
    } finally {
      setReportUploading(false);
    }
  };

  const removeBreakGlassAccount = useCallback((upnToRemove: string) => {
    const lowerUpn = upnToRemove.toLowerCase();
    
    // Calculate updated UPNs
    const updatedUpns = breakGlassInput
      .split(/[\n,;]/)
      .map(u => u.trim())
      .filter(u => u.length > 0 && u.toLowerCase() !== lowerUpn);
    
    // Store previous state for rollback
    const previousInput = breakGlassInput;
    const previousAccounts = breakGlassSettings?.accounts || [];
    
    // Optimistically update UI immediately
    setBreakGlassInput(updatedUpns.join('\n'));
    setBreakGlassSettings(prev => prev ? {
      ...prev,
      accounts: prev.accounts.filter(a => a.userPrincipalName.toLowerCase() !== lowerUpn),
    } : null);
    
    // Save in background
    getAccessToken().then(async (token) => {
      try {
        const response = await fetch('/api/settings/breakglass', {
          method: 'PUT',
          headers: {
            'Authorization': `Bearer ${token}`,
            'Content-Type': 'application/json',
          },
          body: JSON.stringify({ userPrincipalNames: updatedUpns }),
        });
        if (!response.ok) throw new Error('Failed to save');
        showNotification('Account removed');
      } catch (error) {
        console.error('Failed to remove account:', error);
        showNotification('Failed to remove account', 'error');
        // Restore on failure
        setBreakGlassInput(previousInput);
        setBreakGlassSettings(prev => prev ? {
          ...prev,
          accounts: previousAccounts,
        } : null);
      }
    });
  }, [breakGlassInput, breakGlassSettings, getAccessToken]);

  const saveBreakGlassSettings = useCallback(() => {
    const previousSettings = breakGlassSettings;
    const previousInput = breakGlassInput;
    const upns = breakGlassInput
      .split(/[\n,;]/)
      .map(upn => upn.trim())
      .filter(upn => upn.length > 0);

    // Optimistically update UI - show accounts as "pending"
    setBreakGlassSettings(prev => ({
      accounts: upns.map(upn => ({
        userPrincipalName: upn,
        displayName: null,
        objectId: null,
        isResolved: prev?.accounts.find(a => a.userPrincipalName.toLowerCase() === upn.toLowerCase())?.isResolved ?? false,
      })),
      lastUpdated: new Date().toISOString(),
      lastModifiedBy: 'saving...',
    }));

    // Save in background
    getAccessToken().then(async (token) => {
      try {
        const response = await fetch('/api/settings/breakglass', {
          method: 'PUT',
          headers: {
            'Authorization': `Bearer ${token}`,
            'Content-Type': 'application/json',
          },
          body: JSON.stringify({ userPrincipalNames: upns }),
        });

        if (!response.ok) {
          throw new Error('Failed to save break glass settings');
        }

        const data: BreakGlassSettings = await response.json();
        // Update with resolved accounts from server
        setBreakGlassSettings(data);
        setBreakGlassInput(data.accounts.map(a => a.userPrincipalName).join('\n'));
        showNotification('Break glass accounts saved');
      } catch (error) {
        console.error('Failed to save break glass settings:', error);
        showNotification('Failed to save break glass accounts', 'error');
        // Restore previous state on error
        setBreakGlassSettings(previousSettings);
        setBreakGlassInput(previousInput);
      }
    });
  }, [breakGlassInput, breakGlassSettings, getAccessToken]);

  const checkAzureMaps = async (token: string): Promise<ExternalServiceStatus> => {
    try {
      const response = await fetch('/api/config/azure-maps-key', {
        headers: {
          'Authorization': `Bearer ${token}`,
          'Content-Type': 'application/json',
        },
      });

      if (!response.ok) {
        return {
          name: 'azure-maps',
          displayName: 'Azure Maps',
          description: 'Required for Sign-ins Map visualization',
          isConfigured: false,
          isWorking: false,
          errorMessage: 'Failed to check configuration',
          setupUrl: 'https://portal.azure.com/#create/Microsoft.Maps',
          docsUrl: 'https://learn.microsoft.com/en-us/azure/azure-maps/',
        };
      }

      const data = await response.json();
      const isConfigured = data.configured === true && !!data.key;

      let isWorking = false;
      let errorMessage: string | undefined;

      if (isConfigured) {
        try {
          const testUrl = `https://atlas.microsoft.com/search/address/json?api-version=1.0&subscription-key=${data.key}&query=London`;
          const testResponse = await fetch(testUrl);
          
          if (testResponse.ok) {
            isWorking = true;
          } else if (testResponse.status === 401) {
            errorMessage = 'Invalid subscription key';
          } else if (testResponse.status === 403) {
            errorMessage = 'Key does not have required permissions';
          } else {
            errorMessage = `API returned status ${testResponse.status}`;
          }
        } catch {
          errorMessage = 'Failed to validate key';
        }
      }

      return {
        name: 'azure-maps',
        displayName: 'Azure Maps',
        description: 'Required for Sign-ins Map visualization',
        isConfigured,
        isWorking,
        errorMessage: isConfigured ? errorMessage : 'Subscription key not configured',
        setupUrl: 'https://portal.azure.com/#create/Microsoft.Maps',
        docsUrl: 'https://learn.microsoft.com/en-us/azure/azure-maps/',
      };
    } catch {
      return {
        name: 'azure-maps',
        displayName: 'Azure Maps',
        description: 'Required for Sign-ins Map visualization',
        isConfigured: false,
        isWorking: false,
        errorMessage: 'Failed to check configuration',
        setupUrl: 'https://portal.azure.com/#create/Microsoft.Maps',
        docsUrl: 'https://learn.microsoft.com/en-us/azure/azure-maps/',
      };
    }
  };

  useEffect(() => {
    // Load data based on selected tab
    if (selectedTab === 'general') {
      // Settings are loaded by the context
    } else if (selectedTab === 'security') {
      loadPermissionsStatus();
      loadBreakGlassSettings();
    } else if (selectedTab === 'system') {
      loadExternalServices();
      loadSkuMappingStatus();
      checkForUpdates();
    } else if (selectedTab === 'reports' && !hasLoadedReport) {
      loadReportSettings();
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [selectedTab]);

  const handleTabSelect = (_event: SelectTabEvent, data: SelectTabData) => {
    setSelectedTab(data.value as SettingsTab);
  };

  const handleThemeChange = async (newTheme: 'light' | 'dark' | 'system') => {
    setIsSaving(true);
    try {
      await updateSettings({ theme: newTheme });
      setTheme(newTheme);
      showToast('Theme updated');
    } catch {
      showToast('Failed to update theme', 'error');
    } finally {
      setIsSaving(false);
    }
  };

  const handleRefreshIntervalChange = async (seconds: number) => {
    setIsSaving(true);
    try {
      await updateSettings({ refreshIntervalSeconds: seconds });
      showToast('Refresh interval updated');
    } catch {
      showToast('Failed to update refresh interval', 'error');
    } finally {
      setIsSaving(false);
    }
  };

  const handleDateRangeChange = async (range: DateRangePreset) => {
    setIsSaving(true);
    try {
      await updateSettings({ dateRangePreference: range });
      showToast('Date range preference updated');
    } catch {
      showToast('Failed to update date range', 'error');
    } finally {
      setIsSaving(false);
    }
  };

  const handleWidgetToggle = async (widgetType: string, enabled: boolean) => {
    setIsSaving(true);
    try {
      await updateWidget(widgetType, { isEnabled: enabled });
      showToast(`Widget ${enabled ? 'enabled' : 'disabled'}`);
    } catch {
      showToast('Failed to update widget', 'error');
    } finally {
      setIsSaving(false);
    }
  };

  const handleResetWidgets = async () => {
    if (!confirm('Reset all widgets to default configuration?')) return;
    
    setIsSaving(true);
    try {
      await resetWidgets();
      showToast('Widgets reset to defaults');
    } catch {
      showToast('Failed to reset widgets', 'error');
    } finally {
      setIsSaving(false);
    }
  };

  if (isLoading) {
    return (
      <div className="flex items-center justify-center min-h-[400px]">
        <Spinner size="large" label="Loading settings..." />
      </div>
    );
  }

  // Group permissions by category
  const groupedPermissions = permissionsStatus?.permissions.reduce((acc, perm) => {
    if (!acc[perm.category]) {
      acc[perm.category] = [];
    }
    acc[perm.category].push(perm);
    return acc;
  }, {} as Record<string, PermissionStatus[]>) ?? {};

  const categoryOrder = ['Core', 'Devices', 'Mail & Reports', 'Security', 'Exchange Online', 'Defender for Endpoint', 'SharePoint', 'Teams Phone'];

  const configuredServices = externalServices.filter(s => s.isConfigured && s.isWorking).length;
  const totalServices = externalServices.length;

  return (
    <div className="max-w-4xl mx-auto p-6">
      <Toaster toasterId={toasterId} />

      {/* Header */}
      <div className="mb-6">
        <h1 className="text-2xl font-bold text-gray-900 dark:text-white">Settings</h1>
        <p className="mt-1 text-sm text-gray-500 dark:text-gray-400">
          Customize your dashboard experience
        </p>
      </div>

      {/* Tabs */}
      <div className="mb-6">
        <TabList selectedValue={selectedTab} onTabSelect={handleTabSelect} size="large">
          <Tab value="general" icon={<Options24Regular />}>
            General
          </Tab>
          <Tab value="security" icon={<Shield24Regular />}>
            Security
          </Tab>
          <Tab value="system" icon={<Server24Regular />}>
            System
          </Tab>
          <Tab value="reports" icon={<DocumentRegular />}>
            Reports
          </Tab>
        </TabList>
      </div>

      {/* Tab Content */}
      <div className="space-y-6">
        {/* General Tab */}
        {selectedTab === 'general' && (
          <>
            {/* Appearance Section */}
            <SettingsSection
              icon={Color24Regular}
              title="Appearance"
              description="Customize the look and feel"
            >
              <SettingsRow label="Theme" description="Choose your preferred color scheme">
                <Dropdown
                  selectedOptions={[theme]}
                  value={{ light: 'Light', dark: 'Dark', system: 'System' }[theme]}
                  onOptionSelect={(_, data) => handleThemeChange(data.optionValue as 'light' | 'dark' | 'system')}
                  disabled={isSaving}
                >
                  <Option value="light">Light</Option>
                  <Option value="dark">Dark</Option>
                  <Option value="system">System</Option>
                </Dropdown>
              </SettingsRow>

              <SettingsRow label="Compact Mode" description="Use a more condensed layout">
                <Switch
                  checked={settings?.compactMode ?? false}
                  onChange={async (_, data) => {
                    setIsSaving(true);
                    try {
                      await updateSettings({ compactMode: data.checked });
                      showToast('Compact mode updated');
                    } catch {
                      showToast('Failed to update', 'error');
                    } finally {
                      setIsSaving(false);
                    }
                  }}
                  disabled={isSaving}
                />
              </SettingsRow>
            </SettingsSection>

            {/* Data & Refresh Section */}
            <SettingsSection
              icon={Clock24Regular}
              title="Data & Refresh"
              description="Control how data is loaded and refreshed"
            >
              <SettingsRow label="Auto-refresh Interval" description="How often to refresh dashboard data">
                <Dropdown
                  selectedOptions={[`${settings?.refreshIntervalSeconds ?? 300}`]}
                  value={({'60':'1 minute','120':'2 minutes','300':'5 minutes','600':'10 minutes','900':'15 minutes','0':'Manual only'} as Record<string,string>)[`${settings?.refreshIntervalSeconds ?? 300}`] ?? '5 minutes'}
                  onOptionSelect={(_, data) => handleRefreshIntervalChange(Number(data.optionValue))}
                  disabled={isSaving}
                >
                  <Option value="60">1 minute</Option>
                  <Option value="120">2 minutes</Option>
                  <Option value="300">5 minutes</Option>
                  <Option value="600">10 minutes</Option>
                  <Option value="900">15 minutes</Option>
                  <Option value="0">Manual only</Option>
                </Dropdown>
              </SettingsRow>

              <SettingsRow label="Default Date Range" description="Default time period for reports">
                <Dropdown
                  selectedOptions={[settings?.dateRangePreference ?? 'last30days']}
                  value={({'last7days':'Last 7 days','last30days':'Last 30 days','last90days':'Last 90 days','thismonth':'This month','lastmonth':'Last month','custom':'Custom'} as Record<string,string>)[settings?.dateRangePreference ?? 'last30days'] ?? 'Last 30 days'}
                  onOptionSelect={(_, data) => handleDateRangeChange(data.optionValue as DateRangePreset)}
                  disabled={isSaving}
                >
                  <Option value="last7days">Last 7 days</Option>
                  <Option value="last30days">Last 30 days</Option>
                  <Option value="last90days">Last 90 days</Option>
                  <Option value="thismonth">This month</Option>
                  <Option value="lastmonth">Last month</Option>
                </Dropdown>
              </SettingsRow>
            </SettingsSection>

            {/* Widgets Section */}
            <SettingsSection
              icon={Grid24Regular}
              title="Dashboard Widgets"
              description="Choose which widgets to display on your dashboard"
            >
              <div className="space-y-3">
                {widgets.map((widget) => (
                  <WidgetToggle
                    key={widget.widgetType}
                    widget={widget}
                    onToggle={handleWidgetToggle}
                    disabled={isSaving}
                  />
                ))}
              </div>

              <div className="pt-4 border-t border-gray-200 dark:border-gray-700">
                <Button
                  appearance="secondary"
                  icon={<ArrowReset24Regular />}
                  onClick={handleResetWidgets}
                  disabled={isSaving}
                >
                  Reset to Defaults
                </Button>
              </div>
            </SettingsSection>
          </>
        )}

        {/* Security Tab */}
        {selectedTab === 'security' && (
          <>
            {/* Break Glass Accounts Section */}
            <SettingsSection
              icon={PersonKey24Regular}
              title="Break Glass Accounts"
              description="Emergency access accounts that should be excluded from all Conditional Access policies"
              headerAction={
                <Button
                  appearance="subtle"
                  icon={<ArrowSync24Regular />}
                  onClick={() => loadBreakGlassSettings(false)}
                  disabled={isLoadingBreakGlass}
                  size="small"
                >
                  Refresh
                </Button>
              }
            >
              {isLoadingBreakGlass && !hasLoadedBreakGlass ? (
                <div className="flex items-center justify-center py-8">
                  <Spinner size="medium" label="Loading break glass settings..." />
                </div>
              ) : (
                <div className="space-y-4">
                  {/* Notification Banner */}
                  {notification && (
                    <div className={`p-3 rounded-lg flex items-center justify-between ${
                      notification.type === 'success' 
                        ? 'bg-green-50 dark:bg-green-900/20 border border-green-200 dark:border-green-700 text-green-800 dark:text-green-200'
                        : 'bg-red-50 dark:bg-red-900/20 border border-red-200 dark:border-red-700 text-red-800 dark:text-red-200'
                    }`}>
                      <span className="text-sm font-medium">{notification.message}</span>
                      <button 
                        onClick={() => setNotification(null)}
                        className="p-1 hover:opacity-70"
                      >
                        <Dismiss16Regular className="w-4 h-4" />
                      </button>
                    </div>
                  )}

                  {/* Info Banner */}
                  <div className="p-4 bg-blue-50 dark:bg-blue-900/20 border border-blue-200 dark:border-blue-700 rounded-lg">
                    <p className="text-sm text-blue-800 dark:text-blue-200">
                      <strong>What are break glass accounts?</strong> Emergency access accounts that bypass Conditional Access policies
                      to prevent lockouts during emergencies or misconfigurations.{' '}
                      <a
                        href="https://learn.microsoft.com/en-us/entra/identity/role-based-access-control/security-emergency-access"
                        target="_blank"
                        rel="noopener noreferrer"
                        className="underline hover:no-underline inline-flex items-center gap-1"
                      >
                        Learn more <Open16Regular className="w-3 h-3" />
                      </a>
                    </p>
                  </div>

                  {/* Main Content Grid */}
                  <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
                    {/* Left Column - Input */}
                    <div className="space-y-4">
                      <div className="p-4 bg-gray-50 dark:bg-gray-700/50 rounded-lg border border-gray-200 dark:border-gray-600">
                        <label className="block text-sm font-semibold text-gray-900 dark:text-white mb-3">
                          Break Glass Account UPNs
                        </label>
                        <textarea
                          value={breakGlassInput}
                          onChange={(e) => setBreakGlassInput(e.target.value)}
                          placeholder="Enter user principal names (one per line)&#10;&#10;Example:&#10;breakglass1@contoso.com&#10;breakglass2@contoso.com"
                          className="w-full h-48 px-3 py-2 text-sm font-mono bg-white dark:bg-gray-800 border border-gray-300 dark:border-gray-600 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-blue-500 dark:text-white placeholder-gray-400 dark:placeholder-gray-500 resize-none"
                          disabled={isSavingBreakGlass}
                        />
                        <p className="mt-2 text-xs text-gray-500 dark:text-gray-400">
                          Enter one UPN per line, or separate with commas/semicolons.
                        </p>
                        
                        {/* Save Button */}
                        <div className="mt-4 pt-4 border-t border-gray-200 dark:border-gray-600">
                          <Button
                            appearance="primary"
                            onClick={saveBreakGlassSettings}
                            disabled={isSavingBreakGlass}
                            icon={isSavingBreakGlass ? <Spinner size="tiny" /> : undefined}
                            className="w-full sm:w-auto"
                          >
                            {isSavingBreakGlass ? 'Saving...' : 'Save Break Glass Accounts'}
                          </Button>
                        </div>
                      </div>
                    </div>

                    {/* Right Column - Configured Accounts */}
                    <div className="space-y-4">
                      <div className="p-4 bg-gray-50 dark:bg-gray-700/50 rounded-lg border border-gray-200 dark:border-gray-600 h-full">
                        <div className="flex items-center justify-between mb-3">
                          <h4 className="text-sm font-semibold text-gray-900 dark:text-white">
                            Configured Accounts
                          </h4>
                          {breakGlassSettings && breakGlassSettings.accounts.length > 0 && (
                            <Badge appearance="filled" color="brand">
                              {breakGlassSettings.accounts.length} account{breakGlassSettings.accounts.length !== 1 ? 's' : ''}
                            </Badge>
                          )}
                        </div>
                        
                        {breakGlassSettings && breakGlassSettings.accounts.length > 0 ? (
                          <>
                            <div className="space-y-2 max-h-64 overflow-y-auto">
                              {breakGlassSettings.accounts.map((account) => (
                                <div
                                  key={account.userPrincipalName}
                                  className={`flex items-center justify-between p-3 rounded-lg transition-colors ${
                                    account.isResolved
                                      ? 'bg-green-50 dark:bg-green-900/30 border border-green-200 dark:border-green-700'
                                      : 'bg-red-50 dark:bg-red-900/30 border border-red-200 dark:border-red-700'
                                  }`}
                                >
                                  <div className="flex items-center gap-3 min-w-0">
                                    <div
                                      className={`flex-shrink-0 w-8 h-8 rounded-full flex items-center justify-center ${
                                        account.isResolved
                                          ? 'bg-green-500 text-white'
                                          : 'bg-red-500 text-white'
                                      }`}
                                    >
                                      {account.isResolved ? (
                                        <Checkmark16Regular className="w-4 h-4" />
                                      ) : (
                                        <Dismiss16Regular className="w-4 h-4" />
                                      )}
                                    </div>
                                    <div className="min-w-0">
                                      <p
                                        className={`text-sm font-medium truncate ${
                                          account.isResolved
                                            ? 'text-green-800 dark:text-green-200'
                                            : 'text-red-800 dark:text-red-200'
                                        }`}
                                      >
                                        {account.userPrincipalName}
                                      </p>
                                      {account.isResolved && account.displayName ? (
                                        <p className="text-xs text-green-600 dark:text-green-400 truncate">
                                          {account.displayName}
                                        </p>
                                      ) : !account.isResolved ? (
                                        <p className="text-xs text-red-600 dark:text-red-400">
                                          User not found in directory
                                        </p>
                                      ) : null}
                                    </div>
                                  </div>
                                  <div className="flex items-center gap-2 flex-shrink-0 ml-2">
                                    <Badge
                                      appearance="tint"
                                      color={account.isResolved ? 'success' : 'danger'}
                                      size="small"
                                    >
                                      {account.isResolved ? 'Verified' : 'Not Found'}
                                    </Badge>
                                    <button
                                      onClick={() => removeBreakGlassAccount(account.userPrincipalName)}
                                      className="p-1 rounded hover:bg-red-100 dark:hover:bg-red-900/50 text-gray-400 hover:text-red-600 dark:hover:text-red-400 transition-colors"
                                      title="Remove account"
                                    >
                                      <Delete16Regular className="w-4 h-4" />
                                    </button>
                                  </div>
                                </div>
                              ))}
                            </div>
                            {breakGlassSettings.lastUpdated && (
                              <p className="mt-3 pt-3 border-t border-gray-200 dark:border-gray-600 text-xs text-gray-500 dark:text-gray-400">
                                Last updated: {new Date(breakGlassSettings.lastUpdated).toLocaleString()}
                                {breakGlassSettings.lastModifiedBy && 
                                 breakGlassSettings.lastModifiedBy !== 'unknown' && 
                                 breakGlassSettings.lastModifiedBy !== 'saving...' && (
                                  <span className="block sm:inline sm:ml-1">by {breakGlassSettings.lastModifiedBy}</span>
                                )}
                              </p>
                            )}
                          </>
                        ) : (
                          <div className="flex flex-col items-center justify-center py-8 text-center">
                            <div className="w-12 h-12 rounded-full bg-gray-200 dark:bg-gray-600 flex items-center justify-center mb-3">
                              <PersonKey24Regular className="w-6 h-6 text-gray-400 dark:text-gray-500" />
                            </div>
                            <p className="text-sm font-medium text-gray-600 dark:text-gray-400">
                              No accounts configured
                            </p>
                            <p className="text-xs text-gray-500 dark:text-gray-500 mt-1">
                              Add UPNs in the input field and save
                            </p>
                          </div>
                        )}
                      </div>
                    </div>
                  </div>

                  {/* Footer Note */}
                  <div className="p-3 bg-amber-50 dark:bg-amber-900/20 border border-amber-200 dark:border-amber-700 rounded-lg">
                    <p className="text-xs text-amber-800 dark:text-amber-200">
                      <strong>Next step:</strong> After saving, generate the "Conditional Access Break Glass Report" from the Reports page to verify these accounts are excluded from all CA policies.
                    </p>
                  </div>
                </div>
              )}
            </SettingsSection>

            {/* Permissions Section */}
            <SettingsSection
              icon={ShieldCheckmark24Regular}
              title="API Permissions"
              description="Microsoft Graph API permissions required by this application"
              headerAction={
                <div className="flex items-center gap-3">
                  {permissionsStatus && (
                    <Badge 
                      appearance="filled" 
                      color={permissionsStatus.allPermissionsGranted ? 'success' : 'warning'}
                    >
                      {permissionsStatus.grantedPermissions}/{permissionsStatus.totalPermissions} granted
                    </Badge>
                  )}
                  <Button
                    appearance="subtle"
                    icon={<ArrowSync24Regular />}
                    onClick={loadPermissionsStatus}
                    disabled={isLoadingPermissions}
                    size="small"
                  >
                    Refresh
                  </Button>
                </div>
              }
            >
              {isLoadingPermissions ? (
                <div className="flex items-center justify-center py-8">
                  <Spinner size="medium" label="Checking permissions..." />
                </div>
              ) : permissionsStatus ? (
                <div className="space-y-6">
                  {/* Summary */}
                  {!permissionsStatus.allPermissionsGranted && (
                    <div className="p-4 bg-amber-50 dark:bg-amber-900/20 border border-amber-200 dark:border-amber-800 rounded-lg">
                      <p className="text-sm text-amber-800 dark:text-amber-200">
                        <strong>{permissionsStatus.missingPermissions} permission(s)</strong> are not granted. 
                        Some features may be limited. Grant missing permissions in the{' '}
                        <a 
                          href="https://entra.microsoft.com/#view/Microsoft_AAD_IAM/StartboardApplicationsMenuBlade/~/AppAppsPreview"
                          target="_blank"
                          rel="noopener noreferrer"
                          className="underline hover:no-underline inline-flex items-center gap-1"
                        >
                          Entra Admin Center <Open16Regular className="w-3 h-3" />
                        </a>
                      </p>
                    </div>
                  )}

                  {/* Permissions by category */}
                  {categoryOrder.map(category => {
                    const perms = groupedPermissions[category];
                    if (!perms || perms.length === 0) return null;

                    const grantedInCategory = perms.filter(p => p.isGranted).length;
                    const allGranted = grantedInCategory === perms.length;

                    return (
                      <div key={category} className="space-y-2">
                        <div className="flex items-center justify-between">
                          <h4 className="text-sm font-semibold text-gray-700 dark:text-gray-300">
                            {category}
                          </h4>
                          <span className={`text-xs ${allGranted ? 'text-green-600 dark:text-green-400' : 'text-amber-600 dark:text-amber-400'}`}>
                            {grantedInCategory}/{perms.length}
                          </span>
                        </div>
                        <div className="space-y-1">
                          {perms.map(perm => (
                            <PermissionRow key={perm.permissionName} permission={perm} />
                          ))}
                        </div>
                      </div>
                    );
                  })}

                  {/* Last checked */}
                  <p className="text-xs text-gray-400 dark:text-gray-500 text-right">
                    Last checked: {new Date(permissionsStatus.lastChecked).toLocaleString()}
                  </p>
                </div>
              ) : (
                <p className="text-gray-500 dark:text-gray-400 text-center py-4">
                  Unable to load permissions status
                </p>
              )}
            </SettingsSection>
          </>
        )}

        {/* Reports Tab */}
        {selectedTab === 'reports' && (
          <>
            {reportLoading ? (
              <div className="flex items-center justify-center py-16">
                <Spinner size="medium" label="Loading report settings..." />
              </div>
            ) : (
              <>
                {reportMessage && (
                  <div className={`p-4 rounded-lg flex items-center gap-3 ${
                    reportMessage.type === 'success'
                      ? 'bg-green-50 dark:bg-green-900/20 border border-green-200 dark:border-green-700 text-green-700 dark:text-green-400'
                      : 'bg-red-50 dark:bg-red-900/20 border border-red-200 dark:border-red-700 text-red-700 dark:text-red-400'
                  }`}>
                    {reportMessage.type === 'success'
                      ? <CheckmarkCircleFilled className="w-5 h-5 flex-shrink-0" />
                      : <DismissCircleFilled className="w-5 h-5 flex-shrink-0" />}
                    {reportMessage.text}
                  </div>
                )}

                {/* Logo */}
                <SettingsSection icon={ImageRegular} title="Company Logo" description="Appears on the cover page and footer of generated reports">
                  <div className="flex items-start gap-6">
                    <div className="w-48 h-32 border-2 border-dashed border-gray-300 dark:border-gray-600 rounded-lg flex items-center justify-center bg-gray-50 dark:bg-gray-700/50 overflow-hidden flex-shrink-0">
                      {reportSettings.logoBase64 ? (
                        <img src={`data:${reportSettings.logoContentType};base64,${reportSettings.logoBase64}`} alt="Logo" className="max-w-full max-h-full object-contain" />
                      ) : (
                        <div className="text-center text-gray-400">
                          <ImageRegular className="w-8 h-8 mx-auto mb-2" />
                          <span className="text-sm">No logo uploaded</span>
                        </div>
                      )}
                    </div>
                    <div className="flex-1">
                      <p className="text-sm text-gray-600 dark:text-gray-400 mb-4">
                        Recommended size: 300×100 pixels. Supported formats: PNG, JPEG, GIF, SVG.
                      </p>
                      <input type="file" ref={fileInputRef} onChange={(e) => { const f = e.target.files?.[0]; if (f) uploadReportLogo(f); }} accept="image/png,image/jpeg,image/gif,image/svg+xml" className="hidden" />
                      <div className="flex flex-wrap gap-3">
                        <Button appearance="primary" icon={<ImageRegular />} onClick={() => fileInputRef.current?.click()} disabled={reportUploading}>
                          {reportUploading ? 'Uploading...' : 'Upload Logo'}
                        </Button>
                        <Button appearance="secondary" icon={<BuildingRegular />} onClick={importEntraLogo} disabled={reportUploading} title="Import the banner logo from your Entra organisational branding">
                          Use Entra Branding
                        </Button>
                        {reportSettings.logoBase64 && (
                          <Button appearance="secondary" icon={<DeleteRegular />} onClick={removeReportLogo} disabled={reportUploading}>
                            Remove
                          </Button>
                        )}
                      </div>
                      <p className="mt-2 text-xs text-gray-500 dark:text-gray-400">
                        "Use Entra Branding" imports your tenant's banner logo from Entra company branding. Requires Entra ID P1/P2 with branding configured.
                      </p>
                    </div>
                  </div>
                </SettingsSection>

                {/* Text */}
                <SettingsSection icon={DocumentRegular} title="Report Text" description="Labels and text shown on the cover page and footer">
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                    <div>
                      <label className="block text-sm font-medium text-gray-700 dark:text-gray-300 mb-2">Company Name</label>
                      <input type="text" value={reportSettings.companyName} onChange={(e) => setReportSettings(p => ({ ...p, companyName: e.target.value }))} placeholder="Your Company Name" className="w-full px-4 py-2 border border-gray-300 dark:border-gray-600 rounded-lg bg-white dark:bg-gray-700 text-gray-900 dark:text-white focus:ring-2 focus:ring-blue-500" />
                      <p className="mt-1 text-xs text-gray-500">Displayed on the cover page and footer</p>
                    </div>
                    <div>
                      <label className="block text-sm font-medium text-gray-700 dark:text-gray-300 mb-2">Report Title</label>
                      <input type="text" value={reportSettings.reportTitle} onChange={(e) => setReportSettings(p => ({ ...p, reportTitle: e.target.value }))} placeholder="Security Assessment Report" className="w-full px-4 py-2 border border-gray-300 dark:border-gray-600 rounded-lg bg-white dark:bg-gray-700 text-gray-900 dark:text-white focus:ring-2 focus:ring-blue-500" />
                      <p className="mt-1 text-xs text-gray-500">Main title on the cover page</p>
                    </div>
                    <div className="md:col-span-2">
                      <label className="block text-sm font-medium text-gray-700 dark:text-gray-300 mb-2">Footer Text (Optional)</label>
                      <input type="text" value={reportSettings.footerText || ''} onChange={(e) => setReportSettings(p => ({ ...p, footerText: e.target.value || null }))} placeholder="Confidential — For internal use only" className="w-full px-4 py-2 border border-gray-300 dark:border-gray-600 rounded-lg bg-white dark:bg-gray-700 text-gray-900 dark:text-white focus:ring-2 focus:ring-blue-500" />
                      <p className="mt-1 text-xs text-gray-500">Additional confidentiality notice</p>
                    </div>
                  </div>
                </SettingsSection>

                {/* Colours */}
                <SettingsSection icon={ColorRegular} title="Colours" description="Brand colours used in generated reports">
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                    <div>
                      <label className="block text-sm font-medium text-gray-700 dark:text-gray-300 mb-2">Primary Colour</label>
                      <div className="flex items-center gap-3">
                        <input type="color" value={reportSettings.primaryColor} onChange={(e) => setReportSettings(p => ({ ...p, primaryColor: e.target.value }))} className="w-12 h-10 rounded cursor-pointer border border-gray-300 dark:border-gray-600" />
                        <input type="text" value={reportSettings.primaryColor} onChange={(e) => setReportSettings(p => ({ ...p, primaryColor: e.target.value }))} className="flex-1 px-4 py-2 border border-gray-300 dark:border-gray-600 rounded-lg bg-white dark:bg-gray-700 text-gray-900 dark:text-white font-mono text-sm" />
                      </div>
                      <p className="mt-1 text-xs text-gray-500">Used for headers and section backgrounds</p>
                    </div>
                    <div>
                      <label className="block text-sm font-medium text-gray-700 dark:text-gray-300 mb-2">Accent Colour</label>
                      <div className="flex items-center gap-3">
                        <input type="color" value={reportSettings.accentColor} onChange={(e) => setReportSettings(p => ({ ...p, accentColor: e.target.value }))} className="w-12 h-10 rounded cursor-pointer border border-gray-300 dark:border-gray-600" />
                        <input type="text" value={reportSettings.accentColor} onChange={(e) => setReportSettings(p => ({ ...p, accentColor: e.target.value }))} className="flex-1 px-4 py-2 border border-gray-300 dark:border-gray-600 rounded-lg bg-white dark:bg-gray-700 text-gray-900 dark:text-white font-mono text-sm" />
                      </div>
                      <p className="mt-1 text-xs text-gray-500">Used for highlights and emphasis</p>
                    </div>
                  </div>
                </SettingsSection>

                {/* Options + Preview */}
                <SettingsSection icon={Options24Regular} title="Options &amp; Preview" description="Additional report options and a live cover page preview">
                  <label className="flex items-center gap-3 cursor-pointer mb-6">
                    <input type="checkbox" checked={reportSettings.showInfoGraphics} onChange={(e) => setReportSettings(p => ({ ...p, showInfoGraphics: e.target.checked }))} className="w-5 h-5 rounded border-gray-300 text-blue-600" />
                    <div>
                      <span className="text-gray-900 dark:text-white font-medium">Include Infographic Pages</span>
                      <p className="text-sm text-gray-500 dark:text-gray-400">Show statistics and quotes between report sections (recommended for client-facing reports)</p>
                    </div>
                  </label>
                  <div className="border border-gray-200 dark:border-gray-700 rounded-lg overflow-hidden">
                    <div className="h-48 p-6 flex flex-col justify-between text-white" style={{ background: `linear-gradient(135deg, ${reportSettings.primaryColor} 0%, #2d3748 100%)` }}>
                      <div>
                        <div className="text-lg font-light tracking-wider uppercase">MICROSOFT 365</div>
                        <div className="text-sm font-light uppercase tracking-widest" style={{ color: reportSettings.accentColor }}>
                          {reportSettings.reportTitle.toUpperCase().replace('MICROSOFT 365 ', '')}
                        </div>
                      </div>
                      <div className="flex items-end justify-between">
                        <div className="text-sm opacity-80">Sample Tenant Name</div>
                        {reportSettings.logoBase64 ? (
                          <img src={`data:${reportSettings.logoContentType};base64,${reportSettings.logoBase64}`} alt="Logo" className="h-8 object-contain" />
                        ) : (
                          <span className="font-semibold" style={{ color: reportSettings.accentColor }}>{reportSettings.companyName}</span>
                        )}
                      </div>
                    </div>
                  </div>
                </SettingsSection>

                {/* Save */}
                <div className="flex justify-end">
                  <Button appearance="primary" icon={<SaveRegular />} onClick={saveReportSettings} disabled={reportSaving} size="large">
                    {reportSaving ? 'Saving...' : 'Save Report Settings'}
                  </Button>
                </div>
              </>
            )}
          </>
        )}

        {/* System Tab */}
        {selectedTab === 'system' && (
          <>
            {/* External Services Section */}
            <SettingsSection
              icon={PlugConnected24Regular}
              title="External Services"
              description="Third-party services and integrations"
              headerAction={
                <div className="flex items-center gap-3">
                  {totalServices > 0 && (
                    <Badge 
                      appearance="filled" 
                      color={configuredServices === totalServices ? 'success' : 'warning'}
                    >
                      {configuredServices}/{totalServices} configured
                    </Badge>
                  )}
                  <Button
                    appearance="subtle"
                    icon={<ArrowSync24Regular />}
                    onClick={loadExternalServices}
                    disabled={isLoadingServices}
                    size="small"
                  >
                    Refresh
                  </Button>
                </div>
              }
            >
              {isLoadingServices ? (
                <div className="flex items-center justify-center py-8">
                  <Spinner size="medium" label="Checking services..." />
                </div>
              ) : externalServices.length > 0 ? (
                <div className="space-y-3">
                  {externalServices.map(service => (
                    <ExternalServiceRow key={service.name} service={service} />
                  ))}
                </div>
              ) : (
                <p className="text-gray-500 dark:text-gray-400 text-center py-4">
                  No external services configured
                </p>
              )}
            </SettingsSection>

            {/* System Info Section */}
            <SettingsSection
              icon={Info24Regular}
              title="System Information"
              description="Background services and data synchronization status"
              headerAction={
                <Button
                  appearance="subtle"
                  icon={<ArrowSync24Regular />}
                  onClick={loadSkuMappingStatus}
                  disabled={isLoadingSkuStatus}
                  size="small"
                >
                  Refresh
                </Button>
              }
            >
              {isLoadingSkuStatus ? (
                <div className="flex items-center justify-center py-8">
                  <Spinner size="medium" label="Loading status..." />
                </div>
              ) : skuMappingStatus ? (
                <div className="space-y-4">
                  {/* SKU Mapping Service */}
                  <div className="p-4 bg-gray-50 dark:bg-gray-700/50 rounded-lg">
                    <div className="flex items-center justify-between mb-3">
                      <div>
                        <h4 className="font-medium text-gray-900 dark:text-white">License SKU Mappings</h4>
                        <p className="text-sm text-gray-500 dark:text-gray-400">
                          Translates Microsoft SKU codes to friendly names
                        </p>
                      </div>
                      <Badge 
                        appearance="tint" 
                        color={skuMappingStatus.lastRefreshed ? 'success' : 'warning'}
                      >
                        {skuMappingStatus.isRefreshing ? 'Refreshing...' : skuMappingStatus.lastRefreshed ? 'Active' : 'Pending'}
                      </Badge>
                    </div>
                    
                    <div className="grid grid-cols-2 gap-4 mb-4">
                      <div className="p-3 bg-white dark:bg-gray-800 rounded border border-gray-200 dark:border-gray-600">
                        <p className="text-2xl font-bold text-gray-900 dark:text-white">
                          {skuMappingStatus.totalMappings.toLocaleString()}
                        </p>
                        <p className="text-xs text-gray-500 dark:text-gray-400">Total SKU Mappings</p>
                      </div>
                      <div className="p-3 bg-white dark:bg-gray-800 rounded border border-gray-200 dark:border-gray-600">
                        <p className="text-2xl font-bold text-gray-900 dark:text-white">
                          {skuMappingStatus.freeTrialSkusCount.toLocaleString()}
                        </p>
                        <p className="text-xs text-gray-500 dark:text-gray-400">Free/Trial SKUs Detected</p>
                      </div>
                    </div>

                    <div className="space-y-2 text-sm">
                      <div className="flex justify-between">
                        <span className="text-gray-600 dark:text-gray-400">Last Refreshed:</span>
                        <span className="text-gray-900 dark:text-white font-medium">
                          {skuMappingStatus.lastRefreshed 
                            ? new Date(skuMappingStatus.lastRefreshed).toLocaleString()
                            : 'Never'}
                        </span>
                      </div>
                      <div className="flex justify-between">
                        <span className="text-gray-600 dark:text-gray-400">Next Scheduled Refresh:</span>
                        <span className="text-gray-900 dark:text-white font-medium">
                          {skuMappingStatus.nextScheduledRefresh 
                            ? new Date(skuMappingStatus.nextScheduledRefresh).toLocaleString()
                            : 'Not scheduled'}
                        </span>
                      </div>
                      <div className="flex justify-between">
                        <span className="text-gray-600 dark:text-gray-400">Data Source:</span>
                        <a 
                          href="https://learn.microsoft.com/en-us/entra/identity/users/licensing-service-plan-reference"
                          target="_blank"
                          rel="noopener noreferrer"
                          className="text-blue-600 dark:text-blue-400 hover:underline flex items-center gap-1"
                        >
                          Microsoft Docs <Open16Regular className="w-3 h-3" />
                        </a>
                      </div>
                    </div>

                    {skuMappingStatus.lastError && (
                      <div className="mt-3 p-3 bg-red-50 dark:bg-red-900/20 border border-red-200 dark:border-red-800 rounded">
                        <p className="text-sm text-red-800 dark:text-red-200">
                          ⚠ Last error: {skuMappingStatus.lastError}
                        </p>
                      </div>
                    )}

                    <div className="mt-4 pt-4 border-t border-gray-200 dark:border-gray-600">
                      <Button
                        appearance="secondary"
                        icon={<ArrowSync24Regular />}
                        onClick={refreshSkuMappings}
                        disabled={isRefreshingSkuMappings || skuMappingStatus.isRefreshing}
                      >
                        {isRefreshingSkuMappings ? 'Refreshing...' : 'Refresh Now'}
                      </Button>
                      <p className="text-xs text-gray-500 dark:text-gray-400 mt-2">
                        Mappings are automatically refreshed daily at 3:00 AM from Microsoft's official CSV.
                      </p>
                    </div>
                  </div>
                </div>
              ) : (
                <p className="text-gray-500 dark:text-gray-400 text-center py-4">
                  Unable to load system status
                </p>
              )}
            </SettingsSection>

            {/* Application Updates Section */}
            <SettingsSection
              icon={ArrowDownload24Regular}
              title="Application Updates"
              description="Check for and apply new releases from GitHub"
              headerAction={
                <Button
                  appearance="subtle"
                  icon={<ArrowSync24Regular />}
                  onClick={checkForUpdates}
                  disabled={isCheckingUpdate}
                  size="small"
                >
                  {isCheckingUpdate ? 'Checking...' : 'Check Now'}
                </Button>
              }
            >
              {/* Result / progress message */}
              {updateMessage && (
                <div className={`p-3 rounded-lg flex items-center gap-3 text-sm ${
                  updateMessage.type === 'success'
                    ? 'bg-green-50 dark:bg-green-900/20 border border-green-200 dark:border-green-700 text-green-800 dark:text-green-200'
                    : 'bg-red-50 dark:bg-red-900/20 border border-red-200 dark:border-red-700 text-red-800 dark:text-red-200'
                }`}>
                  {updateMessage.type === 'success'
                    ? <CheckmarkCircleFilled className="w-5 h-5 flex-shrink-0" />
                    : <DismissCircleFilled className="w-5 h-5 flex-shrink-0" />}
                  <span>{updateMessage.text}</span>
                </div>
              )}

              {isCheckingUpdate && !updateStatus ? (
                <div className="flex items-center justify-center py-6">
                  <Spinner size="medium" label="Checking for updates..." />
                </div>
              ) : updateStatus ? (
                <div className="space-y-4">
                  {/* Version row */}
                  <div className="grid grid-cols-2 gap-4">
                    <div className="p-4 bg-gray-50 dark:bg-gray-700/50 rounded-lg">
                      <p className="text-xs text-gray-500 dark:text-gray-400 mb-1">Installed version</p>
                      <p className="text-lg font-bold text-gray-900 dark:text-white font-mono">
                        {updateStatus.currentVersion}
                      </p>
                    </div>
                    <div className={`p-4 rounded-lg ${
                      updateStatus.updateAvailable
                        ? 'bg-blue-50 dark:bg-blue-900/20 border border-blue-200 dark:border-blue-700'
                        : 'bg-green-50 dark:bg-green-900/20'
                    }`}>
                      <p className="text-xs text-gray-500 dark:text-gray-400 mb-1">Latest release</p>
                      <p className={`text-lg font-bold font-mono ${
                        updateStatus.updateAvailable
                          ? 'text-blue-700 dark:text-blue-300'
                          : 'text-green-700 dark:text-green-300'
                      }`}>
                        {updateStatus.latestVersion ?? '—'}
                      </p>
                      {updateStatus.publishedAt && (
                        <p className="text-xs text-gray-500 dark:text-gray-400 mt-1">
                          {new Date(updateStatus.publishedAt).toLocaleDateString()}
                        </p>
                      )}
                    </div>
                  </div>

                  {/* Status banner */}
                  {updateStatus.error ? (
                    <div className="p-3 bg-amber-50 dark:bg-amber-900/20 border border-amber-200 dark:border-amber-700 rounded-lg text-sm text-amber-800 dark:text-amber-200">
                      ⚠ {updateStatus.error}
                    </div>
                  ) : updateStatus.updateAvailable ? (
                    <div className="p-4 bg-blue-50 dark:bg-blue-900/20 border border-blue-200 dark:border-blue-700 rounded-lg">
                      <div className="flex items-start justify-between gap-4">
                        <div className="flex-1 min-w-0">
                          <p className="font-semibold text-blue-800 dark:text-blue-200 mb-1">
                            Update available — {updateStatus.latestVersion}
                          </p>
                          {updateStatus.releaseNotes && (
                            <p className="text-sm text-blue-700 dark:text-blue-300 line-clamp-3 whitespace-pre-line">
                              {updateStatus.releaseNotes.slice(0, 300)}{updateStatus.releaseNotes.length > 300 ? '…' : ''}
                            </p>
                          )}
                          {updateStatus.releaseUrl && (
                            <a
                              href={updateStatus.releaseUrl}
                              target="_blank"
                              rel="noopener noreferrer"
                              className="inline-flex items-center gap-1 text-xs text-blue-600 dark:text-blue-400 hover:underline mt-2"
                            >
                              Full release notes <Open16Regular className="w-3 h-3" />
                            </a>
                          )}
                        </div>
                        <div className="flex-shrink-0">
                          {updateStatus.updateConfigured ? (
                            <Button
                              appearance="primary"
                              icon={isApplyingUpdate ? <Spinner size="tiny" /> : <ArrowDownload24Regular />}
                              onClick={() => updateStatus.latestVersion && applyUpdate(updateStatus.latestVersion)}
                              disabled={isApplyingUpdate}
                            >
                              {isApplyingUpdate ? 'Updating…' : `Update to ${updateStatus.latestVersion}`}
                            </Button>
                          ) : (
                            <div className="text-right">
                              <a
                                href={updateStatus.releaseUrl ?? 'https://github.com/Alex-C1/m365-dashboard/releases'}
                                target="_blank"
                                rel="noopener noreferrer"
                              >
                                <Button appearance="primary" icon={<Open16Regular />}>
                                  View on GitHub
                                </Button>
                              </a>
                              <p className="text-xs text-gray-500 dark:text-gray-400 mt-1 max-w-[180px]">
                                One-click update requires Managed Identity configuration
                              </p>
                            </div>
                          )}
                        </div>
                      </div>
                    </div>
                  ) : (
                    <div className="p-3 bg-green-50 dark:bg-green-900/20 border border-green-200 dark:border-green-700 rounded-lg flex items-center gap-3">
                      <CheckmarkCircleFilled className="w-5 h-5 text-green-600 dark:text-green-400 flex-shrink-0" />
                      <p className="text-sm text-green-800 dark:text-green-200">
                        You are running the latest version.
                      </p>
                    </div>
                  )}

                  {/* One-click update not configured note */}
                  {!updateStatus.updateConfigured && (
                    <div className="p-3 bg-gray-50 dark:bg-gray-700/50 rounded-lg border border-gray-200 dark:border-gray-600">
                      <p className="text-xs text-gray-600 dark:text-gray-400">
                        <strong>One-click update not configured.</strong>{' '}
                        To enable in-app updates, set <code className="bg-gray-100 dark:bg-gray-800 px-1 rounded">ContainerApp:SubscriptionId</code>,{' '}
                        <code className="bg-gray-100 dark:bg-gray-800 px-1 rounded">ContainerApp:ResourceGroup</code>, and{' '}
                        <code className="bg-gray-100 dark:bg-gray-800 px-1 rounded">ContainerApp:Name</code> in your app configuration,
                        and assign Contributor role to the Container App's Managed Identity.
                        Alternatively, run <code className="bg-gray-100 dark:bg-gray-800 px-1 rounded">Update-M365Dashboard.ps1</code> from the scripts folder.
                      </p>
                    </div>
                  )}
                </div>
              ) : (
                <p className="text-sm text-gray-500 dark:text-gray-400 text-center py-4">
                  Click "Check Now" to check for available updates.
                </p>
              )}
            </SettingsSection>

            {/* Admin Section */}
            {isAdmin && (
              <SettingsSection
                icon={Settings24Regular}
                title="Admin Settings"
                description="Additional options for administrators"
              >
                <SettingsRow 
                  label="Clear Cache" 
                  description="Force refresh all cached data"
                >
                  <Button
                    appearance="secondary"
                    onClick={async () => {
                      setIsSaving(true);
                      try {
                        const { dashboardApi } = await import('../../services/api');
                        await dashboardApi.refreshCache('all');
                        showToast('Cache cleared successfully');
                      } catch {
                        showToast('Failed to clear cache', 'error');
                      } finally {
                        setIsSaving(false);
                      }
                    }}
                    disabled={isSaving}
                  >
                    Clear Cache
                  </Button>
                </SettingsRow>
              </SettingsSection>
            )}
          </>
        )}
      </div>
    </div>
  );
}

// Helper Components
interface SettingsSectionProps {
  icon: React.ComponentType<{ className?: string }>;
  title: string;
  description: string;
  children: React.ReactNode;
  headerAction?: React.ReactNode;
}

function SettingsSection({ icon: Icon, title, description, children, headerAction }: SettingsSectionProps) {
  return (
    <section className="bg-white dark:bg-gray-800 rounded-xl border border-gray-200 dark:border-gray-700 overflow-hidden">
      <div className="px-6 py-4 border-b border-gray-200 dark:border-gray-700">
        <div className="flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="p-2 bg-gray-100 dark:bg-gray-700 rounded-lg">
              <Icon className="w-5 h-5 text-gray-600 dark:text-gray-400" />
            </div>
            <div>
              <h2 className="font-semibold text-gray-900 dark:text-white">{title}</h2>
              <p className="text-sm text-gray-500 dark:text-gray-400">{description}</p>
            </div>
          </div>
          {headerAction}
        </div>
      </div>
      <div className="p-6 space-y-4">
        {children}
      </div>
    </section>
  );
}

interface SettingsRowProps {
  label: string;
  description?: string;
  children: React.ReactNode;
}

function SettingsRow({ label, description, children }: SettingsRowProps) {
  return (
    <div className="flex items-center justify-between gap-4">
      <div className="flex-1 min-w-0">
        <p className="font-medium text-gray-900 dark:text-white">{label}</p>
        {description && (
          <p className="text-sm text-gray-500 dark:text-gray-400">{description}</p>
        )}
      </div>
      <div className="flex-shrink-0">
        {children}
      </div>
    </div>
  );
}

// Fix instructions keyed by permissionName
const permissionFixInstructions: Record<string, { steps: string[]; links: { label: string; url: string }[] }> = {
  'Security Reader (Exchange)': {
    steps: [
      'Open the Exchange Admin Centre (admin.cloud.microsoft/exchange)',
      'Go to Roles → Admin roles',
      'Open the "View-Only Organization Management" role group',
      'Click "Edit" then go to the Members tab',
      'Click "Add" and search for your app registration by name and select it',
      'Save the role group',
      'Allow a few minutes for the role to propagate, then refresh the permissions page',
    ],
    links: [
      { label: 'Exchange Admin Centre', url: 'https://admin.cloud.microsoft/exchange#/adminRoles' },
      { label: 'Microsoft docs', url: 'https://learn.microsoft.com/en-us/microsoft-365/security/office-365-security/mdo-about' },
    ],
  },
  'Exchange Recipient Administrator': {
    steps: [
      'Open the Microsoft Entra admin centre (entra.microsoft.com)',
      'Go to Roles & admins → Roles & admins',
      'Search for and open the "Exchange Recipient Administrator" role',
      'Click "Add assignments" and select your app registration',
      'Alternatively, assign "Exchange Administrator" for full access',
      'Restart the app to pick up the new role — in Azure App Service: go to your App Service in the Azure portal and click Restart. Running locally: stop and re-run dotnet run.',
    ],
    links: [
      { label: 'Entra Roles & admins', url: 'https://entra.microsoft.com/#view/Microsoft_AAD_IAM/RolesManagementMenuBlade/~/AllRoles' },
    ],
  },
  'Exchange.ManageAsApp': {
    steps: [
      'Open the Azure portal (portal.azure.com)',
      'Go to Azure Active Directory → App registrations → your app',
      'Click "API permissions" → "Add a permission"',
      'Choose "APIs my organization uses" and search for "Office 365 Exchange Online"',
      'Select "Application permissions" and add "Exchange.ManageAsApp"',
      'Click "Grant admin consent" at the top of the permissions list',
    ],
    links: [
      { label: 'App registrations', url: 'https://portal.azure.com/#view/Microsoft_AAD_IAM/ActiveDirectoryMenuBlade/~/RegisteredApps' },
      { label: 'Microsoft docs', url: 'https://learn.microsoft.com/en-us/powershell/exchange/app-only-auth-powershell-v2' },
    ],
  },
  'Machine.Read.All': {
    steps: [
      'Open the Azure portal and navigate to your app registration',
      'Go to "API permissions" → "Add a permission"',
      'Choose "APIs my organization uses" and search for "WindowsDefenderATP"',
      'Select "Application permissions" and add "Machine.Read.All"',
      'Click "Grant admin consent"',
    ],
    links: [
      { label: 'App registrations', url: 'https://portal.azure.com/#view/Microsoft_AAD_IAM/ActiveDirectoryMenuBlade/~/RegisteredApps' },
    ],
  },
};

const defaultFixInstructions = {
  steps: [
    'Open the Azure portal and navigate to your app registration',
    'Go to "API permissions" → "Add a permission" → "Microsoft Graph"',
    'Select "Application permissions" and search for the permission name',
    'Add the permission and click "Grant admin consent"',
  ],
  links: [
    { label: 'App registrations', url: 'https://portal.azure.com/#view/Microsoft_AAD_IAM/ActiveDirectoryMenuBlade/~/RegisteredApps' },
  ],
};

interface PermissionRowProps {
  permission: PermissionStatus;
}

function PermissionRow({ permission }: PermissionRowProps) {
  const [expanded, setExpanded] = useState(false);
  const canExpand = !permission.isGranted;

  // Extract optional app name embedded in description as "| appName:XXX"
  const appNameMatch = permission.description.match(/\| appName:(.+)$/);
  const appName = appNameMatch ? appNameMatch[1].trim() : null;
  const cleanDescription = permission.description.replace(/\s*\|\s*appName:.+$/, '');

  const rawFix = permissionFixInstructions[permission.permissionName] ?? defaultFixInstructions;
  const fix = appName
    ? { ...rawFix, steps: rawFix.steps.map(s => s.replace('your app registration', `"${appName}"`))
    }
    : rawFix;

  return (
    <div className={`rounded-lg border transition-colors ${
      permission.isGranted
        ? 'bg-green-50 dark:bg-green-900/20 border-green-200 dark:border-green-800'
        : expanded
          ? 'bg-red-50 dark:bg-red-900/20 border-red-300 dark:border-red-700'
          : 'bg-red-50 dark:bg-red-900/20 border-red-200 dark:border-red-800'
    }`}>
      {/* Main row */}
      <button
        className={`w-full flex items-center justify-between p-3 text-left ${
          canExpand ? 'cursor-pointer hover:bg-red-100/50 dark:hover:bg-red-900/30' : 'cursor-default'
        } rounded-lg transition-colors`}
        onClick={() => canExpand && setExpanded(e => !e)}
        disabled={!canExpand}
      >
        <div className="flex items-center gap-3 flex-1 min-w-0">
          <div className={`flex-shrink-0 w-5 h-5 rounded-full flex items-center justify-center ${
            permission.isGranted ? 'bg-green-500 text-white' : 'bg-red-500 text-white'
          }`}>
            {permission.isGranted
              ? <Checkmark16Regular className="w-3 h-3" />
              : <Dismiss16Regular className="w-3 h-3" />}
          </div>
          <div className="min-w-0 flex-1">
            <p className={`text-sm font-medium ${
              permission.isGranted ? 'text-green-800 dark:text-green-200' : 'text-red-800 dark:text-red-200'
            }`}>
              {permission.permissionName}
            </p>
            <p className={`text-xs ${
              permission.isGranted ? 'text-green-600 dark:text-green-400' : 'text-red-600 dark:text-red-400'
            }`}>
              {cleanDescription}
            </p>
          </div>
        </div>
        <div className="flex items-center gap-2 flex-shrink-0 ml-2">
          <Badge appearance="tint" color={permission.isGranted ? 'success' : 'danger'} size="small">
            {permission.isGranted ? 'Granted' : 'Missing'}
          </Badge>
          {canExpand && (
            <span className="text-xs text-red-500 dark:text-red-400 flex items-center gap-1">
              {expanded ? <ChevronUp20Regular className="w-3.5 h-3.5" /> : <ChevronDown20Regular className="w-3.5 h-3.5" />}
            </span>
          )}
        </div>
      </button>

      {/* Expanded fix instructions */}
      {expanded && (
        <div className="px-3 pb-3 border-t border-red-200 dark:border-red-700 mt-0 pt-3">
          <p className="text-xs font-semibold text-red-800 dark:text-red-200 mb-2">How to fix:</p>
          <ul className="space-y-1.5 mb-3">
            {fix.steps.map((step, i) => (
              <li key={i} className="flex items-start gap-2 text-xs text-red-700 dark:text-red-300">
                <span className="flex-shrink-0 w-4 h-4 rounded-full bg-red-200 dark:bg-red-800 text-red-800 dark:text-red-200 flex items-center justify-center font-medium text-[10px] mt-0.5">
                  {i + 1}
                </span>
                {step}
              </li>
            ))}
          </ul>
          <div className="flex flex-wrap gap-2">
            {fix.links.map(link => (
              <a
                key={link.url}
                href={link.url}
                target="_blank"
                rel="noopener noreferrer"
                className="inline-flex items-center gap-1 px-2 py-1 text-xs bg-white dark:bg-gray-800 border border-red-300 dark:border-red-600 text-red-700 dark:text-red-300 rounded hover:bg-red-50 dark:hover:bg-red-900/40 transition-colors"
              >
                <Open16Regular className="w-3 h-3" />
                {link.label}
              </a>
            ))}
          </div>
          {permission.errorMessage && (
            <p className="mt-2 text-xs text-red-500 dark:text-red-400 font-mono bg-red-100 dark:bg-red-900/30 rounded p-2 break-words">
              {permission.errorMessage}
            </p>
          )}
        </div>
      )}
    </div>
  );
}

interface ExternalServiceRowProps {
  service: ExternalServiceStatus;
}

function ExternalServiceRow({ service }: ExternalServiceRowProps) {
  const [showScript, setShowScript] = useState(false);
  const [copied, setCopied] = useState(false);

  const getStatus = () => {
    if (service.isConfigured && service.isWorking) {
      return { color: 'success' as const, label: 'Working', bgClass: 'bg-green-50 dark:bg-green-900/20' };
    }
    if (service.isConfigured && !service.isWorking) {
      return { color: 'warning' as const, label: 'Error', bgClass: 'bg-amber-50 dark:bg-amber-900/20' };
    }
    return { color: 'danger' as const, label: 'Not Configured', bgClass: 'bg-red-50 dark:bg-red-900/20' };
  };

  const status = getStatus();

  const handleDownloadScript = () => {
    const link = document.createElement('a');
    link.href = '/api/config/setup-script/azure-maps';
    link.download = 'Setup-AzureMaps.ps1';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    setShowScript(true);
  };

  const handleCopyCommand = () => {
    const command = '.\\Setup-AzureMaps.ps1 -ResourceGroupName "rg-m365dashboard"';
    navigator.clipboard.writeText(command);
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  };

  return (
    <div className={`p-4 rounded-lg ${status.bgClass}`}>
      <div className="flex items-start justify-between gap-4">
        <div className="flex items-start gap-3 flex-1 min-w-0">
          <div className={`flex-shrink-0 w-10 h-10 rounded-lg flex items-center justify-center ${
            service.isConfigured && service.isWorking
              ? 'bg-green-500 text-white'
              : service.isConfigured
              ? 'bg-amber-500 text-white'
              : 'bg-gray-300 dark:bg-gray-600 text-gray-600 dark:text-gray-300'
          }`}>
            <Globe24Regular className="w-5 h-5" />
          </div>
          <div className="min-w-0 flex-1">
            <div className="flex items-center gap-2">
              <p className="font-medium text-gray-900 dark:text-white">
                {service.displayName}
              </p>
              <Badge appearance="tint" color={status.color} size="small">
                {status.label}
              </Badge>
            </div>
            <p className="text-sm text-gray-600 dark:text-gray-400 mt-0.5">
              {service.description}
            </p>
            {service.errorMessage && (
              <p className="text-sm text-red-600 dark:text-red-400 mt-1">
                ⚠ {service.errorMessage}
              </p>
            )}
            {!service.isConfigured && service.name === 'azure-maps' && (
              <div className="mt-3 space-y-3">
                {/* Automated Setup */}
                <div className="p-3 bg-blue-50 dark:bg-blue-900/20 rounded border border-blue-200 dark:border-blue-700">
                  <p className="text-xs font-medium text-blue-800 dark:text-blue-200 mb-2 flex items-center gap-1">
                    <ArrowSync24Regular className="w-4 h-4" />
                    Automated Setup (Recommended)
                  </p>
                  <p className="text-xs text-blue-700 dark:text-blue-300 mb-3">
                    Download and run the PowerShell script to automatically create an Azure Maps account and configure this application.
                  </p>
                  <div className="flex flex-wrap gap-2">
                    <Button
                      appearance="primary"
                      size="small"
                      onClick={handleDownloadScript}
                    >
                      Download Setup Script
                    </Button>
                  </div>
                  {showScript && (
                    <div className="mt-3 p-2 bg-gray-900 rounded">
                      <p className="text-xs text-gray-400 mb-1">Run in PowerShell:</p>
                      <div className="flex items-center gap-2">
                        <code className="text-xs text-green-400 flex-1 font-mono">
                          .\Setup-AzureMaps.ps1 -ResourceGroupName "rg-m365dashboard"
                        </code>
                        <button
                          onClick={handleCopyCommand}
                          className="text-xs text-blue-400 hover:text-blue-300 px-2 py-1 rounded hover:bg-gray-800"
                        >
                          {copied ? '✓ Copied' : 'Copy'}
                        </button>
                      </div>
                    </div>
                  )}
                </div>

                {/* Manual Setup */}
                <div className="p-3 bg-white dark:bg-gray-800 rounded border border-gray-200 dark:border-gray-600">
                  <p className="text-xs font-medium text-gray-700 dark:text-gray-300 mb-2">Manual Setup:</p>
                  <ol className="text-xs text-gray-600 dark:text-gray-400 space-y-1 list-decimal list-inside">
                    <li>Create an Azure Maps account in the Azure Portal</li>
                    <li>Select Gen2 pricing tier (includes free tier)</li>
                    <li>Copy the Primary Key from Authentication</li>
                    <li>Add to appsettings.json under "AzureMaps"."SubscriptionKey"</li>
                    <li>Restart the application</li>
                  </ol>
                </div>
              </div>
            )}
          </div>
        </div>
        <div className="flex flex-col gap-1">
          {service.setupUrl && (
            <a
              href={service.setupUrl}
              target="_blank"
              rel="noopener noreferrer"
              className="text-xs text-blue-600 hover:text-blue-700 dark:text-blue-400 dark:hover:text-blue-300 flex items-center gap-1"
            >
              Azure Portal <Open16Regular className="w-3 h-3" />
            </a>
          )}
          {service.docsUrl && (
            <a
              href={service.docsUrl}
              target="_blank"
              rel="noopener noreferrer"
              className="text-xs text-gray-500 hover:text-gray-700 dark:text-gray-400 dark:hover:text-gray-300 flex items-center gap-1"
            >
              Docs <Open16Regular className="w-3 h-3" />
            </a>
          )}
        </div>
      </div>
    </div>
  );
}

interface WidgetToggleProps {
  widget: WidgetConfiguration;
  onToggle: (widgetType: string, enabled: boolean) => void;
  disabled: boolean;
}

const widgetLabels: Record<string, { name: string; description: string }> = {
  'active-users': { name: 'Active Users', description: 'Daily, weekly, and monthly user activity' },
  'sign-in-analytics': { name: 'Sign-in Analytics', description: 'Authentication success and failure rates' },
  'license-usage': { name: 'License Usage', description: 'Subscription utilization by SKU' },
  'device-compliance': { name: 'Device Compliance', description: 'Intune device compliance status' },
  'mail-activity': { name: 'Mail Activity', description: 'Email sent and received trends' },
  'teams-activity': { name: 'Teams Activity', description: 'Messages, calls, and meetings' },
};

function WidgetToggle({ widget, onToggle, disabled }: WidgetToggleProps) {
  const info = widgetLabels[widget.widgetType] ?? { name: widget.widgetType, description: '' };

  return (
    <div className="flex items-center justify-between p-3 bg-gray-50 dark:bg-gray-700/50 rounded-lg">
      <div className="flex-1 min-w-0">
        <p className="font-medium text-gray-900 dark:text-white">{info.name}</p>
        <p className="text-xs text-gray-500 dark:text-gray-400">{info.description}</p>
      </div>
      <Switch
        checked={widget.isEnabled}
        onChange={(_, data) => onToggle(widget.widgetType, data.checked)}
        disabled={disabled}
      />
    </div>
  );
}
