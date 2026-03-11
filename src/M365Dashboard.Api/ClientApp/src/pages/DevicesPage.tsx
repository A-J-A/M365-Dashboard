import React, { useState, useEffect, useMemo } from 'react';
import {
  LaptopRegular,
  PhoneRegular,
  TabletRegular,
  DesktopRegular,
  CheckmarkCircleFilled,
  DismissCircleFilled,
  WarningFilled,
  SearchRegular,
  OpenRegular,
  ChevronUpRegular,
  ChevronDownRegular,
  ShieldCheckmarkRegular,
  ShieldErrorRegular,
  BuildingRegular,
  PersonRegular,
  LockClosedRegular,
  CloudRegular,
  CertificateRegular,
  CalendarRegular,
  ClockRegular,
  ErrorCircleRegular,
} from '@fluentui/react-icons';
import { useAppContext } from '../contexts/AppContext';

interface IntuneDevice {
  id: string;
  deviceName: string;
  userDisplayName: string | null;
  userPrincipalName: string | null;
  managedDeviceOwnerType: string | null;
  operatingSystem: string | null;
  osVersion: string | null;
  complianceState: string | null;
  managementState: string | null;
  deviceEnrollmentType: string | null;
  lastSyncDateTime: string | null;
  enrolledDateTime: string | null;
  model: string | null;
  manufacturer: string | null;
  serialNumber: string | null;
  jailBroken: string | null;
  isEncrypted: boolean | null;
  isSupervised: boolean | null;
  deviceRegistrationState: string | null;
  managementAgent: string | null;
  totalStorageSpaceInBytes: number | null;
  freeStorageSpaceInBytes: number | null;
  wiFiMacAddress: string | null;
  ethernetMacAddress: string | null;
  imei: string | null;
  phoneNumber: string | null;
  azureAdDeviceId: string | null;
  azureAdRegistered: boolean | null;
  deviceCategoryDisplayName: string | null;
}

interface DeviceStats {
  totalDevices: number;
  compliantDevices: number;
  nonCompliantDevices: number;
  inGracePeriod: number;
  configurationManagerDevices: number;
  windowsDevices: number;
  macOsDevices: number;
  iosDevices: number;
  androidDevices: number;
  linuxDevices: number;
  corporateDevices: number;
  personalDevices: number;
  managedDevices: number;
  encryptedDevices: number;
  lastUpdated: string;
}

interface DeviceListResult {
  devices: IntuneDevice[];
  totalCount: number;
  filteredCount: number;
  nextLink: string | null;
}

interface ApplePushCertificate {
  isConfigured: boolean;
  appleIdentifier?: string;
  topicIdentifier?: string;
  expirationDateTime?: string;
  lastModifiedDateTime?: string;
  certificateSerialNumber?: string;
  certificateUploadStatus?: string;
  daysUntilExpiry?: number;
  status?: 'Healthy' | 'Warning' | 'Critical' | 'Expired' | 'Unknown';
  message?: string;
  error?: string;
  permissionRequired?: boolean;
  lastUpdated?: string;
}

interface DepToken {
  id: string;
  tokenName: string;
  appleIdentifier?: string;
  tokenExpirationDateTime?: string;
  lastModifiedDateTime?: string;
  lastSuccessfulSyncDateTime?: string;
  lastSyncTriggeredDateTime?: string;
  lastSyncErrorCode?: number;
  dataSharingConsentGranted?: boolean;
  tokenType?: string;
  daysUntilExpiry?: number;
  status?: 'Healthy' | 'Warning' | 'Critical' | 'Expired' | 'Unknown';
}

interface DepTokensResponse {
  tokens: DepToken[];
  totalCount: number;
  error?: string;
  permissionRequired?: boolean;
  lastUpdated?: string;
}

type FilterType = 'all' | 'compliant' | 'nonCompliant' | 'windows' | 'macos' | 'ios' | 'android' | 'corporate' | 'personal';
type SortField = 'deviceName' | 'userDisplayName' | 'operatingSystem' | 'complianceState' | 'lastSyncDateTime' | 'enrolledDateTime';
type SortDirection = 'asc' | 'desc';
type ViewMode = 'devices' | 'certificates';

const DevicesPage: React.FC = () => {
  const { getAccessToken } = useAppContext();
  const [devices, setDevices] = useState<IntuneDevice[]>([]);
  const [stats, setStats] = useState<DeviceStats | null>(null);
  const [applePushCert, setApplePushCert] = useState<ApplePushCertificate | null>(null);
  const [depTokens, setDepTokens] = useState<DepTokensResponse | null>(null);
  const [loading, setLoading] = useState(true);
  const [statsLoading, setStatsLoading] = useState(true);
  const [certLoading, setCertLoading] = useState(true);
  const [depTokensLoading, setDepTokensLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [searchTerm, setSearchTerm] = useState('');
  const [filterType, setFilterType] = useState<FilterType>('all');
  const [sortField, setSortField] = useState<SortField>('deviceName');
  const [sortDirection, setSortDirection] = useState<SortDirection>('asc');
  const [viewMode, setViewMode] = useState<ViewMode>('devices');

  useEffect(() => {
    fetchDevices();
    fetchStats();
    fetchApplePushCertificate();
    fetchDepTokens();
  }, []);

  const fetchDevices = async () => {
    try {
      setLoading(true);
      const token = await getAccessToken();
      const response = await fetch('/api/devices?take=500', {
        headers: {
          'Authorization': `Bearer ${token}`,
          'Content-Type': 'application/json',
        },
      });

      if (!response.ok) {
        throw new Error('Failed to fetch devices');
      }

      const data: DeviceListResult = await response.json();
      setDevices(data.devices);
      setError(null);
    } catch (err) {
      setError(err instanceof Error ? err.message : 'An error occurred');
    } finally {
      setLoading(false);
    }
  };

  const fetchStats = async () => {
    try {
      setStatsLoading(true);
      const token = await getAccessToken();
      const response = await fetch('/api/devices/stats', {
        headers: {
          'Authorization': `Bearer ${token}`,
          'Content-Type': 'application/json',
        },
      });

      if (!response.ok) {
        throw new Error('Failed to fetch device stats');
      }

      const data: DeviceStats = await response.json();
      setStats(data);
    } catch (err) {
      console.error('Failed to fetch stats:', err);
    } finally {
      setStatsLoading(false);
    }
  };

  const fetchApplePushCertificate = async () => {
    try {
      setCertLoading(true);
      const token = await getAccessToken();
      const response = await fetch('/api/devices/apple-push-certificate', {
        headers: {
          'Authorization': `Bearer ${token}`,
          'Content-Type': 'application/json',
        },
      });

      if (!response.ok) {
        throw new Error('Failed to fetch Apple Push certificate');
      }

      const data: ApplePushCertificate = await response.json();
      setApplePushCert(data);
    } catch (err) {
      console.error('Failed to fetch Apple Push certificate:', err);
      setApplePushCert({
        isConfigured: false,
        error: err instanceof Error ? err.message : 'Failed to fetch certificate',
      });
    } finally {
      setCertLoading(false);
    }
  };

  const fetchDepTokens = async () => {
    try {
      setDepTokensLoading(true);
      const token = await getAccessToken();
      const response = await fetch('/api/devices/dep-tokens', {
        headers: {
          'Authorization': `Bearer ${token}`,
          'Content-Type': 'application/json',
        },
      });

      if (!response.ok) {
        throw new Error('Failed to fetch DEP tokens');
      }

      const data: DepTokensResponse = await response.json();
      setDepTokens(data);
    } catch (err) {
      console.error('Failed to fetch DEP tokens:', err);
      setDepTokens({
        tokens: [],
        totalCount: 0,
        error: err instanceof Error ? err.message : 'Failed to fetch tokens',
      });
    } finally {
      setDepTokensLoading(false);
    }
  };

  const filteredAndSortedDevices = useMemo(() => {
    let result = [...devices];

    // Search filter
    if (searchTerm) {
      const term = searchTerm.toLowerCase();
      result = result.filter(
        (device) =>
          device.deviceName.toLowerCase().includes(term) ||
          (device.userDisplayName?.toLowerCase().includes(term) ?? false) ||
          (device.userPrincipalName?.toLowerCase().includes(term) ?? false) ||
          (device.serialNumber?.toLowerCase().includes(term) ?? false) ||
          (device.model?.toLowerCase().includes(term) ?? false)
      );
    }

    // Type filter
    switch (filterType) {
      case 'compliant':
        result = result.filter((d) => d.complianceState === 'Compliant');
        break;
      case 'nonCompliant':
        result = result.filter((d) => d.complianceState === 'Noncompliant');
        break;
      case 'windows':
        result = result.filter((d) => d.operatingSystem?.toLowerCase().includes('windows'));
        break;
      case 'macos':
        result = result.filter((d) => d.operatingSystem?.toLowerCase().includes('macos') || d.operatingSystem?.toLowerCase().includes('mac os'));
        break;
      case 'ios':
        result = result.filter((d) => d.operatingSystem?.toLowerCase().includes('ios') || d.operatingSystem?.toLowerCase().includes('ipados'));
        break;
      case 'android':
        result = result.filter((d) => d.operatingSystem?.toLowerCase().includes('android'));
        break;
      case 'corporate':
        result = result.filter((d) => d.managedDeviceOwnerType === 'Company');
        break;
      case 'personal':
        result = result.filter((d) => d.managedDeviceOwnerType === 'Personal');
        break;
    }

    // Sort
    result.sort((a, b) => {
      let aValue: string | null = null;
      let bValue: string | null = null;

      switch (sortField) {
        case 'deviceName':
          aValue = a.deviceName;
          bValue = b.deviceName;
          break;
        case 'userDisplayName':
          aValue = a.userDisplayName;
          bValue = b.userDisplayName;
          break;
        case 'operatingSystem':
          aValue = a.operatingSystem;
          bValue = b.operatingSystem;
          break;
        case 'complianceState':
          aValue = a.complianceState;
          bValue = b.complianceState;
          break;
        case 'lastSyncDateTime':
          aValue = a.lastSyncDateTime;
          bValue = b.lastSyncDateTime;
          break;
        case 'enrolledDateTime':
          aValue = a.enrolledDateTime;
          bValue = b.enrolledDateTime;
          break;
      }

      if (aValue === null && bValue === null) return 0;
      if (aValue === null) return 1;
      if (bValue === null) return -1;

      const comparison = String(aValue).localeCompare(String(bValue));
      return sortDirection === 'asc' ? comparison : -comparison;
    });

    return result;
  }, [devices, searchTerm, filterType, sortField, sortDirection]);

  const handleSort = (field: SortField) => {
    if (sortField === field) {
      setSortDirection(sortDirection === 'asc' ? 'desc' : 'asc');
    } else {
      setSortField(field);
      setSortDirection('asc');
    }
  };

  const formatDate = (dateString: string | null) => {
    if (!dateString) return 'Never';
    const date = new Date(dateString);
    const now = new Date();
    const diffMs = now.getTime() - date.getTime();
    const diffDays = Math.floor(diffMs / (1000 * 60 * 60 * 24));

    if (diffDays === 0) return 'Today';
    if (diffDays === 1) return 'Yesterday';
    if (diffDays < 7) return `${diffDays}d ago`;
    if (diffDays < 30) return `${Math.floor(diffDays / 7)}w ago`;
    return date.toLocaleDateString();
  };

  const formatFullDate = (dateString: string | null | undefined) => {
    if (!dateString) return 'N/A';
    return new Date(dateString).toLocaleDateString('en-GB', {
      day: 'numeric',
      month: 'short',
      year: 'numeric',
      hour: '2-digit',
      minute: '2-digit',
    });
  };

  const formatStorage = (bytes: number | null) => {
    if (bytes === null) return 'N/A';
    const gb = bytes / (1024 * 1024 * 1024);
    return `${gb.toFixed(1)} GB`;
  };

  const openInIntune = (deviceId: string) => {
    const url = `https://intune.microsoft.com/#view/Microsoft_Intune_Devices/DeviceSettingsMenuBlade/~/overview/mdmDeviceId/${deviceId}`;
    window.open(url, '_blank');
  };

  const getFilterLabel = (filter: FilterType): string => {
    switch (filter) {
      case 'all': return 'All Devices';
      case 'compliant': return 'Compliant';
      case 'nonCompliant': return 'Non-Compliant';
      case 'windows': return 'Windows';
      case 'macos': return 'macOS';
      case 'ios': return 'iOS/iPadOS';
      case 'android': return 'Android';
      case 'corporate': return 'Corporate';
      case 'personal': return 'Personal';
      default: return 'All Devices';
    }
  };

  const getOsIcon = (os: string | null) => {
    if (!os) return <LaptopRegular className="w-4 h-4 text-slate-500" />;
    const osLower = os.toLowerCase();
    if (osLower.includes('windows')) return <DesktopRegular className="w-4 h-4 text-blue-500" />;
    if (osLower.includes('macos') || osLower.includes('mac os')) return <LaptopRegular className="w-4 h-4 text-slate-700" />;
    if (osLower.includes('ios') || osLower.includes('ipados')) return <PhoneRegular className="w-4 h-4 text-slate-600" />;
    if (osLower.includes('android')) return <PhoneRegular className="w-4 h-4 text-green-600" />;
    return <LaptopRegular className="w-4 h-4 text-slate-500" />;
  };

  const getComplianceBadge = (state: string | null) => {
    switch (state) {
      case 'Compliant':
        return (
          <span className="inline-flex items-center gap-1 px-2 py-1 rounded-full text-xs font-medium bg-green-100 text-green-700 dark:bg-green-900/30 dark:text-green-400">
            <CheckmarkCircleFilled className="w-3 h-3" />
            Compliant
          </span>
        );
      case 'Noncompliant':
        return (
          <span className="inline-flex items-center gap-1 px-2 py-1 rounded-full text-xs font-medium bg-red-100 text-red-700 dark:bg-red-900/30 dark:text-red-400">
            <DismissCircleFilled className="w-3 h-3" />
            Non-Compliant
          </span>
        );
      case 'InGracePeriod':
        return (
          <span className="inline-flex items-center gap-1 px-2 py-1 rounded-full text-xs font-medium bg-amber-100 text-amber-700 dark:bg-amber-900/30 dark:text-amber-400">
            <WarningFilled className="w-3 h-3" />
            Grace Period
          </span>
        );
      default:
        return (
          <span className="inline-flex items-center gap-1 px-2 py-1 rounded-full text-xs font-medium bg-slate-100 text-slate-700 dark:bg-slate-900/30 dark:text-slate-400">
            Unknown
          </span>
        );
    }
  };

  const getCertStatusColor = (status: string | undefined) => {
    switch (status) {
      case 'Healthy':
        return 'text-green-600 bg-green-100 dark:bg-green-900/30 dark:text-green-400';
      case 'Warning':
        return 'text-amber-600 bg-amber-100 dark:bg-amber-900/30 dark:text-amber-400';
      case 'Critical':
        return 'text-red-600 bg-red-100 dark:bg-red-900/30 dark:text-red-400';
      case 'Expired':
        return 'text-red-700 bg-red-200 dark:bg-red-900/50 dark:text-red-300';
      default:
        return 'text-slate-600 bg-slate-100 dark:bg-slate-900/30 dark:text-slate-400';
    }
  };

  const getCertStatusIcon = (status: string | undefined) => {
    switch (status) {
      case 'Healthy':
        return <CheckmarkCircleFilled className="w-5 h-5 text-green-600" />;
      case 'Warning':
        return <WarningFilled className="w-5 h-5 text-amber-600" />;
      case 'Critical':
      case 'Expired':
        return <ErrorCircleRegular className="w-5 h-5 text-red-600" />;
      default:
        return <CertificateRegular className="w-5 h-5 text-slate-500" />;
    }
  };

  const SortIcon: React.FC<{ field: SortField }> = ({ field }) => {
    if (sortField !== field) return null;
    return sortDirection === 'asc' ? (
      <ChevronUpRegular className="w-4 h-4 ml-1" />
    ) : (
      <ChevronDownRegular className="w-4 h-4 ml-1" />
    );
  };

  // Stat card component
  const StatCard: React.FC<{
    title: string;
    value: number;
    icon: React.ReactNode;
    color: string;
    isActive: boolean;
    onClick: () => void;
  }> = ({ title, value, icon, color, isActive, onClick }) => (
    <button 
      type="button"
      className={`w-full bg-white dark:bg-slate-800 rounded-lg border-2 transition-all cursor-pointer hover:shadow-md p-2 text-left ${
        isActive 
          ? 'border-blue-500 ring-1 ring-blue-200 dark:ring-blue-800 shadow-md' 
          : 'border-slate-200 dark:border-slate-700 hover:border-slate-300 dark:hover:border-slate-600'
      }`}
      onClick={onClick}
    >
      <div className="flex items-center gap-1.5">
        <div className={`p-1 rounded ${color.replace('text-', 'bg-').replace('-600', '-100')} dark:bg-opacity-20 flex-shrink-0`}>
          {icon}
        </div>
        <div className="min-w-0 flex-1">
          <p className="text-[10px] text-slate-500 dark:text-slate-400 leading-tight truncate">{title}</p>
          <p className={`text-sm font-semibold leading-tight ${color}`}>{value}</p>
        </div>
        {isActive && (
          <CheckmarkCircleFilled className="w-3 h-3 text-blue-500 flex-shrink-0" />
        )}
      </div>
    </button>
  );

  return (
    <div className="p-4 space-y-4 w-full max-w-full overflow-hidden">
      {/* Header */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
        <div>
          <h1 className="text-xl font-semibold text-slate-900 dark:text-white">Devices</h1>
          <p className="text-sm text-slate-500 dark:text-slate-400 hidden sm:block">
            Intune managed devices across your organization
          </p>
        </div>
        <div className="flex items-center gap-2">
          <div className="flex rounded-lg border border-slate-200 dark:border-slate-700 overflow-hidden">
            <button
              onClick={() => setViewMode('devices')}
              className={`px-3 py-1.5 text-sm ${viewMode === 'devices' ? 'bg-blue-600 text-white' : 'bg-white dark:bg-slate-800 text-slate-600 dark:text-slate-300'}`}
            >
              Devices
            </button>
            <button
              onClick={() => setViewMode('certificates')}
              className={`px-3 py-1.5 text-sm ${viewMode === 'certificates' ? 'bg-blue-600 text-white' : 'bg-white dark:bg-slate-800 text-slate-600 dark:text-slate-300'}`}
            >
              Certificates
            </button>
          </div>
          <a
            href="https://intune.microsoft.com/#view/Microsoft_Intune_DeviceSettings/DevicesMenu/~/overview"
            target="_blank"
            rel="noopener noreferrer"
            className="inline-flex items-center justify-center gap-2 px-3 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 transition-colors text-sm whitespace-nowrap"
          >
            <OpenRegular className="w-4 h-4" />
            <span>Intune</span>
          </a>
        </div>
      </div>

      {viewMode === 'devices' ? (
        <>
          {/* Stats Cards - Row 1: Compliance */}
          <div className="grid grid-cols-2 sm:grid-cols-4 gap-2">
            <StatCard
              title="Total"
              value={stats?.totalDevices ?? 0}
              icon={<LaptopRegular className="w-4 h-4 text-blue-600" />}
              color="text-blue-600"
              isActive={filterType === 'all'}
              onClick={() => setFilterType('all')}
            />
            <StatCard
              title="Compliant"
              value={stats?.compliantDevices ?? 0}
              icon={<ShieldCheckmarkRegular className="w-4 h-4 text-green-600" />}
              color="text-green-600"
              isActive={filterType === 'compliant'}
              onClick={() => setFilterType('compliant')}
            />
            <StatCard
              title="Non-Compliant"
              value={stats?.nonCompliantDevices ?? 0}
              icon={<ShieldErrorRegular className="w-4 h-4 text-red-600" />}
              color="text-red-600"
              isActive={filterType === 'nonCompliant'}
              onClick={() => setFilterType('nonCompliant')}
            />
            <StatCard
              title="Encrypted"
              value={stats?.encryptedDevices ?? 0}
              icon={<LockClosedRegular className="w-4 h-4 text-purple-600" />}
              color="text-purple-600"
              isActive={false}
              onClick={() => {}}
            />
          </div>

          {/* Stats Cards - Row 2: Platform & Ownership */}
          <div className="grid grid-cols-4 sm:grid-cols-6 gap-2">
            <StatCard
              title="Windows"
              value={stats?.windowsDevices ?? 0}
              icon={<DesktopRegular className="w-4 h-4 text-blue-500" />}
              color="text-blue-500"
              isActive={filterType === 'windows'}
              onClick={() => setFilterType('windows')}
            />
            <StatCard
              title="macOS"
              value={stats?.macOsDevices ?? 0}
              icon={<LaptopRegular className="w-4 h-4 text-slate-600" />}
              color="text-slate-600"
              isActive={filterType === 'macos'}
              onClick={() => setFilterType('macos')}
            />
            <StatCard
              title="iOS"
              value={stats?.iosDevices ?? 0}
              icon={<PhoneRegular className="w-4 h-4 text-slate-500" />}
              color="text-slate-500"
              isActive={filterType === 'ios'}
              onClick={() => setFilterType('ios')}
            />
            <StatCard
              title="Android"
              value={stats?.androidDevices ?? 0}
              icon={<PhoneRegular className="w-4 h-4 text-green-600" />}
              color="text-green-600"
              isActive={filterType === 'android'}
              onClick={() => setFilterType('android')}
            />
            <StatCard
              title="Corporate"
              value={stats?.corporateDevices ?? 0}
              icon={<BuildingRegular className="w-4 h-4 text-indigo-600" />}
              color="text-indigo-600"
              isActive={filterType === 'corporate'}
              onClick={() => setFilterType('corporate')}
            />
            <StatCard
              title="Personal"
              value={stats?.personalDevices ?? 0}
              icon={<PersonRegular className="w-4 h-4 text-amber-600" />}
              color="text-amber-600"
              isActive={filterType === 'personal'}
              onClick={() => setFilterType('personal')}
            />
          </div>

          {/* Filters and Search */}
          <div className="bg-white dark:bg-slate-800 rounded-lg border border-slate-200 dark:border-slate-700 p-3">
            <div className="flex flex-col sm:flex-row gap-2">
              {/* Search */}
              <div className="relative flex-1">
                <SearchRegular className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-400" />
                <input
                  type="text"
                  placeholder="Search devices..."
                  value={searchTerm}
                  onChange={(e) => setSearchTerm(e.target.value)}
                  className="w-full pl-9 pr-3 py-2 border border-slate-200 dark:border-slate-600 rounded-lg bg-white dark:bg-slate-700 text-slate-900 dark:text-white focus:outline-none focus:ring-2 focus:ring-blue-500 text-sm"
                />
              </div>

              {/* Filter Dropdown */}
              <select
                value={filterType}
                onChange={(e) => setFilterType(e.target.value as FilterType)}
                className="px-3 py-2 border border-slate-200 dark:border-slate-600 rounded-lg bg-white dark:bg-slate-700 text-slate-900 dark:text-white focus:outline-none focus:ring-2 focus:ring-blue-500 text-sm"
              >
                <option value="all">All Devices</option>
                <option value="compliant">Compliant</option>
                <option value="nonCompliant">Non-Compliant</option>
                <option value="windows">Windows</option>
                <option value="macos">macOS</option>
                <option value="ios">iOS/iPadOS</option>
                <option value="android">Android</option>
                <option value="corporate">Corporate</option>
                <option value="personal">Personal</option>
              </select>
            </div>

            <div className="mt-2 text-xs text-slate-500 dark:text-slate-400">
              {filteredAndSortedDevices.length} of {devices.length} devices
              {filterType !== 'all' && (
                <span className="text-blue-600 dark:text-blue-400 ml-1">
                  • {getFilterLabel(filterType)}
                </span>
              )}
            </div>
          </div>

          {/* Devices Table */}
          <div className="bg-white dark:bg-slate-800 rounded-lg border border-slate-200 dark:border-slate-700 overflow-hidden">
            {loading ? (
              <div className="flex items-center justify-center h-64">
                <div className="animate-spin rounded-full h-8 w-8 border-b-2 border-blue-600"></div>
              </div>
            ) : error ? (
              <div className="flex items-center justify-center h-64 text-red-500">
                <p>{error}</p>
              </div>
            ) : (
              <>
                {/* Mobile Card View */}
                <div className="block lg:hidden divide-y divide-slate-200 dark:divide-slate-700">
                  {filteredAndSortedDevices.map((device) => (
                    <div key={device.id} className="p-3 hover:bg-slate-50 dark:hover:bg-slate-700/50">
                      <div className="flex items-start gap-3">
                        <div className="w-10 h-10 rounded-full bg-slate-100 dark:bg-slate-700 flex items-center justify-center flex-shrink-0">
                          {getOsIcon(device.operatingSystem)}
                        </div>
                        <div className="flex-1 min-w-0">
                          <div className="flex items-start justify-between gap-2">
                            <div className="min-w-0">
                              <p className="font-medium text-slate-900 dark:text-white truncate">
                                {device.deviceName}
                              </p>
                              <p className="text-xs text-slate-500 dark:text-slate-400 truncate">
                                {device.userDisplayName || 'No user'}
                              </p>
                            </div>
                            <button
                              onClick={() => openInIntune(device.id)}
                              className="p-1.5 text-slate-400 hover:text-blue-600 hover:bg-blue-50 dark:hover:bg-blue-900/30 rounded-lg transition-colors flex-shrink-0"
                              title="Open in Intune"
                            >
                              <OpenRegular className="w-4 h-4" />
                            </button>
                          </div>
                          
                          <div className="flex flex-wrap gap-1.5 mt-2">
                            {getComplianceBadge(device.complianceState)}
                            <span className="inline-flex items-center gap-1 px-1.5 py-0.5 rounded text-xs font-medium bg-slate-100 text-slate-700 dark:bg-slate-700 dark:text-slate-300">
                              {device.operatingSystem || 'Unknown OS'}
                            </span>
                            {device.managedDeviceOwnerType && (
                              <span className={`inline-flex items-center gap-1 px-1.5 py-0.5 rounded text-xs font-medium ${
                                device.managedDeviceOwnerType === 'Company'
                                  ? 'bg-indigo-100 text-indigo-700 dark:bg-indigo-900/30 dark:text-indigo-400'
                                  : 'bg-amber-100 text-amber-700 dark:bg-amber-900/30 dark:text-amber-400'
                              }`}>
                                {device.managedDeviceOwnerType}
                              </span>
                            )}
                          </div>

                          <div className="mt-2 text-xs text-slate-500 dark:text-slate-400">
                            Last sync: {formatDate(device.lastSyncDateTime)}
                          </div>
                        </div>
                      </div>
                    </div>
                  ))}
                </div>

                {/* Desktop Table View */}
                <div className="hidden lg:block overflow-x-auto">
                  <table className="w-full">
                    <thead className="bg-slate-50 dark:bg-slate-700">
                      <tr>
                        <th
                          className="px-4 py-3 text-left text-sm font-medium text-slate-600 dark:text-slate-300 cursor-pointer hover:bg-slate-100 dark:hover:bg-slate-600"
                          onClick={() => handleSort('deviceName')}
                        >
                          <div className="flex items-center">
                            Device
                            <SortIcon field="deviceName" />
                          </div>
                        </th>
                        <th
                          className="px-4 py-3 text-left text-sm font-medium text-slate-600 dark:text-slate-300 cursor-pointer hover:bg-slate-100 dark:hover:bg-slate-600"
                          onClick={() => handleSort('userDisplayName')}
                        >
                          <div className="flex items-center">
                            User
                            <SortIcon field="userDisplayName" />
                          </div>
                        </th>
                        <th
                          className="px-4 py-3 text-left text-sm font-medium text-slate-600 dark:text-slate-300 cursor-pointer hover:bg-slate-100 dark:hover:bg-slate-600"
                          onClick={() => handleSort('operatingSystem')}
                        >
                          <div className="flex items-center">
                            OS
                            <SortIcon field="operatingSystem" />
                          </div>
                        </th>
                        <th
                          className="px-4 py-3 text-left text-sm font-medium text-slate-600 dark:text-slate-300 cursor-pointer hover:bg-slate-100 dark:hover:bg-slate-600"
                          onClick={() => handleSort('complianceState')}
                        >
                          <div className="flex items-center">
                            Compliance
                            <SortIcon field="complianceState" />
                          </div>
                        </th>
                        <th className="px-4 py-3 text-left text-sm font-medium text-slate-600 dark:text-slate-300">
                          Ownership
                        </th>
                        <th
                          className="px-4 py-3 text-left text-sm font-medium text-slate-600 dark:text-slate-300 cursor-pointer hover:bg-slate-100 dark:hover:bg-slate-600"
                          onClick={() => handleSort('lastSyncDateTime')}
                        >
                          <div className="flex items-center">
                            Last Sync
                            <SortIcon field="lastSyncDateTime" />
                          </div>
                        </th>
                        <th className="px-4 py-3 text-right text-sm font-medium text-slate-600 dark:text-slate-300">
                          Actions
                        </th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-200 dark:divide-slate-700">
                      {filteredAndSortedDevices.map((device) => (
                        <tr
                          key={device.id}
                          className="hover:bg-slate-50 dark:hover:bg-slate-700/50 transition-colors"
                        >
                          <td className="px-4 py-3">
                            <div className="flex items-center gap-3">
                              <div className="w-10 h-10 rounded-full bg-slate-100 dark:bg-slate-700 flex items-center justify-center flex-shrink-0">
                                {getOsIcon(device.operatingSystem)}
                              </div>
                              <div>
                                <p className="font-medium text-slate-900 dark:text-white">
                                  {device.deviceName}
                                </p>
                                <p className="text-sm text-slate-500 dark:text-slate-400">
                                  {device.manufacturer} {device.model}
                                </p>
                              </div>
                            </div>
                          </td>
                          <td className="px-4 py-3">
                            <p className="text-sm text-slate-900 dark:text-white">
                              {device.userDisplayName || '-'}
                            </p>
                            <p className="text-xs text-slate-500 dark:text-slate-400 truncate max-w-xs">
                              {device.userPrincipalName || ''}
                            </p>
                          </td>
                          <td className="px-4 py-3">
                            <p className="text-sm text-slate-900 dark:text-white">
                              {device.operatingSystem || 'Unknown'}
                            </p>
                            <p className="text-xs text-slate-500 dark:text-slate-400">
                              {device.osVersion || ''}
                            </p>
                          </td>
                          <td className="px-4 py-3">
                            {getComplianceBadge(device.complianceState)}
                          </td>
                          <td className="px-4 py-3">
                            {device.managedDeviceOwnerType && (
                              <span className={`inline-flex items-center gap-1 px-2 py-1 rounded-full text-xs font-medium ${
                                device.managedDeviceOwnerType === 'Company'
                                  ? 'bg-indigo-100 text-indigo-700 dark:bg-indigo-900/30 dark:text-indigo-400'
                                  : 'bg-amber-100 text-amber-700 dark:bg-amber-900/30 dark:text-amber-400'
                              }`}>
                                {device.managedDeviceOwnerType === 'Company' ? <BuildingRegular className="w-3 h-3" /> : <PersonRegular className="w-3 h-3" />}
                                {device.managedDeviceOwnerType}
                              </span>
                            )}
                          </td>
                          <td className="px-4 py-3">
                            <span className="text-sm text-slate-600 dark:text-slate-300">
                              {formatDate(device.lastSyncDateTime)}
                            </span>
                          </td>
                          <td className="px-4 py-3 text-right">
                            <button
                              onClick={() => openInIntune(device.id)}
                              className="p-2 text-slate-400 hover:text-blue-600 hover:bg-blue-50 dark:hover:bg-blue-900/30 rounded-lg transition-colors"
                              title="Open in Intune"
                            >
                              <OpenRegular className="w-5 h-5" />
                            </button>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </>
            )}
          </div>
        </>
      ) : (
        /* Certificates View */
        <div className="space-y-4">
          {/* Apple Push Certificate Card */}
          <div className="bg-white dark:bg-slate-800 rounded-lg border border-slate-200 dark:border-slate-700 overflow-hidden">
            <div className="p-4 border-b border-slate-200 dark:border-slate-700">
              <div className="flex items-center gap-3">
                <div className="w-10 h-10 rounded-full bg-slate-100 dark:bg-slate-700 flex items-center justify-center">
                  <CertificateRegular className="w-5 h-5 text-slate-600 dark:text-slate-400" />
                </div>
                <div>
                  <h2 className="text-lg font-semibold text-slate-900 dark:text-white">Apple Push Notification Certificate</h2>
                  <p className="text-sm text-slate-500 dark:text-slate-400">Required for iOS/iPadOS and macOS device management</p>
                </div>
              </div>
            </div>

            {certLoading ? (
              <div className="flex items-center justify-center h-48">
                <div className="animate-spin rounded-full h-8 w-8 border-b-2 border-blue-600"></div>
              </div>
            ) : applePushCert?.permissionRequired ? (
              <div className="p-6">
                <div className="flex flex-col items-center justify-center text-center py-8">
                  <div className="w-16 h-16 rounded-full bg-amber-100 dark:bg-amber-900/30 flex items-center justify-center mb-4">
                    <WarningFilled className="w-8 h-8 text-amber-600" />
                  </div>
                  <h3 className="text-lg font-medium text-slate-900 dark:text-white mb-2">Permission Required</h3>
                  <p className="text-sm text-slate-500 dark:text-slate-400 max-w-md mb-4">
                    The <code className="bg-slate-100 dark:bg-slate-700 px-1 py-0.5 rounded text-xs">DeviceManagementServiceConfig.Read.All</code> permission is required to view the Apple Push certificate status.
                  </p>
                  <div className="space-y-2">
                    <a
                      href="https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/~/CallAnAPI/appId/7d9d26f4-9919-4e1d-adc6-1be2d4f60167"
                      target="_blank"
                      rel="noopener noreferrer"
                      className="inline-flex items-center gap-2 px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 transition-colors text-sm"
                    >
                      <OpenRegular className="w-4 h-4" />
                      Add Permission in Azure AD
                    </a>
                    <p className="text-xs text-slate-400">After adding, grant admin consent and sign out/in</p>
                  </div>
                </div>
              </div>
            ) : applePushCert?.error ? (
              <div className="p-6">
                <div className="flex items-center gap-3 text-red-600 dark:text-red-400">
                  <ErrorCircleRegular className="w-6 h-6" />
                  <div>
                    <p className="font-medium">Error loading certificate</p>
                    <p className="text-sm text-slate-500 dark:text-slate-400">{applePushCert.error}</p>
                  </div>
                </div>
              </div>
            ) : !applePushCert?.isConfigured ? (
              <div className="p-6">
                <div className="flex flex-col items-center justify-center text-center py-8">
                  <div className="w-16 h-16 rounded-full bg-slate-100 dark:bg-slate-700 flex items-center justify-center mb-4">
                    <CertificateRegular className="w-8 h-8 text-slate-400" />
                  </div>
                  <h3 className="text-lg font-medium text-slate-900 dark:text-white mb-2">Certificate Not Configured</h3>
                  <p className="text-sm text-slate-500 dark:text-slate-400 max-w-md mb-4">
                    An Apple Push Notification certificate is required to manage iOS, iPadOS, and macOS devices in Intune.
                  </p>
                  <a
                    href="https://intune.microsoft.com/#view/Microsoft_Intune_DeviceSettings/DevicesAppleMenu/~/appleEnrollment"
                    target="_blank"
                    rel="noopener noreferrer"
                    className="inline-flex items-center gap-2 px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 transition-colors text-sm"
                  >
                    <OpenRegular className="w-4 h-4" />
                    Configure in Intune
                  </a>
                </div>
              </div>
            ) : (
              <div className="p-6">
                {/* Status Banner */}
                <div className={`flex items-center gap-3 p-4 rounded-lg mb-6 ${getCertStatusColor(applePushCert.status)}`}>
                  {getCertStatusIcon(applePushCert.status)}
                  <div className="flex-1">
                    <p className="font-medium">
                      {applePushCert.status === 'Healthy' && 'Certificate is valid'}
                      {applePushCert.status === 'Warning' && 'Certificate expires soon'}
                      {applePushCert.status === 'Critical' && 'Certificate expires very soon!'}
                      {applePushCert.status === 'Expired' && 'Certificate has expired!'}
                    </p>
                    <p className="text-sm opacity-80">
                      {applePushCert.daysUntilExpiry !== undefined && applePushCert.daysUntilExpiry >= 0
                        ? `${applePushCert.daysUntilExpiry} days until expiry`
                        : applePushCert.daysUntilExpiry !== undefined
                        ? `Expired ${Math.abs(applePushCert.daysUntilExpiry)} days ago`
                        : 'Expiry date unknown'}
                    </p>
                  </div>
                  {(applePushCert.status === 'Warning' || applePushCert.status === 'Critical' || applePushCert.status === 'Expired') && (
                    <a
                      href="https://intune.microsoft.com/#view/Microsoft_Intune_DeviceSettings/DevicesAppleMenu/~/appleEnrollment"
                      target="_blank"
                      rel="noopener noreferrer"
                      className="px-3 py-1.5 bg-white dark:bg-slate-800 text-slate-900 dark:text-white rounded-lg hover:bg-slate-50 dark:hover:bg-slate-700 transition-colors text-sm font-medium"
                    >
                      Renew Certificate
                    </a>
                  )}
                </div>

                {/* Certificate Details */}
                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                  <div className="space-y-4">
                    <div>
                      <label className="text-xs font-medium text-slate-500 dark:text-slate-400 uppercase tracking-wider">Apple ID</label>
                      <p className="text-sm text-slate-900 dark:text-white mt-1">{applePushCert.appleIdentifier || 'N/A'}</p>
                    </div>
                    <div>
                      <label className="text-xs font-medium text-slate-500 dark:text-slate-400 uppercase tracking-wider">Topic Identifier</label>
                      <p className="text-sm text-slate-900 dark:text-white mt-1 font-mono text-xs break-all">{applePushCert.topicIdentifier || 'N/A'}</p>
                    </div>
                    <div>
                      <label className="text-xs font-medium text-slate-500 dark:text-slate-400 uppercase tracking-wider">Serial Number</label>
                      <p className="text-sm text-slate-900 dark:text-white mt-1 font-mono text-xs">{applePushCert.certificateSerialNumber || 'N/A'}</p>
                    </div>
                  </div>
                  <div className="space-y-4">
                    <div>
                      <label className="text-xs font-medium text-slate-500 dark:text-slate-400 uppercase tracking-wider">Expiration Date</label>
                      <div className="flex items-center gap-2 mt-1">
                        <CalendarRegular className="w-4 h-4 text-slate-400" />
                        <p className="text-sm text-slate-900 dark:text-white">{formatFullDate(applePushCert.expirationDateTime)}</p>
                      </div>
                    </div>
                    <div>
                      <label className="text-xs font-medium text-slate-500 dark:text-slate-400 uppercase tracking-wider">Last Modified</label>
                      <div className="flex items-center gap-2 mt-1">
                        <ClockRegular className="w-4 h-4 text-slate-400" />
                        <p className="text-sm text-slate-900 dark:text-white">{formatFullDate(applePushCert.lastModifiedDateTime)}</p>
                      </div>
                    </div>
                    <div>
                      <label className="text-xs font-medium text-slate-500 dark:text-slate-400 uppercase tracking-wider">Upload Status</label>
                      <p className="text-sm text-slate-900 dark:text-white mt-1">{applePushCert.certificateUploadStatus || 'N/A'}</p>
                    </div>
                  </div>
                </div>

                {/* Actions */}
                <div className="flex items-center gap-3 mt-6 pt-6 border-t border-slate-200 dark:border-slate-700">
                  <a
                    href="https://intune.microsoft.com/#view/Microsoft_Intune_DeviceSettings/DevicesAppleMenu/~/appleEnrollment"
                    target="_blank"
                    rel="noopener noreferrer"
                    className="inline-flex items-center gap-2 px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 transition-colors text-sm"
                  >
                    <OpenRegular className="w-4 h-4" />
                    Manage in Intune
                  </a>
                  <a
                    href="https://identity.apple.com/pushcert/"
                    target="_blank"
                    rel="noopener noreferrer"
                    className="inline-flex items-center gap-2 px-4 py-2 border border-slate-200 dark:border-slate-600 text-slate-700 dark:text-slate-300 rounded-lg hover:bg-slate-50 dark:hover:bg-slate-700 transition-colors text-sm"
                  >
                    <OpenRegular className="w-4 h-4" />
                    Apple Push Certificates Portal
                  </a>
                </div>
              </div>
            )}
          </div>

          {/* Info Card */}
          <div className="bg-blue-50 dark:bg-blue-900/20 rounded-lg border border-blue-200 dark:border-blue-800 p-4">
            <div className="flex gap-3">
              <div className="flex-shrink-0">
                <CertificateRegular className="w-5 h-5 text-blue-600 dark:text-blue-400" />
              </div>
              <div>
                <h3 className="text-sm font-medium text-blue-900 dark:text-blue-100">About Apple Push Notification Certificate</h3>
                <p className="text-sm text-blue-700 dark:text-blue-300 mt-1">
                  The Apple Push Notification (APN) certificate is required to enroll and manage iOS, iPadOS, and macOS devices in Microsoft Intune. 
                  The certificate must be renewed annually before it expires, or all enrolled Apple devices will lose management capabilities.
                </p>
              </div>
            </div>
          </div>

          {/* Apple Enrollment Program Tokens Card */}
          <div className="bg-white dark:bg-slate-800 rounded-lg border border-slate-200 dark:border-slate-700 overflow-hidden">
            <div className="p-4 border-b border-slate-200 dark:border-slate-700">
              <div className="flex items-center gap-3">
                <div className="w-10 h-10 rounded-full bg-slate-100 dark:bg-slate-700 flex items-center justify-center">
                  <PhoneRegular className="w-5 h-5 text-slate-600 dark:text-slate-400" />
                </div>
                <div>
                  <h2 className="text-lg font-semibold text-slate-900 dark:text-white">Apple Enrollment Program Tokens</h2>
                  <p className="text-sm text-slate-500 dark:text-slate-400">Automated Device Enrollment (ADE) tokens for zero-touch deployment</p>
                </div>
              </div>
            </div>

            {depTokensLoading ? (
              <div className="flex items-center justify-center h-48">
                <div className="animate-spin rounded-full h-8 w-8 border-b-2 border-blue-600"></div>
              </div>
            ) : depTokens?.permissionRequired ? (
              <div className="p-6">
                <div className="flex flex-col items-center justify-center text-center py-8">
                  <div className="w-16 h-16 rounded-full bg-amber-100 dark:bg-amber-900/30 flex items-center justify-center mb-4">
                    <WarningFilled className="w-8 h-8 text-amber-600" />
                  </div>
                  <h3 className="text-lg font-medium text-slate-900 dark:text-white mb-2">Permission Required</h3>
                  <p className="text-sm text-slate-500 dark:text-slate-400 max-w-md mb-4">
                    The <code className="bg-slate-100 dark:bg-slate-700 px-1 py-0.5 rounded text-xs">DeviceManagementServiceConfig.Read.All</code> permission is required to view enrollment tokens.
                  </p>
                </div>
              </div>
            ) : depTokens?.error ? (
              <div className="p-6">
                <div className="flex items-center gap-3 text-red-600 dark:text-red-400">
                  <ErrorCircleRegular className="w-6 h-6" />
                  <div>
                    <p className="font-medium">Error loading tokens</p>
                    <p className="text-sm text-slate-500 dark:text-slate-400">{depTokens.error}</p>
                  </div>
                </div>
              </div>
            ) : depTokens?.tokens.length === 0 ? (
              <div className="p-6">
                <div className="flex flex-col items-center justify-center text-center py-8">
                  <div className="w-16 h-16 rounded-full bg-slate-100 dark:bg-slate-700 flex items-center justify-center mb-4">
                    <PhoneRegular className="w-8 h-8 text-slate-400" />
                  </div>
                  <h3 className="text-lg font-medium text-slate-900 dark:text-white mb-2">No Enrollment Tokens</h3>
                  <p className="text-sm text-slate-500 dark:text-slate-400 max-w-md mb-4">
                    No Apple Enrollment Program tokens are configured. Add a token to enable automated device enrollment.
                  </p>
                  <a
                    href="https://intune.microsoft.com/#view/Microsoft_Intune_DeviceSettings/DevicesAppleMenu/~/iosEnrollment"
                    target="_blank"
                    rel="noopener noreferrer"
                    className="inline-flex items-center gap-2 px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 transition-colors text-sm"
                  >
                    <OpenRegular className="w-4 h-4" />
                    Configure in Intune
                  </a>
                </div>
              </div>
            ) : (
              <div className="p-6">
                <div className="space-y-4">
                  {depTokens?.tokens.map((token) => (
                    <div key={token.id} className="border border-slate-200 dark:border-slate-700 rounded-lg overflow-hidden">
                      {/* Token Status Banner */}
                      <div className={`flex items-center gap-3 p-3 ${getCertStatusColor(token.status)}`}>
                        {getCertStatusIcon(token.status)}
                        <div className="flex-1">
                          <p className="font-medium">{token.tokenName || 'Unnamed Token'}</p>
                          <p className="text-sm opacity-80">
                            {token.daysUntilExpiry !== undefined && token.daysUntilExpiry >= 0
                              ? `${token.daysUntilExpiry} days until expiry`
                              : token.daysUntilExpiry !== undefined
                              ? `Expired ${Math.abs(token.daysUntilExpiry)} days ago`
                              : 'Expiry date unknown'}
                          </p>
                        </div>
                        {(token.status === 'Warning' || token.status === 'Critical' || token.status === 'Expired') && (
                          <a
                            href="https://intune.microsoft.com/#view/Microsoft_Intune_DeviceSettings/DevicesAppleMenu/~/iosEnrollment"
                            target="_blank"
                            rel="noopener noreferrer"
                            className="px-3 py-1.5 bg-white dark:bg-slate-800 text-slate-900 dark:text-white rounded-lg hover:bg-slate-50 dark:hover:bg-slate-700 transition-colors text-sm font-medium"
                          >
                            Renew Token
                          </a>
                        )}
                      </div>
                      
                      {/* Token Details */}
                      <div className="p-4 bg-white dark:bg-slate-800">
                        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4 text-sm">
                          <div>
                            <label className="text-xs font-medium text-slate-500 dark:text-slate-400 uppercase tracking-wider">Apple ID</label>
                            <p className="text-slate-900 dark:text-white mt-1">{token.appleIdentifier || 'N/A'}</p>
                          </div>
                          <div>
                            <label className="text-xs font-medium text-slate-500 dark:text-slate-400 uppercase tracking-wider">Token Type</label>
                            <p className="text-slate-900 dark:text-white mt-1">{token.tokenType || 'N/A'}</p>
                          </div>
                          <div>
                            <label className="text-xs font-medium text-slate-500 dark:text-slate-400 uppercase tracking-wider">Expiration Date</label>
                            <div className="flex items-center gap-2 mt-1">
                              <CalendarRegular className="w-4 h-4 text-slate-400" />
                              <p className="text-slate-900 dark:text-white">{formatFullDate(token.tokenExpirationDateTime)}</p>
                            </div>
                          </div>
                          <div>
                            <label className="text-xs font-medium text-slate-500 dark:text-slate-400 uppercase tracking-wider">Last Sync</label>
                            <div className="flex items-center gap-2 mt-1">
                              <ClockRegular className="w-4 h-4 text-slate-400" />
                              <p className="text-slate-900 dark:text-white">{formatFullDate(token.lastSuccessfulSyncDateTime)}</p>
                            </div>
                          </div>
                          <div>
                            <label className="text-xs font-medium text-slate-500 dark:text-slate-400 uppercase tracking-wider">Data Sharing</label>
                            <p className="text-slate-900 dark:text-white mt-1">
                              {token.dataSharingConsentGranted ? (
                                <span className="inline-flex items-center gap-1 text-green-600 dark:text-green-400">
                                  <CheckmarkCircleFilled className="w-4 h-4" /> Granted
                                </span>
                              ) : (
                                <span className="inline-flex items-center gap-1 text-slate-500">
                                  <DismissCircleFilled className="w-4 h-4" /> Not Granted
                                </span>
                              )}
                            </p>
                          </div>
                          {token.lastSyncErrorCode !== undefined && token.lastSyncErrorCode !== 0 && (
                            <div>
                              <label className="text-xs font-medium text-slate-500 dark:text-slate-400 uppercase tracking-wider">Sync Error</label>
                              <p className="text-red-600 dark:text-red-400 mt-1">Error code: {token.lastSyncErrorCode}</p>
                            </div>
                          )}
                        </div>
                      </div>
                    </div>
                  ))}
                </div>

                {/* Actions */}
                <div className="flex items-center gap-3 mt-6 pt-6 border-t border-slate-200 dark:border-slate-700">
                  <a
                    href="https://intune.microsoft.com/#view/Microsoft_Intune_DeviceSettings/DevicesAppleMenu/~/iosEnrollment"
                    target="_blank"
                    rel="noopener noreferrer"
                    className="inline-flex items-center gap-2 px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 transition-colors text-sm"
                  >
                    <OpenRegular className="w-4 h-4" />
                    Manage in Intune
                  </a>
                  <a
                    href="https://business.apple.com/"
                    target="_blank"
                    rel="noopener noreferrer"
                    className="inline-flex items-center gap-2 px-4 py-2 border border-slate-200 dark:border-slate-600 text-slate-700 dark:text-slate-300 rounded-lg hover:bg-slate-50 dark:hover:bg-slate-700 transition-colors text-sm"
                  >
                    <OpenRegular className="w-4 h-4" />
                    Apple Business Manager
                  </a>
                </div>
              </div>
            )}
          </div>

          {/* DEP Info Card */}
          <div className="bg-purple-50 dark:bg-purple-900/20 rounded-lg border border-purple-200 dark:border-purple-800 p-4">
            <div className="flex gap-3">
              <div className="flex-shrink-0">
                <PhoneRegular className="w-5 h-5 text-purple-600 dark:text-purple-400" />
              </div>
              <div>
                <h3 className="text-sm font-medium text-purple-900 dark:text-purple-100">About Apple Enrollment Program</h3>
                <p className="text-sm text-purple-700 dark:text-purple-300 mt-1">
                  Apple Automated Device Enrollment (ADE), formerly known as DEP, enables zero-touch deployment of iOS, iPadOS, and macOS devices. 
                  Tokens must be renewed annually in Apple Business Manager or Apple School Manager and re-uploaded to Intune.
                </p>
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

export default DevicesPage;
