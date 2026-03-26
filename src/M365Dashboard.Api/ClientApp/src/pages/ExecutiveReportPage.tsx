import React, { useState, useEffect, useCallback } from 'react';
import {
  DocumentTextRegular,
  ShieldCheckmarkRegular,
  LaptopRegular,
  PersonRegular,
  MailRegular,
  CloudRegular,
  WarningRegular,
  CheckmarkCircleFilled,
  DismissCircleFilled,
  InfoRegular,
  ArrowSyncRegular,
} from '@fluentui/react-icons';
import { useAppContext } from '../contexts/AppContext';

interface SecureScoreData {
  currentScore: number;
  maxScore: number;
  percentageScore: number;
  identityScore: number;
  identityMaxScore: number;
  identityPercentage: number;
  deviceScore: number;
  deviceMaxScore: number;
  devicePercentage: number;
  appsScore: number;
  appsMaxScore: number;
  appsPercentage: number;
}

interface DeviceStatsData {
  totalDevices: number;
  windowsDevices: number;
  macOsDevices: number;
  iosDevices: number;
  androidDevices: number;
  compliantDevices: number;
  nonCompliantDevices: number;
  complianceRate: number;
}

interface UserStatsData {
  totalUsers: number;
  guestUsers: number;
  deletedUsers: number;
  mfaRegistered: number;
  mfaNotRegistered: number;
}

interface DefenderStatsData {
  exposureScore?: string;
  exposureScoreNumeric?: number;
  vulnerabilitiesDetected: number;
  criticalVulnerabilities: number;
  highVulnerabilities: number;
  mediumVulnerabilities: number;
  lowVulnerabilities: number;
  onboardedMachines?: number;
  onboardedMachineNames?: string[];
  note?: string;
}

interface MailboxStatsData {
  totalMailboxes: number;
  activeMailboxes: number;
  totalStorageUsedGB: number;
  averageStorageGB: number;
}

interface SharePointStatsData {
  totalSites: number;
  activeSites: number;
  totalStorageUsedGB: number;
}

interface AttackSimulationData {
  totalSimulations: number;
  completedSimulations: number;
  averageCompromiseRate: number;
  note?: string;
}

interface EmailSecurityData {
  totalMessages: number;
  spamMessages: number;
  malwareMessages: number;
  phishingMessages: number;
  bulkMessages?: number;
  note?: string;
}

interface WindowsUpdateStatsData {
  totalWindowsDevices: number;
  upToDate: number;
  needsUpdate: number;
  complianceRate: number;
  note?: string;
}

interface CloudAppDiscoveryData {
  discoveredApps: number;
  sanctionedApps: number;
  unsanctionedApps: number;
  note?: string;
}

interface UserSignInDetailData {
  displayName?: string;
  userPrincipalName?: string;
  lastInteractiveSignIn?: string;
  lastNonInteractiveSignIn?: string;
  defaultMfaMethod?: string;
  isMfaRegistered: boolean;
  accountEnabled: boolean;
}

interface DeletedUserData {
  displayName?: string;
  userPrincipalName?: string;
  mail?: string;
  deletedDateTime?: string;
  jobTitle?: string;
  department?: string;
}

interface MailboxDetailData {
  displayName?: string;
  userPrincipalName?: string;
  recipientType?: string;
  storageUsedBytes: number;
  storageUsedGB: number;
  quotaBytes?: number;
  quotaGB?: number;
  percentUsed?: number;
  lastActivityDate?: string;
  itemCount?: number;
}

interface WindowsDeviceDetailData {
  deviceName?: string;
  lastCheckIn?: string;
  osVersion?: string;
  complianceState?: string;
  managementAgent?: string;
  ownership?: string;
  skuFamily?: string;
  osVersionStatus?: 'Current' | 'Warning' | 'Critical' | 'Unknown';
  osVersionStatusMessage?: string;
}

interface IosDeviceDetailData {
  deviceName?: string;
  complianceState?: string;
  managementAgent?: string;
  ownership?: string;
  os?: string;
  osVersion?: string;
  lastCheckIn?: string;
  osVersionStatus?: 'Current' | 'Warning' | 'Critical' | 'Unknown';
  osVersionStatusMessage?: string;
}

interface AndroidDeviceDetailData {
  deviceName?: string;
  complianceState?: string;
  managementAgent?: string;
  os?: string;
  osVersion?: string;
  lastCheckIn?: string;
  securityPatchLevel?: string;
  osVersionStatus?: 'Current' | 'Warning' | 'Critical' | 'Unknown';
  osVersionStatusMessage?: string;
}

interface MacDeviceDetailData {
  deviceName?: string;
  lastCheckIn?: string;
  osVersion?: string;
  complianceState?: string;
  managementAgent?: string;
  ownership?: string;
  osVersionStatus?: 'Current' | 'Warning' | 'Critical' | 'Unknown';
  osVersionStatusMessage?: string;
}

interface AppCredentialStatusData {
  totalApps: number;
  appsWithExpiringSecrets: number;
  appsWithExpiredSecrets: number;
  appsWithExpiringCertificates: number;
  appsWithExpiredCertificates: number;
  thresholdDays: number;
  expiringSecrets: AppCredentialDetail[];
  expiredSecrets: AppCredentialDetail[];
  expiringCertificates: AppCredentialDetail[];
  expiredCertificates: AppCredentialDetail[];
}

interface AppCredentialDetail {
  appName?: string;
  appId?: string;
  credentialType?: string;
  description?: string;
  expiryDate?: string;
  daysUntilExpiry: number;
  status?: string;
}

interface DeviceDetailsData {
  windowsDevices: WindowsDeviceDetailData[];
  iosDevices: IosDeviceDetailData[];
  androidDevices: AndroidDeviceDetailData[];
  macDevices: MacDeviceDetailData[];
}

// Exposure Score Gauge Component - clean Microsoft Defender style
const ExposureScoreGauge: React.FC<{ score: number }> = ({ score }) => {
  // Score ranges: Low 0-29, Medium 30-69, High 70-100
  const getScoreColor = (s: number) => {
    if (s <= 29) return '#f97316'; // Orange
    if (s <= 69) return '#dc2626'; // Red
    return '#7f1d1d'; // Dark red
  };

  const color = getScoreColor(score);
  const needleAngle = -90 + (score / 100) * 180;

  // Arc parameters
  const cx = 100, cy = 90, r = 70;
  
  // Helper to calculate arc endpoint
  const polarToCartesian = (angle: number) => {
    const rad = (angle - 90) * Math.PI / 180;
    return {
      x: cx + r * Math.cos(rad),
      y: cy + r * Math.sin(rad)
    };
  };

  // Create arc path
  const createArc = (startAngle: number, endAngle: number) => {
    const start = polarToCartesian(startAngle);
    const end = polarToCartesian(endAngle);
    const largeArc = endAngle - startAngle > 180 ? 1 : 0;
    return `M ${start.x} ${start.y} A ${r} ${r} 0 ${largeArc} 1 ${end.x} ${end.y}`;
  };

  return (
    <div className="flex flex-col items-center">
      <div style={{ width: '200px', height: '160px' }}>
        <svg width="200" height="160" viewBox="0 0 200 160">
          {/* Background track */}
          <path
            d={createArc(-90, 90)}
            fill="none"
            stroke="#e5e7eb"
            strokeWidth="14"
            className="dark:stroke-slate-700"
          />
          
          {/* Low segment: 0-29 (-90° to -36°) */}
          <path
            d={createArc(-90, -36)}
            fill="none"
            stroke="#f97316"
            strokeWidth="14"
            strokeLinecap="round"
          />
          
          {/* Medium segment: 30-69 (-36° to 36°) */}
          <path
            d={createArc(-33, 33)}
            fill="none"
            stroke="#dc2626"
            strokeWidth="14"
          />
          
          {/* High segment: 70-100 (36° to 90°) */}
          <path
            d={createArc(36, 90)}
            fill="none"
            stroke="#7f1d1d"
            strokeWidth="14"
            strokeLinecap="round"
          />
          
          {/* Needle */}
          <g transform={`rotate(${needleAngle} ${cx} ${cy})`}>
            <line
              x1={cx}
              y1={cy}
              x2={cx}
              y2={cy - r + 15}
              stroke={color}
              strokeWidth="3"
              strokeLinecap="round"
            />
            <circle cx={cx} cy={cy} r="6" fill={color} />
            <circle cx={cx} cy={cy} r="3" fill="white" />
          </g>
          
          {/* Score text - positioned well below the gauge */}
          <text
            x={cx}
            y="145"
            textAnchor="middle"
            style={{ fill: color }}
            fontSize="28"
            fontWeight="bold"
          >
            {score}
            <tspan fill="#9ca3af" fontSize="20" fontWeight="normal">/100</tspan>
          </text>
        </svg>
      </div>
      
      {/* Legend row */}
      <div className="flex items-center justify-center gap-5 text-xs text-slate-500 dark:text-slate-400">
        <div className="flex items-center gap-1.5">
          <div className="w-2.5 h-2.5 rounded-full bg-orange-500"></div>
          <span>Low 0-29</span>
        </div>
        <div className="flex items-center gap-1.5">
          <div className="w-2.5 h-2.5 rounded-full bg-red-600"></div>
          <span>Medium 30-69</span>
        </div>
        <div className="flex items-center gap-1.5">
          <div className="w-2.5 h-2.5 rounded-full bg-red-900"></div>
          <span>High 70-100</span>
        </div>
      </div>
    </div>
  );
};

interface ExecutiveReportData {
  reportMonth: string;
  generatedAt: string;
  startDate: string;
  endDate: string;
  secureScore?: SecureScoreData;
  deviceStats?: DeviceStatsData;
  userStats?: UserStatsData;
  defenderStats?: DefenderStatsData;
  mailboxStats?: MailboxStatsData;
  sharePointStats?: SharePointStatsData;
  attackSimulation?: AttackSimulationData;
  emailSecurity?: EmailSecurityData;
  windowsUpdateStats?: WindowsUpdateStatsData;
  cloudAppDiscovery?: CloudAppDiscoveryData;
  riskyUsersCount: number;
  highRiskUsers?: string[];
  userSignInDetails?: UserSignInDetailData[];
  deletedUsersInPeriod?: DeletedUserData[];
  mailboxDetails?: MailboxDetailData[];
  deviceDetails?: DeviceDetailsData;
  appCredentialStatus?: AppCredentialStatusData;
}

const ExecutiveReportPage: React.FC = () => {
  const { getAccessToken } = useAppContext();
  const [reportData, setReportData] = useState<ExecutiveReportData | null>(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const fetchReportData = useCallback(async () => {
    try {
      setLoading(true);
      setError(null);
      const token = await getAccessToken();
      
      const response = await fetch(`/api/executivereport/data`, {
        headers: {
          'Authorization': `Bearer ${token}`,
          'Content-Type': 'application/json',
        },
      });

      if (!response.ok) {
        throw new Error('Failed to fetch report data');
      }

      const data = await response.json();
      setReportData(data);
    } catch (err) {
      setError(err instanceof Error ? err.message : 'An error occurred');
    } finally {
      setLoading(false);
    }
  }, [getAccessToken]);

  const downloadReport = async () => {
    try {
      const token = await getAccessToken();
      
      const response = await fetch(`/api/executivereport/download`, {
        headers: {
          'Authorization': `Bearer ${token}`,
        },
      });

      if (!response.ok) {
        throw new Error('Failed to download report');
      }

      // Get filename from Content-Disposition header or use default
      const contentDisposition = response.headers.get('Content-Disposition');
      // Backend returns either PDF or Word depending on platform support
      let filename = `Executive_Summary_${reportData?.reportMonth?.replace(' ', '_') || 'Report'}`;
      if (contentDisposition) {
        const filenameMatch = contentDisposition.match(/filename="?([^"]+)"?/);
        if (filenameMatch && filenameMatch[1]) {
          filename = filenameMatch[1];
        }
      }

      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      window.URL.revokeObjectURL(url);
      document.body.removeChild(a);
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Failed to download report');
    }
  };

  useEffect(() => {
    fetchReportData();
  }, []);

  const getScoreColor = (score: number) => {
    if (score >= 70) return 'text-green-600';
    if (score >= 50) return 'text-amber-600';
    return 'text-red-600';
  };

  const getScoreBgColor = (score: number) => {
    if (score >= 70) return 'bg-green-100 dark:bg-green-900/30';
    if (score >= 50) return 'bg-amber-100 dark:bg-amber-900/30';
    return 'bg-red-100 dark:bg-red-900/30';
  };

  const StatCard: React.FC<{
    title: string;
    value: string | number;
    subtitle?: string;
    icon: React.ReactNode;
    color?: string;
  }> = ({ title, value, subtitle, icon, color = 'text-blue-600' }) => (
    <div className="bg-white dark:bg-slate-800 rounded-lg border border-slate-200 dark:border-slate-700 p-4">
      <div className="flex items-center gap-3">
        <div className={`p-2 rounded-lg bg-slate-100 dark:bg-slate-700 ${color}`}>
          {icon}
        </div>
        <div>
          <p className="text-sm text-slate-500 dark:text-slate-400">{title}</p>
          <p className={`text-xl font-semibold ${color}`}>{value}</p>
          {subtitle && <p className="text-xs text-slate-400">{subtitle}</p>}
        </div>
      </div>
    </div>
  );

  const SectionCard: React.FC<{
    title: string;
    icon: React.ReactNode;
    children: React.ReactNode;
    note?: string;
  }> = ({ title, icon, children, note }) => (
    <div className="bg-white dark:bg-slate-800 rounded-lg border border-slate-200 dark:border-slate-700 overflow-hidden">
      <div className="p-4 border-b border-slate-200 dark:border-slate-700 flex items-center gap-2">
        {icon}
        <h3 className="font-semibold text-slate-900 dark:text-white">{title}</h3>
      </div>
      <div className="p-4">
        {children}
        {note && (
          <div className="mt-3 flex items-start gap-2 text-xs text-slate-500 dark:text-slate-400 bg-slate-50 dark:bg-slate-900/50 p-2 rounded">
            <InfoRegular className="w-4 h-4 flex-shrink-0 mt-0.5" />
            <span>{note}</span>
          </div>
        )}
      </div>
    </div>
  );

  // Version Status Badge Component
  const VersionStatusBadge: React.FC<{ status?: string; message?: string }> = ({ status, message }) => {
    const getStatusStyle = () => {
      switch (status) {
        case 'Current':
          return 'bg-green-100 text-green-800 dark:bg-green-900/30 dark:text-green-400';
        case 'Warning':
          return 'bg-amber-100 text-amber-800 dark:bg-amber-900/30 dark:text-amber-400';
        case 'Critical':
          return 'bg-red-100 text-red-800 dark:bg-red-900/30 dark:text-red-400';
        default:
          return 'bg-slate-100 text-slate-600 dark:bg-slate-700 dark:text-slate-400';
      }
    };

    const getIcon = () => {
      switch (status) {
        case 'Current':
          return '✓';
        case 'Warning':
          return '⚠';
        case 'Critical':
          return '❌';
        default:
          return '?';
      }
    };

    return (
      <span className={`inline-flex items-center gap-1 px-2 py-0.5 rounded text-xs font-medium whitespace-nowrap ${getStatusStyle()}`}
            title={message || status || 'Unknown'}>
        <span>{getIcon()}</span>
        <span>{message || status || 'Unknown'}</span>
      </span>
    );
  };

  const DataRow: React.FC<{
    label: string;
    value: string | number;
    color?: 'green' | 'amber' | 'red' | 'default';
  }> = ({ label, value, color = 'default' }) => {
    const colorClasses = {
      green: 'text-green-600 dark:text-green-400',
      amber: 'text-amber-600 dark:text-amber-400',
      red: 'text-red-600 dark:text-red-400',
      default: 'text-slate-900 dark:text-white',
    };

    return (
      <div className="flex justify-between py-1.5 border-b border-slate-100 dark:border-slate-700 last:border-0">
        <span className="text-sm text-slate-600 dark:text-slate-400">{label}</span>
        <span className={`text-sm font-medium ${colorClasses[color]}`}>{value}</span>
      </div>
    );
  };

  return (
    <div className="p-4 space-y-4 w-full max-w-full overflow-hidden">
      {/* Header */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
        <div>
          <h1 className="text-xl font-semibold text-slate-900 dark:text-white">Executive Summary Report</h1>
          <p className="text-sm text-slate-500 dark:text-slate-400">
            Monthly security and usage overview for your Microsoft 365 tenant
          </p>
        </div>
        <div className="flex items-center gap-2">
          <button
            onClick={() => fetchReportData()}
            disabled={loading}
            className="px-3 py-2 bg-slate-100 dark:bg-slate-700 text-slate-700 dark:text-slate-300 rounded-lg hover:bg-slate-200 dark:hover:bg-slate-600 transition-colors disabled:opacity-50"
          >
            <ArrowSyncRegular className={`w-4 h-4 ${loading ? 'animate-spin' : ''}`} />
          </button>
          <button
            onClick={() => downloadReport()}
            disabled={!reportData}
            className="inline-flex items-center gap-2 px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 transition-colors disabled:opacity-50"
          >
            <DocumentTextRegular className="w-4 h-4" />
            <span>Download Report</span>
          </button>
        </div>
      </div>

      {/* Loading State */}
      {loading && (
        <div className="flex items-center justify-center h-64">
          <div className="text-center">
            <div className="animate-spin rounded-full h-8 w-8 border-b-2 border-blue-600 mx-auto mb-4"></div>
            <p className="text-slate-500 dark:text-slate-400">Generating report data...</p>
          </div>
        </div>
      )}

      {/* Error State */}
      {error && (
        <div className="bg-red-50 dark:bg-red-900/20 border border-red-200 dark:border-red-800 rounded-lg p-4">
          <p className="text-red-700 dark:text-red-400">{error}</p>
        </div>
      )}

      {/* Report Content */}
      {reportData && !loading && (
        <>
          {/* Report Header */}
          <div className="bg-gradient-to-r from-blue-600 to-blue-700 rounded-lg p-6 text-white">
            <div className="flex items-center gap-3 mb-2">
              <DocumentTextRegular className="w-8 h-8" />
              <div>
                <h2 className="text-2xl font-bold">Microsoft 365 Executive Summary</h2>
                <p className="text-blue-100">{reportData.reportMonth}</p>
              </div>
            </div>
            <p className="text-sm text-blue-200 mt-2">
              Generated: {new Date(reportData.generatedAt).toLocaleString()}
            </p>
          </div>

          {/* Security Score */}
          <div className="grid grid-cols-1 gap-4">
            <div className={`rounded-lg p-6 ${getScoreBgColor(reportData.secureScore?.percentageScore || 0)}`}>
              <div className="flex items-center justify-between">
                <div className="flex-1">
                  <p className="text-sm font-medium text-slate-600 dark:text-slate-400">Microsoft Secure Score</p>
                  <p className={`text-4xl font-bold ${getScoreColor(reportData.secureScore?.percentageScore || 0)}`}>
                    {reportData.secureScore?.percentageScore || 0}%
                  </p>
                  <p className="text-sm text-slate-500 dark:text-slate-400">
                    {reportData.secureScore?.currentScore || 0} / {reportData.secureScore?.maxScore || 0} points
                  </p>

                </div>
                <ShieldCheckmarkRegular className={`w-16 h-16 ${getScoreColor(reportData.secureScore?.percentageScore || 0)} opacity-50 ml-4 self-start`} />
              </div>
            </div>
          </div>

          {/* Main Content Grid */}
          <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
            {/* Intune Managed Devices */}
            <SectionCard 
              title="Intune Managed Devices" 
              icon={<LaptopRegular className="w-5 h-5 text-blue-600" />}
            >
              <div className="grid grid-cols-2 gap-4 mb-4">
                <div className="text-center p-3 bg-slate-50 dark:bg-slate-900/50 rounded-lg">
                  <p className="text-2xl font-bold text-blue-600">{reportData.deviceStats?.totalDevices || 0}</p>
                  <p className="text-xs text-slate-500">Total Devices</p>
                </div>
                <div className="text-center p-3 bg-slate-50 dark:bg-slate-900/50 rounded-lg">
                  <p className={`text-2xl font-bold ${(reportData.deviceStats?.complianceRate || 0) >= 90 ? 'text-green-600' : 'text-amber-600'}`}>
                    {reportData.deviceStats?.complianceRate || 0}%
                  </p>
                  <p className="text-xs text-slate-500">Compliance Rate</p>
                </div>
              </div>
              <DataRow label="Windows" value={reportData.deviceStats?.windowsDevices || 0} />
              <DataRow label="macOS" value={reportData.deviceStats?.macOsDevices || 0} />
              <DataRow label="iOS/iPadOS" value={reportData.deviceStats?.iosDevices || 0} />
              <DataRow label="Android" value={reportData.deviceStats?.androidDevices || 0} />
              <DataRow label="Compliant" value={reportData.deviceStats?.compliantDevices || 0} color="green" />
              <DataRow label="Non-Compliant" value={reportData.deviceStats?.nonCompliantDevices || 0} color="red" />
            </SectionCard>



            {/* Microsoft Defender */}
            <SectionCard 
              title="Microsoft Defender for Endpoint" 
              icon={<ShieldCheckmarkRegular className="w-5 h-5 text-red-600" />}
              note={reportData.defenderStats?.note}
            >
              {/* Exposure Score Gauge */}
              <div className="flex flex-col items-center">
                <ExposureScoreGauge 
                  score={reportData.defenderStats?.exposureScoreNumeric ?? 0} 
                />
              </div>

            </SectionCard>

            {/* User Accounts */}
            <SectionCard 
              title="User Accounts" 
              icon={<PersonRegular className="w-5 h-5 text-purple-600" />}
            >
              <DataRow label="Total Users" value={reportData.userStats?.totalUsers || 0} />
              <DataRow label="Guest Users" value={reportData.userStats?.guestUsers || 0} />
              <DataRow label="Deleted Users (Soft)" value={reportData.userStats?.deletedUsers || 0} />
              <DataRow label="MFA Registered" value={reportData.userStats?.mfaRegistered || 0} color="green" />
              <DataRow label="MFA Not Registered" value={reportData.userStats?.mfaNotRegistered || 0} color="amber" />
              <div className="mt-3 pt-3 border-t border-slate-200 dark:border-slate-700">
                <DataRow label="Risky Users" value={reportData.riskyUsersCount || 0} color={reportData.riskyUsersCount > 0 ? 'red' : 'default'} />
                {reportData.highRiskUsers && reportData.highRiskUsers.length > 0 && (
                  <div className="mt-2 p-2 bg-red-50 dark:bg-red-900/20 rounded text-xs text-red-700 dark:text-red-400">
                    <strong>High Risk:</strong> {reportData.highRiskUsers.join(', ')}
                  </div>
                )}
              </div>
            </SectionCard>

            {/* Attack Simulation Training */}
            <SectionCard 
              title="Attack Simulation Training" 
              icon={<WarningRegular className="w-5 h-5 text-orange-600" />}
              note={reportData.attackSimulation?.note}
            >
              <DataRow label="Total Simulations" value={reportData.attackSimulation?.totalSimulations || 0} />
              <DataRow label="Completed" value={reportData.attackSimulation?.completedSimulations || 0} />
              <DataRow 
                label="Average Compromise Rate" 
                value={`${reportData.attackSimulation?.averageCompromiseRate || 0}%`} 
                color={(reportData.attackSimulation?.averageCompromiseRate || 0) > 20 ? 'red' : 'green'}
              />
            </SectionCard>

            {/* Shadow IT */}
            <SectionCard 
              title="Shadow IT (Cloud App Discovery)" 
              icon={<CloudRegular className="w-5 h-5 text-slate-600" />}
              note={reportData.cloudAppDiscovery?.note}
            >
              <DataRow label="Discovered Apps" value={reportData.cloudAppDiscovery?.discoveredApps || 0} />
              <DataRow label="Sanctioned" value={reportData.cloudAppDiscovery?.sanctionedApps || 0} color="green" />
              <DataRow label="Unsanctioned" value={reportData.cloudAppDiscovery?.unsanctionedApps || 0} color="amber" />
            </SectionCard>

            {/* Mailbox Usage */}
            <SectionCard 
              title="Mailbox Usage" 
              icon={<MailRegular className="w-5 h-5 text-blue-600" />}
            >
              <DataRow label="Total Mailboxes" value={reportData.mailboxStats?.totalMailboxes || 0} />
              <DataRow label="Active Mailboxes" value={reportData.mailboxStats?.activeMailboxes || 0} />
              <DataRow label="Total Storage Used" value={`${reportData.mailboxStats?.totalStorageUsedGB || 0} GB`} />
              <DataRow label="Average Storage" value={`${reportData.mailboxStats?.averageStorageGB || 0} GB`} />
            </SectionCard>

            {/* SharePoint Usage */}
            <SectionCard 
              title="SharePoint Usage" 
              icon={<CloudRegular className="w-5 h-5 text-teal-600" />}
            >
              <DataRow label="Total Sites" value={reportData.sharePointStats?.totalSites || 0} />
              <DataRow label="Active Sites" value={reportData.sharePointStats?.activeSites || 0} />
              <DataRow label="Storage Used" value={`${reportData.sharePointStats?.totalStorageUsedGB || 0} GB`} />
            </SectionCard>

            {/* Email Security */}
            <SectionCard 
              title="Email Security (Last 30 Days)" 
              icon={<ShieldCheckmarkRegular className="w-5 h-5 text-green-600" />}
              note={reportData.emailSecurity?.note}
            >
              <DataRow label="Total Messages" value={reportData.emailSecurity?.totalMessages?.toLocaleString() || 0} />
            </SectionCard>

            {/* App Registration Credentials */}
            <SectionCard 
              title={`App Secrets & Certificates (${reportData.appCredentialStatus?.thresholdDays || 45} day threshold)`}
              icon={<WarningRegular className="w-5 h-5 text-amber-600" />}
            >
              <DataRow label="Total App Registrations" value={reportData.appCredentialStatus?.totalApps || 0} />
              <DataRow 
                label="Apps with Expiring Secrets" 
                value={reportData.appCredentialStatus?.appsWithExpiringSecrets || 0} 
                color={(reportData.appCredentialStatus?.appsWithExpiringSecrets || 0) > 0 ? 'amber' : 'default'}
              />
              <DataRow 
                label="Apps with Expired Secrets" 
                value={reportData.appCredentialStatus?.appsWithExpiredSecrets || 0} 
                color={(reportData.appCredentialStatus?.appsWithExpiredSecrets || 0) > 0 ? 'red' : 'default'}
              />
              <DataRow 
                label="Apps with Expiring Certificates" 
                value={reportData.appCredentialStatus?.appsWithExpiringCertificates || 0} 
                color={(reportData.appCredentialStatus?.appsWithExpiringCertificates || 0) > 0 ? 'amber' : 'default'}
              />
              <DataRow 
                label="Apps with Expired Certificates" 
                value={reportData.appCredentialStatus?.appsWithExpiredCertificates || 0} 
                color={(reportData.appCredentialStatus?.appsWithExpiredCertificates || 0) > 0 ? 'red' : 'default'}
              />
            </SectionCard>
          </div>

          {/* Windows Devices */}
          {reportData.deviceDetails?.windowsDevices && reportData.deviceDetails.windowsDevices.length > 0 && (
            <div className="mt-6">
              <div className="bg-white dark:bg-slate-800 rounded-lg border border-slate-200 dark:border-slate-700 overflow-hidden">
                <div className="px-4 py-3 bg-blue-50 dark:bg-blue-900/20 border-b border-slate-200 dark:border-slate-700">
                  <div className="flex items-center gap-2">
                    <LaptopRegular className="w-5 h-5 text-blue-600" />
                    <h3 className="font-semibold text-slate-700 dark:text-slate-200">Windows Devices</h3>
                    <span className="text-xs text-blue-600 dark:text-blue-400">({reportData.deviceDetails.windowsDevices.length} devices)</span>
                  </div>
                </div>
                <div className="overflow-x-auto">
                  <table className="w-full text-sm">
                    <thead className="bg-slate-100 dark:bg-slate-900">
                      <tr>
                        <th className="px-4 py-2 text-left font-medium text-slate-600 dark:text-slate-400">Device Name</th>
                        <th className="px-4 py-2 text-left font-medium text-slate-600 dark:text-slate-400">OS Version</th>
                        <th className="px-4 py-2 text-left font-medium text-slate-600 dark:text-slate-400">Update Status</th>
                        <th className="px-4 py-2 text-left font-medium text-slate-600 dark:text-slate-400">Compliance</th>
                        <th className="px-4 py-2 text-left font-medium text-slate-600 dark:text-slate-400">Model</th>
                        <th className="px-4 py-2 text-left font-medium text-slate-600 dark:text-slate-400">Last Check-in</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-200 dark:divide-slate-700">
                      {reportData.deviceDetails.windowsDevices.map((device, index) => (
                        <tr key={index} className="hover:bg-slate-50 dark:hover:bg-slate-900/30">
                          <td className="px-4 py-2 text-slate-700 dark:text-slate-300 font-medium">{device.deviceName || '-'}</td>
                          <td className="px-4 py-2 text-slate-600 dark:text-slate-400 text-xs">{device.osVersion || '-'}</td>
                          <td className="px-4 py-2">
                            <VersionStatusBadge status={device.osVersionStatus} message={device.osVersionStatusMessage} />
                          </td>
                          <td className="px-4 py-2">
                            <span className={`inline-flex items-center px-2 py-0.5 rounded text-xs font-medium ${
                              device.complianceState === 'Compliant' 
                                ? 'bg-green-100 text-green-800 dark:bg-green-900/30 dark:text-green-400'
                                : device.complianceState === 'Non-Compliant'
                                ? 'bg-red-100 text-red-800 dark:bg-red-900/30 dark:text-red-400'
                                : 'bg-slate-100 text-slate-600 dark:bg-slate-700 dark:text-slate-400'
                            }`}>
                              {device.complianceState || 'Unknown'}
                            </span>
                          </td>
                          <td className="px-4 py-2 text-slate-600 dark:text-slate-400">{device.skuFamily || '-'}</td>
                          <td className="px-4 py-2 text-slate-600 dark:text-slate-400">
                            {device.lastCheckIn 
                              ? new Date(device.lastCheckIn).toLocaleDateString()
                              : <span className="text-slate-400">Never</span>}
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>
          )}

          {/* macOS Devices */}
          {reportData.deviceDetails?.macDevices && reportData.deviceDetails.macDevices.length > 0 && (
            <div className="mt-6">
              <div className="bg-white dark:bg-slate-800 rounded-lg border border-slate-200 dark:border-slate-700 overflow-hidden">
                <div className="px-4 py-3 bg-slate-50 dark:bg-slate-900/50 border-b border-slate-200 dark:border-slate-700">
                  <div className="flex items-center gap-2">
                    <LaptopRegular className="w-5 h-5 text-slate-600" />
                    <h3 className="font-semibold text-slate-700 dark:text-slate-200">macOS Devices</h3>
                    <span className="text-xs text-slate-500">({reportData.deviceDetails.macDevices.length} devices)</span>
                  </div>
                </div>
                <div className="overflow-x-auto">
                  <table className="w-full text-sm">
                    <thead className="bg-slate-100 dark:bg-slate-900">
                      <tr>
                        <th className="px-4 py-2 text-left font-medium text-slate-600 dark:text-slate-400">Device Name</th>
                        <th className="px-4 py-2 text-left font-medium text-slate-600 dark:text-slate-400">OS Version</th>
                        <th className="px-4 py-2 text-left font-medium text-slate-600 dark:text-slate-400">Update Status</th>
                        <th className="px-4 py-2 text-left font-medium text-slate-600 dark:text-slate-400">Compliance</th>
                        <th className="px-4 py-2 text-left font-medium text-slate-600 dark:text-slate-400">Ownership</th>
                        <th className="px-4 py-2 text-left font-medium text-slate-600 dark:text-slate-400">Last Check-in</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-200 dark:divide-slate-700">
                      {reportData.deviceDetails.macDevices.map((device, index) => (
                        <tr key={index} className="hover:bg-slate-50 dark:hover:bg-slate-900/30">
                          <td className="px-4 py-2 text-slate-700 dark:text-slate-300 font-medium">{device.deviceName || '-'}</td>
                          <td className="px-4 py-2 text-slate-600 dark:text-slate-400 text-xs">{device.osVersion || '-'}</td>
                          <td className="px-4 py-2">
                            <VersionStatusBadge status={device.osVersionStatus} message={device.osVersionStatusMessage} />
                          </td>
                          <td className="px-4 py-2">
                            <span className={`inline-flex items-center px-2 py-0.5 rounded text-xs font-medium ${
                              device.complianceState === 'Compliant' 
                                ? 'bg-green-100 text-green-800 dark:bg-green-900/30 dark:text-green-400'
                                : device.complianceState === 'Non-Compliant'
                                ? 'bg-red-100 text-red-800 dark:bg-red-900/30 dark:text-red-400'
                                : 'bg-slate-100 text-slate-600 dark:bg-slate-700 dark:text-slate-400'
                            }`}>
                              {device.complianceState || 'Unknown'}
                            </span>
                          </td>
                          <td className="px-4 py-2 text-slate-600 dark:text-slate-400">{device.ownership || '-'}</td>
                          <td className="px-4 py-2 text-slate-600 dark:text-slate-400">
                            {device.lastCheckIn 
                              ? new Date(device.lastCheckIn).toLocaleDateString()
                              : <span className="text-slate-400">Never</span>}
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>
          )}

          {/* iOS/iPadOS Devices */}
          {reportData.deviceDetails?.iosDevices && reportData.deviceDetails.iosDevices.length > 0 && (
            <div className="mt-6">
              <div className="bg-white dark:bg-slate-800 rounded-lg border border-slate-200 dark:border-slate-700 overflow-hidden">
                <div className="px-4 py-3 bg-purple-50 dark:bg-purple-900/20 border-b border-slate-200 dark:border-slate-700">
                  <div className="flex items-center gap-2">
                    <LaptopRegular className="w-5 h-5 text-purple-600" />
                    <h3 className="font-semibold text-slate-700 dark:text-slate-200">iOS/iPadOS Devices</h3>
                    <span className="text-xs text-purple-600 dark:text-purple-400">({reportData.deviceDetails.iosDevices.length} devices)</span>
                  </div>
                </div>
                <div className="overflow-x-auto">
                  <table className="w-full text-sm">
                    <thead className="bg-slate-100 dark:bg-slate-900">
                      <tr>
                        <th className="px-4 py-2 text-left font-medium text-slate-600 dark:text-slate-400">Device Name</th>
                        <th className="px-4 py-2 text-left font-medium text-slate-600 dark:text-slate-400">OS Version</th>
                        <th className="px-4 py-2 text-left font-medium text-slate-600 dark:text-slate-400">Update Status</th>
                        <th className="px-4 py-2 text-left font-medium text-slate-600 dark:text-slate-400">Compliance</th>
                        <th className="px-4 py-2 text-left font-medium text-slate-600 dark:text-slate-400">Ownership</th>
                        <th className="px-4 py-2 text-left font-medium text-slate-600 dark:text-slate-400">Last Check-in</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-200 dark:divide-slate-700">
                      {reportData.deviceDetails.iosDevices.map((device, index) => (
                        <tr key={index} className="hover:bg-slate-50 dark:hover:bg-slate-900/30">
                          <td className="px-4 py-2 text-slate-700 dark:text-slate-300 font-medium">{device.deviceName || '-'}</td>
                          <td className="px-4 py-2 text-slate-600 dark:text-slate-400">{device.osVersion || '-'}</td>
                          <td className="px-4 py-2">
                            <VersionStatusBadge status={device.osVersionStatus} message={device.osVersionStatusMessage} />
                          </td>
                          <td className="px-4 py-2">
                            <span className={`inline-flex items-center px-2 py-0.5 rounded text-xs font-medium ${
                              device.complianceState === 'Compliant' 
                                ? 'bg-green-100 text-green-800 dark:bg-green-900/30 dark:text-green-400'
                                : device.complianceState === 'Non-Compliant'
                                ? 'bg-red-100 text-red-800 dark:bg-red-900/30 dark:text-red-400'
                                : 'bg-slate-100 text-slate-600 dark:bg-slate-700 dark:text-slate-400'
                            }`}>
                              {device.complianceState || 'Unknown'}
                            </span>
                          </td>
                          <td className="px-4 py-2 text-slate-600 dark:text-slate-400">{device.ownership || '-'}</td>
                          <td className="px-4 py-2 text-slate-600 dark:text-slate-400">
                            {device.lastCheckIn 
                              ? new Date(device.lastCheckIn).toLocaleDateString()
                              : <span className="text-slate-400">Never</span>}
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>
          )}

          {/* Android Devices */}
          {reportData.deviceDetails?.androidDevices && reportData.deviceDetails.androidDevices.length > 0 && (
            <div className="mt-6">
              <div className="bg-white dark:bg-slate-800 rounded-lg border border-slate-200 dark:border-slate-700 overflow-hidden">
                <div className="px-4 py-3 bg-green-50 dark:bg-green-900/20 border-b border-slate-200 dark:border-slate-700">
                  <div className="flex items-center gap-2">
                    <LaptopRegular className="w-5 h-5 text-green-600" />
                    <h3 className="font-semibold text-slate-700 dark:text-slate-200">Android Devices</h3>
                    <span className="text-xs text-green-600 dark:text-green-400">({reportData.deviceDetails.androidDevices.length} devices)</span>
                  </div>
                </div>
                <div className="overflow-x-auto">
                  <table className="w-full text-sm">
                    <thead className="bg-slate-100 dark:bg-slate-900">
                      <tr>
                        <th className="px-4 py-2 text-left font-medium text-slate-600 dark:text-slate-400">Device Name</th>
                        <th className="px-4 py-2 text-left font-medium text-slate-600 dark:text-slate-400">OS Version</th>
                        <th className="px-4 py-2 text-left font-medium text-slate-600 dark:text-slate-400">Update Status</th>
                        <th className="px-4 py-2 text-left font-medium text-slate-600 dark:text-slate-400">Compliance</th>
                        <th className="px-4 py-2 text-left font-medium text-slate-600 dark:text-slate-400">Last Check-in</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-200 dark:divide-slate-700">
                      {reportData.deviceDetails.androidDevices.map((device, index) => (
                        <tr key={index} className="hover:bg-slate-50 dark:hover:bg-slate-900/30">
                          <td className="px-4 py-2 text-slate-700 dark:text-slate-300 font-medium">{device.deviceName || '-'}</td>
                          <td className="px-4 py-2 text-slate-600 dark:text-slate-400">{device.osVersion || '-'}</td>
                          <td className="px-4 py-2">
                            <VersionStatusBadge status={device.osVersionStatus} message={device.osVersionStatusMessage} />
                          </td>
                          <td className="px-4 py-2">
                            <span className={`inline-flex items-center px-2 py-0.5 rounded text-xs font-medium ${
                              device.complianceState === 'Compliant' 
                                ? 'bg-green-100 text-green-800 dark:bg-green-900/30 dark:text-green-400'
                                : device.complianceState === 'Non-Compliant'
                                ? 'bg-red-100 text-red-800 dark:bg-red-900/30 dark:text-red-400'
                                : 'bg-slate-100 text-slate-600 dark:bg-slate-700 dark:text-slate-400'
                            }`}>
                              {device.complianceState || 'Unknown'}
                            </span>
                          </td>
                          <td className="px-4 py-2 text-slate-600 dark:text-slate-400">
                            {device.lastCheckIn 
                              ? new Date(device.lastCheckIn).toLocaleDateString()
                              : <span className="text-slate-400">Never</span>}
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>
          )}

          {/* User Sign-in Details */}
          {reportData.userSignInDetails && reportData.userSignInDetails.length > 0 && (
            <div className="mt-6">
              <div className="bg-white dark:bg-slate-800 rounded-lg border border-slate-200 dark:border-slate-700 overflow-hidden">
                <div className="px-4 py-3 bg-slate-50 dark:bg-slate-900/50 border-b border-slate-200 dark:border-slate-700">
                  <div className="flex items-center gap-2">
                    <PersonRegular className="w-5 h-5 text-purple-600" />
                    <h3 className="font-semibold text-slate-700 dark:text-slate-200">User Sign-in & MFA Details</h3>
                    <span className="text-xs text-slate-500">({reportData.userSignInDetails.length} users)</span>
                  </div>
                </div>
                <div className="overflow-x-auto">
                  <table className="w-full text-sm">
                    <thead className="bg-slate-100 dark:bg-slate-900">
                      <tr>
                        <th className="px-4 py-2 text-left font-medium text-slate-600 dark:text-slate-400">Display Name</th>
                        <th className="px-4 py-2 text-left font-medium text-slate-600 dark:text-slate-400">Email</th>
                        <th className="px-4 py-2 text-left font-medium text-slate-600 dark:text-slate-400">Last Interactive Sign-in</th>
                        <th className="px-4 py-2 text-left font-medium text-slate-600 dark:text-slate-400">Last Non-Interactive Sign-in</th>
                        <th className="px-4 py-2 text-left font-medium text-slate-600 dark:text-slate-400">Default MFA Method</th>
                        <th className="px-4 py-2 text-center font-medium text-slate-600 dark:text-slate-400">MFA</th>
                        <th className="px-4 py-2 text-center font-medium text-slate-600 dark:text-slate-400">Enabled</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-200 dark:divide-slate-700">
                      {reportData.userSignInDetails.map((user, index) => (
                        <tr key={index} className="hover:bg-slate-50 dark:hover:bg-slate-900/30">
                          <td className="px-4 py-2 text-slate-700 dark:text-slate-300">{user.displayName || '-'}</td>
                          <td className="px-4 py-2 text-slate-600 dark:text-slate-400 text-xs">{user.userPrincipalName || '-'}</td>
                          <td className="px-4 py-2 text-slate-600 dark:text-slate-400">
                            {user.lastInteractiveSignIn 
                              ? new Date(user.lastInteractiveSignIn).toLocaleDateString()
                              : <span className="text-slate-400">Never</span>}
                          </td>
                          <td className="px-4 py-2 text-slate-600 dark:text-slate-400">
                            {user.lastNonInteractiveSignIn 
                              ? new Date(user.lastNonInteractiveSignIn).toLocaleDateString()
                              : <span className="text-slate-400">Never</span>}
                          </td>
                          <td className="px-4 py-2 text-slate-600 dark:text-slate-400">
                            {user.defaultMfaMethod || <span className="text-slate-400">None</span>}
                          </td>
                          <td className="px-4 py-2 text-center">
                            {user.isMfaRegistered 
                              ? <span className="inline-flex items-center px-2 py-0.5 rounded text-xs font-medium bg-green-100 text-green-800 dark:bg-green-900/30 dark:text-green-400">Yes</span>
                              : <span className="inline-flex items-center px-2 py-0.5 rounded text-xs font-medium bg-red-100 text-red-800 dark:bg-red-900/30 dark:text-red-400">No</span>}
                          </td>
                          <td className="px-4 py-2 text-center">
                            {user.accountEnabled 
                              ? <span className="inline-flex items-center px-2 py-0.5 rounded text-xs font-medium bg-green-100 text-green-800 dark:bg-green-900/30 dark:text-green-400">Yes</span>
                              : <span className="inline-flex items-center px-2 py-0.5 rounded text-xs font-medium bg-slate-100 text-slate-600 dark:bg-slate-700 dark:text-slate-400">No</span>}
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>
          )}

          {/* Deleted Users in Period */}
          {reportData.deletedUsersInPeriod && reportData.deletedUsersInPeriod.length > 0 && (
            <div className="mt-6">
              <div className="bg-white dark:bg-slate-800 rounded-lg border border-slate-200 dark:border-slate-700 overflow-hidden">
                <div className="px-4 py-3 bg-red-50 dark:bg-red-900/20 border-b border-slate-200 dark:border-slate-700">
                  <div className="flex items-center gap-2">
                    <PersonRegular className="w-5 h-5 text-red-600" />
                    <h3 className="font-semibold text-slate-700 dark:text-slate-200">Deleted Users in Period</h3>
                    <span className="text-xs text-red-600 dark:text-red-400">({reportData.deletedUsersInPeriod.length} users)</span>
                  </div>
                </div>
                <div className="overflow-x-auto">
                  <table className="w-full text-sm">
                    <thead className="bg-slate-100 dark:bg-slate-900">
                      <tr>
                        <th className="px-4 py-2 text-left font-medium text-slate-600 dark:text-slate-400">Display Name</th>
                        <th className="px-4 py-2 text-left font-medium text-slate-600 dark:text-slate-400">Email</th>
                        <th className="px-4 py-2 text-left font-medium text-slate-600 dark:text-slate-400">Deleted Date</th>
                        <th className="px-4 py-2 text-left font-medium text-slate-600 dark:text-slate-400">Job Title</th>
                        <th className="px-4 py-2 text-left font-medium text-slate-600 dark:text-slate-400">Department</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-200 dark:divide-slate-700">
                      {reportData.deletedUsersInPeriod.map((user, index) => (
                        <tr key={index} className="hover:bg-slate-50 dark:hover:bg-slate-900/30">
                          <td className="px-4 py-2 text-slate-700 dark:text-slate-300">{user.displayName || '-'}</td>
                          <td className="px-4 py-2 text-slate-600 dark:text-slate-400 text-xs">{user.userPrincipalName || user.mail || '-'}</td>
                          <td className="px-4 py-2 text-red-600 dark:text-red-400">
                            {user.deletedDateTime 
                              ? new Date(user.deletedDateTime).toLocaleDateString()
                              : '-'}
                          </td>
                          <td className="px-4 py-2 text-slate-600 dark:text-slate-400">{user.jobTitle || '-'}</td>
                          <td className="px-4 py-2 text-slate-600 dark:text-slate-400">{user.department || '-'}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>
          )}

          {/* No Deleted Users Message */}
          {reportData.deletedUsersInPeriod && reportData.deletedUsersInPeriod.length === 0 && (
            <div className="mt-6">
              <div className="bg-white dark:bg-slate-800 rounded-lg border border-slate-200 dark:border-slate-700 p-4">
                <div className="flex items-center gap-2 text-slate-500 dark:text-slate-400">
                  <PersonRegular className="w-5 h-5" />
                  <span>No users were deleted during this period.</span>
                </div>
              </div>
            </div>
          )}

          {/* Mailbox Storage Details */}
          {reportData.mailboxDetails && reportData.mailboxDetails.length > 0 && (
            <div className="mt-6">
              <div className="bg-white dark:bg-slate-800 rounded-lg border border-slate-200 dark:border-slate-700 overflow-hidden">
                <div className="px-4 py-3 bg-blue-50 dark:bg-blue-900/20 border-b border-slate-200 dark:border-slate-700">
                  <div className="flex items-center gap-2">
                    <MailRegular className="w-5 h-5 text-blue-600" />
                    <h3 className="font-semibold text-slate-700 dark:text-slate-200">Mailbox Storage Details</h3>
                    <span className="text-xs text-blue-600 dark:text-blue-400">({reportData.mailboxDetails.length} mailboxes)</span>
                  </div>
                </div>
                <div className="overflow-x-auto">
                  <table className="w-full text-sm">
                    <thead className="bg-slate-100 dark:bg-slate-900">
                      <tr>
                        <th className="px-4 py-2 text-left font-medium text-slate-600 dark:text-slate-400">Display Name</th>
                        <th className="px-4 py-2 text-left font-medium text-slate-600 dark:text-slate-400">Email</th>
                        <th className="px-4 py-2 text-left font-medium text-slate-600 dark:text-slate-400">Type</th>
                        <th className="px-4 py-2 text-right font-medium text-slate-600 dark:text-slate-400">Size (GB)</th>
                        <th className="px-4 py-2 text-right font-medium text-slate-600 dark:text-slate-400">Quota (GB)</th>
                        <th className="px-4 py-2 text-right font-medium text-slate-600 dark:text-slate-400">% Used</th>
                        <th className="px-4 py-2 text-right font-medium text-slate-600 dark:text-slate-400">Items</th>
                        <th className="px-4 py-2 text-left font-medium text-slate-600 dark:text-slate-400">Last Activity</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-200 dark:divide-slate-700">
                      {reportData.mailboxDetails.map((mailbox, index) => (
                        <tr key={index} className="hover:bg-slate-50 dark:hover:bg-slate-900/30">
                          <td className="px-4 py-2 text-slate-700 dark:text-slate-300">{mailbox.displayName || '-'}</td>
                          <td className="px-4 py-2 text-slate-600 dark:text-slate-400 text-xs">{mailbox.userPrincipalName || '-'}</td>
                          <td className="px-4 py-2 text-slate-600 dark:text-slate-400">{mailbox.recipientType || 'User'}</td>
                          <td className="px-4 py-2 text-right font-medium text-slate-700 dark:text-slate-300">{mailbox.storageUsedGB?.toFixed(2) || '0.00'}</td>
                          <td className="px-4 py-2 text-right text-slate-600 dark:text-slate-400">{mailbox.quotaGB?.toFixed(0) || '-'}</td>
                          <td className="px-4 py-2 text-right">
                            {mailbox.percentUsed != null ? (
                              <span className={`font-medium ${
                                mailbox.percentUsed >= 90 ? 'text-red-600 dark:text-red-400' :
                                mailbox.percentUsed >= 80 ? 'text-amber-600 dark:text-amber-400' :
                                'text-green-600 dark:text-green-400'
                              }`}>
                                {mailbox.percentUsed.toFixed(1)}%
                              </span>
                            ) : '-'}
                          </td>
                          <td className="px-4 py-2 text-right text-slate-600 dark:text-slate-400">
                            {mailbox.itemCount?.toLocaleString() || '-'}
                          </td>
                          <td className="px-4 py-2 text-slate-600 dark:text-slate-400">
                            {mailbox.lastActivityDate 
                              ? new Date(mailbox.lastActivityDate).toLocaleDateString()
                              : <span className="text-slate-400">Never</span>}
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>
          )}

          {/* Footer */}
          <div className="text-center text-xs text-slate-500 dark:text-slate-400 py-4">
            <p>This report was automatically generated by M365 Dashboard.</p>
            <p>Some metrics may require additional licensing or API permissions.</p>
          </div>
        </>
      )}
    </div>
  );
};

export default ExecutiveReportPage;
