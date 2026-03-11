import React, { useState, useEffect, useCallback } from 'react';
import {
  ShieldCheckmarkRegular,
  ArrowDownloadRegular,
  ArrowSyncRegular,
  CheckmarkCircleFilled,
  DismissCircleFilled,
  WarningRegular,
  InfoRegular,
  PersonRegular,
  KeyRegular,
  MailRegular,
  ShareScreenStartRegular,
  PeopleRegular,
  PhoneRegular,
  ShieldRegular,
  ChevronDownRegular,
  ChevronRightRegular,
  OpenRegular,
} from '@fluentui/react-icons';
import { useAppContext } from '../contexts/AppContext';

// Types
interface SecurityCheck {
  name: string;
  description: string;
  checkedAt: string;
  status: 'Compliant' | 'NonCompliant' | 'Warning' | 'NotApplicable' | 'Error' | 'Unknown';
  currentValue: string;
  expectedValue: string;
  remediation: string;
  reference: string;
  affectedItems: string[];
  isBeta: boolean;
}

interface ComplianceSection {
  sectionName: string;
  sectionDescription: string;
  totalChecks: number;
  compliantChecks: number;
  nonCompliantChecks: number;
  compliancePercentage: number;
  checks: SecurityCheck[];
}

interface UserStatistics {
  totalUsers: number;
  memberUsers: number;
  guestUsers: number;
  licensedUsers: number;
  unlicensedUsers: number;
  blockedUsers: number;
  blockedUsersWithLicenses: number;
  adminUsers: number;
}

interface AdminRoleAssignment {
  roleName: string;
  memberCount: number;
  members: string[];
}

interface SecurityAssessmentResult {
  reportTitle: string;
  generatedAt: string;
  tenantId: string;
  tenantName: string;
  tenantDomain: string;
  userStats: UserStatistics;
  roleDistribution: AdminRoleAssignment[];
  entraIdCompliance: ComplianceSection;
  exchangeCompliance: ComplianceSection;
  sharePointCompliance: ComplianceSection;
  teamsCompliance: ComplianceSection;
  intuneCompliance: ComplianceSection;
  defenderCompliance: ComplianceSection;
  totalChecks: number;
  compliantChecks: number;
  nonCompliantChecks: number;
  overallCompliancePercentage: number;
}

// Status Badge Component
const StatusBadge: React.FC<{ status: SecurityCheck['status'] }> = ({ status }) => {
  const config: Record<string, { bg: string; text: string; icon: React.ReactNode }> = {
    Compliant: { 
      bg: 'bg-green-100 dark:bg-green-900/30', 
      text: 'text-green-700 dark:text-green-400', 
      icon: <CheckmarkCircleFilled className="w-4 h-4" /> 
    },
    NonCompliant: { 
      bg: 'bg-red-100 dark:bg-red-900/30', 
      text: 'text-red-700 dark:text-red-400', 
      icon: <DismissCircleFilled className="w-4 h-4" /> 
    },
    Warning: { 
      bg: 'bg-amber-100 dark:bg-amber-900/30', 
      text: 'text-amber-700 dark:text-amber-400', 
      icon: <WarningRegular className="w-4 h-4" /> 
    },
    NotApplicable: { 
      bg: 'bg-slate-100 dark:bg-slate-700', 
      text: 'text-slate-600 dark:text-slate-400', 
      icon: <InfoRegular className="w-4 h-4" /> 
    },
    Error: { 
      bg: 'bg-red-100 dark:bg-red-900/30', 
      text: 'text-red-700 dark:text-red-400', 
      icon: <DismissCircleFilled className="w-4 h-4" /> 
    },
    Unknown: { 
      bg: 'bg-slate-100 dark:bg-slate-700', 
      text: 'text-slate-600 dark:text-slate-400', 
      icon: <InfoRegular className="w-4 h-4" /> 
    },
  };

  const { bg, text, icon } = config[status] || config.Unknown;
  const label = status === 'NonCompliant' ? 'Non-Compliant' : status === 'NotApplicable' ? 'N/A' : status;

  return (
    <span className={`inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full text-xs font-medium ${bg} ${text}`}>
      {icon}
      {label}
    </span>
  );
};

// Stat Card Component
const StatCard: React.FC<{
  value: number;
  label: string;
  sublabel?: string;
  icon: React.ReactNode;
  variant?: 'default' | 'success' | 'danger';
}> = ({ value, label, sublabel, icon, variant = 'default' }) => {
  const borderColors = {
    default: 'border-l-blue-500',
    success: 'border-l-green-500',
    danger: 'border-l-red-500',
  };

  return (
    <div className={`bg-white dark:bg-slate-800 rounded-lg shadow-sm border-l-4 ${borderColors[variant]} p-5`}>
      <div className="flex items-center gap-3">
        <div className="p-2 bg-slate-100 dark:bg-slate-700 rounded-lg">
          {icon}
        </div>
        <div>
          <div className="text-3xl font-bold text-slate-900 dark:text-white">{value.toLocaleString()}</div>
          <div className="text-sm text-slate-600 dark:text-slate-400">{label}</div>
          {sublabel && <div className="text-xs text-slate-500 dark:text-slate-500">{sublabel}</div>}
        </div>
      </div>
    </div>
  );
};

// Compliance Section Card Component
const ComplianceSectionCard: React.FC<{
  section: ComplianceSection;
  icon: React.ReactNode;
  defaultExpanded?: boolean;
}> = ({ section, icon, defaultExpanded = false }) => {
  const [expanded, setExpanded] = useState(defaultExpanded);

  return (
    <div className="bg-white dark:bg-slate-800 rounded-lg shadow-sm overflow-hidden">
      <button
        onClick={() => setExpanded(!expanded)}
        className="w-full px-6 py-4 flex items-center justify-between hover:bg-slate-50 dark:hover:bg-slate-700/50 transition-colors"
      >
        <div className="flex items-center gap-4">
          <div className="p-2 bg-blue-100 dark:bg-blue-900/30 rounded-lg">
            {icon}
          </div>
          <div className="text-left">
            <h3 className="font-semibold text-slate-900 dark:text-white">{section.sectionName}</h3>
            <p className="text-sm text-slate-500 dark:text-slate-400">
              {section.compliantChecks}/{section.totalChecks} checks compliant
            </p>
          </div>
        </div>
        <div className="flex items-center gap-4">
          <div className="text-right">
            <div className={`text-2xl font-bold ${
              section.compliancePercentage >= 80 ? 'text-green-600 dark:text-green-400' :
              section.compliancePercentage >= 50 ? 'text-amber-600 dark:text-amber-400' :
              'text-red-600 dark:text-red-400'
            }`}>
              {section.compliancePercentage}%
            </div>
          </div>
          {expanded ? (
            <ChevronDownRegular className="w-5 h-5 text-slate-400" />
          ) : (
            <ChevronRightRegular className="w-5 h-5 text-slate-400" />
          )}
        </div>
      </button>

      {expanded && (
        <div className="border-t border-slate-200 dark:border-slate-700">
          <p className="px-6 py-3 text-sm text-slate-600 dark:text-slate-400 bg-slate-50 dark:bg-slate-800/50">
            {section.sectionDescription}
          </p>
          <div className="overflow-x-auto">
            <table className="w-full">
              <thead className="bg-slate-100 dark:bg-slate-700">
                <tr>
                  <th className="px-4 py-3 text-left text-xs font-semibold text-slate-600 dark:text-slate-300 uppercase tracking-wider">Name</th>
                  <th className="px-4 py-3 text-left text-xs font-semibold text-slate-600 dark:text-slate-300 uppercase tracking-wider">Description</th>
                  <th className="px-4 py-3 text-center text-xs font-semibold text-slate-600 dark:text-slate-300 uppercase tracking-wider">Status</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-200 dark:divide-slate-700">
                {section.checks.map((check, idx) => (
                  <tr key={idx} className="hover:bg-slate-50 dark:hover:bg-slate-700/30">
                    <td className="px-4 py-3">
                      <div className="font-medium text-sm text-slate-900 dark:text-white">
                        {check.name}
                        {check.isBeta && (
                          <span className="ml-2 text-xs bg-purple-100 dark:bg-purple-900/30 text-purple-700 dark:text-purple-400 px-1.5 py-0.5 rounded">
                            Beta
                          </span>
                        )}
                      </div>
                      {check.currentValue && (
                        <div className="text-xs text-slate-500 dark:text-slate-400 mt-1">
                          {check.currentValue}
                        </div>
                      )}
                    </td>
                    <td className="px-4 py-3 text-sm text-slate-600 dark:text-slate-400 max-w-md">
                      {check.description}
                    </td>
                    <td className="px-4 py-3 text-center">
                      <StatusBadge status={check.status} />
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}
    </div>
  );
};

// Main Page Component
const SecurityAssessmentPage: React.FC = () => {
  const { getAccessToken } = useAppContext();
  const [assessment, setAssessment] = useState<SecurityAssessmentResult | null>(null);
  const [loading, setLoading] = useState(false);
  const [downloading, setDownloading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const runAssessment = useCallback(async () => {
    try {
      setLoading(true);
      setError(null);
      const token = await getAccessToken();

      const response = await fetch('/api/securityassessment/run', {
        method: 'GET',
        headers: {
          'Authorization': `Bearer ${token}`,
          'Content-Type': 'application/json',
        },
      });

      if (!response.ok) {
        throw new Error('Failed to run security assessment');
      }

      const data: SecurityAssessmentResult = await response.json();
      setAssessment(data);
    } catch (err) {
      console.error('Error running assessment:', err);
      setError(err instanceof Error ? err.message : 'Failed to run assessment');
    } finally {
      setLoading(false);
    }
  }, [getAccessToken]);

  const downloadReport = useCallback(async () => {
    try {
      setDownloading(true);
      const token = await getAccessToken();

      const response = await fetch('/api/securityassessment/download', {
        method: 'GET',
        headers: {
          'Authorization': `Bearer ${token}`,
        },
      });

      if (!response.ok) {
        throw new Error('Failed to download report');
      }

      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `M365_Security_Assessment_${new Date().toISOString().split('T')[0]}.html`;
      document.body.appendChild(a);
      a.click();
      window.URL.revokeObjectURL(url);
      document.body.removeChild(a);
    } catch (err) {
      console.error('Error downloading report:', err);
      setError(err instanceof Error ? err.message : 'Failed to download report');
    } finally {
      setDownloading(false);
    }
  }, [getAccessToken]);

  // Run assessment on mount
  useEffect(() => {
    runAssessment();
  }, [runAssessment]);

  // Loading state
  if (loading) {
    return (
      <div className="p-6">
        <div className="flex items-center justify-center h-64">
          <div className="text-center">
            <ArrowSyncRegular className="w-8 h-8 animate-spin text-blue-500 mx-auto mb-4" />
            <p className="text-slate-600 dark:text-slate-400">Running security assessment...</p>
            <p className="text-sm text-slate-500 dark:text-slate-500 mt-1">This may take a minute</p>
          </div>
        </div>
      </div>
    );
  }

  // Error state
  if (error) {
    return (
      <div className="p-6">
        <div className="flex items-center justify-center h-64">
          <div className="text-center">
            <DismissCircleFilled className="w-8 h-8 text-red-500 mx-auto mb-4" />
            <p className="text-slate-600 dark:text-slate-400">Failed to load security assessment</p>
            <p className="text-sm text-red-500 mt-1">{error}</p>
            <button
              onClick={runAssessment}
              className="mt-4 px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 transition-colors"
            >
              Retry
            </button>
          </div>
        </div>
      </div>
    );
  }

  // No data state
  if (!assessment) {
    return (
      <div className="p-6">
        <div className="flex items-center justify-center h-64">
          <div className="text-center">
            <ShieldCheckmarkRegular className="w-8 h-8 text-slate-400 mx-auto mb-4" />
            <p className="text-slate-600 dark:text-slate-400">No assessment data available</p>
            <button
              onClick={runAssessment}
              className="mt-4 px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 transition-colors"
            >
              Run Assessment
            </button>
          </div>
        </div>
      </div>
    );
  }

  // Build sections array
  const sections = [
    { data: assessment.entraIdCompliance, icon: <KeyRegular className="w-5 h-5 text-blue-600 dark:text-blue-400" /> },
    { data: assessment.exchangeCompliance, icon: <MailRegular className="w-5 h-5 text-blue-600 dark:text-blue-400" /> },
    { data: assessment.sharePointCompliance, icon: <ShareScreenStartRegular className="w-5 h-5 text-blue-600 dark:text-blue-400" /> },
    { data: assessment.teamsCompliance, icon: <PeopleRegular className="w-5 h-5 text-blue-600 dark:text-blue-400" /> },
    { data: assessment.intuneCompliance, icon: <PhoneRegular className="w-5 h-5 text-blue-600 dark:text-blue-400" /> },
    { data: assessment.defenderCompliance, icon: <ShieldRegular className="w-5 h-5 text-blue-600 dark:text-blue-400" /> },
  ];

  return (
    <div className="p-6 space-y-6">
      {/* Header */}
      <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-4">
        <div>
          <h1 className="text-2xl font-bold text-slate-900 dark:text-white">Security Assessment</h1>
          <p className="text-slate-600 dark:text-slate-400">
            Comprehensive security check for {assessment.tenantName || assessment.tenantDomain}
          </p>
        </div>
        <div className="flex items-center gap-3">
          <button
            onClick={runAssessment}
            disabled={loading}
            className="inline-flex items-center gap-2 px-4 py-2 bg-slate-100 dark:bg-slate-700 text-slate-700 dark:text-slate-200 rounded-lg hover:bg-slate-200 dark:hover:bg-slate-600 transition-colors disabled:opacity-50"
          >
            <ArrowSyncRegular className={`w-4 h-4 ${loading ? 'animate-spin' : ''}`} />
            Refresh
          </button>
          <button
            onClick={downloadReport}
            disabled={downloading}
            className="inline-flex items-center gap-2 px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 transition-colors disabled:opacity-50"
          >
            <ArrowDownloadRegular className="w-4 h-4" />
            {downloading ? 'Downloading...' : 'Download Report'}
          </button>
        </div>
      </div>

      {/* Overall Score */}
      <div className="bg-gradient-to-r from-blue-600 to-blue-700 rounded-xl p-6 text-white">
        <div className="flex flex-col md:flex-row md:items-center md:justify-between gap-4">
          <div>
            <h2 className="text-lg font-medium opacity-90">Overall Compliance Score</h2>
            <p className="text-sm opacity-75 mt-1">
              {assessment.compliantChecks} of {assessment.totalChecks} checks passed
            </p>
          </div>
          <div className="flex items-center gap-6">
            <div className="text-center">
              <div className="text-5xl font-bold">{assessment.overallCompliancePercentage}%</div>
            </div>
            <div className="hidden sm:block w-32 h-32">
              <svg viewBox="0 0 36 36" className="transform -rotate-90">
                <path
                  d="M18 2.0845 a 15.9155 15.9155 0 0 1 0 31.831 a 15.9155 15.9155 0 0 1 0 -31.831"
                  fill="none"
                  stroke="rgba(255,255,255,0.2)"
                  strokeWidth="3"
                />
                <path
                  d="M18 2.0845 a 15.9155 15.9155 0 0 1 0 31.831 a 15.9155 15.9155 0 0 1 0 -31.831"
                  fill="none"
                  stroke="white"
                  strokeWidth="3"
                  strokeDasharray={`${assessment.overallCompliancePercentage}, 100`}
                  strokeLinecap="round"
                />
              </svg>
            </div>
          </div>
        </div>
      </div>

      {/* User Stats */}
      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4">
        <StatCard
          value={assessment.userStats.totalUsers}
          label="Total Users"
          sublabel={`Including ${assessment.userStats.guestUsers} guest users`}
          icon={<PersonRegular className="w-5 h-5 text-slate-600 dark:text-slate-400" />}
        />
        <StatCard
          value={assessment.userStats.licensedUsers}
          label="Licensed Users"
          sublabel={`${assessment.userStats.unlicensedUsers} unlicensed users`}
          icon={<KeyRegular className="w-5 h-5 text-slate-600 dark:text-slate-400" />}
          variant="success"
        />
        <StatCard
          value={assessment.userStats.blockedUsers}
          label="Blocked Users"
          sublabel={`${assessment.userStats.blockedUsersWithLicenses} blocked & licensed`}
          icon={<DismissCircleFilled className="w-5 h-5 text-slate-600 dark:text-slate-400" />}
          variant="danger"
        />
      </div>

      {/* Role Distribution */}
      {assessment.roleDistribution.length > 0 && (
        <div className="bg-white dark:bg-slate-800 rounded-lg shadow-sm p-6">
          <h3 className="font-semibold text-slate-900 dark:text-white mb-4">Admin Role Distribution</h3>
          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
            {assessment.roleDistribution.slice(0, 8).map((role, idx) => (
              <div key={idx} className="flex items-center justify-between p-3 bg-slate-50 dark:bg-slate-700/50 rounded-lg">
                <span className="text-sm text-slate-700 dark:text-slate-300 truncate">{role.roleName}</span>
                <span className="ml-2 px-2 py-0.5 text-sm font-medium bg-blue-100 dark:bg-blue-900/30 text-blue-700 dark:text-blue-400 rounded">
                  {role.memberCount}
                </span>
              </div>
            ))}
          </div>
        </div>
      )}

      {/* Compliance Sections */}
      <div className="space-y-4">
        <h2 className="text-lg font-semibold text-slate-900 dark:text-white">Compliance by Service Area</h2>
        {sections.map(({ data, icon }, idx) => (
          <ComplianceSectionCard
            key={idx}
            section={data}
            icon={icon}
            defaultExpanded={idx === 0}
          />
        ))}
      </div>

      {/* Footer */}
      <div className="text-center text-sm text-slate-500 dark:text-slate-400 py-4">
        <p>
          Report generated on {new Date(assessment.generatedAt).toLocaleString()} •{' '}
          <a 
            href="https://www.cisecurity.org/benchmark/microsoft_365" 
            target="_blank" 
            rel="noopener noreferrer" 
            className="text-blue-600 dark:text-blue-400 hover:underline inline-flex items-center gap-1"
          >
            CIS Benchmarks <OpenRegular className="w-3 h-3" />
          </a>
        </p>
      </div>
    </div>
  );
};

export default SecurityAssessmentPage;
