import React, { useState, useEffect } from 'react';
import { Link } from 'react-router-dom';
import {
  ShieldCheckmarkRegular,
  WarningRegular,
  PersonRegular,
  LockClosedRegular,
  LockOpenRegular,
  OpenRegular,
  CheckmarkCircleFilled,
  DismissCircleFilled,
  InfoRegular,
} from '@fluentui/react-icons';
import { useAppContext } from '../contexts/AppContext';

interface SecurityScore {
  currentScore: number;
  maxScore: number;
  percentageScore: number;
  controlScores: SecurityControlScore[];
  lastUpdated: string;
}

interface SecurityControlScore {
  controlName: string;
  controlCategory: string;
  description: string | null;
  score: number;
  maxScore: number;
}

interface RiskyUser {
  id: string;
  userPrincipalName: string;
  displayName: string | null;
  riskLevel: string;
  riskState: string;
  riskDetail: string | null;
  riskLastUpdatedDateTime: string | null;
  isDeleted: boolean;
  isProcessing: boolean;
}

interface RiskySignIn {
  id: string;
  userPrincipalName: string;
  displayName: string | null;
  createdDateTime: string | null;
  ipAddress: string | null;
  location: string | null;
  riskLevel: string;
  riskState: string;
  riskDetail: string | null;
  clientAppUsed: string | null;
  deviceDetail: string | null;
}

interface SecurityStats {
  totalRiskyUsers: number;
  highRiskUsers: number;
  mediumRiskUsers: number;
  lowRiskUsers: number;
  usersAtRisk: number;
  riskySignInsLast24Hours: number;
  activeAlerts: number;
  highSeverityAlerts: number;
  mediumSeverityAlerts: number;
  lowSeverityAlerts: number;
  mfaRegisteredUsers: number;
  mfaNotRegisteredUsers: number;
  mfaRegistrationPercentage: number;
  lastUpdated: string;
}

interface SecurityOverview {
  secureScore: SecurityScore | null;
  stats: SecurityStats;
  riskyUsers: RiskyUser[];
  riskySignIns: RiskySignIn[];
  recentAlerts: unknown[];
  lastUpdated: string;
}

const SecurityPage: React.FC = () => {
  const { getAccessToken } = useAppContext();
  const [overview, setOverview] = useState<SecurityOverview | null>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    fetchSecurityOverview();
  }, []);

  const fetchSecurityOverview = async () => {
    try {
      setLoading(true);
      const token = await getAccessToken();
      const response = await fetch('/api/security/overview', {
        headers: {
          'Authorization': `Bearer ${token}`,
          'Content-Type': 'application/json',
        },
      });

      if (!response.ok) {
        throw new Error('Failed to fetch security overview');
      }

      const data: SecurityOverview = await response.json();
      setOverview(data);
      setError(null);
    } catch (err) {
      setError(err instanceof Error ? err.message : 'An error occurred');
    } finally {
      setLoading(false);
    }
  };

  const formatDate = (dateString: string | null) => {
    if (!dateString) return 'N/A';
    const date = new Date(dateString);
    const now = new Date();
    const diffMs = now.getTime() - date.getTime();
    const diffMins = Math.floor(diffMs / 60000);
    const diffHours = Math.floor(diffMs / 3600000);
    const diffDays = Math.floor(diffMs / 86400000);

    if (diffMins < 60) return `${diffMins}m ago`;
    if (diffHours < 24) return `${diffHours}h ago`;
    if (diffDays < 7) return `${diffDays}d ago`;
    return date.toLocaleDateString();
  };

  const getRiskLevelColor = (level: string) => {
    switch (level.toLowerCase()) {
      case 'high':
        return 'text-red-600 bg-red-100 dark:bg-red-900/30 dark:text-red-400';
      case 'medium':
        return 'text-amber-600 bg-amber-100 dark:bg-amber-900/30 dark:text-amber-400';
      case 'low':
        return 'text-yellow-600 bg-yellow-100 dark:bg-yellow-900/30 dark:text-yellow-400';
      default:
        return 'text-slate-600 bg-slate-100 dark:bg-slate-700 dark:text-slate-400';
    }
  };

  const getScoreColor = (percentage: number) => {
    if (percentage >= 80) return 'text-green-600';
    if (percentage >= 60) return 'text-amber-600';
    return 'text-red-600';
  };

  const getScoreRingColor = (percentage: number) => {
    if (percentage >= 80) return 'stroke-green-500';
    if (percentage >= 60) return 'stroke-amber-500';
    return 'stroke-red-500';
  };

  // Circular progress component for Secure Score
  const SecureScoreRing: React.FC<{ percentage: number; score: number; maxScore: number }> = ({ 
    percentage, score, maxScore 
  }) => {
    const radius = 60;
    const strokeWidth = 10;
    const normalizedRadius = radius - strokeWidth / 2;
    const circumference = normalizedRadius * 2 * Math.PI;
    const strokeDashoffset = circumference - (percentage / 100) * circumference;

    return (
      <div className="relative inline-flex items-center justify-center">
        <svg height={radius * 2} width={radius * 2} className="transform -rotate-90">
          <circle
            stroke="currentColor"
            className="text-slate-200 dark:text-slate-700"
            fill="transparent"
            strokeWidth={strokeWidth}
            r={normalizedRadius}
            cx={radius}
            cy={radius}
          />
          <circle
            className={getScoreRingColor(percentage)}
            fill="transparent"
            strokeWidth={strokeWidth}
            strokeDasharray={circumference + ' ' + circumference}
            style={{ strokeDashoffset }}
            strokeLinecap="round"
            r={normalizedRadius}
            cx={radius}
            cy={radius}
          />
        </svg>
        <div className="absolute flex flex-col items-center justify-center">
          <span className={`text-2xl font-bold ${getScoreColor(percentage)}`}>
            {percentage.toFixed(2)}%
          </span>
          <span className="text-xs text-slate-500 dark:text-slate-400">
            {score.toFixed(2)}/{maxScore.toFixed(0)}
          </span>
        </div>
      </div>
    );
  };

  // Stat card component
  const StatCard: React.FC<{
    title: string;
    value: number | string;
    icon: React.ReactNode;
    color: string;
    subtitle?: string;
  }> = ({ title, value, icon, color, subtitle }) => (
    <div className="bg-white dark:bg-slate-800 rounded-lg border border-slate-200 dark:border-slate-700 p-3">
      <div className="flex items-center gap-2">
        <div className={`p-1.5 rounded ${color.replace('text-', 'bg-').replace('-600', '-100').replace('-500', '-100')} dark:bg-opacity-20 flex-shrink-0`}>
          {icon}
        </div>
        <div className="min-w-0 flex-1">
          <p className="text-xs text-slate-500 dark:text-slate-400 leading-tight">{title}</p>
          <p className={`text-lg font-semibold leading-tight ${color}`}>{value}</p>
          {subtitle && <p className="text-xs text-slate-400">{subtitle}</p>}
        </div>
      </div>
    </div>
  );

  if (loading) {
    return (
      <div className="p-4 flex items-center justify-center h-64">
        <div className="animate-spin rounded-full h-8 w-8 border-b-2 border-blue-600"></div>
      </div>
    );
  }

  if (error) {
    return (
      <div className="p-4">
        <div className="bg-red-50 dark:bg-red-900/20 border border-red-200 dark:border-red-800 rounded-lg p-4">
          <p className="text-red-600 dark:text-red-400">{error}</p>
          <button 
            onClick={fetchSecurityOverview}
            className="mt-2 text-sm text-red-600 dark:text-red-400 underline"
          >
            Retry
          </button>
        </div>
      </div>
    );
  }

  const stats = overview?.stats;
  const secureScore = overview?.secureScore;

  return (
    <div className="p-4 space-y-4 w-full max-w-full overflow-hidden">
      {/* Header */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
        <div>
          <h1 className="text-xl font-semibold text-slate-900 dark:text-white">Security</h1>
          <p className="text-sm text-slate-500 dark:text-slate-400 hidden sm:block">
            Identity protection, secure score, and security insights
          </p>
        </div>
        <a
          href="https://security.microsoft.com"
          target="_blank"
          rel="noopener noreferrer"
          className="inline-flex items-center justify-center gap-2 px-3 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 transition-colors text-sm whitespace-nowrap"
        >
          <OpenRegular className="w-4 h-4" />
          <span className="hidden sm:inline">Microsoft Security Center</span>
          <span className="sm:hidden">Security</span>
        </a>
      </div>

      {/* Top Row: Secure Score + MFA Stats */}
      <div className="grid grid-cols-1 lg:grid-cols-3 gap-4">
        {/* Secure Score Card */}
        <div className="bg-white dark:bg-slate-800 rounded-lg border border-slate-200 dark:border-slate-700 p-4">
          <div className="flex items-center justify-between mb-3">
            <h2 className="text-lg font-semibold text-slate-900 dark:text-white">Secure Score</h2>
            <a
              href="https://security.microsoft.com/securescore"
              target="_blank"
              rel="noopener noreferrer"
              className="text-blue-600 hover:text-blue-700 text-sm flex items-center gap-1"
            >
              Details <OpenRegular className="w-3 h-3" />
            </a>
          </div>
          {secureScore ? (
            <div className="flex items-center justify-center">
              <SecureScoreRing 
                percentage={secureScore.percentageScore} 
                score={secureScore.currentScore}
                maxScore={secureScore.maxScore}
              />
            </div>
          ) : (
            <div className="flex flex-col items-center justify-center py-4 text-slate-500 dark:text-slate-400">
              <InfoRegular className="w-8 h-8 mb-2" />
              <p className="text-sm text-center">Secure Score not available</p>
              <p className="text-xs text-center mt-1">Requires SecurityEvents.Read.All permission</p>
            </div>
          )}
        </div>

        {/* MFA Registration */}
        <div className="bg-white dark:bg-slate-800 rounded-lg border border-slate-200 dark:border-slate-700 p-4">
          <div className="flex items-center justify-between mb-3">
            <h2 className="text-lg font-semibold text-slate-900 dark:text-white">MFA Registration</h2>
            <Link
              to="/security/mfa"
              className="text-blue-600 hover:text-blue-700 text-sm flex items-center gap-1"
            >
              Details <OpenRegular className="w-3 h-3" />
            </Link>
          </div>
          <div className="space-y-3">
            <div className="flex items-center justify-between">
              <span className="text-sm text-slate-600 dark:text-slate-300">Registration Rate</span>
              <span className={`text-2xl font-bold ${(stats?.mfaRegistrationPercentage ?? 0) >= 90 ? 'text-green-600' : (stats?.mfaRegistrationPercentage ?? 0) >= 70 ? 'text-amber-600' : 'text-red-600'}`}>
                {stats?.mfaRegistrationPercentage ?? 0}%
              </span>
            </div>
            <div className="w-full bg-slate-200 dark:bg-slate-700 rounded-full h-2">
              <div 
                className={`h-2 rounded-full ${(stats?.mfaRegistrationPercentage ?? 0) >= 90 ? 'bg-green-500' : (stats?.mfaRegistrationPercentage ?? 0) >= 70 ? 'bg-amber-500' : 'bg-red-500'}`}
                style={{ width: `${stats?.mfaRegistrationPercentage ?? 0}%` }}
              />
            </div>
            <div className="flex justify-between text-xs text-slate-500 dark:text-slate-400">
              <span className="flex items-center gap-1">
                <CheckmarkCircleFilled className="w-3 h-3 text-green-500" />
                {stats?.mfaRegisteredUsers ?? 0} registered
              </span>
              <span className="flex items-center gap-1">
                <DismissCircleFilled className="w-3 h-3 text-red-500" />
                {stats?.mfaNotRegisteredUsers ?? 0} not registered
              </span>
            </div>
          </div>
        </div>

        {/* Risk Summary */}
        <div className="bg-white dark:bg-slate-800 rounded-lg border border-slate-200 dark:border-slate-700 p-4">
          <div className="flex items-center justify-between mb-3">
            <h2 className="text-lg font-semibold text-slate-900 dark:text-white">Risk Summary</h2>
            <a
              href="https://entra.microsoft.com/#view/Microsoft_AAD_IAM/IdentityProtectionMenuBlade/~/Overview"
              target="_blank"
              rel="noopener noreferrer"
              className="text-blue-600 hover:text-blue-700 text-sm flex items-center gap-1"
            >
              Details <OpenRegular className="w-3 h-3" />
            </a>
          </div>
          <div className="grid grid-cols-2 gap-2">
            <div className="bg-red-50 dark:bg-red-900/20 rounded-lg p-2 text-center">
              <p className="text-2xl font-bold text-red-600 dark:text-red-400">{stats?.highRiskUsers ?? 0}</p>
              <p className="text-xs text-red-600 dark:text-red-400">High Risk</p>
            </div>
            <div className="bg-amber-50 dark:bg-amber-900/20 rounded-lg p-2 text-center">
              <p className="text-2xl font-bold text-amber-600 dark:text-amber-400">{stats?.mediumRiskUsers ?? 0}</p>
              <p className="text-xs text-amber-600 dark:text-amber-400">Medium Risk</p>
            </div>
            <div className="bg-yellow-50 dark:bg-yellow-900/20 rounded-lg p-2 text-center">
              <p className="text-2xl font-bold text-yellow-600 dark:text-yellow-400">{stats?.lowRiskUsers ?? 0}</p>
              <p className="text-xs text-yellow-600 dark:text-yellow-400">Low Risk</p>
            </div>
            <div className="bg-slate-50 dark:bg-slate-700 rounded-lg p-2 text-center">
              <p className="text-2xl font-bold text-slate-600 dark:text-slate-300">{stats?.riskySignInsLast24Hours ?? 0}</p>
              <p className="text-xs text-slate-600 dark:text-slate-400">Risky Sign-ins (24h)</p>
            </div>
          </div>
        </div>
      </div>

      {/* Stats Cards Row */}
      <div className="grid grid-cols-2 sm:grid-cols-4 gap-2">
        <StatCard
          title="Risky Users"
          value={stats?.totalRiskyUsers ?? 0}
          icon={<WarningRegular className="w-4 h-4 text-red-600" />}
          color="text-red-600"
        />
        <StatCard
          title="Users at Risk"
          value={stats?.usersAtRisk ?? 0}
          icon={<PersonRegular className="w-4 h-4 text-amber-600" />}
          color="text-amber-600"
        />
        <StatCard
          title="MFA Enabled"
          value={stats?.mfaRegisteredUsers ?? 0}
          icon={<LockClosedRegular className="w-4 h-4 text-green-600" />}
          color="text-green-600"
        />
        <StatCard
          title="No MFA"
          value={stats?.mfaNotRegisteredUsers ?? 0}
          icon={<LockOpenRegular className="w-4 h-4 text-red-500" />}
          color="text-red-500"
        />
      </div>

      {/* Risky Users and Sign-ins Tables */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
        {/* Risky Users */}
        <div className="bg-white dark:bg-slate-800 rounded-lg border border-slate-200 dark:border-slate-700 overflow-hidden">
          <div className="px-4 py-3 border-b border-slate-200 dark:border-slate-700 flex items-center justify-between">
            <h2 className="font-semibold text-slate-900 dark:text-white">Risky Users</h2>
            <a
              href="https://entra.microsoft.com/#view/Microsoft_AAD_IAM/RiskyUsersV2Blade"
              target="_blank"
              rel="noopener noreferrer"
              className="text-blue-600 hover:text-blue-700 text-sm flex items-center gap-1"
            >
              View all <OpenRegular className="w-3 h-3" />
            </a>
          </div>
          {overview?.riskyUsers && overview.riskyUsers.length > 0 ? (
            <div className="divide-y divide-slate-200 dark:divide-slate-700 max-h-80 overflow-y-auto">
              {overview.riskyUsers.slice(0, 10).map((user) => (
                <div key={user.id} className="px-4 py-3 hover:bg-slate-50 dark:hover:bg-slate-700/50">
                  <div className="flex items-center justify-between">
                    <div className="min-w-0 flex-1">
                      <p className="font-medium text-slate-900 dark:text-white truncate">
                        {user.displayName || user.userPrincipalName}
                      </p>
                      <p className="text-xs text-slate-500 dark:text-slate-400 truncate">
                        {user.userPrincipalName}
                      </p>
                    </div>
                    <div className="flex items-center gap-2 ml-2">
                      <span className={`px-2 py-0.5 rounded text-xs font-medium ${getRiskLevelColor(user.riskLevel)}`}>
                        {user.riskLevel}
                      </span>
                      <span className="text-xs text-slate-400">
                        {formatDate(user.riskLastUpdatedDateTime)}
                      </span>
                    </div>
                  </div>
                  {user.riskDetail && user.riskDetail !== 'None' && (
                    <p className="text-xs text-slate-500 dark:text-slate-400 mt-1">
                      {user.riskDetail}
                    </p>
                  )}
                </div>
              ))}
            </div>
          ) : (
            <div className="px-4 py-8 text-center text-slate-500 dark:text-slate-400">
              <ShieldCheckmarkRegular className="w-12 h-12 mx-auto mb-2 text-green-500" />
              <p className="font-medium text-green-600 dark:text-green-400">No risky users detected</p>
              <p className="text-sm mt-1">All users are in good standing</p>
            </div>
          )}
        </div>

        {/* Risky Sign-ins */}
        <div className="bg-white dark:bg-slate-800 rounded-lg border border-slate-200 dark:border-slate-700 overflow-hidden">
          <div className="px-4 py-3 border-b border-slate-200 dark:border-slate-700 flex items-center justify-between">
            <h2 className="font-semibold text-slate-900 dark:text-white">Risky Sign-ins (24h)</h2>
            <a
              href="https://entra.microsoft.com/#view/Microsoft_AAD_IAM/RiskySignInsBlade"
              target="_blank"
              rel="noopener noreferrer"
              className="text-blue-600 hover:text-blue-700 text-sm flex items-center gap-1"
            >
              View all <OpenRegular className="w-3 h-3" />
            </a>
          </div>
          {overview?.riskySignIns && overview.riskySignIns.length > 0 ? (
            <div className="divide-y divide-slate-200 dark:divide-slate-700 max-h-80 overflow-y-auto">
              {overview.riskySignIns.slice(0, 10).map((signIn) => (
                <div key={signIn.id} className="px-4 py-3 hover:bg-slate-50 dark:hover:bg-slate-700/50">
                  <div className="flex items-center justify-between">
                    <div className="min-w-0 flex-1">
                      <p className="font-medium text-slate-900 dark:text-white truncate">
                        {signIn.displayName || signIn.userPrincipalName}
                      </p>
                      <div className="flex items-center gap-2 text-xs text-slate-500 dark:text-slate-400">
                        {signIn.ipAddress && <span>{signIn.ipAddress}</span>}
                        {signIn.location && <span>• {signIn.location}</span>}
                      </div>
                    </div>
                    <div className="flex items-center gap-2 ml-2">
                      <span className={`px-2 py-0.5 rounded text-xs font-medium ${getRiskLevelColor(signIn.riskLevel)}`}>
                        {signIn.riskLevel}
                      </span>
                      <span className="text-xs text-slate-400">
                        {formatDate(signIn.createdDateTime)}
                      </span>
                    </div>
                  </div>
                  {signIn.clientAppUsed && (
                    <p className="text-xs text-slate-500 dark:text-slate-400 mt-1">
                      {signIn.clientAppUsed} {signIn.deviceDetail && `• ${signIn.deviceDetail}`}
                    </p>
                  )}
                </div>
              ))}
            </div>
          ) : (
            <div className="px-4 py-8 text-center text-slate-500 dark:text-slate-400">
              <ShieldCheckmarkRegular className="w-12 h-12 mx-auto mb-2 text-green-500" />
              <p className="font-medium text-green-600 dark:text-green-400">No risky sign-ins detected</p>
              <p className="text-sm mt-1">No suspicious activity in the last 24 hours</p>
            </div>
          )}
        </div>
      </div>


    </div>
  );
};

export default SecurityPage;
