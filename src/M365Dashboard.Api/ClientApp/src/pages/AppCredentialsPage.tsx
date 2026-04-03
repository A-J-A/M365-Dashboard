import React, { useState, useEffect } from 'react';
import {
  KeyRegular,
  CertificateRegular,
  WarningRegular,
  ErrorCircleRegular,
  CheckmarkCircleRegular,
  ArrowSyncRegular,
  ArrowDownloadRegular,
  FilterRegular,
  SearchRegular,
  ChevronDownRegular,
  ChevronUpRegular,
  AppsRegular,
} from '@fluentui/react-icons';
import { useAppContext } from '../contexts/AppContext';

interface AppCredentialDetail {
  appId: string;
  appDisplayName: string;
  credentialType: string;
  keyId?: string;
  displayName?: string;
  startDateTime?: string;
  endDateTime: string;
  daysUntilExpiry: number;
  status: string;
}

interface AppCredentialStatus {
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
  lastUpdated: string;
}

const AppCredentialsPage: React.FC = () => {
  const { getAccessToken } = useAppContext();
  const [data, setData] = useState<AppCredentialStatus | null>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [searchTerm, setSearchTerm] = useState('');
  const [filterType, setFilterType] = useState<'all' | 'secrets' | 'certificates'>('all');
  const [filterStatus, setFilterStatus] = useState<'all' | 'expiring' | 'expired'>('all');
  const [expandedSections, setExpandedSections] = useState({
    expiringSecrets: true,
    expiredSecrets: true,
    expiringCertificates: true,
    expiredCertificates: true,
  });

  useEffect(() => {
    fetchData();
  }, []);

  const fetchData = async () => {
    try {
      setLoading(true);
      setError(null);
      const token = await getAccessToken();
      const response = await fetch('/api/security/app-credentials', {
        headers: {
          'Authorization': `Bearer ${token}`,
        },
      });

      if (!response.ok) {
        throw new Error(`Failed to fetch app credentials: ${response.statusText}`);
      }

      const result = await response.json();
      setData(result);
    } catch (err) {
      setError(err instanceof Error ? err.message : 'An error occurred');
    } finally {
      setLoading(false);
    }
  };

  const exportToCsv = () => {
    if (!data) return;

    const allCredentials = [
      ...data.expiringSecrets,
      ...data.expiredSecrets,
      ...data.expiringCertificates,
      ...data.expiredCertificates,
    ];

    const headers = ['App Name', 'App ID', 'Credential Type', 'Display Name', 'Expiry Date', 'Days Until Expiry', 'Status'];
    const rows = allCredentials.map(cred => [
      cred.appDisplayName,
      cred.appId,
      cred.credentialType,
      cred.displayName || '',
      new Date(cred.endDateTime).toLocaleDateString(),
      cred.daysUntilExpiry.toString(),
      cred.status,
    ]);

    const csv = [headers.join(','), ...rows.map(row => row.map(cell => `"${cell}"`).join(','))].join('\n');
    const blob = new Blob([csv], { type: 'text/csv' });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `app-credentials-${new Date().toISOString().split('T')[0]}.csv`;
    document.body.appendChild(a);
    a.click();
    window.URL.revokeObjectURL(url);
    a.remove();
  };

  const filterCredentials = (credentials: AppCredentialDetail[]) => {
    return credentials.filter(cred => {
      const matchesSearch = searchTerm === '' ||
        cred.appDisplayName.toLowerCase().includes(searchTerm.toLowerCase()) ||
        cred.appId.toLowerCase().includes(searchTerm.toLowerCase()) ||
        (cred.displayName?.toLowerCase().includes(searchTerm.toLowerCase()) ?? false);
      return matchesSearch;
    });
  };

  const toggleSection = (section: keyof typeof expandedSections) => {
    setExpandedSections(prev => ({
      ...prev,
      [section]: !prev[section],
    }));
  };

  const getStatusBadge = (status: string, daysUntilExpiry: number) => {
    if (status === 'Expired') {
      return (
        <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-xs font-medium bg-red-100 text-red-700 dark:bg-red-900/30 dark:text-red-400">
          <ErrorCircleRegular className="w-3 h-3" />
          Expired ({Math.abs(daysUntilExpiry)} days ago)
        </span>
      );
    }
    if (status === 'Critical' || daysUntilExpiry <= 7) {
      return (
        <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-xs font-medium bg-red-100 text-red-700 dark:bg-red-900/30 dark:text-red-400">
          <ErrorCircleRegular className="w-3 h-3" />
          Critical ({daysUntilExpiry} days)
        </span>
      );
    }
    if (status === 'Warning' || daysUntilExpiry <= 14) {
      return (
        <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-xs font-medium bg-amber-100 text-amber-700 dark:bg-amber-900/30 dark:text-amber-400">
          <WarningRegular className="w-3 h-3" />
          Warning ({daysUntilExpiry} days)
        </span>
      );
    }
    return (
      <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-xs font-medium bg-amber-100 text-amber-700 dark:bg-amber-900/30 dark:text-amber-400">
        <WarningRegular className="w-3 h-3" />
        Expiring ({daysUntilExpiry} days)
      </span>
    );
  };

  const CredentialTable: React.FC<{
    credentials: AppCredentialDetail[];
    title: string;
    icon: React.ReactNode;
    iconColor: string;
    sectionKey: keyof typeof expandedSections;
  }> = ({ credentials, title, icon, iconColor, sectionKey }) => {
    const filtered = filterCredentials(credentials);
    const isExpanded = expandedSections[sectionKey];

    if (credentials.length === 0) return null;

    return (
      <div className="bg-white dark:bg-slate-800 rounded-xl border border-slate-200 dark:border-slate-700 overflow-hidden">
        <button
          onClick={() => toggleSection(sectionKey)}
          className="w-full px-6 py-4 flex items-center justify-between hover:bg-slate-50 dark:hover:bg-slate-700/50 transition-colors"
        >
          <div className="flex items-center gap-3">
            <span className={`p-2 rounded-lg ${iconColor}`}>
              {icon}
            </span>
            <div className="text-left">
              <h3 className="font-semibold text-slate-900 dark:text-white">{title}</h3>
              <p className="text-sm text-slate-500 dark:text-slate-400">
                {filtered.length} of {credentials.length} credentials
              </p>
            </div>
          </div>
          {isExpanded ? (
            <ChevronUpRegular className="w-5 h-5 text-slate-400" />
          ) : (
            <ChevronDownRegular className="w-5 h-5 text-slate-400" />
          )}
        </button>

        {isExpanded && (
          <div className="border-t border-slate-200 dark:border-slate-700 overflow-x-auto">
            <table className="w-full">
              <thead>
                <tr className="bg-slate-50 dark:bg-slate-700/50">
                  <th className="px-4 py-3 text-left text-xs font-medium text-slate-500 dark:text-slate-400 uppercase tracking-wider">
                    Application
                  </th>
                  <th className="px-4 py-3 text-left text-xs font-medium text-slate-500 dark:text-slate-400 uppercase tracking-wider">
                    Credential Name
                  </th>
                  <th className="px-4 py-3 text-left text-xs font-medium text-slate-500 dark:text-slate-400 uppercase tracking-wider">
                    Type
                  </th>
                  <th className="px-4 py-3 text-left text-xs font-medium text-slate-500 dark:text-slate-400 uppercase tracking-wider">
                    Expiry Date
                  </th>
                  <th className="px-4 py-3 text-left text-xs font-medium text-slate-500 dark:text-slate-400 uppercase tracking-wider">
                    Status
                  </th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-200 dark:divide-slate-700">
                {filtered.map((cred, index) => (
                  <tr key={`${cred.appId}-${cred.keyId}-${index}`} className="hover:bg-slate-50 dark:hover:bg-slate-700/30">
                    <td className="px-4 py-3">
                      <div>
                        <div className="font-medium text-slate-900 dark:text-white">{cred.appDisplayName}</div>
                        <div className="text-xs text-slate-500 dark:text-slate-400">{cred.appId}</div>
                      </div>
                    </td>
                    <td className="px-4 py-3 text-sm text-slate-600 dark:text-slate-300">
                      {cred.displayName || '-'}
                    </td>
                    <td className="px-4 py-3">
                      <span className={`inline-flex items-center gap-1 px-2 py-0.5 rounded text-xs font-medium ${
                        cred.credentialType === 'Secret'
                          ? 'bg-purple-100 text-purple-700 dark:bg-purple-900/30 dark:text-purple-400'
                          : 'bg-blue-100 text-blue-700 dark:bg-blue-900/30 dark:text-blue-400'
                      }`}>
                        {cred.credentialType === 'Secret' ? (
                          <KeyRegular className="w-3 h-3" />
                        ) : (
                          <CertificateRegular className="w-3 h-3" />
                        )}
                        {cred.credentialType}
                      </span>
                    </td>
                    <td className="px-4 py-3 text-sm text-slate-600 dark:text-slate-300">
                      {new Date(cred.endDateTime).toLocaleDateString()}
                    </td>
                    <td className="px-4 py-3">
                      {getStatusBadge(cred.status, cred.daysUntilExpiry)}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>
    );
  };

  if (loading) {
    return (
      <div className="p-4 flex items-center justify-center h-64">
        <div className="flex flex-col items-center gap-3">
          <div className="animate-spin rounded-full h-8 w-8 border-b-2 border-blue-600"></div>
          <p className="text-slate-500 dark:text-slate-400">Loading app credentials...</p>
        </div>
      </div>
    );
  }

  if (error) {
    return (
      <div className="p-4">
        <div className="bg-red-50 dark:bg-red-900/20 border border-red-200 dark:border-red-800 rounded-xl p-6 text-center">
          <ErrorCircleRegular className="w-12 h-12 mx-auto text-red-500 mb-3" />
          <h3 className="text-lg font-semibold text-red-900 dark:text-red-200 mb-2">Error Loading Data</h3>
          <p className="text-red-700 dark:text-red-300 mb-4">{error}</p>
          <button
            onClick={fetchData}
            className="px-4 py-2 bg-red-600 text-white rounded-lg hover:bg-red-700 transition-colors"
          >
            Try Again
          </button>
        </div>
      </div>
    );
  }

  if (!data) return null;

  const totalIssues = data.expiringSecrets.length + data.expiredSecrets.length +
    data.expiringCertificates.length + data.expiredCertificates.length;

  return (
    <div className="p-4 space-y-6 max-w-7xl mx-auto">
      {/* Header */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4">
        <div>
          <h1 className="text-2xl font-bold text-slate-900 dark:text-white">App Registration Credentials</h1>
          <p className="text-sm text-slate-500 dark:text-slate-400">
            Monitor secrets and certificates expiring within {data.thresholdDays} days
          </p>
        </div>
        <div className="flex items-center gap-2">
          <button
            onClick={fetchData}
            className="flex items-center gap-2 px-4 py-2 text-slate-600 dark:text-slate-400 hover:bg-slate-100 dark:hover:bg-slate-700 rounded-lg transition-colors"
          >
            <ArrowSyncRegular className="w-4 h-4" />
            Refresh
          </button>
          <button
            onClick={exportToCsv}
            className="flex items-center gap-2 px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 transition-colors"
          >
            <ArrowDownloadRegular className="w-4 h-4" />
            Export CSV
          </button>
        </div>
      </div>

      {/* Summary Cards */}
      <div className="grid grid-cols-2 md:grid-cols-5 gap-4">
        <div className="bg-white dark:bg-slate-800 rounded-xl border border-slate-200 dark:border-slate-700 p-4">
          <div className="flex items-center gap-3">
            <div className="p-2 rounded-lg bg-blue-100 dark:bg-blue-900/30">
              <AppsRegular className="w-5 h-5 text-blue-600 dark:text-blue-400" />
            </div>
            <div>
              <p className="text-2xl font-bold text-slate-900 dark:text-white">{data.totalApps}</p>
              <p className="text-xs text-slate-500 dark:text-slate-400">Total Apps</p>
            </div>
          </div>
        </div>

        <div className="bg-white dark:bg-slate-800 rounded-xl border border-slate-200 dark:border-slate-700 p-4">
          <div className="flex items-center gap-3">
            <div className="p-2 rounded-lg bg-amber-100 dark:bg-amber-900/30">
              <KeyRegular className="w-5 h-5 text-amber-600 dark:text-amber-400" />
            </div>
            <div>
              <p className={`text-2xl font-bold ${data.appsWithExpiringSecrets > 0 ? 'text-amber-600 dark:text-amber-400' : 'text-slate-900 dark:text-white'}`}>
                {data.appsWithExpiringSecrets}
              </p>
              <p className="text-xs text-slate-500 dark:text-slate-400">Expiring Secrets</p>
            </div>
          </div>
        </div>

        <div className="bg-white dark:bg-slate-800 rounded-xl border border-slate-200 dark:border-slate-700 p-4">
          <div className="flex items-center gap-3">
            <div className="p-2 rounded-lg bg-red-100 dark:bg-red-900/30">
              <KeyRegular className="w-5 h-5 text-red-600 dark:text-red-400" />
            </div>
            <div>
              <p className={`text-2xl font-bold ${data.appsWithExpiredSecrets > 0 ? 'text-red-600 dark:text-red-400' : 'text-slate-900 dark:text-white'}`}>
                {data.appsWithExpiredSecrets}
              </p>
              <p className="text-xs text-slate-500 dark:text-slate-400">Expired Secrets</p>
            </div>
          </div>
        </div>

        <div className="bg-white dark:bg-slate-800 rounded-xl border border-slate-200 dark:border-slate-700 p-4">
          <div className="flex items-center gap-3">
            <div className="p-2 rounded-lg bg-amber-100 dark:bg-amber-900/30">
              <CertificateRegular className="w-5 h-5 text-amber-600 dark:text-amber-400" />
            </div>
            <div>
              <p className={`text-2xl font-bold ${data.appsWithExpiringCertificates > 0 ? 'text-amber-600 dark:text-amber-400' : 'text-slate-900 dark:text-white'}`}>
                {data.appsWithExpiringCertificates}
              </p>
              <p className="text-xs text-slate-500 dark:text-slate-400">Expiring Certs</p>
            </div>
          </div>
        </div>

        <div className="bg-white dark:bg-slate-800 rounded-xl border border-slate-200 dark:border-slate-700 p-4">
          <div className="flex items-center gap-3">
            <div className="p-2 rounded-lg bg-red-100 dark:bg-red-900/30">
              <CertificateRegular className="w-5 h-5 text-red-600 dark:text-red-400" />
            </div>
            <div>
              <p className={`text-2xl font-bold ${data.appsWithExpiredCertificates > 0 ? 'text-red-600 dark:text-red-400' : 'text-slate-900 dark:text-white'}`}>
                {data.appsWithExpiredCertificates}
              </p>
              <p className="text-xs text-slate-500 dark:text-slate-400">Expired Certs</p>
            </div>
          </div>
        </div>
      </div>

      {/* Search */}
      <div className="flex items-center gap-4">
        <div className="flex-1 relative">
          <SearchRegular className="absolute left-3 top-1/2 -translate-y-1/2 w-5 h-5 text-slate-400" />
          <input
            type="text"
            value={searchTerm}
            onChange={(e) => setSearchTerm(e.target.value)}
            placeholder="Search by app name or ID..."
            className="w-full pl-10 pr-4 py-2 border border-slate-300 dark:border-slate-600 rounded-lg bg-white dark:bg-slate-700 text-slate-900 dark:text-white placeholder-slate-400"
          />
        </div>
      </div>

      {/* Status Summary */}
      {totalIssues === 0 ? (
        <div className="bg-green-50 dark:bg-green-900/20 border border-green-200 dark:border-green-800 rounded-xl p-6 text-center">
          <CheckmarkCircleRegular className="w-12 h-12 mx-auto text-green-500 mb-3" />
          <h3 className="text-lg font-semibold text-green-900 dark:text-green-200 mb-2">All Clear!</h3>
          <p className="text-green-700 dark:text-green-300">
            No secrets or certificates are expiring within the next {data.thresholdDays} days.
          </p>
        </div>
      ) : (
        <div className="space-y-4">
          {/* Expiring Secrets */}
          <CredentialTable
            credentials={data.expiringSecrets}
            title="Expiring Secrets"
            icon={<KeyRegular className="w-5 h-5 text-amber-600 dark:text-amber-400" />}
            iconColor="bg-amber-100 dark:bg-amber-900/30"
            sectionKey="expiringSecrets"
          />

          {/* Expired Secrets */}
          <CredentialTable
            credentials={data.expiredSecrets}
            title="Expired Secrets"
            icon={<KeyRegular className="w-5 h-5 text-red-600 dark:text-red-400" />}
            iconColor="bg-red-100 dark:bg-red-900/30"
            sectionKey="expiredSecrets"
          />

          {/* Expiring Certificates */}
          <CredentialTable
            credentials={data.expiringCertificates}
            title="Expiring Certificates"
            icon={<CertificateRegular className="w-5 h-5 text-amber-600 dark:text-amber-400" />}
            iconColor="bg-amber-100 dark:bg-amber-900/30"
            sectionKey="expiringCertificates"
          />

          {/* Expired Certificates */}
          <CredentialTable
            credentials={data.expiredCertificates}
            title="Expired Certificates"
            icon={<CertificateRegular className="w-5 h-5 text-red-600 dark:text-red-400" />}
            iconColor="bg-red-100 dark:bg-red-900/30"
            sectionKey="expiredCertificates"
          />
        </div>
      )}

      {/* Last Updated */}
      <div className="text-center text-sm text-slate-500 dark:text-slate-400">
        Last updated: {new Date(data.lastUpdated).toLocaleString()}
      </div>
    </div>
  );
};

export default AppCredentialsPage;
