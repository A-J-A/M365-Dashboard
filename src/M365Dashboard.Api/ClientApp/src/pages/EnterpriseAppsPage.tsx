import React, { useState, useEffect, useCallback } from 'react';
import {
  AppsRegular,
  ArrowSyncRegular,
  ArrowDownloadRegular,
  SearchRegular,
  KeyRegular,
  CertificateRegular,
  CalendarRegular,
  SparkleRegular,
  WarningRegular,
  ErrorCircleRegular,
  CheckmarkCircleRegular,
  InfoRegular,
  ChevronUpRegular,
  ChevronDownRegular,
} from '@fluentui/react-icons';
import { useAppContext } from '../contexts/AppContext';

interface AppAuditItem {
  id: string;
  appId: string;
  displayName: string;
  description?: string;
  createdDateTime?: string;
  createdDaysAgo?: number;
  isNew: boolean;
  signInAudience?: string;
  publisherDomain?: string;
  hasCredentials: boolean;
  noCredentials: boolean;
  allCredentialsExpired: boolean;
  secretCount: number;
  certCount: number;
  nextExpiry?: string;
  daysUntilNextExpiry?: number;
  requiresResourceAccess: boolean;
  resourceAccessCount: number;
}

interface AuditData {
  summary: {
    totalApps: number;
    newLast30Days: number;
    noCredentials: number;
    allCredentialsExpired: number;
    appsWithPermissions: number;
  };
  recentApps: AppAuditItem[];
  noCredentialsApps: AppAuditItem[];
  expiredCredentialsApps: AppAuditItem[];
  allAppsByDate: AppAuditItem[];
  lastUpdated: string;
}

type TabKey = 'overview' | 'new' | 'noCredentials' | 'expiredCredentials' | 'allApps';

const EnterpriseAppsPage: React.FC = () => {
  const { getAccessToken } = useAppContext();
  const [data, setData] = useState<AuditData | null>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [activeTab, setActiveTab] = useState<TabKey>('overview');
  const [search, setSearch] = useState('');
  const [sortCol, setSortCol] = useState<'displayName' | 'createdDateTime' | 'daysUntilNextExpiry'>('createdDateTime');
  const [sortAsc, setSortAsc] = useState(false);

  const fetchData = useCallback(async () => {
    try {
      setLoading(true);
      setError(null);
      const token = await getAccessToken();
      const res = await fetch('/api/applicationconsent/enterprise-app-audit', {
        headers: { Authorization: `Bearer ${token}` },
      });
      if (!res.ok) throw new Error(`Failed to load: ${res.statusText}`);
      setData(await res.json());
    } catch (e) {
      setError(e instanceof Error ? e.message : 'An error occurred');
    } finally {
      setLoading(false);
    }
  }, [getAccessToken]);

  useEffect(() => { fetchData(); }, [fetchData]);

  const fmtDate = (d?: string) =>
    d ? new Date(d).toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' }) : '—';

  const filtered = (apps: AppAuditItem[]) => {
    const q = search.toLowerCase();
    return apps
      .filter(a => !q || a.displayName.toLowerCase().includes(q) || (a.appId?.toLowerCase().includes(q)))
      .sort((a, b) => {
        let av: any, bv: any;
        if (sortCol === 'displayName') { av = a.displayName; bv = b.displayName; }
        else if (sortCol === 'createdDateTime') { av = a.createdDateTime ?? ''; bv = b.createdDateTime ?? ''; }
        else { av = a.daysUntilNextExpiry ?? 9999; bv = b.daysUntilNextExpiry ?? 9999; }
        if (av < bv) return sortAsc ? -1 : 1;
        if (av > bv) return sortAsc ? 1 : -1;
        return 0;
      });
  };

  const exportCsv = (apps: AppAuditItem[], filename: string) => {
    const headers = ['Name', 'App ID', 'Created', 'Secrets', 'Certs', 'Next Expiry', 'Permissions', 'Audience'];
    const rows = apps.map(a => [
      a.displayName, a.appId ?? '', fmtDate(a.createdDateTime),
      a.secretCount, a.certCount, fmtDate(a.nextExpiry),
      a.resourceAccessCount, a.signInAudience ?? '',
    ]);
    const csv = [headers, ...rows].map(r => r.map(c => `"${c}"`).join(',')).join('\n');
    const url = URL.createObjectURL(new Blob([csv], { type: 'text/csv' }));
    const a = document.createElement('a');
    a.href = url; a.download = filename;
    document.body.appendChild(a); a.click();
    URL.revokeObjectURL(url); a.remove();
  };

  const SortTh: React.FC<{ col: typeof sortCol; label: string }> = ({ col, label }) => (
    <th
      onClick={() => { if (sortCol === col) setSortAsc(!sortAsc); else { setSortCol(col); setSortAsc(true); } }}
      className="px-4 py-3 text-left text-xs font-medium text-slate-500 dark:text-slate-400 uppercase tracking-wider cursor-pointer hover:text-slate-700 dark:hover:text-slate-200 select-none"
    >
      <span className="flex items-center gap-1">
        {label}
        {sortCol === col ? (sortAsc ? <ChevronUpRegular className="w-3 h-3" /> : <ChevronDownRegular className="w-3 h-3" />) : null}
      </span>
    </th>
  );

  const CredBadge: React.FC<{ app: AppAuditItem }> = ({ app }) => {
    if (app.noCredentials) return (
      <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-xs font-medium bg-slate-100 text-slate-600 dark:bg-slate-700 dark:text-slate-400">
        <InfoRegular className="w-3 h-3" /> None
      </span>
    );
    if (app.allCredentialsExpired) return (
      <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-xs font-medium bg-red-100 text-red-700 dark:bg-red-900/30 dark:text-red-400">
        <ErrorCircleRegular className="w-3 h-3" /> All Expired
      </span>
    );
    if (app.daysUntilNextExpiry !== undefined && app.daysUntilNextExpiry <= 30) return (
      <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-xs font-medium bg-amber-100 text-amber-700 dark:bg-amber-900/30 dark:text-amber-400">
        <WarningRegular className="w-3 h-3" /> Expiring {app.daysUntilNextExpiry}d
      </span>
    );
    return (
      <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-xs font-medium bg-green-100 text-green-700 dark:bg-green-900/30 dark:text-green-400">
        <CheckmarkCircleRegular className="w-3 h-3" /> Valid
      </span>
    );
  };

  const AppTable: React.FC<{ apps: AppAuditItem[]; exportName: string }> = ({ apps, exportName }) => {
    const rows = filtered(apps);
    return (
      <>
        <div className="flex items-center justify-between mb-3">
          <span className="text-sm text-slate-500 dark:text-slate-400">{rows.length} app{rows.length !== 1 ? 's' : ''}</span>
          <button
            onClick={() => exportCsv(rows, exportName)}
            className="flex items-center gap-2 px-3 py-1.5 text-sm text-slate-600 dark:text-slate-400 border border-slate-300 dark:border-slate-600 rounded-lg hover:bg-slate-50 dark:hover:bg-slate-700 transition-colors"
          >
            <ArrowDownloadRegular className="w-4 h-4" /> Export CSV
          </button>
        </div>
        {rows.length === 0 ? (
          <div className="text-center py-12 text-slate-500 dark:text-slate-400">
            <CheckmarkCircleRegular className="w-10 h-10 mx-auto mb-2 text-green-500" />
            <p>No apps found</p>
          </div>
        ) : (
          <div className="overflow-x-auto rounded-lg border border-slate-200 dark:border-slate-700">
            <table className="w-full text-sm">
              <thead className="bg-slate-50 dark:bg-slate-700/50">
                <tr>
                  <SortTh col="displayName" label="Application" />
                  <SortTh col="createdDateTime" label="Created" />
                  <th className="px-4 py-3 text-left text-xs font-medium text-slate-500 dark:text-slate-400 uppercase tracking-wider">Credentials</th>
                  <SortTh col="daysUntilNextExpiry" label="Next Expiry" />
                  <th className="px-4 py-3 text-left text-xs font-medium text-slate-500 dark:text-slate-400 uppercase tracking-wider">Permissions</th>
                  <th className="px-4 py-3 text-left text-xs font-medium text-slate-500 dark:text-slate-400 uppercase tracking-wider">Audience</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-200 dark:divide-slate-700">
                {rows.map(app => (
                  <tr key={app.id} className="hover:bg-slate-50 dark:hover:bg-slate-700/30">
                    <td className="px-4 py-3">
                      <div className="font-medium text-slate-900 dark:text-white flex items-center gap-2">
                        {app.displayName}
                        {app.isNew && (
                          <span className="inline-flex items-center gap-1 px-1.5 py-0.5 rounded text-xs font-medium bg-blue-100 text-blue-700 dark:bg-blue-900/30 dark:text-blue-400">
                            <SparkleRegular className="w-3 h-3" /> New
                          </span>
                        )}
                      </div>
                      <div className="text-xs text-slate-500 dark:text-slate-400 font-mono mt-0.5">{app.appId}</div>
                    </td>
                    <td className="px-4 py-3 text-slate-600 dark:text-slate-300">
                      <div>{fmtDate(app.createdDateTime)}</div>
                      {app.createdDaysAgo !== undefined && (
                        <div className="text-xs text-slate-400">{app.createdDaysAgo}d ago</div>
                      )}
                    </td>
                    <td className="px-4 py-3">
                      <CredBadge app={app} />
                      {app.hasCredentials && (
                        <div className="flex items-center gap-2 mt-1">
                          {app.secretCount > 0 && (
                            <span className="flex items-center gap-1 text-xs text-slate-500 dark:text-slate-400">
                              <KeyRegular className="w-3 h-3" /> {app.secretCount}
                            </span>
                          )}
                          {app.certCount > 0 && (
                            <span className="flex items-center gap-1 text-xs text-slate-500 dark:text-slate-400">
                              <CertificateRegular className="w-3 h-3" /> {app.certCount}
                            </span>
                          )}
                        </div>
                      )}
                    </td>
                    <td className="px-4 py-3 text-slate-600 dark:text-slate-300">
                      {app.nextExpiry ? (
                        <>
                          <div>{fmtDate(app.nextExpiry)}</div>
                          <div className={`text-xs ${(app.daysUntilNextExpiry ?? 999) <= 30 ? 'text-amber-600 dark:text-amber-400' : 'text-slate-400'}`}>
                            {app.daysUntilNextExpiry}d
                          </div>
                        </>
                      ) : '—'}
                    </td>
                    <td className="px-4 py-3">
                      {app.requiresResourceAccess ? (
                        <span className="text-xs text-slate-600 dark:text-slate-300">{app.resourceAccessCount} permission{app.resourceAccessCount !== 1 ? 's' : ''}</span>
                      ) : (
                        <span className="text-xs text-slate-400">None</span>
                      )}
                    </td>
                    <td className="px-4 py-3">
                      <span className="text-xs text-slate-600 dark:text-slate-300">
                        {app.signInAudience === 'AzureADMyOrg' ? 'Single tenant'
                          : app.signInAudience === 'AzureADMultipleOrgs' ? 'Multi-tenant'
                          : app.signInAudience ?? '—'}
                      </span>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </>
    );
  };

  const tabs: { key: TabKey; label: string; icon: React.ReactNode; count?: number; color?: string }[] = [
    { key: 'overview', label: 'Overview', icon: <AppsRegular className="w-4 h-4" /> },
    { key: 'new', label: 'New (30d)', icon: <SparkleRegular className="w-4 h-4" />, count: data?.summary.newLast30Days, color: 'text-blue-600' },
    { key: 'noCredentials', label: 'No Credentials', icon: <InfoRegular className="w-4 h-4" />, count: data?.summary.noCredentials, color: 'text-slate-500' },
    { key: 'expiredCredentials', label: 'Expired Credentials', icon: <ErrorCircleRegular className="w-4 h-4" />, count: data?.summary.allCredentialsExpired, color: 'text-red-600' },
    { key: 'allApps', label: 'All Apps', icon: <CalendarRegular className="w-4 h-4" />, count: data?.summary.totalApps },
  ];

  if (loading) return (
    <div className="p-4 flex items-center justify-center h-64">
      <div className="flex flex-col items-center gap-3">
        <div className="animate-spin rounded-full h-8 w-8 border-b-2 border-blue-600"></div>
        <p className="text-slate-500 dark:text-slate-400">Loading enterprise apps...</p>
      </div>
    </div>
  );

  if (error) return (
    <div className="p-4">
      <div className="bg-red-50 dark:bg-red-900/20 border border-red-200 dark:border-red-800 rounded-xl p-6 text-center">
        <ErrorCircleRegular className="w-12 h-12 mx-auto text-red-500 mb-3" />
        <p className="text-red-700 dark:text-red-300 mb-4">{error}</p>
        <button onClick={fetchData} className="px-4 py-2 bg-red-600 text-white rounded-lg hover:bg-red-700 transition-colors">Try Again</button>
      </div>
    </div>
  );

  if (!data) return null;

  return (
    <div className="p-4 space-y-4 max-w-7xl mx-auto">
      {/* Header */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
        <div>
          <h1 className="text-xl font-semibold text-slate-900 dark:text-white">Enterprise Applications</h1>
          <p className="text-sm text-slate-500 dark:text-slate-400">App registration audit — credentials, usage and creation timeline</p>
        </div>
        <div className="flex items-center gap-2">
          <button onClick={fetchData} className="px-3 py-2 bg-slate-100 dark:bg-slate-700 text-slate-700 dark:text-slate-300 rounded-lg hover:bg-slate-200 dark:hover:bg-slate-600 transition-colors">
            <ArrowSyncRegular className="w-4 h-4" />
          </button>
        </div>
      </div>

      {/* Summary cards */}
      <div className="grid grid-cols-2 md:grid-cols-5 gap-3">
        {[
          { label: 'Total Apps', value: data.summary.totalApps, color: 'text-blue-600', bg: 'bg-blue-100 dark:bg-blue-900/30' },
          { label: 'New (30d)', value: data.summary.newLast30Days, color: 'text-blue-600', bg: 'bg-blue-100 dark:bg-blue-900/30' },
          { label: 'No Credentials', value: data.summary.noCredentials, color: data.summary.noCredentials > 0 ? 'text-slate-600' : 'text-green-600', bg: 'bg-slate-100 dark:bg-slate-700' },
          { label: 'Expired Creds', value: data.summary.allCredentialsExpired, color: data.summary.allCredentialsExpired > 0 ? 'text-red-600' : 'text-green-600', bg: data.summary.allCredentialsExpired > 0 ? 'bg-red-100 dark:bg-red-900/30' : 'bg-green-100 dark:bg-green-900/30' },
          { label: 'With Permissions', value: data.summary.appsWithPermissions, color: 'text-purple-600', bg: 'bg-purple-100 dark:bg-purple-900/30' },
        ].map(c => (
          <div key={c.label} className="bg-white dark:bg-slate-800 rounded-lg border border-slate-200 dark:border-slate-700 p-4">
            <p className={`text-2xl font-bold ${c.color}`}>{c.value}</p>
            <p className="text-xs text-slate-500 dark:text-slate-400 mt-0.5">{c.label}</p>
          </div>
        ))}
      </div>

      {/* Tabs */}
      <div className="border-b border-slate-200 dark:border-slate-700">
        <nav className="flex gap-1 overflow-x-auto">
          {tabs.map(t => (
            <button
              key={t.key}
              onClick={() => setActiveTab(t.key)}
              className={`flex items-center gap-2 px-4 py-2.5 text-sm font-medium whitespace-nowrap border-b-2 transition-colors ${
                activeTab === t.key
                  ? 'border-blue-600 text-blue-600 dark:text-blue-400'
                  : 'border-transparent text-slate-600 dark:text-slate-400 hover:text-slate-900 dark:hover:text-slate-200'
              }`}
            >
              {t.icon}
              {t.label}
              {t.count !== undefined && (
                <span className={`px-1.5 py-0.5 rounded-full text-xs font-medium bg-slate-100 dark:bg-slate-700 ${t.color ?? 'text-slate-600 dark:text-slate-300'}`}>
                  {t.count}
                </span>
              )}
            </button>
          ))}
        </nav>
      </div>

      {/* Search bar (all tabs except overview) */}
      {activeTab !== 'overview' && (
        <div className="relative">
          <SearchRegular className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-400" />
          <input
            type="text"
            value={search}
            onChange={e => setSearch(e.target.value)}
            placeholder="Search by name or App ID..."
            className="w-full pl-9 pr-4 py-2 border border-slate-300 dark:border-slate-600 rounded-lg bg-white dark:bg-slate-800 text-slate-900 dark:text-white placeholder-slate-400 text-sm"
          />
        </div>
      )}

      {/* Tab content */}
      {activeTab === 'overview' && (
        <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
          {/* New apps */}
          <div className="bg-white dark:bg-slate-800 rounded-lg border border-slate-200 dark:border-slate-700 p-4">
            <div className="flex items-center gap-2 mb-3">
              <SparkleRegular className="w-5 h-5 text-blue-600" />
              <h3 className="font-semibold text-slate-900 dark:text-white">Recently Created</h3>
              <span className="ml-auto text-sm text-slate-500">{data.recentApps.length} apps in last 30 days</span>
            </div>
            {data.recentApps.length === 0 ? (
              <p className="text-sm text-slate-500 dark:text-slate-400 text-center py-4">No apps created in the last 30 days</p>
            ) : (
              <div className="space-y-2">
                {data.recentApps.slice(0, 5).map(app => (
                  <div key={app.id} className="flex items-center justify-between py-1.5 border-b border-slate-100 dark:border-slate-700 last:border-0">
                    <div>
                      <p className="text-sm font-medium text-slate-900 dark:text-white">{app.displayName}</p>
                      <p className="text-xs text-slate-500">{fmtDate(app.createdDateTime)}</p>
                    </div>
                    <CredBadge app={app} />
                  </div>
                ))}
                {data.recentApps.length > 5 && (
                  <button onClick={() => setActiveTab('new')} className="text-xs text-blue-600 dark:text-blue-400 hover:underline">
                    View all {data.recentApps.length} →
                  </button>
                )}
              </div>
            )}
          </div>

          {/* Expired credentials */}
          <div className="bg-white dark:bg-slate-800 rounded-lg border border-slate-200 dark:border-slate-700 p-4">
            <div className="flex items-center gap-2 mb-3">
              <ErrorCircleRegular className="w-5 h-5 text-red-600" />
              <h3 className="font-semibold text-slate-900 dark:text-white">Expired Credentials</h3>
              <span className="ml-auto text-sm text-slate-500">{data.expiredCredentialsApps.length} apps</span>
            </div>
            {data.expiredCredentialsApps.length === 0 ? (
              <div className="flex items-center gap-2 py-4 justify-center">
                <CheckmarkCircleRegular className="w-5 h-5 text-green-500" />
                <p className="text-sm text-slate-500 dark:text-slate-400">No apps with all-expired credentials</p>
              </div>
            ) : (
              <div className="space-y-2">
                {data.expiredCredentialsApps.slice(0, 5).map(app => (
                  <div key={app.id} className="flex items-center justify-between py-1.5 border-b border-slate-100 dark:border-slate-700 last:border-0">
                    <div>
                      <p className="text-sm font-medium text-slate-900 dark:text-white">{app.displayName}</p>
                      <p className="text-xs text-slate-500">{app.secretCount} secret{app.secretCount !== 1 ? 's' : ''}, {app.certCount} cert{app.certCount !== 1 ? 's' : ''}</p>
                    </div>
                    <span className="text-xs text-red-600 dark:text-red-400 font-medium">All expired</span>
                  </div>
                ))}
                {data.expiredCredentialsApps.length > 5 && (
                  <button onClick={() => setActiveTab('expiredCredentials')} className="text-xs text-blue-600 dark:text-blue-400 hover:underline">
                    View all {data.expiredCredentialsApps.length} →
                  </button>
                )}
              </div>
            )}
          </div>

          {/* No credentials */}
          <div className="bg-white dark:bg-slate-800 rounded-lg border border-slate-200 dark:border-slate-700 p-4">
            <div className="flex items-center gap-2 mb-3">
              <InfoRegular className="w-5 h-5 text-slate-500" />
              <h3 className="font-semibold text-slate-900 dark:text-white">No Credentials</h3>
              <span className="ml-auto text-sm text-slate-500">{data.noCredentialsApps.length} apps</span>
            </div>
            <p className="text-xs text-slate-500 dark:text-slate-400 mb-3">Apps with no secrets or certificates configured — may be unused or use managed identity/federated credentials.</p>
            {data.noCredentialsApps.length === 0 ? (
              <p className="text-sm text-slate-500 text-center py-2">None found</p>
            ) : (
              <div className="space-y-1">
                {data.noCredentialsApps.slice(0, 5).map(app => (
                  <div key={app.id} className="flex items-center justify-between py-1.5 border-b border-slate-100 dark:border-slate-700 last:border-0">
                    <p className="text-sm text-slate-900 dark:text-white">{app.displayName}</p>
                    <p className="text-xs text-slate-500">{fmtDate(app.createdDateTime)}</p>
                  </div>
                ))}
                {data.noCredentialsApps.length > 5 && (
                  <button onClick={() => setActiveTab('noCredentials')} className="text-xs text-blue-600 dark:text-blue-400 hover:underline">
                    View all {data.noCredentialsApps.length} →
                  </button>
                )}
              </div>
            )}
          </div>

          {/* Last updated */}
          <div className="bg-white dark:bg-slate-800 rounded-lg border border-slate-200 dark:border-slate-700 p-4 flex flex-col justify-between">
            <div className="flex items-center gap-2 mb-3">
              <CalendarRegular className="w-5 h-5 text-slate-500" />
              <h3 className="font-semibold text-slate-900 dark:text-white">App Timeline</h3>
            </div>
            <div className="space-y-1 flex-1">
              {data.allAppsByDate.slice(0, 6).map(app => (
                <div key={app.id} className="flex items-center justify-between text-sm py-1 border-b border-slate-100 dark:border-slate-700 last:border-0">
                  <span className="text-slate-900 dark:text-white truncate max-w-[180px]">{app.displayName}</span>
                  <span className="text-xs text-slate-500 ml-2 shrink-0">{fmtDate(app.createdDateTime)}</span>
                </div>
              ))}
            </div>
            <button onClick={() => setActiveTab('allApps')} className="text-xs text-blue-600 dark:text-blue-400 hover:underline mt-3 text-left">
              View all {data.summary.totalApps} apps →
            </button>
          </div>
        </div>
      )}

      {activeTab === 'new' && <AppTable apps={data.recentApps} exportName="new-apps-30d.csv" />}
      {activeTab === 'noCredentials' && <AppTable apps={data.noCredentialsApps} exportName="apps-no-credentials.csv" />}
      {activeTab === 'expiredCredentials' && <AppTable apps={data.expiredCredentialsApps} exportName="apps-expired-credentials.csv" />}
      {activeTab === 'allApps' && <AppTable apps={data.allAppsByDate} exportName="all-apps-by-date.csv" />}

      <p className="text-xs text-slate-400 text-right">Last updated: {new Date(data.lastUpdated).toLocaleString()}</p>
    </div>
  );
};

export default EnterpriseAppsPage;
