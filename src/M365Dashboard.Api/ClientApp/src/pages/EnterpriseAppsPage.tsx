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
  BuildingRegular,
  LockClosedRegular,
  PersonRegular,
} from '@fluentui/react-icons';
import { useAppContext } from '../contexts/AppContext';

interface AppRegistration {
  id: string; appId: string; displayName: string; description?: string;
  createdDateTime?: string; createdDaysAgo?: number; isNew: boolean;
  signInAudience?: string; publisherDomain?: string;
  hasCredentials: boolean; noCredentials: boolean; allCredentialsExpired: boolean;
  secretCount: number; certCount: number; nextExpiry?: string;
  daysUntilNextExpiry?: number; requiresResourceAccess: boolean; resourceAccessCount: number;
}
interface EnterpriseApp {
  id: string; appId: string; displayName: string; description?: string;
  createdDateTime?: string; createdDaysAgo?: number; isNew: boolean;
  accountEnabled?: boolean; signInAudience?: string;
  isOwnRegistration: boolean; isMicrosoftApp: boolean; isVerified: boolean;
  publisherName?: string; homepage?: string; tags?: string[];
}
interface AuditData {
  summary: {
    totalRegistrations: number; totalEnterpriseApps: number; thirdPartyApps: number;
    newRegistrations30d: number; newEnterpriseApps30d: number;
    noCredentials: number; allCredentialsExpired: number; disabledEnterpriseApps: number;
  };
  recentRegistrations: AppRegistration[]; noCredentialsApps: AppRegistration[];
  expiredCredentialsApps: AppRegistration[]; allRegistrationsByDate: AppRegistration[];
  recentEnterpriseApps: EnterpriseApp[]; allEnterpriseApps: EnterpriseApp[];
  thirdPartyEnterpriseApps: EnterpriseApp[]; disabledEnterpriseApps: EnterpriseApp[];
  lastUpdated: string;
}
type TabKey = 'overview'|'registrations'|'newRegistrations'|'noCredentials'|'expiredCredentials'|'enterpriseApps'|'thirdParty'|'newEnterprise';

const EnterpriseAppsPage: React.FC = () => {
  const { getAccessToken } = useAppContext();
  const [data, setData] = useState<AuditData | null>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [activeTab, setActiveTab] = useState<TabKey>('overview');
  const [search, setSearch] = useState('');
  const [sortCol, setSortCol] = useState('createdDateTime');
  const [sortAsc, setSortAsc] = useState(false);

  const fetchData = useCallback(async () => {
    try {
      setLoading(true); setError(null);
      const token = await getAccessToken();
      const res = await fetch('/api/applicationconsent/enterprise-app-audit', { headers: { Authorization: `Bearer ${token}` } });
      if (!res.ok) throw new Error(`Failed: ${res.statusText}`);
      setData(await res.json());
    } catch (e) { setError(e instanceof Error ? e.message : 'An error occurred'); }
    finally { setLoading(false); }
  }, [getAccessToken]);

  useEffect(() => { fetchData(); }, [fetchData]);

  const fmtDate = (d?: string) => d ? new Date(d).toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' }) : '—';

  const filterAndSort = <T extends { displayName: string; appId?: string; createdDateTime?: string }>(items: T[]): T[] => {
    const q = search.toLowerCase();
    return items
      .filter(a => !q || a.displayName.toLowerCase().includes(q) || (a.appId?.toLowerCase().includes(q) ?? false))
      .sort((a, b) => {
        const av: any = sortCol === 'displayName' ? a.displayName : (a.createdDateTime ?? '');
        const bv: any = sortCol === 'displayName' ? b.displayName : (b.createdDateTime ?? '');
        return sortAsc ? (av < bv ? -1 : av > bv ? 1 : 0) : (av > bv ? -1 : av < bv ? 1 : 0);
      });
  };

  const exportCsv = (rows: any[], fields: string[], filename: string) => {
    const csv = [fields, ...rows.map(r => fields.map(f => `"${r[f] ?? ''}"`))]
      .map(r => r.join(',')).join('\n');
    const url = URL.createObjectURL(new Blob([csv], { type: 'text/csv' }));
    const a = document.createElement('a'); a.href = url; a.download = filename;
    document.body.appendChild(a); a.click(); URL.revokeObjectURL(url); a.remove();
  };

  const SortTh: React.FC<{ col: string; label: string }> = ({ col, label }) => (
    <th onClick={() => { sortCol === col ? setSortAsc(!sortAsc) : (setSortCol(col), setSortAsc(true)); }}
      className="px-4 py-3 text-left text-xs font-medium text-slate-500 dark:text-slate-400 uppercase tracking-wider cursor-pointer hover:text-slate-700 dark:hover:text-slate-200 select-none">
      <span className="flex items-center gap-1">
        {label}
        {sortCol === col && (sortAsc ? <ChevronUpRegular className="w-3 h-3" /> : <ChevronDownRegular className="w-3 h-3" />)}
      </span>
    </th>
  );

  const CredBadge: React.FC<{ app: AppRegistration }> = ({ app }) => {
    if (app.noCredentials) return <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-xs font-medium bg-slate-100 text-slate-600 dark:bg-slate-700 dark:text-slate-400"><InfoRegular className="w-3 h-3" /> None</span>;
    if (app.allCredentialsExpired) return <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-xs font-medium bg-red-100 text-red-700 dark:bg-red-900/30 dark:text-red-400"><ErrorCircleRegular className="w-3 h-3" /> All Expired</span>;
    if ((app.daysUntilNextExpiry ?? 999) <= 30) return <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-xs font-medium bg-amber-100 text-amber-700 dark:bg-amber-900/30 dark:text-amber-400"><WarningRegular className="w-3 h-3" /> Expiring {app.daysUntilNextExpiry}d</span>;
    return <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-xs font-medium bg-green-100 text-green-700 dark:bg-green-900/30 dark:text-green-400"><CheckmarkCircleRegular className="w-3 h-3" /> Valid</span>;
  };

  const RegistrationTable: React.FC<{ apps: AppRegistration[]; exportName: string }> = ({ apps, exportName }) => {
    const rows = filterAndSort(apps);
    return <>
      <div className="flex items-center justify-between mb-3">
        <span className="text-sm text-slate-500 dark:text-slate-400">{rows.length} app registration{rows.length !== 1 ? 's' : ''}</span>
        <button onClick={() => exportCsv(rows, ['displayName','appId','createdDateTime','secretCount','certCount','nextExpiry','resourceAccessCount','signInAudience'], exportName)}
          className="flex items-center gap-2 px-3 py-1.5 text-sm text-slate-600 dark:text-slate-400 border border-slate-300 dark:border-slate-600 rounded-lg hover:bg-slate-50 dark:hover:bg-slate-700 transition-colors">
          <ArrowDownloadRegular className="w-4 h-4" /> Export CSV
        </button>
      </div>
      {rows.length === 0 ? (
        <div className="text-center py-12"><CheckmarkCircleRegular className="w-10 h-10 mx-auto mb-2 text-green-500" /><p className="text-slate-500">No app registrations found</p></div>
      ) : (
        <div className="overflow-x-auto rounded-lg border border-slate-200 dark:border-slate-700">
          <table className="w-full text-sm">
            <thead className="bg-slate-50 dark:bg-slate-700/50">
              <tr><SortTh col="displayName" label="Application" /><SortTh col="createdDateTime" label="Created" />
                <th className="px-4 py-3 text-left text-xs font-medium text-slate-500 dark:text-slate-400 uppercase tracking-wider">Credentials</th>
                <th className="px-4 py-3 text-left text-xs font-medium text-slate-500 dark:text-slate-400 uppercase tracking-wider">Next Expiry</th>
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
                      {app.isNew && <span className="inline-flex items-center gap-1 px-1.5 py-0.5 rounded text-xs font-medium bg-blue-100 text-blue-700 dark:bg-blue-900/30 dark:text-blue-400"><SparkleRegular className="w-3 h-3" /> New</span>}
                    </div>
                    <div className="text-xs text-slate-500 font-mono mt-0.5">{app.appId}</div>
                  </td>
                  <td className="px-4 py-3 text-slate-600 dark:text-slate-300">
                    <div>{fmtDate(app.createdDateTime)}</div>
                    {app.createdDaysAgo !== undefined && <div className="text-xs text-slate-400">{app.createdDaysAgo}d ago</div>}
                  </td>
                  <td className="px-4 py-3">
                    <CredBadge app={app} />
                    {app.hasCredentials && (
                      <div className="flex gap-2 mt-1">
                        {app.secretCount > 0 && <span className="flex items-center gap-1 text-xs text-slate-500"><KeyRegular className="w-3 h-3" />{app.secretCount}</span>}
                        {app.certCount > 0 && <span className="flex items-center gap-1 text-xs text-slate-500"><CertificateRegular className="w-3 h-3" />{app.certCount}</span>}
                      </div>
                    )}
                  </td>
                  <td className="px-4 py-3 text-slate-600 dark:text-slate-300">
                    {app.nextExpiry ? <><div>{fmtDate(app.nextExpiry)}</div><div className={`text-xs ${(app.daysUntilNextExpiry ?? 999) <= 30 ? 'text-amber-600' : 'text-slate-400'}`}>{app.daysUntilNextExpiry}d</div></> : '—'}
                  </td>
                  <td className="px-4 py-3">
                    {app.requiresResourceAccess ? <span className="text-xs text-slate-600 dark:text-slate-300">{app.resourceAccessCount} permission{app.resourceAccessCount !== 1 ? 's' : ''}</span> : <span className="text-xs text-slate-400">None</span>}
                  </td>
                  <td className="px-4 py-3 text-xs text-slate-600 dark:text-slate-300">
                    {app.signInAudience === 'AzureADMyOrg' ? 'Single tenant' : app.signInAudience === 'AzureADMultipleOrgs' ? 'Multi-tenant' : app.signInAudience ?? '—'}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </>;
  };

  const EnterpriseTable: React.FC<{ apps: EnterpriseApp[]; exportName: string }> = ({ apps, exportName }) => {
    const rows = filterAndSort(apps);
    return <>
      <div className="flex items-center justify-between mb-3">
        <span className="text-sm text-slate-500 dark:text-slate-400">{rows.length} enterprise app{rows.length !== 1 ? 's' : ''}</span>
        <button onClick={() => exportCsv(rows, ['displayName','appId','createdDateTime','publisherName','accountEnabled','isMicrosoftApp','isOwnRegistration'], exportName)}
          className="flex items-center gap-2 px-3 py-1.5 text-sm text-slate-600 dark:text-slate-400 border border-slate-300 dark:border-slate-600 rounded-lg hover:bg-slate-50 dark:hover:bg-slate-700 transition-colors">
          <ArrowDownloadRegular className="w-4 h-4" /> Export CSV
        </button>
      </div>
      {rows.length === 0 ? (
        <div className="text-center py-12"><CheckmarkCircleRegular className="w-10 h-10 mx-auto mb-2 text-green-500" /><p className="text-slate-500">No enterprise apps found</p></div>
      ) : (
        <div className="overflow-x-auto rounded-lg border border-slate-200 dark:border-slate-700">
          <table className="w-full text-sm">
            <thead className="bg-slate-50 dark:bg-slate-700/50">
              <tr><SortTh col="displayName" label="Application" /><SortTh col="createdDateTime" label="Added" />
                <th className="px-4 py-3 text-left text-xs font-medium text-slate-500 dark:text-slate-400 uppercase tracking-wider">Publisher</th>
                <th className="px-4 py-3 text-left text-xs font-medium text-slate-500 dark:text-slate-400 uppercase tracking-wider">Type</th>
                <th className="px-4 py-3 text-left text-xs font-medium text-slate-500 dark:text-slate-400 uppercase tracking-wider">Status</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-200 dark:divide-slate-700">
              {rows.map(app => (
                <tr key={app.id} className="hover:bg-slate-50 dark:hover:bg-slate-700/30">
                  <td className="px-4 py-3">
                    <div className="font-medium text-slate-900 dark:text-white flex items-center gap-2">
                      {app.displayName}
                      {app.isNew && <span className="inline-flex items-center gap-1 px-1.5 py-0.5 rounded text-xs font-medium bg-blue-100 text-blue-700 dark:bg-blue-900/30 dark:text-blue-400"><SparkleRegular className="w-3 h-3" /> New</span>}
                    </div>
                    <div className="text-xs text-slate-500 font-mono mt-0.5">{app.appId}</div>
                  </td>
                  <td className="px-4 py-3 text-slate-600 dark:text-slate-300">
                    <div>{fmtDate(app.createdDateTime)}</div>
                    {app.createdDaysAgo !== undefined && <div className="text-xs text-slate-400">{app.createdDaysAgo}d ago</div>}
                  </td>
                  <td className="px-4 py-3">
                    <div className="text-sm text-slate-600 dark:text-slate-300">{app.publisherName ?? '—'}</div>
                    {app.isVerified && <span className="text-xs text-blue-600 dark:text-blue-400">✓ Verified</span>}
                  </td>
                  <td className="px-4 py-3">
                    {app.isMicrosoftApp
                      ? <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-xs font-medium bg-blue-100 text-blue-700 dark:bg-blue-900/30 dark:text-blue-400"><BuildingRegular className="w-3 h-3" /> Microsoft</span>
                      : app.isOwnRegistration
                        ? <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-xs font-medium bg-purple-100 text-purple-700 dark:bg-purple-900/30 dark:text-purple-400"><PersonRegular className="w-3 h-3" /> Own App</span>
                        : <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-xs font-medium bg-amber-100 text-amber-700 dark:bg-amber-900/30 dark:text-amber-400"><LockClosedRegular className="w-3 h-3" /> Third-party</span>}
                  </td>
                  <td className="px-4 py-3">
                    {app.accountEnabled === false
                      ? <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-xs font-medium bg-slate-100 text-slate-600 dark:bg-slate-700 dark:text-slate-400">Disabled</span>
                      : <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-xs font-medium bg-green-100 text-green-700 dark:bg-green-900/30 dark:text-green-400">Enabled</span>}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </>;
  };

  const tabs: { key: TabKey; label: string; icon: React.ReactNode; count?: number }[] = [
    { key: 'overview',           label: 'Overview',            icon: <AppsRegular className="w-4 h-4" /> },
    { key: 'registrations',      label: 'All Registrations',   icon: <LockClosedRegular className="w-4 h-4" />,  count: data?.summary.totalRegistrations },
    { key: 'newRegistrations',   label: 'New (30d)',            icon: <SparkleRegular className="w-4 h-4" />,    count: data?.summary.newRegistrations30d },
    { key: 'noCredentials',      label: 'No Credentials',      icon: <InfoRegular className="w-4 h-4" />,       count: data?.summary.noCredentials },
    { key: 'expiredCredentials', label: 'Expired Creds',       icon: <ErrorCircleRegular className="w-4 h-4" />,count: data?.summary.allCredentialsExpired },
    { key: 'enterpriseApps',     label: 'All Enterprise Apps', icon: <BuildingRegular className="w-4 h-4" />,   count: data?.summary.totalEnterpriseApps },
    { key: 'thirdParty',         label: 'Third-party',         icon: <LockClosedRegular className="w-4 h-4" />, count: data?.summary.thirdPartyApps },
    { key: 'newEnterprise',      label: 'New (30d)',            icon: <SparkleRegular className="w-4 h-4" />,   count: data?.summary.newEnterpriseApps30d },
  ];

  if (loading) return <div className="p-4 flex items-center justify-center h-64"><div className="flex flex-col items-center gap-3"><div className="animate-spin rounded-full h-8 w-8 border-b-2 border-blue-600"></div><p className="text-slate-500 dark:text-slate-400">Loading enterprise apps...</p></div></div>;
  if (error) return <div className="p-4"><div className="bg-red-50 dark:bg-red-900/20 border border-red-200 dark:border-red-800 rounded-xl p-6 text-center"><ErrorCircleRegular className="w-12 h-12 mx-auto text-red-500 mb-3" /><p className="text-red-700 dark:text-red-300 mb-4">{error}</p><button onClick={fetchData} className="px-4 py-2 bg-red-600 text-white rounded-lg hover:bg-red-700 transition-colors">Try Again</button></div></div>;
  if (!data) return null;

  const isRegTab = ['registrations','newRegistrations','noCredentials','expiredCredentials'].includes(activeTab);
  const isEntTab = ['enterpriseApps','thirdParty','newEnterprise'].includes(activeTab);

  return (
    <div className="p-4 space-y-4 max-w-7xl mx-auto">
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
        <div>
          <h1 className="text-xl font-semibold text-slate-900 dark:text-white">Enterprise Applications</h1>
          <p className="text-sm text-slate-500 dark:text-slate-400">App registrations and enterprise apps — credentials, usage and creation timeline</p>
        </div>
        <button onClick={fetchData} className="px-3 py-2 bg-slate-100 dark:bg-slate-700 text-slate-700 dark:text-slate-300 rounded-lg hover:bg-slate-200 dark:hover:bg-slate-600 transition-colors self-start"><ArrowSyncRegular className="w-4 h-4" /></button>
      </div>

      {/* Summary */}
      <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
        {[
          { label: 'App Registrations',        value: data.summary.totalRegistrations,    color: 'text-blue-600' },
          { label: 'Enterprise Apps',           value: data.summary.totalEnterpriseApps,   color: 'text-purple-600' },
          { label: 'Third-party Apps',          value: data.summary.thirdPartyApps,        color: 'text-amber-600' },
          { label: 'Expired Credentials',       value: data.summary.allCredentialsExpired, color: data.summary.allCredentialsExpired > 0 ? 'text-red-600' : 'text-green-600' },
          { label: 'No Credentials',            value: data.summary.noCredentials,         color: 'text-slate-600' },
          { label: 'New Registrations (30d)',   value: data.summary.newRegistrations30d,   color: 'text-blue-600' },
          { label: 'New Enterprise Apps (30d)', value: data.summary.newEnterpriseApps30d,  color: 'text-purple-600' },
          { label: 'Disabled Enterprise Apps',  value: data.summary.disabledEnterpriseApps,color: data.summary.disabledEnterpriseApps > 0 ? 'text-slate-600' : 'text-green-600' },
        ].map(c => (
          <div key={c.label} className="bg-white dark:bg-slate-800 rounded-lg border border-slate-200 dark:border-slate-700 p-4">
            <p className={`text-2xl font-bold ${c.color}`}>{c.value}</p>
            <p className="text-xs text-slate-500 dark:text-slate-400 mt-0.5">{c.label}</p>
          </div>
        ))}
      </div>

      {/* Tabs with section labels */}
      <div className="border-b border-slate-200 dark:border-slate-700">
        <div className="flex items-center gap-4 px-1 pb-1">
          <span className="text-xs font-semibold text-blue-600 uppercase tracking-wider">App Registrations</span>
          <div className="flex-1 border-t border-dashed border-slate-200 dark:border-slate-700 mx-2" />
          <span className="text-xs font-semibold text-purple-600 uppercase tracking-wider">Enterprise Apps</span>
        </div>
        <nav className="flex gap-1 overflow-x-auto">
          {tabs.map(t => (
            <button key={t.key} onClick={() => setActiveTab(t.key)}
              className={`flex items-center gap-2 px-3 py-2.5 text-sm font-medium whitespace-nowrap border-b-2 transition-colors ${activeTab === t.key ? 'border-blue-600 text-blue-600 dark:text-blue-400' : 'border-transparent text-slate-600 dark:text-slate-400 hover:text-slate-900 dark:hover:text-slate-200'}`}>
              {t.icon} {t.label}
              {t.count !== undefined && <span className="px-1.5 py-0.5 rounded-full text-xs font-medium bg-slate-100 dark:bg-slate-700 text-slate-600 dark:text-slate-300">{t.count}</span>}
            </button>
          ))}
        </nav>
      </div>

      {/* Search */}
      {activeTab !== 'overview' && (
        <div className="relative">
          <SearchRegular className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-400" />
          <input type="text" value={search} onChange={e => setSearch(e.target.value)} placeholder="Search by name or App ID..."
            className="w-full pl-9 pr-4 py-2 border border-slate-300 dark:border-slate-600 rounded-lg bg-white dark:bg-slate-800 text-slate-900 dark:text-white placeholder-slate-400 text-sm" />
        </div>
      )}

      {/* Overview */}
      {activeTab === 'overview' && (
        <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
          <div className="bg-white dark:bg-slate-800 rounded-lg border border-slate-200 dark:border-slate-700 p-4">
            <div className="flex items-center gap-2 mb-3"><SparkleRegular className="w-5 h-5 text-blue-600" /><h3 className="font-semibold text-slate-900 dark:text-white">New App Registrations (30d)</h3><span className="ml-auto text-sm text-slate-500">{data.recentRegistrations.length}</span></div>
            {data.recentRegistrations.length === 0 ? <p className="text-sm text-slate-500 text-center py-4">None in the last 30 days</p> : (
              <div className="space-y-2">
                {data.recentRegistrations.slice(0,5).map(app => (
                  <div key={app.id} className="flex items-center justify-between py-1.5 border-b border-slate-100 dark:border-slate-700 last:border-0">
                    <div><p className="text-sm font-medium text-slate-900 dark:text-white">{app.displayName}</p><p className="text-xs text-slate-500">{fmtDate(app.createdDateTime)}</p></div>
                    <CredBadge app={app} />
                  </div>
                ))}
                {data.recentRegistrations.length > 5 && <button onClick={() => setActiveTab('newRegistrations')} className="text-xs text-blue-600 dark:text-blue-400 hover:underline">View all {data.recentRegistrations.length} →</button>}
              </div>
            )}
          </div>
          <div className="bg-white dark:bg-slate-800 rounded-lg border border-slate-200 dark:border-slate-700 p-4">
            <div className="flex items-center gap-2 mb-3"><ErrorCircleRegular className="w-5 h-5 text-red-600" /><h3 className="font-semibold text-slate-900 dark:text-white">Expired Credentials</h3><span className="ml-auto text-sm text-slate-500">{data.expiredCredentialsApps.length}</span></div>
            {data.expiredCredentialsApps.length === 0 ? <div className="flex items-center gap-2 py-4 justify-center"><CheckmarkCircleRegular className="w-5 h-5 text-green-500" /><p className="text-sm text-slate-500">None found</p></div> : (
              <div className="space-y-2">
                {data.expiredCredentialsApps.slice(0,5).map(app => (
                  <div key={app.id} className="flex items-center justify-between py-1.5 border-b border-slate-100 dark:border-slate-700 last:border-0">
                    <div><p className="text-sm font-medium text-slate-900 dark:text-white">{app.displayName}</p><p className="text-xs text-slate-500">{app.secretCount}s / {app.certCount}c</p></div>
                    <span className="text-xs text-red-600 dark:text-red-400 font-medium">All expired</span>
                  </div>
                ))}
                {data.expiredCredentialsApps.length > 5 && <button onClick={() => setActiveTab('expiredCredentials')} className="text-xs text-blue-600 dark:text-blue-400 hover:underline">View all {data.expiredCredentialsApps.length} →</button>}
              </div>
            )}
          </div>
          <div className="bg-white dark:bg-slate-800 rounded-lg border border-slate-200 dark:border-slate-700 p-4">
            <div className="flex items-center gap-2 mb-3"><LockClosedRegular className="w-5 h-5 text-amber-600" /><h3 className="font-semibold text-slate-900 dark:text-white">Third-party Enterprise Apps</h3><span className="ml-auto text-sm text-slate-500">{data.thirdPartyEnterpriseApps.length}</span></div>
            <p className="text-xs text-slate-500 dark:text-slate-400 mb-3">Non-Microsoft apps with access to your tenant. Review regularly.</p>
            {data.thirdPartyEnterpriseApps.length === 0 ? <p className="text-sm text-slate-500 text-center py-2">None found</p> : (
              <div className="space-y-1">
                {data.thirdPartyEnterpriseApps.slice(0,5).map(app => (
                  <div key={app.id} className="flex items-center justify-between py-1.5 border-b border-slate-100 dark:border-slate-700 last:border-0">
                    <p className="text-sm text-slate-900 dark:text-white">{app.displayName}</p>
                    <p className="text-xs text-slate-500">{app.isVerified ? '✓ Verified' : 'Unverified'}</p>
                  </div>
                ))}
                {data.thirdPartyEnterpriseApps.length > 5 && <button onClick={() => setActiveTab('thirdParty')} className="text-xs text-blue-600 dark:text-blue-400 hover:underline">View all {data.thirdPartyEnterpriseApps.length} →</button>}
              </div>
            )}
          </div>
          <div className="bg-white dark:bg-slate-800 rounded-lg border border-slate-200 dark:border-slate-700 p-4">
            <div className="flex items-center gap-2 mb-3"><SparkleRegular className="w-5 h-5 text-purple-600" /><h3 className="font-semibold text-slate-900 dark:text-white">New Enterprise Apps (30d)</h3><span className="ml-auto text-sm text-slate-500">{data.recentEnterpriseApps.length}</span></div>
            {data.recentEnterpriseApps.length === 0 ? <p className="text-sm text-slate-500 text-center py-4">None in the last 30 days</p> : (
              <div className="space-y-1">
                {data.recentEnterpriseApps.slice(0,6).map(app => (
                  <div key={app.id} className="flex items-center justify-between py-1.5 border-b border-slate-100 dark:border-slate-700 last:border-0">
                    <p className="text-sm text-slate-900 dark:text-white truncate max-w-[200px]">{app.displayName}</p>
                    <span className="text-xs text-slate-500 ml-2 shrink-0">{fmtDate(app.createdDateTime)}</span>
                  </div>
                ))}
                {data.recentEnterpriseApps.length > 6 && <button onClick={() => setActiveTab('newEnterprise')} className="text-xs text-blue-600 dark:text-blue-400 hover:underline">View all {data.recentEnterpriseApps.length} →</button>}
              </div>
            )}
          </div>
        </div>
      )}

      {activeTab === 'registrations'      && <RegistrationTable apps={data.allRegistrationsByDate}     exportName="all-registrations.csv" />}
      {activeTab === 'newRegistrations'   && <RegistrationTable apps={data.recentRegistrations}         exportName="new-registrations-30d.csv" />}
      {activeTab === 'noCredentials'      && <RegistrationTable apps={data.noCredentialsApps}           exportName="registrations-no-credentials.csv" />}
      {activeTab === 'expiredCredentials' && <RegistrationTable apps={data.expiredCredentialsApps}      exportName="registrations-expired-credentials.csv" />}
      {activeTab === 'enterpriseApps'     && <EnterpriseTable   apps={data.allEnterpriseApps}           exportName="all-enterprise-apps.csv" />}
      {activeTab === 'thirdParty'         && <EnterpriseTable   apps={data.thirdPartyEnterpriseApps}    exportName="third-party-apps.csv" />}
      {activeTab === 'newEnterprise'      && <EnterpriseTable   apps={data.recentEnterpriseApps}        exportName="new-enterprise-apps-30d.csv" />}

      <p className="text-xs text-slate-400 text-right">Last updated: {new Date(data.lastUpdated).toLocaleString()}</p>
    </div>
  );
};

export default EnterpriseAppsPage;
