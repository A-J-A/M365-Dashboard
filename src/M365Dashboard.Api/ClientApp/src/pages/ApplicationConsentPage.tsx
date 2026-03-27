import React, { useState, useEffect, useCallback } from 'react';
import {
  Apps24Regular,
  Warning24Regular,
  ArrowSync24Regular,
  ShieldCheckmark24Regular,
  Clock24Regular,
} from '@fluentui/react-icons';
import {
  AppsRegular,
  ArrowDownloadRegular,
  SearchRegular,
  KeyRegular,
  CertificateRegular,
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
import { Badge } from '@fluentui/react-components';
import { PieChart, Pie, Cell, ResponsiveContainer, Tooltip, Legend } from 'recharts';
import { useAppContext } from '../contexts/AppContext';

// ── Types ─────────────────────────────────────────────────────────────────────

interface AppRegistration {
  id: string; appId: string; displayName: string; description?: string;
  createdDateTime?: string; createdDaysAgo?: number; isNew: boolean;
  signInAudience?: string; publisherDomain?: string;
  hasCredentials: boolean; noCredentials: boolean; allCredentialsExpired: boolean;
  secretCount: number; certCount: number; nextExpiry?: string;
  daysUntilNextExpiry?: number; requiresResourceAccess: boolean; resourceAccessCount: number;
  accountEnabled: boolean;
}
interface EnterpriseAppAudit {
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
  recentEnterpriseApps: EnterpriseAppAudit[]; allEnterpriseApps: EnterpriseAppAudit[];
  thirdPartyEnterpriseApps: EnterpriseAppAudit[]; disabledEnterpriseApps: EnterpriseAppAudit[];
  lastUpdated: string;
}

type TabKey =
  | 'overview' | 'risky' | 'grants' | 'enterpriseApps'
  | 'registrations' | 'noCredentials' | 'expiredCredentials'
  | 'thirdParty' | 'newApps';

// ── Component ─────────────────────────────────────────────────────────────────

const ApplicationConsentPage: React.FC = () => {
  const { getAccessToken } = useAppContext();

  // Consent/grants data
  const [oauth2Grants, setOauth2Grants]       = useState<any>(null);
  const [riskyConsents, setRiskyConsents]     = useState<any>(null);

  // Audit data (registrations + enterprise app hygiene)
  const [auditData, setAuditData]             = useState<AuditData | null>(null);

  const [loading, setLoading]   = useState(true);
  const [error, setError]       = useState<string | null>(null);
  const [activeTab, setActiveTab] = useState<TabKey>('overview');
  const [search, setSearch]     = useState('');
  const [sortCol, setSortCol]   = useState('createdDateTime');
  const [sortAsc, setSortAsc]   = useState(false);

  const fetchData = useCallback(async () => {
    try {
      setLoading(true); setError(null);
      const token = await getAccessToken();
      const h = { Authorization: `Bearer ${token}` };

      const [grantsRes, riskyRes, auditRes] = await Promise.all([
        fetch('/api/applicationconsent/oauth2-grants', { headers: h }),
        fetch('/api/applicationconsent/risky-consents', { headers: h }),
        fetch('/api/applicationconsent/enterprise-app-audit', { headers: h }),
      ]);

      if (grantsRes.ok) setOauth2Grants(await grantsRes.json());
      if (riskyRes.ok)  setRiskyConsents(await riskyRes.json());
      if (auditRes.ok)  setAuditData(await auditRes.json());
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Failed to load data');
    } finally {
      setLoading(false);
    }
  }, [getAccessToken]);

  useEffect(() => { fetchData(); }, [fetchData]);

  // ── Helpers ────────────────────────────────────────────────────────────────

  const fmtDate = (d?: string) =>
    d ? new Date(d).toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' }) : '—';

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

  // ── Sub-components ─────────────────────────────────────────────────────────

  const SortTh: React.FC<{ col: string; label: string }> = ({ col, label }) => (
    <th
      onClick={() => { if (sortCol === col) setSortAsc(!sortAsc); else { setSortCol(col); setSortAsc(true); } }}
      className="px-4 py-3 text-left text-xs font-medium text-slate-500 dark:text-slate-400 uppercase tracking-wider cursor-pointer hover:text-slate-700 dark:hover:text-slate-200 select-none">
      <span className="flex items-center gap-1">
        {label}
        {sortCol === col && (sortAsc ? <ChevronUpRegular className="w-3 h-3" /> : <ChevronDownRegular className="w-3 h-3" />)}
      </span>
    </th>
  );

  const CredBadge: React.FC<{ app: AppRegistration }> = ({ app }) => {
    if (app.noCredentials)
      return <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-xs font-medium bg-slate-100 text-slate-600 dark:bg-slate-700 dark:text-slate-400"><InfoRegular className="w-3 h-3" /> None</span>;
    if (app.allCredentialsExpired)
      return <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-xs font-medium bg-red-100 text-red-700 dark:bg-red-900/30 dark:text-red-400"><ErrorCircleRegular className="w-3 h-3" /> All Expired</span>;
    if ((app.daysUntilNextExpiry ?? 999) <= 30)
      return <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-xs font-medium bg-amber-100 text-amber-700 dark:bg-amber-900/30 dark:text-amber-400"><WarningRegular className="w-3 h-3" /> Expiring {app.daysUntilNextExpiry}d</span>;
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
      {rows.length === 0
        ? <div className="text-center py-12"><CheckmarkCircleRegular className="w-10 h-10 mx-auto mb-2 text-green-500" /><p className="text-slate-500">None found</p></div>
        : <div className="overflow-x-auto rounded-lg border border-slate-200 dark:border-slate-700">
            <table className="w-full text-sm">
              <thead className="bg-slate-50 dark:bg-slate-700/50">
                <tr>
                  <SortTh col="displayName" label="Application" />
                  <SortTh col="createdDateTime" label="Created" />
                  <th className="px-4 py-3 text-left text-xs font-medium text-slate-500 dark:text-slate-400 uppercase tracking-wider">Credentials</th>
                  <th className="px-4 py-3 text-left text-xs font-medium text-slate-500 dark:text-slate-400 uppercase tracking-wider">Next Expiry</th>
                  <th className="px-4 py-3 text-left text-xs font-medium text-slate-500 dark:text-slate-400 uppercase tracking-wider">Permissions</th>
                  <th className="px-4 py-3 text-left text-xs font-medium text-slate-500 dark:text-slate-400 uppercase tracking-wider">Audience</th>
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
                      {app.requiresResourceAccess
                        ? <span className="text-xs text-slate-600 dark:text-slate-300">{app.resourceAccessCount} permission{app.resourceAccessCount !== 1 ? 's' : ''}</span>
                        : <span className="text-xs text-slate-400">None</span>}
                    </td>
                    <td className="px-4 py-3 text-xs text-slate-600 dark:text-slate-300">
                      {app.signInAudience === 'AzureADMyOrg' ? 'Single tenant' : app.signInAudience === 'AzureADMultipleOrgs' ? 'Multi-tenant' : app.signInAudience ?? '—'}
                    </td>
                    <td className="px-4 py-3">
                      {app.accountEnabled
                        ? <span className="inline-flex items-center px-2 py-0.5 rounded-full text-xs font-medium bg-green-100 text-green-700 dark:bg-green-900/30 dark:text-green-400">Activated</span>
                        : <span className="inline-flex items-center px-2 py-0.5 rounded-full text-xs font-medium bg-slate-100 text-slate-600 dark:bg-slate-700 dark:text-slate-400">Deactivated</span>}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>}
    </>;
  };

  const EnterpriseAuditTable: React.FC<{ apps: EnterpriseAppAudit[]; exportName: string }> = ({ apps, exportName }) => {
    const rows = filterAndSort(apps);
    return <>
      <div className="flex items-center justify-between mb-3">
        <span className="text-sm text-slate-500 dark:text-slate-400">{rows.length} enterprise app{rows.length !== 1 ? 's' : ''}</span>
        <button onClick={() => exportCsv(rows, ['displayName','appId','createdDateTime','publisherName','accountEnabled','isMicrosoftApp','isOwnRegistration'], exportName)}
          className="flex items-center gap-2 px-3 py-1.5 text-sm text-slate-600 dark:text-slate-400 border border-slate-300 dark:border-slate-600 rounded-lg hover:bg-slate-50 dark:hover:bg-slate-700 transition-colors">
          <ArrowDownloadRegular className="w-4 h-4" /> Export CSV
        </button>
      </div>
      {rows.length === 0
        ? <div className="text-center py-12"><CheckmarkCircleRegular className="w-10 h-10 mx-auto mb-2 text-green-500" /><p className="text-slate-500">None found</p></div>
        : <div className="overflow-x-auto rounded-lg border border-slate-200 dark:border-slate-700">
            <table className="w-full text-sm">
              <thead className="bg-slate-50 dark:bg-slate-700/50">
                <tr>
                  <SortTh col="displayName" label="Application" />
                  <SortTh col="createdDateTime" label="Added" />
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
                    <td className="px-4 py-3 text-sm text-slate-600 dark:text-slate-300">{app.publisherName ?? '—'}</td>
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
          </div>}
    </>;
  };

  // ── Tabs config ────────────────────────────────────────────────────────────

  const tabs: { key: TabKey; label: string; count?: number; section: 'consent' | 'hygiene' }[] = [
    { key: 'overview',           label: 'Overview',            section: 'consent' },
    { key: 'risky',              label: 'Risky Consents',      count: riskyConsents?.summary?.totalRiskyConsents, section: 'consent' },
    { key: 'grants',             label: 'OAuth Grants',        count: oauth2Grants?.summary?.totalGrants, section: 'consent' },
    { key: 'enterpriseApps',     label: 'Enterprise Apps',     count: auditData?.summary.totalEnterpriseApps, section: 'consent' },
    { key: 'registrations',      label: 'App Registrations',   count: auditData?.summary.totalRegistrations, section: 'hygiene' },
    { key: 'noCredentials',      label: 'No Credentials',      count: auditData?.summary.noCredentials, section: 'hygiene' },
    { key: 'expiredCredentials', label: 'Expired Credentials', count: auditData?.summary.allCredentialsExpired, section: 'hygiene' },
    { key: 'thirdParty',         label: 'Third-party',         count: auditData?.summary.thirdPartyApps, section: 'hygiene' },
    { key: 'newApps',            label: 'New (30d)',            count: (auditData?.summary.newRegistrations30d ?? 0) + (auditData?.summary.newEnterpriseApps30d ?? 0), section: 'hygiene' },
  ];

  const riskLevelData = riskyConsents?.summary ? [
    { name: 'High',   value: riskyConsents.summary.highRisk,   color: '#ef4444' },
    { name: 'Medium', value: riskyConsents.summary.mediumRisk, color: '#f59e0b' },
    { name: 'Low',    value: riskyConsents.summary.lowRisk,    color: '#22c55e' },
  ].filter(d => d.value > 0) : [];

  // ── Render ─────────────────────────────────────────────────────────────────

  if (loading) return (
    <div className="p-4 flex items-center justify-center h-64">
      <div className="flex flex-col items-center gap-3">
        <div className="animate-spin rounded-full h-8 w-8 border-b-2 border-blue-600"></div>
        <p className="text-slate-500 dark:text-slate-400">Loading application data...</p>
      </div>
    </div>
  );

  return (
    <div className="p-4 space-y-4 max-w-7xl mx-auto">

      {/* Header */}
      <div className="flex items-start justify-between">
        <div>
          <h1 className="text-xl font-semibold text-slate-900 dark:text-white flex items-center gap-2">
            <Apps24Regular className="w-6 h-6" /> Application Management
          </h1>
          <p className="text-sm text-slate-500 dark:text-slate-400">OAuth consents, permissions, credentials and app lifecycle</p>
        </div>
        <button onClick={fetchData} className="px-3 py-2 bg-slate-100 dark:bg-slate-700 text-slate-700 dark:text-slate-300 rounded-lg hover:bg-slate-200 dark:hover:bg-slate-600 transition-colors">
          <ArrowSync24Regular className="w-4 h-4" />
        </button>
      </div>

      {error && <div className="bg-red-50 dark:bg-red-900/20 border border-red-200 dark:border-red-800 rounded-lg p-4 text-red-700 dark:text-red-300">{error}</div>}

      {/* Summary cards */}
      <div className="grid grid-cols-2 md:grid-cols-5 gap-3">
        {[
          { label: 'OAuth Grants',        value: oauth2Grants?.summary?.totalGrants ?? 0,              color: 'text-slate-900 dark:text-white' },
          { label: 'Enterprise Apps',     value: auditData?.summary.totalEnterpriseApps ?? 0,          color: 'text-blue-600' },
          { label: 'App Registrations',   value: auditData?.summary.totalRegistrations ?? 0,           color: 'text-purple-600' },
          { label: 'Risky Consents',      value: riskyConsents?.summary?.totalRiskyConsents ?? 0,      color: (riskyConsents?.summary?.totalRiskyConsents ?? 0) > 0 ? 'text-red-600' : 'text-green-600' },
          { label: 'Expired Credentials', value: auditData?.summary.allCredentialsExpired ?? 0,        color: (auditData?.summary.allCredentialsExpired ?? 0) > 0 ? 'text-red-600' : 'text-green-600' },
        ].map(c => (
          <div key={c.label} className="bg-white dark:bg-slate-800 rounded-lg border border-slate-200 dark:border-slate-700 p-4">
            <p className={`text-2xl font-bold ${c.color}`}>{c.value}</p>
            <p className="text-xs text-slate-500 dark:text-slate-400 mt-0.5">{c.label}</p>
          </div>
        ))}
      </div>

      {/* Alerts */}
      {((riskyConsents?.summary?.highRisk ?? 0) > 0 || (auditData?.summary.allCredentialsExpired ?? 0) > 0) && (
        <div className="space-y-2">
          {(riskyConsents?.summary?.highRisk ?? 0) > 0 && (
            <button onClick={() => setActiveTab('risky')}
              className="w-full text-left bg-red-50 dark:bg-red-900/20 border border-red-200 dark:border-red-800 rounded-lg p-4 flex items-start gap-3 hover:bg-red-100 dark:hover:bg-red-900/30 transition-colors">
              <Warning24Regular className="w-5 h-5 text-red-600 flex-shrink-0 mt-0.5" />
              <div>
                <p className="font-medium text-red-800 dark:text-red-200">High-Risk App Consents Detected</p>
                <p className="text-sm text-red-700 dark:text-red-300">{riskyConsents.summary.highRisk} applications have been granted high-risk permissions.</p>
              </div>
            </button>
          )}
          {(auditData?.summary.allCredentialsExpired ?? 0) > 0 && (
            <button onClick={() => setActiveTab('expiredCredentials')}
              className="w-full text-left bg-amber-50 dark:bg-amber-900/20 border border-amber-200 dark:border-amber-800 rounded-lg p-4 flex items-start gap-3 hover:bg-amber-100 dark:hover:bg-amber-900/30 transition-colors">
              <Clock24Regular className="w-5 h-5 text-amber-600 flex-shrink-0 mt-0.5" />
              <div>
                <p className="font-medium text-amber-800 dark:text-amber-200">Expired Application Credentials</p>
                <p className="text-sm text-amber-700 dark:text-amber-300">{auditData!.summary.allCredentialsExpired} app registrations have all credentials expired.</p>
              </div>
            </button>
          )}
        </div>
      )}

      {/* Tabs */}
      <div className="border-b border-slate-200 dark:border-slate-700">
        <nav className="flex gap-1 overflow-x-auto">
          {tabs.map(t => (
            <button key={t.key} onClick={() => setActiveTab(t.key)}
              className={`flex items-center gap-2 px-3 py-2.5 text-sm font-medium whitespace-nowrap border-b-2 transition-colors ${
                activeTab === t.key
                  ? 'border-blue-600 text-blue-600 dark:text-blue-400'
                  : 'border-transparent text-slate-600 dark:text-slate-400 hover:text-slate-900 dark:hover:text-slate-200'
              }`}>
              {t.label}
              {t.count !== undefined && (
                <span className={`px-1.5 py-0.5 rounded-full text-xs font-medium bg-slate-100 dark:bg-slate-700 ${
                  t.count > 0 && (t.key === 'risky' || t.key === 'expiredCredentials') ? 'text-red-600 dark:text-red-400' : 'text-slate-600 dark:text-slate-300'
                }`}>{t.count}</span>
              )}
            </button>
          ))}
        </nav>
      </div>

      {/* Search (hygiene tabs) */}
      {['registrations','noCredentials','expiredCredentials','thirdParty','newApps','enterpriseApps'].includes(activeTab) && (
        <div className="relative">
          <SearchRegular className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-400" />
          <input type="text" value={search} onChange={e => setSearch(e.target.value)} placeholder="Search by name or App ID..."
            className="w-full pl-9 pr-4 py-2 border border-slate-300 dark:border-slate-600 rounded-lg bg-white dark:bg-slate-800 text-slate-900 dark:text-white placeholder-slate-400 text-sm" />
        </div>
      )}

      {/* ── OVERVIEW TAB ──────────────────────────────────────────────────── */}
      {activeTab === 'overview' && (
        <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
          {/* Risky consents chart */}
          <div className="bg-white dark:bg-slate-800 rounded-lg border border-slate-200 dark:border-slate-700 p-5">
            <h3 className="font-semibold mb-4 text-slate-900 dark:text-white">Risky Consents by Level</h3>
            {riskLevelData.length === 0
              ? <div className="flex items-center justify-center h-48 gap-2"><CheckmarkCircleRegular className="w-6 h-6 text-green-500" /><p className="text-slate-500">No risky consents</p></div>
              : <div className="h-48">
                  <ResponsiveContainer width="100%" height="100%">
                    <PieChart>
                      <Pie data={riskLevelData} dataKey="value" nameKey="name" cx="50%" cy="50%" outerRadius={70} label>
                        {riskLevelData.map((e, i) => <Cell key={i} fill={e.color} />)}
                      </Pie>
                      <Legend /><Tooltip />
                    </PieChart>
                  </ResponsiveContainer>
                </div>}
          </div>

          {/* Top apps by permissions */}
          <div className="bg-white dark:bg-slate-800 rounded-lg border border-slate-200 dark:border-slate-700 p-5">
            <h3 className="font-semibold mb-4 text-slate-900 dark:text-white">Top Apps by Permissions</h3>
            <div className="space-y-2">
              {oauth2Grants?.topAppsByPermissions?.slice(0, 6).map((app: any, i: number) => (
                <div key={i} className="flex items-center justify-between p-2 bg-slate-50 dark:bg-slate-700/50 rounded-lg">
                  <span className="font-medium text-sm text-slate-900 dark:text-white truncate max-w-[200px]">{app.appName}</span>
                  <div className="flex items-center gap-2 shrink-0">
                    <span className="text-xs text-slate-500">{app.totalScopes} scopes</span>
                    {app.hasHighRiskPermissions && <Warning24Regular className="w-4 h-4 text-red-500" />}
                  </div>
                </div>
              ))}
            </div>
          </div>

          {/* Credential status */}
          <div className="bg-white dark:bg-slate-800 rounded-lg border border-slate-200 dark:border-slate-700 p-5">
            <h3 className="font-semibold mb-4 text-slate-900 dark:text-white">Credential Status</h3>
            {auditData && (
              <div className="space-y-2">
                {[
                  { label: 'Valid credentials',    value: (auditData.summary.totalRegistrations - auditData.summary.noCredentials - auditData.summary.allCredentialsExpired), color: 'bg-green-500' },
                  { label: 'No credentials',        value: auditData.summary.noCredentials,         color: 'bg-slate-400' },
                  { label: 'All credentials expired', value: auditData.summary.allCredentialsExpired, color: 'bg-red-500' },
                ].map(r => (
                  <div key={r.label} className="flex items-center justify-between">
                    <div className="flex items-center gap-2">
                      <div className={`w-2.5 h-2.5 rounded-full ${r.color}`} />
                      <span className="text-sm text-slate-600 dark:text-slate-300">{r.label}</span>
                    </div>
                    <span className="text-sm font-medium text-slate-900 dark:text-white">{r.value}</span>
                  </div>
                ))}
              </div>
            )}
          </div>

          {/* New apps */}
          <div className="bg-white dark:bg-slate-800 rounded-lg border border-slate-200 dark:border-slate-700 p-5">
            <h3 className="font-semibold mb-4 text-slate-900 dark:text-white">New Apps (Last 30 days)</h3>
            {auditData && (
              <div className="space-y-2">
                {[...( auditData.recentRegistrations.slice(0,3).map(a => ({ name: a.displayName, date: a.createdDateTime, type: 'Registration' }))),
                  ...( auditData.recentEnterpriseApps.slice(0,3).map(a => ({ name: a.displayName, date: a.createdDateTime, type: 'Enterprise' })))
                ].sort((a, b) => (b.date ?? '') > (a.date ?? '') ? 1 : -1).slice(0, 6).map((a, i) => (
                  <div key={i} className="flex items-center justify-between py-1 border-b border-slate-100 dark:border-slate-700 last:border-0">
                    <span className="text-sm text-slate-900 dark:text-white truncate max-w-[180px]">{a.name}</span>
                    <div className="flex items-center gap-2 shrink-0">
                      <span className="text-xs text-slate-400">{fmtDate(a.date)}</span>
                      <span className={`text-xs px-1.5 py-0.5 rounded ${a.type === 'Registration' ? 'bg-purple-100 text-purple-700 dark:bg-purple-900/30 dark:text-purple-400' : 'bg-blue-100 text-blue-700 dark:bg-blue-900/30 dark:text-blue-400'}`}>{a.type}</span>
                    </div>
                  </div>
                ))}
                {(auditData.recentRegistrations.length + auditData.recentEnterpriseApps.length) === 0 && <p className="text-sm text-slate-500 text-center py-4">None in the last 30 days</p>}
              </div>
            )}
          </div>
        </div>
      )}

      {/* ── RISKY TAB ─────────────────────────────────────────────────────── */}
      {activeTab === 'risky' && (
        <div className="bg-white dark:bg-slate-800 rounded-lg border border-slate-200 dark:border-slate-700 overflow-hidden">
          <div className="p-4 border-b border-slate-200 dark:border-slate-700 flex items-center gap-2">
            <Warning24Regular className="w-5 h-5 text-red-500" />
            <h3 className="font-semibold text-slate-900 dark:text-white">Risky Application Consents</h3>
          </div>
          {!riskyConsents?.riskyConsents?.length
            ? <p className="text-slate-500 text-center py-12">No risky consents detected</p>
            : <div className="overflow-x-auto">
                <table className="w-full text-sm">
                  <thead className="bg-slate-50 dark:bg-slate-700/50">
                    <tr>
                      {['Application','Risk Level','Risk Factors','High-Risk Scopes','Consent Type'].map(h => (
                        <th key={h} className="px-4 py-3 text-left text-xs font-medium text-slate-500 dark:text-slate-400 uppercase tracking-wider">{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-200 dark:divide-slate-700">
                    {riskyConsents.riskyConsents.map((c: any, i: number) => (
                      <tr key={i} className="hover:bg-slate-50 dark:hover:bg-slate-700/30">
                        <td className="px-4 py-3">
                          <div className="flex items-center gap-2">
                            {c.isVerified ? <ShieldCheckmark24Regular className="w-5 h-5 text-green-500 shrink-0" /> : <Warning24Regular className="w-5 h-5 text-amber-500 shrink-0" />}
                            <div><p className="font-medium text-slate-900 dark:text-white">{c.appName}</p><p className="text-xs text-slate-500">{c.publisherName || 'Unknown'}</p></div>
                          </div>
                        </td>
                        <td className="px-4 py-3">
                          <Badge appearance="filled" color={c.riskLevel === 'High' ? 'danger' : c.riskLevel === 'Medium' ? 'warning' : 'success'}>{c.riskLevel}</Badge>
                        </td>
                        <td className="px-4 py-3">
                          <div className="flex flex-wrap gap-1">
                            {c.riskFactors?.map((f: string, j: number) => <Badge key={j} appearance="tint" size="small">{f}</Badge>)}
                          </div>
                        </td>
                        <td className="px-4 py-3">
                          <div className="flex flex-wrap gap-1 max-w-xs">
                            {c.riskyScopes?.slice(0,3).map((s: string, j: number) => <Badge key={j} appearance="outline" color="danger" size="small">{s}</Badge>)}
                            {c.riskyScopes?.length > 3 && <Badge appearance="tint" size="small">+{c.riskyScopes.length - 3}</Badge>}
                          </div>
                        </td>
                        <td className="px-4 py-3">
                          <Badge appearance="tint" color={c.consentType === 'AllPrincipals' ? 'warning' : 'brand'} size="small">
                            {c.consentType === 'AllPrincipals' ? 'Admin (All Users)' : 'User'}
                          </Badge>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>}
        </div>
      )}

      {/* ── OAUTH GRANTS TAB ──────────────────────────────────────────────── */}
      {activeTab === 'grants' && (
        <div className="bg-white dark:bg-slate-800 rounded-lg border border-slate-200 dark:border-slate-700 overflow-hidden">
          <div className="p-4 border-b border-slate-200 dark:border-slate-700">
            <h3 className="font-semibold text-slate-900 dark:text-white">OAuth2 Permission Grants</h3>
          </div>
          <div className="overflow-x-auto">
            <table className="w-full text-sm">
              <thead className="bg-slate-50 dark:bg-slate-700/50">
                <tr>
                  {['Application','Consent Type','Scopes'].map(h => (
                    <th key={h} className="px-4 py-3 text-left text-xs font-medium text-slate-500 dark:text-slate-400 uppercase tracking-wider">{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-200 dark:divide-slate-700">
                {oauth2Grants?.grants?.slice(0, 50).map((g: any, i: number) => (
                  <tr key={i} className="hover:bg-slate-50 dark:hover:bg-slate-700/30">
                    <td className="px-4 py-3 font-medium text-slate-900 dark:text-white">{g.clientDisplayName}</td>
                    <td className="px-4 py-3">
                      <Badge appearance="tint" color={g.consentType === 'AllPrincipals' ? 'warning' : 'brand'} size="small">
                        {g.consentType === 'AllPrincipals' ? 'Admin' : 'User'}
                      </Badge>
                    </td>
                    <td className="px-4 py-3">
                      <div className="flex flex-wrap gap-1 max-w-md">
                        {g.scopes?.slice(0,5).map((s: string, j: number) => <Badge key={j} appearance="outline" size="small">{s}</Badge>)}
                        {g.scopes?.length > 5 && <Badge appearance="tint" size="small">+{g.scopes.length - 5}</Badge>}
                      </div>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {/* ── ENTERPRISE APPS TAB ───────────────────────────────────────── */}
      {activeTab === 'enterpriseApps' && auditData && <EnterpriseAuditTable apps={auditData.allEnterpriseApps} exportName="all-enterprise-apps.csv" />}

      {/* ── HYGIENE TABS ──────────────────────────────────────────────────── */}
      {activeTab === 'registrations'      && auditData && <RegistrationTable apps={auditData.allRegistrationsByDate}    exportName="all-registrations.csv" />}
      {activeTab === 'noCredentials'      && auditData && <RegistrationTable apps={auditData.noCredentialsApps}          exportName="no-credentials.csv" />}
      {activeTab === 'expiredCredentials' && auditData && <RegistrationTable apps={auditData.expiredCredentialsApps}     exportName="expired-credentials.csv" />}
      {activeTab === 'thirdParty'         && auditData && <EnterpriseAuditTable apps={auditData.thirdPartyEnterpriseApps} exportName="third-party-apps.csv" />}
      {activeTab === 'newApps'            && auditData && (
        <div className="space-y-6">
          <div>
            <h3 className="text-sm font-semibold text-slate-700 dark:text-slate-300 mb-3 flex items-center gap-2">
              <SparkleRegular className="w-4 h-4 text-purple-600" /> New App Registrations ({auditData.recentRegistrations.length})
            </h3>
            <RegistrationTable apps={auditData.recentRegistrations} exportName="new-registrations-30d.csv" />
          </div>
          <div>
            <h3 className="text-sm font-semibold text-slate-700 dark:text-slate-300 mb-3 flex items-center gap-2">
              <SparkleRegular className="w-4 h-4 text-blue-600" /> New Enterprise Apps ({auditData.recentEnterpriseApps.length})
            </h3>
            <EnterpriseAuditTable apps={auditData.recentEnterpriseApps} exportName="new-enterprise-apps-30d.csv" />
          </div>
        </div>
      )}

      <p className="text-xs text-slate-400 text-right">
        Last updated: {auditData ? new Date(auditData.lastUpdated).toLocaleString() : '—'}
      </p>
    </div>
  );
};

export default ApplicationConsentPage;
