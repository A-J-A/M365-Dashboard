import React, { useState, useEffect } from 'react';
import {
  SendRegular,
  MailInboxRegular,
  OpenRegular,
  CheckmarkCircleFilled,
  ArrowTrendingRegular,
  WarningRegular,
  ErrorCircleRegular,
  StorageRegular,
  ArchiveRegular,
  ClockRegular,
  ArrowForwardRegular,
  GlobeRegular,
  BuildingRegular,
  InfoRegular,
  ArrowDownloadRegular,
  FilterRegular,
  PersonRegular,
  LockOpenRegular,
  PeopleRegular,
  SearchRegular,
  ArrowSwapRegular,
} from '@fluentui/react-icons';
import { useAppContext } from '../contexts/AppContext';

interface MailTrafficReport {
  date: string;
  messagesSent: number;
  messagesReceived: number;
  spamReceived: number;
  malwareReceived: number;
  goodMail: number;
}

interface TopSender {
  userPrincipalName: string;
  displayName: string | null;
  messageCount: number;
}

interface TopRecipient {
  userPrincipalName: string;
  displayName: string | null;
  messageCount: number;
}

interface MailflowSummary {
  totalMessagesSent: number;
  totalMessagesReceived: number;
  totalSpamBlocked: number;
  totalMalwareBlocked: number;
  averageMessagesPerDay: number;
  dailyTraffic: MailTrafficReport[];
  topSenders: TopSender[];
  topRecipients: TopRecipient[];
  lastUpdated: string;
}

interface StorageMailbox {
  userPrincipalName: string;
  displayName: string;
  storageUsedBytes: number;
  storageUsedGB: number;
  prohibitSendQuotaGB: number;
  percentUsed: number;
  itemCount: number;
  status: 'Full' | 'Critical' | 'Warning' | 'OK';
  lastActivityDate: string | null;
  hasArchive: boolean;
  isDeleted: boolean;
}

interface StorageUsageData {
  nearQuotaMailboxes: StorageMailbox[];
  allMailboxes: StorageMailbox[];
  summary: {
    totalMailboxes: number;
    criticalCount: number;
    warningCount: number;
    okCount: number;
    totalStorageUsedGB: number;
    averagePercentUsed: number;
    archiveEnabledCount: number;
  };
  lastUpdated?: string;
  error?: string;
}

interface ForwardingMailbox {
  userId: string;
  userPrincipalName: string;
  displayName: string;
  mail: string;
  ruleId: string;
  ruleName: string;
  ruleEnabled: boolean;
  forwardingType: 'Forward' | 'Redirect' | 'ForwardAsAttachment';
  forwardingTarget: string;
  forwardingTargetDomain: string;
  isExternal: boolean;
  deliverToMailbox: boolean;
  riskLevel: 'High' | 'Low';
}

interface ForwardingReportData {
  forwardingRules: ForwardingMailbox[];
  totalMailboxesScanned: number;
  mailboxesWithForwarding: number;
  totalForwardingRules: number;
  externalForwardingCount: number;
  internalForwardingCount: number;
  tenantDomains: string[];
  lastUpdated?: string;
  error?: string;
}

// Exchange Mailbox Forwarding interfaces
interface ExchangeMailboxForwarding {
  id: string;
  displayName: string;
  userPrincipalName: string;
  primarySmtpAddress: string;
  forwardingAddress: string | null;
  forwardingSmtpAddress: string | null;
  forwardingTarget: string;
  forwardingTargetDomain: string;
  deliverToMailboxAndForward: boolean;
  isExternal: boolean;
  recipientTypeDetails: string;
}

interface ExchangeMailboxForwardingResult {
  mailboxes: ExchangeMailboxForwarding[];
  totalCount: number;
  externalCount: number;
  internalCount: number;
  lastUpdated: string;
}

type ViewMode = 'traffic' | 'storage' | 'forwarding' | 'mailboxFwd' | 'access';
type ForwardingFilter = 'all' | 'external' | 'internal';

const MailflowPage: React.FC = () => {
  const { getAccessToken } = useAppContext();
  const [summary, setSummary] = useState<MailflowSummary | null>(null);
  const [storageData, setStorageData] = useState<StorageUsageData | null>(null);
  const [forwardingData, setForwardingData] = useState<ForwardingReportData | null>(null);
  const [exchangeForwardingData, setExchangeForwardingData] = useState<ExchangeMailboxForwardingResult | null>(null);
  const [summaryLoading, setSummaryLoading] = useState(true);
  const [storageLoading, setStorageLoading] = useState(true);
  const [forwardingLoading, setForwardingLoading] = useState(false);
  const [exchangeForwardingLoading, setExchangeForwardingLoading] = useState(false);
  const [viewMode, setViewMode] = useState<ViewMode>('traffic');
  const [period, setPeriod] = useState(7);
  const [forwardingFilter, setForwardingFilter] = useState<ForwardingFilter>('all');
  const [exchangeForwardingFilter, setExchangeForwardingFilter] = useState<ForwardingFilter>('all');

  // Mailbox Access state
  type AccessQueryMode = 'by-user' | 'delegates';
  type MailboxAccessEntry = {
    mailboxEmail: string; mailboxDisplayName: string | null; mailboxType: string;
    permission: string; grantedTo: string; isInherited: boolean;
  };
  type MailboxAccessResult = {
    subjectEmail: string; queryType: string;
    fullAccessMailboxes: MailboxAccessEntry[];
    sendAsMailboxes: MailboxAccessEntry[];
    sendOnBehalfMailboxes: MailboxAccessEntry[];
    totalCount: number; mailboxesChecked: number; lastUpdated: string;
  };
  const [accessEmail, setAccessEmail] = useState('');
  const [accessMode, setAccessMode] = useState<AccessQueryMode>('by-user');
  const [accessLoading, setAccessLoading] = useState(false);
  const [accessResult, setAccessResult] = useState<MailboxAccessResult | null>(null);
  const [accessError, setAccessError] = useState<string | null>(null);

  const handleAccessSearch = async () => {
    const trimmed = accessEmail.trim();
    if (!trimmed || !trimmed.includes('@')) { setAccessError('Please enter a valid email address.'); return; }
    setAccessLoading(true); setAccessResult(null); setAccessError(null);
    try {
      const token = await getAccessToken();
      const endpoint = accessMode === 'by-user'
        ? `/api/exchange/mailbox-access/by-user?email=${encodeURIComponent(trimmed)}`
        : `/api/exchange/mailbox-access/delegates?email=${encodeURIComponent(trimmed)}`;
      const response = await fetch(endpoint, { headers: { Authorization: `Bearer ${token}` } });
      if (!response.ok) { const d = await response.json().catch(() => ({})); throw new Error(d.message ?? d.error ?? `HTTP ${response.status}`); }
      setAccessResult(await response.json());
    } catch (err) {
      setAccessError(err instanceof Error ? err.message : 'An unexpected error occurred.');
    } finally { setAccessLoading(false); }
  };

  useEffect(() => {
    fetchSummary();
    fetchStorageUsage();
  }, []);

  useEffect(() => {
    if (viewMode === 'forwarding' && !forwardingData && !forwardingLoading) {
      fetchForwardingReport();
    }
    if (viewMode === 'mailboxFwd' && !exchangeForwardingData && !exchangeForwardingLoading) {
      fetchExchangeForwarding();
    }
  }, [viewMode]);

  useEffect(() => {
    fetchSummary();
  }, [period]);

  const fetchSummary = async () => {
    try {
      setSummaryLoading(true);
      const token = await getAccessToken();
      const response = await fetch(`/api/mailflow/summary?days=${period}`, {
        headers: {
          'Authorization': `Bearer ${token}`,
          'Content-Type': 'application/json',
        },
      });

      if (!response.ok) {
        throw new Error('Failed to fetch mailflow summary');
      }

      const data: MailflowSummary = await response.json();
      setSummary(data);
    } catch (err) {
      console.error('Failed to fetch summary:', err);
    } finally {
      setSummaryLoading(false);
    }
  };

  const fetchStorageUsage = async () => {
    try {
      setStorageLoading(true);
      const token = await getAccessToken();
      const response = await fetch('/api/mailflow/storage-usage?thresholdPercent=80', {
        headers: {
          'Authorization': `Bearer ${token}`,
          'Content-Type': 'application/json',
        },
      });

      if (!response.ok) {
        throw new Error('Failed to fetch storage usage');
      }

      const data: StorageUsageData = await response.json();
      setStorageData(data);
    } catch (err) {
      console.error('Failed to fetch storage usage:', err);
    } finally {
      setStorageLoading(false);
    }
  };

  const fetchForwardingReport = async () => {
    try {
      setForwardingLoading(true);
      const token = await getAccessToken();
      const response = await fetch('/api/exchange/inbox-rules-forwarding?take=100', {
        headers: {
          'Authorization': `Bearer ${token}`,
          'Content-Type': 'application/json',
        },
      });

      if (!response.ok) {
        throw new Error('Failed to fetch forwarding report');
      }

      const data: ForwardingReportData = await response.json();
      setForwardingData(data);
    } catch (err) {
      console.error('Failed to fetch forwarding report:', err);
    } finally {
      setForwardingLoading(false);
    }
  };

  const fetchExchangeForwarding = async () => {
    try {
      setExchangeForwardingLoading(true);
      const token = await getAccessToken();
      const response = await fetch('/api/exchange/mailbox-forwarding?take=500', {
        headers: {
          'Authorization': `Bearer ${token}`,
          'Content-Type': 'application/json',
        },
      });

      if (!response.ok) {
        throw new Error('Failed to fetch Exchange mailbox forwarding');
      }

      const data: ExchangeMailboxForwardingResult = await response.json();
      setExchangeForwardingData(data);
    } catch (err) {
      console.error('Failed to fetch Exchange mailbox forwarding:', err);
    } finally {
      setExchangeForwardingLoading(false);
    }
  };

  const exportForwardingToCsv = () => {
    if (!forwardingData?.forwardingRules) return;

    const filteredData = getFilteredForwardingData();
    const headers = ['Display Name', 'Email', 'Forwarding Type', 'Forwarding Target', 'Target Domain', 'Internal/External', 'Rule Name', 'Rule Enabled', 'Keeps Copy', 'Risk Level'];
    const rows = filteredData.map(m => [
      m.displayName,
      m.mail,
      m.forwardingType,
      m.forwardingTarget,
      m.forwardingTargetDomain,
      m.isExternal ? 'External' : 'Internal',
      m.ruleName,
      m.ruleEnabled ? 'Yes' : 'No',
      m.deliverToMailbox ? 'Yes' : 'No',
      m.riskLevel
    ]);

    const csvContent = [headers, ...rows].map(row => row.map(cell => `"${cell}"`).join(',')).join('\n');
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = `inbox-rules-forwarding-report-${new Date().toISOString().split('T')[0]}.csv`;
    link.click();
  };

  const getFilteredForwardingData = (): ForwardingMailbox[] => {
    if (!forwardingData?.forwardingRules) return [];
    
    switch (forwardingFilter) {
      case 'external':
        return forwardingData.forwardingRules.filter(m => m.isExternal);
      case 'internal':
        return forwardingData.forwardingRules.filter(m => !m.isExternal);
      default:
        return forwardingData.forwardingRules;
    }
  };

  const getFilteredExchangeForwardingData = (): ExchangeMailboxForwarding[] => {
    if (!exchangeForwardingData?.mailboxes) return [];
    
    switch (exchangeForwardingFilter) {
      case 'external':
        return exchangeForwardingData.mailboxes.filter(m => m.isExternal);
      case 'internal':
        return exchangeForwardingData.mailboxes.filter(m => !m.isExternal);
      default:
        return exchangeForwardingData.mailboxes;
    }
  };

  const exportExchangeForwardingToCsv = () => {
    if (!exchangeForwardingData?.mailboxes) return;

    const filteredData = getFilteredExchangeForwardingData();
    const headers = ['Display Name', 'Email', 'Forwarding Target', 'Target Domain', 'Internal/External', 'Keeps Copy', 'Recipient Type'];
    const rows = filteredData.map(m => [
      m.displayName,
      m.primarySmtpAddress,
      m.forwardingTarget,
      m.forwardingTargetDomain,
      m.isExternal ? 'External' : 'Internal',
      m.deliverToMailboxAndForward ? 'Yes' : 'No',
      m.recipientTypeDetails
    ]);

    const csvContent = [headers, ...rows].map(row => row.map(cell => `"${cell}"`).join(',')).join('\n');
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = `exchange-mailbox-forwarding-${new Date().toISOString().split('T')[0]}.csv`;
    link.click();
  };

  const formatNumber = (num: number) => {
    if (num >= 1000000) return `${(num / 1000000).toFixed(1)}M`;
    if (num >= 1000) return `${(num / 1000).toFixed(1)}K`;
    return num.toString();
  };

  // Stat card component
  const StatCard: React.FC<{
    title: string;
    value: number | string;
    icon: React.ReactNode;
    color: string;
  }> = ({ title, value, icon, color }) => (
    <div className="w-full bg-white dark:bg-slate-800 rounded-lg border-2 border-slate-200 dark:border-slate-700 p-2">
      <div className="flex items-center gap-1.5">
        <div className={`p-1 rounded ${color.replace('text-', 'bg-').replace('-600', '-100')} dark:bg-opacity-20 flex-shrink-0`}>
          {icon}
        </div>
        <div className="min-w-0 flex-1">
          <p className="text-[10px] text-slate-500 dark:text-slate-400 leading-tight truncate">{title}</p>
          <p className={`text-sm font-semibold leading-tight ${color}`}>{typeof value === 'number' ? formatNumber(value) : value}</p>
        </div>
      </div>
    </div>
  );

  // Simple bar chart for daily traffic
  const TrafficChart: React.FC<{ data: MailTrafficReport[] }> = ({ data }) => {
    if (data.length === 0) return <p className="text-slate-500">No data available</p>;

    // Sort by date ascending and take the last 14 days for display
    const sortedData = [...data].sort((a, b) => a.date.localeCompare(b.date)).slice(-14);
    const maxValue = Math.max(...sortedData.map(d => Math.max(d.messagesSent, d.messagesReceived)), 1);

    // Helper to format date string - handles yyyy-MM-dd format from API
    const formatDateLabel = (dateStr: string) => {
      try {
        // Handle ISO date format (yyyy-MM-dd)
        if (dateStr.includes('-')) {
          const [year, month, day] = dateStr.split('T')[0].split('-');
          const date = new Date(parseInt(year), parseInt(month) - 1, parseInt(day));
          return date.toLocaleDateString('en-US', { month: 'short', day: 'numeric' });
        }
        // Fallback to direct parsing
        const date = new Date(dateStr);
        if (!isNaN(date.getTime())) {
          return date.toLocaleDateString('en-US', { month: 'short', day: 'numeric' });
        }
        return dateStr;
      } catch {
        return dateStr;
      }
    };
    
    return (
      <div className="space-y-1">
        {sortedData.map((day, index) => (
          <div key={index} className="flex items-center gap-2 text-xs">
            <span className="w-16 text-slate-500 dark:text-slate-400 flex-shrink-0">
              {formatDateLabel(day.date)}
            </span>
            <div className="flex-1 flex gap-1">
              <div 
                className="h-4 bg-blue-500 rounded-sm"
                style={{ width: `${(day.messagesSent / maxValue) * 50}%`, minWidth: day.messagesSent > 0 ? '2px' : '0' }}
                title={`Sent: ${formatNumber(day.messagesSent)}`}
              />
              <div 
                className="h-4 bg-green-500 rounded-sm"
                style={{ width: `${(day.messagesReceived / maxValue) * 50}%`, minWidth: day.messagesReceived > 0 ? '2px' : '0' }}
                title={`Received: ${formatNumber(day.messagesReceived)}`}
              />
            </div>
            <span className="w-20 text-right text-slate-600 dark:text-slate-300">
              {formatNumber(day.messagesSent + day.messagesReceived)}
            </span>
          </div>
        ))}
        <div className="flex items-center gap-4 mt-2 text-xs text-slate-500">
          <span className="flex items-center gap-1"><span className="w-3 h-3 bg-blue-500 rounded-sm"></span> Sent</span>
          <span className="flex items-center gap-1"><span className="w-3 h-3 bg-green-500 rounded-sm"></span> Received</span>
        </div>
      </div>
    );
  };

  return (
    <div className="p-4 space-y-4 w-full max-w-full overflow-hidden">
      {/* Header */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
        <div>
          <h1 className="text-xl font-semibold text-slate-900 dark:text-white">Exchange Online</h1>
          <p className="text-sm text-slate-500 dark:text-slate-400 hidden sm:block">
            Email traffic and storage
          </p>
        </div>
        <div className="flex items-center gap-2">
          <div className="flex rounded-lg border border-slate-200 dark:border-slate-700 overflow-hidden">
            <button
              onClick={() => setViewMode('traffic')}
              className={`px-3 py-1.5 text-sm ${viewMode === 'traffic' ? 'bg-blue-600 text-white' : 'bg-white dark:bg-slate-800 text-slate-600 dark:text-slate-300'}`}
            >
              Traffic
            </button>
            <button
              onClick={() => setViewMode('storage')}
              className={`px-3 py-1.5 text-sm ${viewMode === 'storage' ? 'bg-blue-600 text-white' : 'bg-white dark:bg-slate-800 text-slate-600 dark:text-slate-300'}`}
            >
              Storage
            </button>
            <button
              onClick={() => setViewMode('forwarding')}
              className={`px-3 py-1.5 text-sm ${viewMode === 'forwarding' ? 'bg-blue-600 text-white' : 'bg-white dark:bg-slate-800 text-slate-600 dark:text-slate-300'}`}
            >
              Rules Fwd
            </button>
            <button
              onClick={() => setViewMode('mailboxFwd')}
              className={`px-3 py-1.5 text-sm ${viewMode === 'mailboxFwd' ? 'bg-blue-600 text-white' : 'bg-white dark:bg-slate-800 text-slate-600 dark:text-slate-300'}`}
            >
              Mailbox Fwd
            </button>
            <button
              onClick={() => setViewMode('access')}
              className={`px-3 py-1.5 text-sm ${viewMode === 'access' ? 'bg-blue-600 text-white' : 'bg-white dark:bg-slate-800 text-slate-600 dark:text-slate-300'}`}
            >
              Access
            </button>
          </div>
          <a
            href="https://admin.exchange.microsoft.com"
            target="_blank"
            rel="noopener noreferrer"
            className="inline-flex items-center justify-center gap-2 px-3 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 transition-colors text-sm whitespace-nowrap"
          >
            <OpenRegular className="w-4 h-4" />
            <span>Exchange Admin Center</span>
          </a>
        </div>
      </div>

      {/* Period Filter and Stats Cards */}
      <div className="flex items-center gap-2 mb-2">
        <span className="text-sm text-slate-500 dark:text-slate-400">Period:</span>
        <select
          value={period}
          onChange={(e) => setPeriod(Number(e.target.value))}
          className="px-2 py-1 text-sm border border-slate-200 dark:border-slate-600 rounded bg-white dark:bg-slate-700 text-slate-900 dark:text-white"
        >
          <option value={7}>Last 7 days</option>
          <option value={14}>Last 14 days</option>
        </select>
      </div>
      <div className="grid grid-cols-2 sm:grid-cols-3 gap-2">
        <StatCard
          title="Messages Sent"
          value={summary?.totalMessagesSent ?? 0}
          icon={<SendRegular className="w-4 h-4 text-blue-600" />}
          color="text-blue-600"
        />
        <StatCard
          title="Messages Received"
          value={summary?.totalMessagesReceived ?? 0}
          icon={<MailInboxRegular className="w-4 h-4 text-green-600" />}
          color="text-green-600"
        />
        <StatCard
          title="Avg/Day"
          value={summary?.averageMessagesPerDay ?? 0}
          icon={<ArrowTrendingRegular className="w-4 h-4 text-purple-600" />}
          color="text-purple-600"
        />
      </div>

      {viewMode === 'traffic' ? (
        /* Traffic View */
        <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
          {/* Daily Traffic Chart */}
          <div className="bg-white dark:bg-slate-800 rounded-lg border border-slate-200 dark:border-slate-700 p-4">
            <h2 className="text-lg font-semibold text-slate-900 dark:text-white mb-4">Daily Traffic</h2>
            {summaryLoading ? (
              <div className="flex items-center justify-center h-48">
                <div className="animate-spin rounded-full h-8 w-8 border-b-2 border-blue-600"></div>
              </div>
            ) : (
              <TrafficChart data={summary?.dailyTraffic ?? []} />
            )}
          </div>

          {/* Top Senders & Recipients */}
          <div className="space-y-4">
            {/* Top Senders */}
            <div className="bg-white dark:bg-slate-800 rounded-lg border border-slate-200 dark:border-slate-700 p-4">
              <h2 className="text-lg font-semibold text-slate-900 dark:text-white mb-3">Top Senders</h2>
              {summaryLoading ? (
                <div className="flex items-center justify-center h-32">
                  <div className="animate-spin rounded-full h-6 w-6 border-b-2 border-blue-600"></div>
                </div>
              ) : summary?.topSenders && summary.topSenders.length > 0 ? (
                <div className="space-y-2">
                  {summary.topSenders.slice(0, 5).map((sender, index) => (
                    <div key={index} className="flex items-center justify-between text-sm">
                      <div className="flex items-center gap-2 min-w-0">
                        <span className="w-5 h-5 rounded-full bg-blue-100 dark:bg-blue-900/30 text-blue-600 dark:text-blue-400 flex items-center justify-center text-xs font-medium">
                          {index + 1}
                        </span>
                        <span className="truncate text-slate-700 dark:text-slate-300">
                          {sender.displayName || sender.userPrincipalName}
                        </span>
                      </div>
                      <span className="font-medium text-slate-900 dark:text-white ml-2">
                        {formatNumber(sender.messageCount)}
                      </span>
                    </div>
                  ))}
                </div>
              ) : (
                <p className="text-sm text-slate-500">No data available</p>
              )}
            </div>

            {/* Top Recipients */}
            <div className="bg-white dark:bg-slate-800 rounded-lg border border-slate-200 dark:border-slate-700 p-4">
              <h2 className="text-lg font-semibold text-slate-900 dark:text-white mb-3">Top Recipients</h2>
              {summaryLoading ? (
                <div className="flex items-center justify-center h-32">
                  <div className="animate-spin rounded-full h-6 w-6 border-b-2 border-blue-600"></div>
                </div>
              ) : summary?.topRecipients && summary.topRecipients.length > 0 ? (
                <div className="space-y-2">
                  {summary.topRecipients.slice(0, 5).map((recipient, index) => (
                    <div key={index} className="flex items-center justify-between text-sm">
                      <div className="flex items-center gap-2 min-w-0">
                        <span className="w-5 h-5 rounded-full bg-green-100 dark:bg-green-900/30 text-green-600 dark:text-green-400 flex items-center justify-center text-xs font-medium">
                          {index + 1}
                        </span>
                        <span className="truncate text-slate-700 dark:text-slate-300">
                          {recipient.displayName || recipient.userPrincipalName}
                        </span>
                      </div>
                      <span className="font-medium text-slate-900 dark:text-white ml-2">
                        {formatNumber(recipient.messageCount)}
                      </span>
                    </div>
                  ))}
                </div>
              ) : (
                <p className="text-sm text-slate-500">No data available</p>
              )}
            </div>
          </div>
        </div>
      ) : viewMode === 'storage' ? (
        /* Storage View */
        <div className="space-y-4">
          {/* Storage Summary Cards */}
          <div className="grid grid-cols-2 sm:grid-cols-5 gap-2">
            <StatCard
              title="Total Storage Used"
              value={`${storageData?.summary?.totalStorageUsedGB?.toFixed(1) ?? 0} GB`}
              icon={<StorageRegular className="w-4 h-4 text-blue-600" />}
              color="text-blue-600"
            />
            <StatCard
              title="Critical (>90%)"
              value={storageData?.summary?.criticalCount ?? 0}
              icon={<ErrorCircleRegular className="w-4 h-4 text-red-600" />}
              color="text-red-600"
            />
            <StatCard
              title="Warning (>80%)"
              value={storageData?.summary?.warningCount ?? 0}
              icon={<WarningRegular className="w-4 h-4 text-amber-600" />}
              color="text-amber-600"
            />
            <StatCard
              title="Archive Enabled"
              value={storageData?.summary?.archiveEnabledCount ?? 0}
              icon={<ArchiveRegular className="w-4 h-4 text-indigo-600" />}
              color="text-indigo-600"
            />
            <StatCard
              title="Avg Usage"
              value={`${storageData?.summary?.averagePercentUsed?.toFixed(1) ?? 0}%`}
              icon={<ArrowTrendingRegular className="w-4 h-4 text-purple-600" />}
              color="text-purple-600"
            />
          </div>

          {/* Mailboxes Near Quota */}
          <div className="bg-white dark:bg-slate-800 rounded-lg border border-slate-200 dark:border-slate-700 overflow-hidden">
            <div className="p-4 border-b border-slate-200 dark:border-slate-700">
              <h2 className="text-lg font-semibold text-slate-900 dark:text-white">Mailboxes Near Quota</h2>
              <p className="text-sm text-slate-500 dark:text-slate-400">Mailboxes using 80% or more of their quota</p>
            </div>
            
            {storageLoading ? (
              <div className="flex items-center justify-center h-64">
                <div className="animate-spin rounded-full h-8 w-8 border-b-2 border-blue-600"></div>
              </div>
            ) : storageData?.error ? (
              <div className="flex items-center justify-center h-64 text-red-500 p-4">
                <p>{storageData.error}</p>
              </div>
            ) : storageData?.nearQuotaMailboxes && storageData.nearQuotaMailboxes.length > 0 ? (
              <div className="overflow-x-auto">
                <table className="w-full">
                  <thead className="bg-slate-50 dark:bg-slate-700">
                    <tr>
                      <th className="px-4 py-3 text-left text-sm font-medium text-slate-600 dark:text-slate-300">User</th>
                      <th className="px-4 py-3 text-left text-sm font-medium text-slate-600 dark:text-slate-300">Size (Used / Total)</th>
                      <th className="px-4 py-3 text-left text-sm font-medium text-slate-600 dark:text-slate-300">Usage</th>
                      <th className="px-4 py-3 text-center text-sm font-medium text-slate-600 dark:text-slate-300">Archive</th>
                      <th className="px-4 py-3 text-left text-sm font-medium text-slate-600 dark:text-slate-300">Status</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-200 dark:divide-slate-700">
                    {storageData.nearQuotaMailboxes.map((mailbox, index) => (
                      <tr key={index} className="hover:bg-slate-50 dark:hover:bg-slate-700/50">
                        <td className="px-4 py-3">
                          <div className="flex items-center gap-3">
                            <div className={`w-8 h-8 rounded-full flex items-center justify-center flex-shrink-0 ${
                              mailbox.status === 'Critical' || mailbox.status === 'Full'
                                ? 'bg-red-100 dark:bg-red-900/30'
                                : 'bg-amber-100 dark:bg-amber-900/30'
                            }`}>
                              {mailbox.status === 'Critical' || mailbox.status === 'Full' ? (
                                <ErrorCircleRegular className="w-4 h-4 text-red-600 dark:text-red-400" />
                              ) : (
                                <WarningRegular className="w-4 h-4 text-amber-600 dark:text-amber-400" />
                              )}
                            </div>
                            <div className="min-w-0">
                              <p className="font-medium text-slate-900 dark:text-white truncate">
                                {mailbox.displayName}
                              </p>
                              <p className="text-xs text-slate-500 dark:text-slate-400 truncate">
                                {mailbox.userPrincipalName}
                              </p>
                            </div>
                          </div>
                        </td>
                        <td className="px-4 py-3">
                          <span className="text-sm text-slate-600 dark:text-slate-300">
                            {mailbox.storageUsedGB.toFixed(2)} GB / {mailbox.prohibitSendQuotaGB > 0 ? `${mailbox.prohibitSendQuotaGB.toFixed(0)} GB` : 'N/A'}
                          </span>
                        </td>
                        <td className="px-4 py-3">
                          <div className="flex items-center gap-2">
                            <div className="flex-1 h-2 bg-slate-200 dark:bg-slate-600 rounded-full overflow-hidden max-w-[100px]">
                              <div
                                className={`h-full rounded-full ${
                                  mailbox.percentUsed >= 90 ? 'bg-red-500' : 'bg-amber-500'
                                }`}
                                style={{ width: `${Math.min(mailbox.percentUsed, 100)}%` }}
                              />
                            </div>
                            <span className={`text-sm font-medium ${
                              mailbox.percentUsed >= 90 ? 'text-red-600 dark:text-red-400' : 'text-amber-600 dark:text-amber-400'
                            }`}>
                              {mailbox.percentUsed.toFixed(1)}%
                            </span>
                          </div>
                        </td>
                        <td className="px-4 py-3 text-center">
                          {mailbox.hasArchive ? (
                            <span className="inline-flex items-center gap-1 px-2 py-1 rounded-full text-xs font-medium bg-blue-100 text-blue-700 dark:bg-blue-900/30 dark:text-blue-400">
                              <ArchiveRegular className="w-3 h-3" />
                              Yes
                            </span>
                          ) : (
                            <span className="text-xs text-slate-400">No</span>
                          )}
                        </td>
                        <td className="px-4 py-3">
                          <span className={`inline-flex items-center px-2 py-1 rounded-full text-xs font-medium ${
                            mailbox.status === 'Full'
                              ? 'bg-red-100 text-red-700 dark:bg-red-900/30 dark:text-red-400'
                              : mailbox.status === 'Critical'
                              ? 'bg-red-100 text-red-700 dark:bg-red-900/30 dark:text-red-400'
                              : 'bg-amber-100 text-amber-700 dark:bg-amber-900/30 dark:text-amber-400'
                          }`}>
                            {mailbox.status}
                          </span>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            ) : (
              <div className="flex flex-col items-center justify-center h-64 text-slate-500">
                <CheckmarkCircleFilled className="w-12 h-12 text-green-500 mb-2" />
                <p className="text-lg font-medium text-slate-900 dark:text-white">All mailboxes healthy</p>
                <p className="text-sm">No mailboxes are using more than 80% of their quota</p>
              </div>
            )}
          </div>

          {/* All Mailboxes Storage Chart */}
          {storageData?.allMailboxes && storageData.allMailboxes.length > 0 && (
            <div className="bg-white dark:bg-slate-800 rounded-lg border border-slate-200 dark:border-slate-700 p-4">
              <h2 className="text-lg font-semibold text-slate-900 dark:text-white mb-4">Top Mailboxes by Storage</h2>
              <div className="space-y-2">
                {storageData.allMailboxes.slice(0, 10).map((mailbox, index) => (
                  <div key={index} className="flex items-center gap-3">
                    <span className="w-5 h-5 rounded-full bg-slate-100 dark:bg-slate-700 text-slate-600 dark:text-slate-400 flex items-center justify-center text-xs font-medium flex-shrink-0">
                      {index + 1}
                    </span>
                    <div className="flex-1 min-w-0">
                      <div className="flex items-center justify-between mb-1">
                        <div className="flex items-center gap-2 min-w-0">
                          <span className="text-sm font-medium text-slate-700 dark:text-slate-300 truncate">
                            {mailbox.displayName}
                          </span>
                          {mailbox.hasArchive && (
                            <ArchiveRegular className="w-3.5 h-3.5 text-blue-500 flex-shrink-0" title="Archive enabled" />
                          )}
                        </div>
                        <span className="text-xs text-slate-500 ml-2 flex-shrink-0">
                          {mailbox.storageUsedGB.toFixed(2)} / {mailbox.prohibitSendQuotaGB > 0 ? `${mailbox.prohibitSendQuotaGB.toFixed(0)} GB` : 'N/A'}
                        </span>
                      </div>
                      <div className="h-2 bg-slate-200 dark:bg-slate-600 rounded-full overflow-hidden">
                        <div
                          className={`h-full rounded-full ${
                            mailbox.percentUsed >= 90
                              ? 'bg-red-500'
                              : mailbox.percentUsed >= 80
                              ? 'bg-amber-500'
                              : 'bg-blue-500'
                          }`}
                          style={{ width: `${Math.min(mailbox.percentUsed, 100)}%` }}
                        />
                      </div>
                    </div>
                    <span className={`text-xs font-medium w-12 text-right ${
                      mailbox.percentUsed >= 90
                        ? 'text-red-600'
                        : mailbox.percentUsed >= 80
                        ? 'text-amber-600'
                        : 'text-slate-600 dark:text-slate-400'
                    }`}>
                      {mailbox.percentUsed.toFixed(0)}%
                    </span>
                  </div>
                ))}
              </div>
            </div>
          )}

          {/* Last Updated */}
          {storageData?.lastUpdated && (
            <div className="flex items-center justify-center gap-2 text-xs text-slate-400 dark:text-slate-500 pt-2">
              <ClockRegular className="w-3.5 h-3.5" />
              <span>Last updated: {new Date(storageData.lastUpdated).toLocaleString()}</span>
            </div>
          )}
        </div>
      ) : viewMode === 'forwarding' ? (
        /* Forwarding View */
        <div className="space-y-4">
          {/* Forwarding Summary Cards */}
          <div className="grid grid-cols-2 sm:grid-cols-5 gap-2">
            <StatCard
              title="Mailboxes Scanned"
              value={forwardingData?.totalMailboxesScanned ?? 0}
              icon={<MailInboxRegular className="w-4 h-4 text-blue-600" />}
              color="text-blue-600"
            />
            <StatCard
              title="With Forwarding"
              value={forwardingData?.mailboxesWithForwarding ?? 0}
              icon={<SendRegular className="w-4 h-4 text-purple-600" />}
              color="text-purple-600"
            />
            <StatCard
              title="External Forwarding"
              value={forwardingData?.externalForwardingCount ?? 0}
              icon={<GlobeRegular className="w-4 h-4 text-red-600" />}
              color="text-red-600"
            />
            <StatCard
              title="Internal Forwarding"
              value={forwardingData?.internalForwardingCount ?? 0}
              icon={<BuildingRegular className="w-4 h-4 text-green-600" />}
              color="text-green-600"
            />
            <StatCard
              title="% With Forwarding"
              value={forwardingData?.totalMailboxesScanned ? `${((forwardingData.mailboxesWithForwarding / forwardingData.totalMailboxesScanned) * 100).toFixed(1)}%` : '0%'}
              icon={<ArrowTrendingRegular className="w-4 h-4 text-amber-600" />}
              color="text-amber-600"
            />
          </div>

          {/* Forwarding Table */}
          <div className="bg-white dark:bg-slate-800 rounded-lg border border-slate-200 dark:border-slate-700 overflow-hidden">
            <div className="p-4 border-b border-slate-200 dark:border-slate-700 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
              <div>
                <h2 className="text-lg font-semibold text-slate-900 dark:text-white">Inbox Rules Forwarding</h2>
                <p className="text-sm text-slate-500 dark:text-slate-400">Inbox rules configured to forward or redirect email</p>
              </div>
              <div className="flex items-center gap-2">
                {/* Filter */}
                <div className="flex items-center gap-2">
                  <FilterRegular className="w-4 h-4 text-slate-400" />
                  <select
                    value={forwardingFilter}
                    onChange={(e) => setForwardingFilter(e.target.value as ForwardingFilter)}
                    className="px-2 py-1 text-sm border border-slate-200 dark:border-slate-600 rounded bg-white dark:bg-slate-700 text-slate-900 dark:text-white"
                  >
                    <option value="all">All ({forwardingData?.totalForwardingRules ?? 0})</option>
                    <option value="external">External ({forwardingData?.externalForwardingCount ?? 0})</option>
                    <option value="internal">Internal ({forwardingData?.internalForwardingCount ?? 0})</option>
                  </select>
                </div>
                {/* Export */}
                <button
                  onClick={exportForwardingToCsv}
                  disabled={!forwardingData?.forwardingRules?.length}
                  className="inline-flex items-center gap-1.5 px-3 py-1.5 text-sm bg-slate-100 dark:bg-slate-700 text-slate-700 dark:text-slate-300 rounded hover:bg-slate-200 dark:hover:bg-slate-600 disabled:opacity-50 disabled:cursor-not-allowed"
                >
                  <ArrowDownloadRegular className="w-4 h-4" />
                  Export CSV
                </button>
                {/* Refresh */}
                <button
                  onClick={fetchForwardingReport}
                  disabled={forwardingLoading}
                  className="inline-flex items-center gap-1.5 px-3 py-1.5 text-sm bg-blue-600 text-white rounded hover:bg-blue-700 disabled:opacity-50"
                >
                  {forwardingLoading ? (
                    <div className="animate-spin rounded-full h-4 w-4 border-b-2 border-white"></div>
                  ) : (
                    <ArrowForwardRegular className="w-4 h-4" />
                  )}
                  {forwardingLoading ? 'Scanning...' : 'Refresh'}
                </button>
              </div>
            </div>

            {forwardingLoading ? (
              <div className="flex flex-col items-center justify-center h-64 gap-3">
                <div className="animate-spin rounded-full h-8 w-8 border-b-2 border-blue-600"></div>
                <p className="text-sm text-slate-500 dark:text-slate-400">Scanning mailboxes for forwarding rules...</p>
                <p className="text-xs text-slate-400 dark:text-slate-500">This may take a minute for large tenants</p>
              </div>
            ) : forwardingData?.error ? (
              <div className="flex items-center justify-center h-64 text-red-500 p-4">
                <p>{forwardingData.error}</p>
              </div>
            ) : getFilteredForwardingData().length > 0 ? (
              <div className="overflow-x-auto">
                <table className="w-full">
                  <thead className="bg-slate-50 dark:bg-slate-700">
                    <tr>
                      <th className="px-4 py-3 text-left text-sm font-medium text-slate-600 dark:text-slate-300">User</th>
                      <th className="px-4 py-3 text-left text-sm font-medium text-slate-600 dark:text-slate-300">Forwarding Target</th>
                      <th className="px-4 py-3 text-left text-sm font-medium text-slate-600 dark:text-slate-300">Type</th>
                      <th className="px-4 py-3 text-left text-sm font-medium text-slate-600 dark:text-slate-300">Rule</th>
                      <th className="px-4 py-3 text-center text-sm font-medium text-slate-600 dark:text-slate-300">Keeps Copy</th>
                      <th className="px-4 py-3 text-left text-sm font-medium text-slate-600 dark:text-slate-300">Risk</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-200 dark:divide-slate-700">
                    {getFilteredForwardingData().map((mailbox, index) => (
                      <tr key={`${mailbox.userId}-${mailbox.ruleName}-${index}`} className="hover:bg-slate-50 dark:hover:bg-slate-700/50">
                        <td className="px-4 py-3">
                          <div className="flex items-center gap-3">
                            <div className={`w-8 h-8 rounded-full flex items-center justify-center flex-shrink-0 ${
                              mailbox.isExternal
                                ? 'bg-red-100 dark:bg-red-900/30'
                                : 'bg-green-100 dark:bg-green-900/30'
                            }`}>
                              {mailbox.isExternal ? (
                                <GlobeRegular className="w-4 h-4 text-red-600 dark:text-red-400" />
                              ) : (
                                <BuildingRegular className="w-4 h-4 text-green-600 dark:text-green-400" />
                              )}
                            </div>
                            <div className="min-w-0">
                              <p className="font-medium text-slate-900 dark:text-white truncate">
                                {mailbox.displayName}
                              </p>
                              <p className="text-xs text-slate-500 dark:text-slate-400 truncate">
                                {mailbox.mail}
                              </p>
                            </div>
                          </div>
                        </td>
                        <td className="px-4 py-3">
                          <div className="min-w-0">
                            <p className="text-sm text-slate-900 dark:text-white truncate" title={mailbox.forwardingTarget}>
                              {mailbox.forwardingTarget}
                            </p>
                            <p className="text-xs text-slate-500 dark:text-slate-400">
                              {mailbox.isExternal ? 'External domain' : 'Internal domain'}
                            </p>
                          </div>
                        </td>
                        <td className="px-4 py-3">
                          <span className={`inline-flex items-center gap-1 px-2 py-1 rounded-full text-xs font-medium ${
                            mailbox.forwardingType === 'Redirect'
                              ? 'bg-amber-100 text-amber-700 dark:bg-amber-900/30 dark:text-amber-400'
                              : mailbox.forwardingType === 'ForwardAsAttachment'
                              ? 'bg-purple-100 text-purple-700 dark:bg-purple-900/30 dark:text-purple-400'
                              : 'bg-blue-100 text-blue-700 dark:bg-blue-900/30 dark:text-blue-400'
                          }`}>
                            <ArrowForwardRegular className="w-3 h-3" />
                            {mailbox.forwardingType}
                          </span>
                        </td>
                        <td className="px-4 py-3">
                          <div className="flex items-center gap-2">
                            <span className="text-sm text-slate-700 dark:text-slate-300 truncate max-w-[150px]" title={mailbox.ruleName}>
                              {mailbox.ruleName}
                            </span>
                            {!mailbox.ruleEnabled && (
                              <span className="text-xs text-slate-400">(Disabled)</span>
                            )}
                          </div>
                        </td>
                        <td className="px-4 py-3 text-center">
                          {mailbox.deliverToMailbox ? (
                            <span className="inline-flex items-center gap-1 px-2 py-1 rounded-full text-xs font-medium bg-green-100 text-green-700 dark:bg-green-900/30 dark:text-green-400">
                              <CheckmarkCircleFilled className="w-3 h-3" />
                              Yes
                            </span>
                          ) : (
                            <span className="inline-flex items-center gap-1 px-2 py-1 rounded-full text-xs font-medium bg-red-100 text-red-700 dark:bg-red-900/30 dark:text-red-400">
                              <ErrorCircleRegular className="w-3 h-3" />
                              No
                            </span>
                          )}
                        </td>
                        <td className="px-4 py-3">
                          <span className={`inline-flex items-center gap-1 px-2 py-1 rounded-full text-xs font-medium ${
                            mailbox.riskLevel === 'High'
                              ? 'bg-red-100 text-red-700 dark:bg-red-900/30 dark:text-red-400'
                              : 'bg-green-100 text-green-700 dark:bg-green-900/30 dark:text-green-400'
                          }`}>
                            {mailbox.riskLevel === 'High' && <WarningRegular className="w-3 h-3" />}
                            {mailbox.riskLevel}
                          </span>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            ) : (
              <div className="flex flex-col items-center justify-center h-64 text-slate-500">
                <CheckmarkCircleFilled className="w-12 h-12 text-green-500 mb-2" />
                <p className="text-lg font-medium text-slate-900 dark:text-white">No forwarding rules found</p>
                <p className="text-sm text-center max-w-md">
                  {forwardingFilter !== 'all' 
                    ? `No ${forwardingFilter} forwarding rules detected. Try changing the filter.`
                    : 'No inbox rules with forwarding actions were detected in scanned mailboxes.'
                  }
                </p>
              </div>
            )}
          </div>

          {/* Tenant Domains */}
          {forwardingData?.tenantDomains && forwardingData.tenantDomains.length > 0 && (
            <div className="bg-white dark:bg-slate-800 rounded-lg border border-slate-200 dark:border-slate-700 p-4">
              <h3 className="text-sm font-semibold text-slate-900 dark:text-white mb-2">Recognized Tenant Domains</h3>
              <div className="flex flex-wrap gap-2">
                {forwardingData.tenantDomains.map((domain, index) => (
                  <span 
                    key={index}
                    className="inline-flex items-center gap-1 px-2 py-1 rounded-full text-xs font-medium bg-slate-100 dark:bg-slate-700 text-slate-600 dark:text-slate-300"
                  >
                    <BuildingRegular className="w-3 h-3" />
                    {domain}
                  </span>
                ))}
              </div>
              <p className="text-xs text-slate-400 dark:text-slate-500 mt-2">
                Forwarding to these domains is classified as internal. All other domains are classified as external.
              </p>
            </div>
          )}

          {/* Last Updated */}
          {forwardingData?.lastUpdated && (
            <div className="flex items-center justify-center gap-2 text-xs text-slate-400 dark:text-slate-500 pt-2">
              <ClockRegular className="w-3.5 h-3.5" />
              <span>Last updated: {new Date(forwardingData.lastUpdated).toLocaleString()}</span>
            </div>
          )}
        </div>
      ) : viewMode === 'mailboxFwd' ? (
        /* Exchange Mailbox Forwarding View */
        <div className="space-y-4">
          {/* Summary Cards */}
          <div className="grid grid-cols-2 sm:grid-cols-4 gap-2">
            <StatCard
              title="With Forwarding"
              value={exchangeForwardingData?.totalCount ?? 0}
              icon={<SendRegular className="w-4 h-4 text-purple-600" />}
              color="text-purple-600"
            />
            <StatCard
              title="External"
              value={exchangeForwardingData?.externalCount ?? 0}
              icon={<GlobeRegular className="w-4 h-4 text-red-600" />}
              color="text-red-600"
            />
            <StatCard
              title="Internal"
              value={exchangeForwardingData?.internalCount ?? 0}
              icon={<BuildingRegular className="w-4 h-4 text-green-600" />}
              color="text-green-600"
            />
            <StatCard
              title="Keeps Copy"
              value={exchangeForwardingData?.mailboxes?.filter(m => m.deliverToMailboxAndForward).length ?? 0}
              icon={<ArchiveRegular className="w-4 h-4 text-blue-600" />}
              color="text-blue-600"
            />
          </div>

          {/* Info Banner */}
          <div className="bg-amber-50 dark:bg-amber-900/20 border border-amber-200 dark:border-amber-800 rounded-lg p-3 flex items-start gap-3">
            <InfoRegular className="w-5 h-5 text-amber-600 dark:text-amber-400 flex-shrink-0 mt-0.5" />
            <div className="text-sm text-amber-700 dark:text-amber-300">
              <p>This shows <strong>Exchange mailbox-level forwarding</strong> configured via ForwardingAddress or ForwardingSmtpAddress properties. This is different from inbox rules which are shown in the "Rules Fwd" tab.</p>
            </div>
          </div>

          {/* Forwarding Table */}
          <div className="bg-white dark:bg-slate-800 rounded-lg border border-slate-200 dark:border-slate-700 overflow-hidden">
            <div className="p-4 border-b border-slate-200 dark:border-slate-700 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
              <div>
                <h2 className="text-lg font-semibold text-slate-900 dark:text-white">Exchange Mailbox Forwarding</h2>
                <p className="text-sm text-slate-500 dark:text-slate-400">Mailboxes with forwarding configured at the Exchange level</p>
              </div>
              <div className="flex items-center gap-2">
                {/* Filter */}
                <div className="flex items-center gap-2">
                  <FilterRegular className="w-4 h-4 text-slate-400" />
                  <select
                    value={exchangeForwardingFilter}
                    onChange={(e) => setExchangeForwardingFilter(e.target.value as ForwardingFilter)}
                    className="px-2 py-1 text-sm border border-slate-200 dark:border-slate-600 rounded bg-white dark:bg-slate-700 text-slate-900 dark:text-white"
                  >
                    <option value="all">All ({exchangeForwardingData?.totalCount ?? 0})</option>
                    <option value="external">External ({exchangeForwardingData?.externalCount ?? 0})</option>
                    <option value="internal">Internal ({exchangeForwardingData?.internalCount ?? 0})</option>
                  </select>
                </div>
                {/* Export */}
                <button
                  onClick={exportExchangeForwardingToCsv}
                  disabled={!exchangeForwardingData?.mailboxes?.length}
                  className="inline-flex items-center gap-1.5 px-3 py-1.5 text-sm bg-slate-100 dark:bg-slate-700 text-slate-700 dark:text-slate-300 rounded hover:bg-slate-200 dark:hover:bg-slate-600 disabled:opacity-50 disabled:cursor-not-allowed"
                >
                  <ArrowDownloadRegular className="w-4 h-4" />
                  Export CSV
                </button>
                {/* Refresh */}
                <button
                  onClick={fetchExchangeForwarding}
                  disabled={exchangeForwardingLoading}
                  className="inline-flex items-center gap-1.5 px-3 py-1.5 text-sm bg-blue-600 text-white rounded hover:bg-blue-700 disabled:opacity-50"
                >
                  {exchangeForwardingLoading ? (
                    <div className="animate-spin rounded-full h-4 w-4 border-b-2 border-white"></div>
                  ) : (
                    <ArrowForwardRegular className="w-4 h-4" />
                  )}
                  {exchangeForwardingLoading ? 'Loading...' : 'Refresh'}
                </button>
              </div>
            </div>

            {exchangeForwardingLoading ? (
              <div className="flex flex-col items-center justify-center h-64 gap-3">
                <div className="animate-spin rounded-full h-8 w-8 border-b-2 border-blue-600"></div>
                <p className="text-sm text-slate-500 dark:text-slate-400">Fetching mailbox forwarding configuration...</p>
              </div>
            ) : getFilteredExchangeForwardingData().length > 0 ? (
              <div className="overflow-x-auto">
                <table className="w-full">
                  <thead className="bg-slate-50 dark:bg-slate-700">
                    <tr>
                      <th className="px-4 py-3 text-left text-sm font-medium text-slate-600 dark:text-slate-300">User</th>
                      <th className="px-4 py-3 text-left text-sm font-medium text-slate-600 dark:text-slate-300">Forwarding Target</th>
                      <th className="px-4 py-3 text-center text-sm font-medium text-slate-600 dark:text-slate-300">Scope</th>
                      <th className="px-4 py-3 text-center text-sm font-medium text-slate-600 dark:text-slate-300">Keeps Copy</th>
                      <th className="px-4 py-3 text-left text-sm font-medium text-slate-600 dark:text-slate-300">Type</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-200 dark:divide-slate-700">
                    {getFilteredExchangeForwardingData().map((mailbox) => (
                      <tr key={mailbox.id} className="hover:bg-slate-50 dark:hover:bg-slate-700/50">
                        <td className="px-4 py-3">
                          <div className="flex items-center gap-3">
                            <div className={`w-8 h-8 rounded-full flex items-center justify-center flex-shrink-0 ${
                              mailbox.isExternal
                                ? 'bg-red-100 dark:bg-red-900/30'
                                : 'bg-green-100 dark:bg-green-900/30'
                            }`}>
                              {mailbox.isExternal ? (
                                <GlobeRegular className="w-4 h-4 text-red-600 dark:text-red-400" />
                              ) : (
                                <BuildingRegular className="w-4 h-4 text-green-600 dark:text-green-400" />
                              )}
                            </div>
                            <div className="min-w-0">
                              <p className="font-medium text-slate-900 dark:text-white truncate">
                                {mailbox.displayName}
                              </p>
                              <p className="text-xs text-slate-500 dark:text-slate-400 truncate">
                                {mailbox.primarySmtpAddress}
                              </p>
                            </div>
                          </div>
                        </td>
                        <td className="px-4 py-3">
                          <div className="min-w-0">
                            <p className="text-sm text-slate-900 dark:text-white truncate" title={mailbox.forwardingTarget}>
                              {mailbox.forwardingTarget}
                            </p>
                            <p className="text-xs text-slate-500 dark:text-slate-400">
                              {mailbox.forwardingTargetDomain}
                            </p>
                          </div>
                        </td>
                        <td className="px-4 py-3 text-center">
                          <span className={`inline-flex items-center gap-1 px-2 py-1 rounded-full text-xs font-medium ${
                            mailbox.isExternal
                              ? 'bg-red-100 text-red-700 dark:bg-red-900/30 dark:text-red-400'
                              : 'bg-green-100 text-green-700 dark:bg-green-900/30 dark:text-green-400'
                          }`}>
                            {mailbox.isExternal ? (
                              <><GlobeRegular className="w-3 h-3" /> External</>
                            ) : (
                              <><BuildingRegular className="w-3 h-3" /> Internal</>
                            )}
                          </span>
                        </td>
                        <td className="px-4 py-3 text-center">
                          {mailbox.deliverToMailboxAndForward ? (
                            <span className="inline-flex items-center gap-1 px-2 py-1 rounded-full text-xs font-medium bg-green-100 text-green-700 dark:bg-green-900/30 dark:text-green-400">
                              <CheckmarkCircleFilled className="w-3 h-3" />
                              Yes
                            </span>
                          ) : (
                            <span className="inline-flex items-center gap-1 px-2 py-1 rounded-full text-xs font-medium bg-red-100 text-red-700 dark:bg-red-900/30 dark:text-red-400">
                              <ErrorCircleRegular className="w-3 h-3" />
                              No
                            </span>
                          )}
                        </td>
                        <td className="px-4 py-3">
                          <span className="inline-flex items-center px-2 py-1 rounded-full text-xs font-medium bg-slate-100 text-slate-700 dark:bg-slate-700 dark:text-slate-300">
                            {mailbox.recipientTypeDetails}
                          </span>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            ) : (
              <div className="flex flex-col items-center justify-center h-64 text-slate-500">
                <CheckmarkCircleFilled className="w-12 h-12 text-green-500 mb-2" />
                <p className="text-lg font-medium text-slate-900 dark:text-white">No mailbox forwarding configured</p>
                <p className="text-sm text-center max-w-md">
                  {exchangeForwardingFilter !== 'all' 
                    ? `No ${exchangeForwardingFilter} mailbox forwarding detected. Try changing the filter.`
                    : 'No mailboxes have Exchange-level forwarding configured.'
                  }
                </p>
              </div>
            )}
          </div>

          {/* Last Updated */}
          {exchangeForwardingData?.lastUpdated && (
            <div className="flex items-center justify-center gap-2 text-xs text-slate-400 dark:text-slate-500 pt-2">
              <ClockRegular className="w-3.5 h-3.5" />
              <span>Last updated: {new Date(exchangeForwardingData.lastUpdated).toLocaleString()}</span>
            </div>
          )}
        </div>
      ) : viewMode === 'access' ? (
        /* Mailbox Access Lookup */
        <div className="space-y-4">
          {/* Search card */}
          <div className="bg-white dark:bg-slate-800 rounded-lg border border-slate-200 dark:border-slate-700 p-5 space-y-4">
            <div>
              <h2 className="text-lg font-semibold text-slate-900 dark:text-white">Mailbox Access Lookup</h2>
              <p className="text-sm text-slate-500 dark:text-slate-400 mt-0.5">Find which mailboxes a user can access, or who has access to a specific mailbox.</p>
            </div>

            {/* Mode toggle */}
            <div className="flex items-center gap-3 flex-wrap">
              <div className="flex rounded-lg overflow-hidden border border-slate-200 dark:border-slate-600 text-sm">
                <button
                  onClick={() => { setAccessMode('by-user'); setAccessResult(null); setAccessError(null); }}
                  className={`px-4 py-2 flex items-center gap-2 transition-colors ${
                    accessMode === 'by-user' ? 'bg-blue-600 text-white' : 'bg-white dark:bg-slate-800 text-slate-700 dark:text-slate-300 hover:bg-slate-50 dark:hover:bg-slate-700'
                  }`}
                >
                  <PersonRegular className="w-4 h-4" />
                  What can this user access?
                </button>
                <button
                  onClick={() => { setAccessMode('delegates'); setAccessResult(null); setAccessError(null); }}
                  className={`px-4 py-2 flex items-center gap-2 transition-colors border-l border-slate-200 dark:border-slate-600 ${
                    accessMode === 'delegates' ? 'bg-blue-600 text-white' : 'bg-white dark:bg-slate-800 text-slate-700 dark:text-slate-300 hover:bg-slate-50 dark:hover:bg-slate-700'
                  }`}
                >
                  <MailInboxRegular className="w-4 h-4" />
                  Who has access to this mailbox?
                </button>
              </div>
              <button
                onClick={() => { setAccessMode(m => m === 'by-user' ? 'delegates' : 'by-user'); setAccessResult(null); setAccessError(null); }}
                title="Swap mode"
                className="p-2 rounded-lg border border-slate-200 dark:border-slate-600 text-slate-500 hover:bg-slate-50 dark:hover:bg-slate-700 transition-colors"
              >
                <ArrowSwapRegular className="w-4 h-4" />
              </button>
            </div>

            {/* Email input */}
            <div className="flex gap-2">
              <div className="relative flex-1">
                <PersonRegular className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-400" />
                <input
                  type="email"
                  value={accessEmail}
                  onChange={e => { setAccessEmail(e.target.value); setAccessError(null); }}
                  onKeyDown={e => e.key === 'Enter' && handleAccessSearch()}
                  placeholder={accessMode === 'by-user' ? 'user@contoso.com' : 'sharedmailbox@contoso.com'}
                  className="w-full pl-9 pr-4 py-2.5 border border-slate-300 dark:border-slate-600 rounded-lg bg-white dark:bg-slate-700 text-slate-900 dark:text-white placeholder-slate-400 focus:ring-2 focus:ring-blue-500 text-sm"
                  disabled={accessLoading}
                />
              </div>
              <button
                onClick={handleAccessSearch}
                disabled={accessLoading || !accessEmail.trim()}
                className="px-5 py-2.5 bg-blue-600 hover:bg-blue-700 disabled:opacity-50 text-white rounded-lg flex items-center gap-2 text-sm font-medium transition-colors"
              >
                {accessLoading ? <span className="animate-spin">⟳</span> : <SearchRegular className="w-4 h-4" />}
                {accessLoading ? 'Searching…' : 'Search'}
              </button>
            </div>

            <p className="text-xs text-slate-500 dark:text-slate-400">
              {accessMode === 'by-user'
                ? 'Scans all mailboxes for Full Access, Send As, and Send on Behalf grants. May take 30–90 seconds on larger tenants.'
                : 'Returns all delegates with any permission on the specified mailbox. Usually completes in a few seconds.'}
            </p>
          </div>

          {/* Error */}
          {accessError && (
            <div className="p-4 bg-red-50 dark:bg-red-900/20 border border-red-200 dark:border-red-800 rounded-lg text-sm text-red-700 dark:text-red-300 flex items-center gap-3">
              <WarningRegular className="w-5 h-5 flex-shrink-0" />
              {accessError}
            </div>
          )}

          {/* Loading */}
          {accessLoading && (
            <div className="bg-white dark:bg-slate-800 rounded-lg border border-slate-200 dark:border-slate-700 p-10 flex flex-col items-center gap-3">
              <div className="animate-spin rounded-full h-8 w-8 border-b-2 border-blue-600" />
              <p className="text-sm text-slate-500">{accessMode === 'by-user' ? 'Scanning all mailboxes… this may take a moment.' : 'Fetching delegates…'}</p>
            </div>
          )}

          {/* Results */}
          {accessResult && !accessLoading && (() => {
            const total = accessResult.fullAccessMailboxes.length + accessResult.sendAsMailboxes.length + accessResult.sendOnBehalfMailboxes.length;
            const PermBadge = ({ p }: { p: string }) => (
              p === 'Full Access'
                ? <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded text-xs font-medium bg-blue-100 dark:bg-blue-900/40 text-blue-800 dark:text-blue-300"><LockOpenRegular className="w-3 h-3" />Full Access</span>
                : p === 'Send As'
                ? <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded text-xs font-medium bg-purple-100 dark:bg-purple-900/40 text-purple-800 dark:text-purple-300"><SendRegular className="w-3 h-3" />Send As</span>
                : <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded text-xs font-medium bg-amber-100 dark:bg-amber-900/40 text-amber-800 dark:text-amber-300"><PeopleRegular className="w-3 h-3" />Send on Behalf</span>
            );
            const ResultTable = ({ entries, label }: { entries: typeof accessResult.fullAccessMailboxes; label: string }) => (
              <div className="bg-white dark:bg-slate-800 rounded-lg border border-slate-200 dark:border-slate-700 overflow-hidden">
                <div className="px-4 py-3 border-b border-slate-200 dark:border-slate-700 flex items-center justify-between">
                  <h3 className="font-semibold text-slate-900 dark:text-white flex items-center gap-2">
                    {label === 'Full Access' && <LockOpenRegular className="w-4 h-4 text-blue-500" />}
                    {label === 'Send As' && <SendRegular className="w-4 h-4 text-purple-500" />}
                    {label === 'Send on Behalf' && <PeopleRegular className="w-4 h-4 text-amber-500" />}
                    {label}
                  </h3>
                  <span className={`text-xs font-semibold px-2 py-0.5 rounded-full ${
                    entries.length > 0 ? 'bg-blue-100 dark:bg-blue-900/40 text-blue-700 dark:text-blue-300' : 'bg-slate-100 dark:bg-slate-700 text-slate-500'
                  }`}>{entries.length}</span>
                </div>
                {entries.length === 0 ? (
                  <p className="px-4 py-3 text-sm text-slate-500 dark:text-slate-400">None found.</p>
                ) : (
                  <table className="min-w-full divide-y divide-slate-200 dark:divide-slate-700 text-sm">
                    <thead className="bg-slate-50 dark:bg-slate-700/50">
                      <tr>
                        <th className="px-4 py-2.5 text-left text-xs font-semibold text-slate-500 uppercase">{accessMode === 'by-user' ? 'Mailbox' : 'Granted To'}</th>
                        <th className="px-4 py-2.5 text-left text-xs font-semibold text-slate-500 uppercase">Type</th>
                        <th className="px-4 py-2.5 text-left text-xs font-semibold text-slate-500 uppercase">Permission</th>
                      </tr>
                    </thead>
                    <tbody className="bg-white dark:bg-slate-800 divide-y divide-slate-100 dark:divide-slate-700">
                      {entries.map((e, i) => (
                        <tr key={i} className="hover:bg-slate-50 dark:hover:bg-slate-700/30">
                          <td className="px-4 py-3">
                            <p className="font-medium text-slate-900 dark:text-white">{accessMode === 'by-user' ? (e.mailboxDisplayName || e.mailboxEmail) : e.grantedTo}</p>
                            {accessMode === 'by-user' && <p className="text-xs text-slate-400">{e.mailboxEmail}</p>}
                          </td>
                          <td className="px-4 py-3 text-xs text-slate-500">{e.mailboxType.replace('Mailbox', '').trim() || e.mailboxType}</td>
                          <td className="px-4 py-3"><PermBadge p={e.permission} /></td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                )}
              </div>
            );
            return (
              <div className="space-y-4">
                {/* Summary */}
                <div className={`rounded-lg border px-4 py-3 flex items-center justify-between flex-wrap gap-3 ${
                  total > 0 ? 'bg-blue-50 dark:bg-blue-900/20 border-blue-200 dark:border-blue-800' : 'bg-green-50 dark:bg-green-900/20 border-green-200 dark:border-green-800'
                }`}>
                  <div>
                    <p className={`font-semibold text-sm ${total > 0 ? 'text-blue-800 dark:text-blue-200' : 'text-green-800 dark:text-green-200'}`}>
                      {accessMode === 'by-user'
                        ? total > 0 ? `${accessResult.subjectEmail} has access to ${total} mailbox${total !== 1 ? 'es' : ''}` : `${accessResult.subjectEmail} has no delegated access to other mailboxes`
                        : total > 0 ? `${total} delegate${total !== 1 ? 's have' : ' has'} access to this mailbox` : 'No delegates found for this mailbox'}
                    </p>
                    <p className="text-xs text-slate-500 mt-0.5">
                      {accessMode === 'by-user' ? `${accessResult.mailboxesChecked.toLocaleString()} mailboxes scanned` : 'Full Access, Send As, Send on Behalf checked'}
                      {' · '}{new Date(accessResult.lastUpdated).toLocaleTimeString()}
                    </p>
                  </div>
                </div>
                <ResultTable entries={accessResult.fullAccessMailboxes} label="Full Access" />
                <ResultTable entries={accessResult.sendAsMailboxes} label="Send As" />
                <ResultTable entries={accessResult.sendOnBehalfMailboxes} label="Send on Behalf" />
              </div>
            );
          })()}
        </div>
      ) : null}
    </div>
  );
};

export default MailflowPage;
