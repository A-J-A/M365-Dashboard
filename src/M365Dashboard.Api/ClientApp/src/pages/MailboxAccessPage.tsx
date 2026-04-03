import React, { useState, useRef } from 'react';
import {
  Spinner,
  Badge,
} from '@fluentui/react-components';
import {
  MailRegular,
  PersonRegular,
  LockOpenRegular,
  SendRegular,
  PeopleRegular,
  SearchRegular,
  InfoRegular,
  WarningRegular,
  ArrowSwapRegular,
} from '@fluentui/react-icons';
import { useAppContext } from '../contexts/AppContext';

// ── Types ─────────────────────────────────────────────────────────────────────

interface MailboxAccessEntry {
  mailboxEmail: string;
  mailboxDisplayName: string | null;
  mailboxType: string;
  permission: 'Full Access' | 'Send As' | 'Send on Behalf';
  grantedTo: string;
  isInherited: boolean;
}

interface MailboxAccessResult {
  subjectEmail: string;
  queryType: 'AccessByUser' | 'DelegatesOnMailbox';
  fullAccessMailboxes: MailboxAccessEntry[];
  sendAsMailboxes: MailboxAccessEntry[];
  sendOnBehalfMailboxes: MailboxAccessEntry[];
  totalCount: number;
  mailboxesChecked: number;
  lastUpdated: string;
}

type QueryMode = 'by-user' | 'delegates';

// ── Permission badge ──────────────────────────────────────────────────────────

const PermissionBadge: React.FC<{ permission: string }> = ({ permission }) => {
  if (permission === 'Full Access') {
    return (
      <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded text-xs font-medium bg-blue-100 dark:bg-blue-900/40 text-blue-800 dark:text-blue-300">
        <LockOpenRegular className="w-3 h-3" /> Full Access
      </span>
    );
  }
  if (permission === 'Send As') {
    return (
      <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded text-xs font-medium bg-purple-100 dark:bg-purple-900/40 text-purple-800 dark:text-purple-300">
        <SendRegular className="w-3 h-3" /> Send As
      </span>
    );
  }
  return (
    <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded text-xs font-medium bg-amber-100 dark:bg-amber-900/40 text-amber-800 dark:text-amber-300">
      <PeopleRegular className="w-3 h-3" /> Send on Behalf
    </span>
  );
};

// ── Mailbox type label ────────────────────────────────────────────────────────

const mailboxTypeLabel = (type: string) => {
  const map: Record<string, string> = {
    UserMailbox:    'User',
    SharedMailbox:  'Shared',
    RoomMailbox:    'Room',
    EquipmentMailbox: 'Equipment',
    GroupMailbox:   'Group',
  };
  return map[type] ?? type;
};

// ── Result table ──────────────────────────────────────────────────────────────

const ResultTable: React.FC<{
  entries: MailboxAccessEntry[];
  mode: QueryMode;
  emptyLabel: string;
}> = ({ entries, mode, emptyLabel }) => {
  if (entries.length === 0) {
    return (
      <p className="text-sm text-slate-500 dark:text-slate-400 py-3 px-1">{emptyLabel}</p>
    );
  }

  return (
    <div className="overflow-x-auto rounded-lg border border-slate-200 dark:border-slate-700">
      <table className="min-w-full divide-y divide-slate-200 dark:divide-slate-700 text-sm">
        <thead className="bg-slate-50 dark:bg-slate-700/50">
          <tr>
            <th className="px-4 py-2.5 text-left text-xs font-semibold text-slate-600 dark:text-slate-300 uppercase tracking-wide">
              {mode === 'by-user' ? 'Mailbox' : 'Granted To'}
            </th>
            <th className="px-4 py-2.5 text-left text-xs font-semibold text-slate-600 dark:text-slate-300 uppercase tracking-wide">Type</th>
            <th className="px-4 py-2.5 text-left text-xs font-semibold text-slate-600 dark:text-slate-300 uppercase tracking-wide">Permission</th>
            {mode === 'delegates' && (
              <th className="px-4 py-2.5 text-left text-xs font-semibold text-slate-600 dark:text-slate-300 uppercase tracking-wide">Inherited</th>
            )}
          </tr>
        </thead>
        <tbody className="bg-white dark:bg-slate-800 divide-y divide-slate-100 dark:divide-slate-700">
          {entries.map((e, i) => (
            <tr key={i} className="hover:bg-slate-50 dark:hover:bg-slate-700/30 transition-colors">
              <td className="px-4 py-3">
                <div>
                  <p className="font-medium text-slate-900 dark:text-white">
                    {mode === 'by-user'
                      ? (e.mailboxDisplayName || e.mailboxEmail)
                      : e.grantedTo}
                  </p>
                  <p className="text-xs text-slate-500 dark:text-slate-400 mt-0.5">
                    {mode === 'by-user' ? e.mailboxEmail : ''}
                  </p>
                </div>
              </td>
              <td className="px-4 py-3">
                <span className="text-xs text-slate-600 dark:text-slate-400">
                  {mailboxTypeLabel(e.mailboxType)}
                </span>
              </td>
              <td className="px-4 py-3">
                <PermissionBadge permission={e.permission} />
              </td>
              {mode === 'delegates' && (
                <td className="px-4 py-3">
                  {e.isInherited
                    ? <span className="text-xs text-slate-400">Inherited</span>
                    : <span className="text-xs text-slate-600 dark:text-slate-400">Direct</span>}
                </td>
              )}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
};

// ── Section card ──────────────────────────────────────────────────────────────

const Section: React.FC<{
  icon: React.ReactNode;
  title: string;
  count: number;
  children: React.ReactNode;
}> = ({ icon, title, count, children }) => (
  <div className="bg-white dark:bg-slate-800 rounded-xl border border-slate-200 dark:border-slate-700 overflow-hidden">
    <div className="px-5 py-4 border-b border-slate-200 dark:border-slate-700 flex items-center justify-between">
      <div className="flex items-center gap-2">
        <span className="text-slate-500 dark:text-slate-400">{icon}</span>
        <h3 className="font-semibold text-slate-900 dark:text-white">{title}</h3>
      </div>
      <Badge appearance="filled" color={count > 0 ? 'brand' : 'subtle'}>
        {count}
      </Badge>
    </div>
    <div className="p-5">{children}</div>
  </div>
);

// ── Main page ─────────────────────────────────────────────────────────────────

const MailboxAccessPage: React.FC = () => {
  const { getAccessToken } = useAppContext();

  const [mode, setMode]           = useState<QueryMode>('by-user');
  const [email, setEmail]         = useState('');
  const [loading, setLoading]     = useState(false);
  const [result, setResult]       = useState<MailboxAccessResult | null>(null);
  const [error, setError]         = useState<string | null>(null);
  const inputRef                  = useRef<HTMLInputElement>(null);

  const [debugResult, setDebugResult] = useState<string | null>(null);
  const [debugLoading, setDebugLoading] = useState(false);

  const handleDebug = async () => {
    const trimmed = email.trim();
    if (!trimmed || !trimmed.includes('@')) return;
    setDebugLoading(true);
    setDebugResult(null);
    try {
      const token = await getAccessToken();
      const response = await fetch(
        `/api/exchange/debug/mailbox-permissions?mailbox=${encodeURIComponent(trimmed)}`,
        { headers: { Authorization: `Bearer ${token}` } }
      );
      const data = await response.json();
      setDebugResult(JSON.stringify(data, null, 2));
    } catch (err) {
      setDebugResult(String(err));
    } finally {
      setDebugLoading(false);
    }
  };

  const handleSearch = async () => {
    const trimmed = email.trim();
    if (!trimmed || !trimmed.includes('@')) {
      setError('Please enter a valid email address.');
      return;
    }

    setLoading(true);
    setResult(null);
    setError(null);

    try {
      const token    = await getAccessToken();
      const endpoint = mode === 'by-user'
        ? `/api/exchange/mailbox-access/by-user?email=${encodeURIComponent(trimmed)}`
        : `/api/exchange/mailbox-access/delegates?email=${encodeURIComponent(trimmed)}`;

      const response = await fetch(endpoint, {
        headers: { Authorization: `Bearer ${token}` },
      });

      if (!response.ok) {
        const data = await response.json().catch(() => ({}));
        throw new Error(data.message ?? data.error ?? `HTTP ${response.status}`);
      }

      setResult(await response.json());
    } catch (err) {
      setError(err instanceof Error ? err.message : 'An unexpected error occurred.');
    } finally {
      setLoading(false);
    }
  };

  const handleKeyDown = (e: React.KeyboardEvent) => {
    if (e.key === 'Enter') handleSearch();
  };

  const handleSwapMode = () => {
    setMode(m => m === 'by-user' ? 'delegates' : 'by-user');
    setResult(null);
    setError(null);
  };

  const totalEntries = result
    ? result.fullAccessMailboxes.length + result.sendAsMailboxes.length + result.sendOnBehalfMailboxes.length
    : 0;

  return (
    <div className="max-w-4xl mx-auto p-6 space-y-6">
      {/* Header */}
      <div>
        <h1 className="text-2xl font-bold text-slate-900 dark:text-white flex items-center gap-2">
          <MailRegular className="w-7 h-7" />
          Mailbox Access
        </h1>
        <p className="mt-1 text-sm text-slate-500 dark:text-slate-400">
          Look up mailbox permissions — find which mailboxes a user can access, or who has access to a specific mailbox.
        </p>
      </div>

      {/* Mode toggle + search */}
      <div className="bg-white dark:bg-slate-800 rounded-xl border border-slate-200 dark:border-slate-700 p-5 space-y-4">
        {/* Mode selector */}
        <div className="flex items-center gap-3 flex-wrap">
          <div className="flex rounded-lg overflow-hidden border border-slate-200 dark:border-slate-600 text-sm">
            <button
              onClick={() => { setMode('by-user'); setResult(null); setError(null); }}
              className={`px-4 py-2 flex items-center gap-2 transition-colors ${
                mode === 'by-user'
                  ? 'bg-blue-600 text-white'
                  : 'bg-white dark:bg-slate-800 text-slate-700 dark:text-slate-300 hover:bg-slate-50 dark:hover:bg-slate-700'
              }`}
            >
              <PersonRegular className="w-4 h-4" />
              What can this user access?
            </button>
            <button
              onClick={() => { setMode('delegates'); setResult(null); setError(null); }}
              className={`px-4 py-2 flex items-center gap-2 transition-colors border-l border-slate-200 dark:border-slate-600 ${
                mode === 'delegates'
                  ? 'bg-blue-600 text-white'
                  : 'bg-white dark:bg-slate-800 text-slate-700 dark:text-slate-300 hover:bg-slate-50 dark:hover:bg-slate-700'
              }`}
            >
              <MailRegular className="w-4 h-4" />
              Who has access to this mailbox?
            </button>
          </div>

          <button
            onClick={handleSwapMode}
            title="Swap mode"
            className="p-2 rounded-lg border border-slate-200 dark:border-slate-600 text-slate-500 dark:text-slate-400 hover:bg-slate-50 dark:hover:bg-slate-700 transition-colors"
          >
            <ArrowSwapRegular className="w-4 h-4" />
          </button>
        </div>

        {/* Input */}
        <div>
          <label className="block text-sm font-medium text-slate-700 dark:text-slate-300 mb-1.5">
            {mode === 'by-user' ? 'User email address' : 'Mailbox email address'}
          </label>
          <div className="flex gap-2">
            <div className="relative flex-1">
              <PersonRegular className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-400" />
              <input
                ref={inputRef}
                type="email"
                value={email}
                onChange={e => { setEmail(e.target.value); setError(null); }}
                onKeyDown={handleKeyDown}
                placeholder={mode === 'by-user' ? 'user@contoso.com' : 'sharedmailbox@contoso.com'}
                className="w-full pl-9 pr-4 py-2.5 border border-slate-300 dark:border-slate-600 rounded-lg bg-white dark:bg-slate-700 text-slate-900 dark:text-white placeholder-slate-400 focus:ring-2 focus:ring-blue-500 focus:border-blue-500 text-sm"
                disabled={loading}
              />
            </div>
            <button
              onClick={handleSearch}
              disabled={loading || !email.trim()}
              className="px-5 py-2.5 bg-blue-600 hover:bg-blue-700 disabled:opacity-50 disabled:cursor-not-allowed text-white rounded-lg flex items-center gap-2 text-sm font-medium transition-colors"
            >
              {loading
                ? <Spinner size="tiny" />
                : <SearchRegular className="w-4 h-4" />}
              {loading ? 'Searching…' : 'Search'}
            </button>
          </div>
        </div>

        {/* Info note */}
        <div className="flex items-start gap-2 text-xs text-slate-500 dark:text-slate-400 bg-slate-50 dark:bg-slate-700/50 rounded-lg px-3 py-2.5">
          <InfoRegular className="w-4 h-4 flex-shrink-0 mt-0.5" />
          <span>
            {mode === 'by-user'
              ? 'Scans all mailboxes in the tenant for Full Access, Send As, and Send on Behalf grants assigned to the specified user. This may take 30–90 seconds on larger tenants.'
              : 'Returns all users and groups with Full Access, Send As, or Send on Behalf permissions on the specified mailbox. Usually completes within a few seconds.'}
          </span>
        </div>

        {/* Debug panel — enter a mailbox you know has permissions and click Debug */}
        <details className="text-xs">
          <summary className="cursor-pointer text-slate-400 hover:text-slate-600 dark:hover:text-slate-300 select-none">
            🔧 Debug: inspect raw Exchange API response for one mailbox
          </summary>
          <div className="mt-2 space-y-2">
            <p className="text-slate-500 dark:text-slate-400">
              Enter a mailbox address above that you know has delegates, then click Debug to see the raw JSON Exchange returns.
            </p>
            <button
              onClick={handleDebug}
              disabled={debugLoading || !email.trim()}
              className="px-3 py-1.5 bg-slate-600 hover:bg-slate-700 disabled:opacity-50 text-white rounded text-xs"
            >
              {debugLoading ? 'Loading…' : 'Debug: Get-MailboxPermission'}
            </button>
            {debugResult && (
              <pre className="mt-2 p-3 bg-slate-900 text-green-400 rounded text-[11px] overflow-x-auto max-h-96 overflow-y-auto whitespace-pre-wrap break-all">
                {debugResult}
              </pre>
            )}
          </div>
        </details>
      </div>

      {/* Error */}
      {error && (
        <div className="flex items-start gap-3 p-4 bg-red-50 dark:bg-red-900/20 border border-red-200 dark:border-red-800 rounded-xl text-sm text-red-700 dark:text-red-300">
          <WarningRegular className="w-5 h-5 flex-shrink-0 mt-0.5" />
          <span>{error}</span>
        </div>
      )}

      {/* Loading state */}
      {loading && (
        <div className="bg-white dark:bg-slate-800 rounded-xl border border-slate-200 dark:border-slate-700 p-10 flex flex-col items-center gap-3">
          <Spinner size="large" />
          <p className="text-slate-500 dark:text-slate-400 text-sm">
            {mode === 'by-user'
              ? 'Scanning all mailboxes for access grants… this may take a moment.'
              : 'Fetching delegates for this mailbox…'}
          </p>
        </div>
      )}

      {/* Results */}
      {result && !loading && (
        <div className="space-y-5">
          {/* Summary banner */}
          <div className={`rounded-xl border px-5 py-4 flex items-center justify-between flex-wrap gap-3 ${
            totalEntries > 0
              ? 'bg-blue-50 dark:bg-blue-900/20 border-blue-200 dark:border-blue-800'
              : 'bg-green-50 dark:bg-green-900/20 border-green-200 dark:border-green-800'
          }`}>
            <div>
              <p className={`font-semibold ${totalEntries > 0 ? 'text-blue-800 dark:text-blue-200' : 'text-green-800 dark:text-green-200'}`}>
                {mode === 'by-user'
                  ? totalEntries > 0
                    ? `${result.subjectEmail} has access to ${totalEntries} mailbox${totalEntries !== 1 ? 'es' : ''}`
                    : `${result.subjectEmail} has no delegated access to other mailboxes`
                  : totalEntries > 0
                    ? `${totalEntries} delegate${totalEntries !== 1 ? 's have' : ' has'} access to this mailbox`
                    : 'No delegates found for this mailbox'}
              </p>
              <p className="text-xs text-slate-500 dark:text-slate-400 mt-1">
                {mode === 'by-user'
                  ? `${result.mailboxesChecked.toLocaleString()} mailboxes scanned`
                  : `Checked Full Access, Send As, and Send on Behalf`}
                {' · '}
                {new Date(result.lastUpdated).toLocaleTimeString()}
              </p>
            </div>
            <div className="flex gap-2 text-xs flex-wrap">
              {result.fullAccessMailboxes.length > 0 && (
                <span className="px-2 py-1 rounded bg-blue-100 dark:bg-blue-900/40 text-blue-800 dark:text-blue-300">
                  {result.fullAccessMailboxes.length} Full Access
                </span>
              )}
              {result.sendAsMailboxes.length > 0 && (
                <span className="px-2 py-1 rounded bg-purple-100 dark:bg-purple-900/40 text-purple-800 dark:text-purple-300">
                  {result.sendAsMailboxes.length} Send As
                </span>
              )}
              {result.sendOnBehalfMailboxes.length > 0 && (
                <span className="px-2 py-1 rounded bg-amber-100 dark:bg-amber-900/40 text-amber-800 dark:text-amber-300">
                  {result.sendOnBehalfMailboxes.length} Send on Behalf
                </span>
              )}
            </div>
          </div>

          {/* Full Access */}
          <Section
            icon={<LockOpenRegular className="w-5 h-5" />}
            title="Full Access"
            count={result.fullAccessMailboxes.length}
          >
            <ResultTable
              entries={result.fullAccessMailboxes}
              mode={mode}
              emptyLabel="No Full Access grants found."
            />
          </Section>

          {/* Send As */}
          <Section
            icon={<SendRegular className="w-5 h-5" />}
            title="Send As"
            count={result.sendAsMailboxes.length}
          >
            <ResultTable
              entries={result.sendAsMailboxes}
              mode={mode}
              emptyLabel="No Send As grants found."
            />
          </Section>

          {/* Send on Behalf */}
          <Section
            icon={<PeopleRegular className="w-5 h-5" />}
            title="Send on Behalf"
            count={result.sendOnBehalfMailboxes.length}
          >
            <ResultTable
              entries={result.sendOnBehalfMailboxes}
              mode={mode}
              emptyLabel="No Send on Behalf grants found."
            />
          </Section>
        </div>
      )}
    </div>
  );
};

export default MailboxAccessPage;
