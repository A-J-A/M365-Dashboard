import React, { useState, useEffect, useCallback } from 'react';
import {
  ShieldCheckmarkRegular,
  ShieldErrorRegular,
  ChevronDownRegular,
  ChevronUpRegular,
  CheckmarkCircleRegular,
  DismissCircleRegular,
  WarningRegular,
  InfoRegular,
  OpenRegular,
  ArrowClockwiseRegular,
  MailRegular,
  LinkRegular,
  DocumentRegular,
  FilterRegular,
} from '@fluentui/react-icons';
import { useAppContext } from '../contexts/AppContext';

// ── Types ─────────────────────────────────────────────────────────────────────

interface AntiPhishPolicy {
  identity: string;
  name: string;
  enabled: boolean;
  isDefault: boolean;
  enableSpoofIntelligence: boolean;
  enableUnauthenticatedSender: boolean;
  enableViaTag: boolean;
  phishThresholdLevel: string;
  enableFirstContactSafetyTips: boolean;
  enableSimilarUsersSafetyTips: boolean;
  enableSimilarDomainsSafetyTips: boolean;
  enableMailboxIntelligence: boolean;
  enableMailboxIntelligenceProtection: boolean;
  mailboxIntelligenceProtectionAction: string;
  enableOrganizationDomainsProtection: boolean;
  enableTargetedDomainsProtection: boolean;
  enableTargetedUserProtection: boolean;
  targetedUserProtectionAction: string;
  targetedDomainProtectionAction: string;
  targetedUsersToProtect: string[];
  targetedDomainsToProtect: string[];
  whenCreated: string | null;
  whenChanged: string | null;
}

interface AntiMalwarePolicy {
  identity: string;
  name: string;
  enabled: boolean;
  isDefault: boolean;
  action: string;
  enableFileFilter: boolean;
  fileTypes: string[];
  zapEnabled: boolean;
  enableInternalSenderAdminNotifications: boolean;
  internalSenderAdminAddress: string | null;
  enableExternalSenderAdminNotifications: boolean;
  externalSenderAdminAddress: string | null;
  whenCreated: string | null;
  whenChanged: string | null;
}

interface AntiSpamPolicy {
  identity: string;
  name: string;
  enabled: boolean;
  isDefault: boolean;
  spamAction: string;
  highConfidenceSpamAction: string;
  phishSpamAction: string;
  highConfidencePhishAction: string;
  bulkSpamAction: string;
  bulkThreshold: number;
  zapEnabled: boolean;
  spamZapEnabled: boolean;
  phishZapEnabled: boolean;
  enableEndUserSpamNotifications: boolean;
  endUserSpamNotificationFrequency: number;
  allowedSenderDomainsPresent: boolean;
  blockedSenderDomainsPresent: boolean;
  quarantineTag: string;
  whenCreated: string | null;
  whenChanged: string | null;
}

interface OutboundSpamPolicy {
  identity: string;
  name: string;
  enabled: boolean;
  isDefault: boolean;
  actionWhenThresholdReached: string;
  recipientLimitExternalPerHour: number;
  recipientLimitInternalPerHour: number;
  recipientLimitPerDay: number;
  autoForwardingMode: boolean;
  autoForwardingModeValue: string;
  whenCreated: string | null;
  whenChanged: string | null;
}

interface SafeAttachmentsPolicy {
  identity: string;
  name: string;
  enabled: boolean;
  isDefault: boolean;
  action: string;
  actionOnError: boolean;
  redirect: boolean;
  redirectAddress: string | null;
  whenCreated: string | null;
  whenChanged: string | null;
}

interface SafeLinksPolicy {
  identity: string;
  name: string;
  enabled: boolean;
  isDefault: boolean;
  allowClickThrough: boolean;
  disableUrlRewrite: boolean;
  enableForInternalSenders: boolean;
  enableSafeLinksForEmail: boolean;
  enableSafeLinksForTeams: boolean;
  enableSafeLinksForOffice: boolean;
  trackClicks: boolean;
  scanUrls: boolean;
  doNotRewriteUrls: string[];
  whenCreated: string | null;
  whenChanged: string | null;
}

interface PolicyResult<T> {
  policies: T[];
  totalCount: number;
  error: string | null;
  isLicensed?: boolean;
}

interface DefenderOverview {
  antiPhish: PolicyResult<AntiPhishPolicy>;
  antiMalware: PolicyResult<AntiMalwarePolicy>;
  antiSpam: PolicyResult<AntiSpamPolicy>;
  outboundSpam: PolicyResult<OutboundSpamPolicy>;
  safeAttachments: PolicyResult<SafeAttachmentsPolicy> & { isLicensed: boolean };
  safeLinks: PolicyResult<SafeLinksPolicy> & { isLicensed: boolean };
  lastUpdated: string;
}

// ── Helper functions ──────────────────────────────────────────────────────────

const formatDate = (dateString: string | null) => {
  if (!dateString) return 'N/A';
  return new Date(dateString).toLocaleDateString('en-GB', {
    day: '2-digit', month: 'short', year: 'numeric',
  });
};

const friendlyAction = (action: string) => {
  const map: Record<string, string> = {
    MoveToJmf: 'Junk Mail',
    Quarantine: 'Quarantine',
    Delete: 'Delete',
    DeleteMessage: 'Delete',
    NoAction: 'No Action',
    Redirect: 'Redirect',
    Block: 'Block',
    DynamicDelivery: 'Dynamic Delivery',
    BlockUserForToday: 'Block User (24h)',
    Alert: 'Alert Only',
    BlockUser: 'Block User',
  };
  return map[action] ?? action;
};

const friendlyAutoForward = (value: string) => {
  const map: Record<string, string> = {
    Automatic: 'Automatic (org policy)',
    On: 'Allowed',
    Off: 'Blocked',
  };
  return map[value] ?? value;
};

const friendlyPhishThreshold = (level: string) => {
  const map: Record<string, string> = {
    '1': '1 – Standard',
    '2': '2 – Aggressive',
    '3': '3 – More Aggressive',
    '4': '4 – Most Aggressive',
  };
  return map[level] ?? `Level ${level}`;
};

// ── Badge components ──────────────────────────────────────────────────────────

const EnabledBadge: React.FC<{ enabled: boolean; label?: string }> = ({ enabled, label }) => (
  <span className={`inline-flex items-center gap-1 px-2 py-0.5 rounded text-xs font-medium
    ${enabled
      ? 'bg-green-100 text-green-700 dark:bg-green-900/30 dark:text-green-400'
      : 'bg-red-100 text-red-700 dark:bg-red-900/30 dark:text-red-400'}`}>
    {enabled
      ? <CheckmarkCircleRegular className="w-3 h-3" />
      : <DismissCircleRegular className="w-3 h-3" />}
    {label ?? (enabled ? 'Enabled' : 'Disabled')}
  </span>
);

const ActionBadge: React.FC<{ action: string }> = ({ action }) => {
  const friendly = friendlyAction(action);
  const colorMap: Record<string, string> = {
    Quarantine: 'bg-amber-100 text-amber-700 dark:bg-amber-900/30 dark:text-amber-400',
    Delete: 'bg-red-100 text-red-700 dark:bg-red-900/30 dark:text-red-400',
    DeleteMessage: 'bg-red-100 text-red-700 dark:bg-red-900/30 dark:text-red-400',
    Block: 'bg-red-100 text-red-700 dark:bg-red-900/30 dark:text-red-400',
    'Junk Mail': 'bg-blue-100 text-blue-700 dark:bg-blue-900/30 dark:text-blue-400',
    'No Action': 'bg-slate-100 text-slate-600 dark:bg-slate-700 dark:text-slate-400',
    'Dynamic Delivery': 'bg-purple-100 text-purple-700 dark:bg-purple-900/30 dark:text-purple-400',
    'Block User (24h)': 'bg-orange-100 text-orange-700 dark:bg-orange-900/30 dark:text-orange-400',
  };
  const cls = colorMap[friendly] ?? 'bg-slate-100 text-slate-600 dark:bg-slate-700 dark:text-slate-400';
  return <span className={`inline-flex items-center px-2 py-0.5 rounded text-xs font-medium ${cls}`}>{friendly}</span>;
};

const SettingRow: React.FC<{ label: string; children: React.ReactNode }> = ({ label, children }) => (
  <div className="flex items-center justify-between py-1.5 border-b border-slate-100 dark:border-slate-700/50 last:border-0">
    <span className="text-xs text-slate-500 dark:text-slate-400 pr-4">{label}</span>
    <div className="flex-shrink-0">{children}</div>
  </div>
);

const BoolRow: React.FC<{ label: string; value: boolean; goodWhenTrue?: boolean }> = ({
  label, value, goodWhenTrue = true,
}) => {
  const isGood = goodWhenTrue ? value : !value;
  return (
    <SettingRow label={label}>
      <span className={`inline-flex items-center gap-1 text-xs font-medium
        ${isGood ? 'text-green-600 dark:text-green-400' : 'text-amber-600 dark:text-amber-400'}`}>
        {value
          ? <CheckmarkCircleRegular className="w-3.5 h-3.5" />
          : <DismissCircleRegular className="w-3.5 h-3.5" />}
        {value ? 'Yes' : 'No'}
      </span>
    </SettingRow>
  );
};

// ── Error / Unlicensed banners ────────────────────────────────────────────────

const ErrorBanner: React.FC<{ message: string }> = ({ message }) => (
  <div className="flex items-start gap-2 p-3 bg-red-50 dark:bg-red-900/20 border border-red-200 dark:border-red-800 rounded-lg text-sm text-red-700 dark:text-red-400">
    <DismissCircleRegular className="w-4 h-4 flex-shrink-0 mt-0.5" />
    <span>{message}</span>
  </div>
);

const UnlicensedBanner: React.FC = () => (
  <div className="flex items-start gap-2 p-3 bg-amber-50 dark:bg-amber-900/20 border border-amber-200 dark:border-amber-800 rounded-lg text-sm text-amber-700 dark:text-amber-400">
    <InfoRegular className="w-4 h-4 flex-shrink-0 mt-0.5" />
    <span>Not available — requires <strong>Microsoft Defender for Office 365 Plan 1</strong> or higher.</span>
  </div>
);

// ── Expandable policy card ────────────────────────────────────────────────────

const PolicyCard: React.FC<{
  title: string;
  isDefault: boolean;
  enabled: boolean;
  lastChanged: string | null;
  children: React.ReactNode;
}> = ({ title, isDefault, enabled, lastChanged, children }) => {
  const [open, setOpen] = useState(isDefault);

  return (
    <div className="border border-slate-200 dark:border-slate-700 rounded-lg overflow-hidden">
      <button
        onClick={() => setOpen(o => !o)}
        className="w-full flex items-center justify-between px-4 py-3 bg-slate-50 dark:bg-slate-800/60 hover:bg-slate-100 dark:hover:bg-slate-700/50 transition-colors"
      >
        <div className="flex items-center gap-2 min-w-0">
          <span className="font-medium text-sm text-slate-900 dark:text-white truncate">{title}</span>
          {isDefault && (
            <span className="px-1.5 py-0.5 rounded text-xs bg-blue-100 text-blue-700 dark:bg-blue-900/30 dark:text-blue-400 font-medium flex-shrink-0">
              Default
            </span>
          )}
          <EnabledBadge enabled={enabled} />
        </div>
        <div className="flex items-center gap-3 ml-3 flex-shrink-0">
          {lastChanged && (
            <span className="text-xs text-slate-400 hidden sm:block">
              Changed {formatDate(lastChanged)}
            </span>
          )}
          {open
            ? <ChevronUpRegular className="w-4 h-4 text-slate-400" />
            : <ChevronDownRegular className="w-4 h-4 text-slate-400" />}
        </div>
      </button>
      {open && (
        <div className="px-4 py-3 bg-white dark:bg-slate-800 space-y-0.5">
          {children}
        </div>
      )}
    </div>
  );
};

// ── Section wrapper ───────────────────────────────────────────────────────────

const Section: React.FC<{
  icon: React.ReactNode;
  title: string;
  subtitle: string;
  count: number;
  error: string | null;
  unlicensed?: boolean;
  children: React.ReactNode;
}> = ({ icon, title, subtitle, count, error, unlicensed = false, children }) => (
  <div className="bg-white dark:bg-slate-800 rounded-lg border border-slate-200 dark:border-slate-700 overflow-hidden">
    <div className="px-4 py-3 border-b border-slate-200 dark:border-slate-700 flex items-center justify-between">
      <div className="flex items-center gap-2">
        <div className="text-blue-600 dark:text-blue-400">{icon}</div>
        <div>
          <h2 className="text-sm font-semibold text-slate-900 dark:text-white">{title}</h2>
          <p className="text-xs text-slate-500 dark:text-slate-400">{subtitle}</p>
        </div>
      </div>
      {!unlicensed && !error && (
        <span className="text-xs text-slate-500 dark:text-slate-400 bg-slate-100 dark:bg-slate-700 px-2 py-0.5 rounded-full">
          {count} {count === 1 ? 'policy' : 'policies'}
        </span>
      )}
    </div>
    <div className="p-4 space-y-3">
      {error && !unlicensed && <ErrorBanner message={error} />}
      {unlicensed && <UnlicensedBanner />}
      {!error && !unlicensed && children}
    </div>
  </div>
);

// ── Main page ─────────────────────────────────────────────────────────────────

const DefenderForOfficePage: React.FC = () => {
  const { getAccessToken } = useAppContext();
  const [overview, setOverview] = useState<DefenderOverview | null>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  const fetchOverview = useCallback(async () => {
    try {
      setLoading(true);
      setError(null);
      const token = await getAccessToken();
      const response = await fetch('/api/defender-office/overview', {
        headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' },
      });
      if (!response.ok) throw new Error(`HTTP ${response.status}: ${response.statusText}`);
      const data: DefenderOverview = await response.json();
      setOverview(data);
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Failed to load Defender data');
    } finally {
      setLoading(false);
    }
  }, [getAccessToken]);

  useEffect(() => { fetchOverview(); }, [fetchOverview]);

  // ── Summary counts for header ────────────────────────────────────────────
  const summaryStats = overview ? [
    {
      label: 'Anti-Phish',
      count: overview.antiPhish.totalCount,
      ok: !overview.antiPhish.error && overview.antiPhish.totalCount > 0,
    },
    {
      label: 'Anti-Malware',
      count: overview.antiMalware.totalCount,
      ok: !overview.antiMalware.error && overview.antiMalware.totalCount > 0,
    },
    {
      label: 'Anti-Spam',
      count: overview.antiSpam.totalCount,
      ok: !overview.antiSpam.error && overview.antiSpam.totalCount > 0,
    },
    {
      label: 'Safe Attachments',
      count: overview.safeAttachments.totalCount,
      ok: overview.safeAttachments.isLicensed && !overview.safeAttachments.error,
      unlicensed: !overview.safeAttachments.isLicensed,
    },
    {
      label: 'Safe Links',
      count: overview.safeLinks.totalCount,
      ok: overview.safeLinks.isLicensed && !overview.safeLinks.error,
      unlicensed: !overview.safeLinks.isLicensed,
    },
  ] : [];

  // ── Loading ──────────────────────────────────────────────────────────────
  if (loading) {
    return (
      <div className="p-4 space-y-4">
        <div className="flex items-center justify-between">
          <div>
            <div className="h-6 w-64 bg-slate-200 dark:bg-slate-700 rounded animate-pulse" />
            <div className="h-4 w-48 bg-slate-200 dark:bg-slate-700 rounded animate-pulse mt-1" />
          </div>
        </div>
        <div className="grid grid-cols-2 sm:grid-cols-5 gap-3">
          {Array.from({ length: 5 }).map((_, i) => (
            <div key={i} className="h-16 bg-slate-200 dark:bg-slate-700 rounded-lg animate-pulse" />
          ))}
        </div>
        {Array.from({ length: 3 }).map((_, i) => (
          <div key={i} className="h-48 bg-slate-200 dark:bg-slate-700 rounded-lg animate-pulse" />
        ))}
      </div>
    );
  }

  if (error) {
    return (
      <div className="p-4">
        <div className="bg-red-50 dark:bg-red-900/20 border border-red-200 dark:border-red-800 rounded-lg p-4">
          <p className="text-red-600 dark:text-red-400 font-medium">Failed to load Defender data</p>
          <p className="text-sm text-red-500 dark:text-red-500 mt-1">{error}</p>
          <button onClick={fetchOverview} className="mt-3 text-sm text-red-600 dark:text-red-400 underline flex items-center gap-1">
            <ArrowClockwiseRegular className="w-3 h-3" /> Retry
          </button>
        </div>
      </div>
    );
  }

  if (!overview) return null;

  return (
    <div className="p-4 space-y-5 w-full max-w-full overflow-hidden">

      {/* Header */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
        <div>
          <h1 className="text-xl font-semibold text-slate-900 dark:text-white">Defender for Office 365</h1>
          <p className="text-sm text-slate-500 dark:text-slate-400 hidden sm:block">
            Email security policies — anti-phish, malware, spam, Safe Attachments &amp; Links
          </p>
        </div>
        <div className="flex items-center gap-2">
          <button
            onClick={fetchOverview}
            className="inline-flex items-center gap-2 px-3 py-2 bg-slate-100 dark:bg-slate-700 text-slate-700 dark:text-slate-300 rounded-lg hover:bg-slate-200 dark:hover:bg-slate-600 transition-colors text-sm"
          >
            <ArrowClockwiseRegular className="w-4 h-4" /> Refresh
          </button>
          <a
            href="https://security.microsoft.com/presets"
            target="_blank"
            rel="noopener noreferrer"
            className="inline-flex items-center gap-2 px-3 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 transition-colors text-sm whitespace-nowrap"
          >
            <OpenRegular className="w-4 h-4" />
            <span className="hidden sm:inline">Security Portal</span>
            <span className="sm:hidden">Portal</span>
          </a>
        </div>
      </div>

      {/* Summary stat strip */}
      <div className="grid grid-cols-2 sm:grid-cols-5 gap-3">
        {summaryStats.map(s => (
          <div key={s.label} className={`rounded-lg border p-3 text-center
            ${s.unlicensed
              ? 'border-amber-200 dark:border-amber-800 bg-amber-50 dark:bg-amber-900/10'
              : s.ok
                ? 'border-green-200 dark:border-green-800 bg-green-50 dark:bg-green-900/10'
                : 'border-red-200 dark:border-red-800 bg-red-50 dark:bg-red-900/10'}`}>
            <p className={`text-lg font-bold leading-tight
              ${s.unlicensed ? 'text-amber-600 dark:text-amber-400'
                : s.ok ? 'text-green-600 dark:text-green-400'
                : 'text-red-600 dark:text-red-400'}`}>
              {s.unlicensed ? '—' : s.count}
            </p>
            <p className="text-xs text-slate-500 dark:text-slate-400 leading-tight mt-0.5">{s.label}</p>
            {s.unlicensed && <p className="text-xs text-amber-500 dark:text-amber-500 mt-0.5">Not licensed</p>}
          </div>
        ))}
      </div>

      {/* Last updated */}
      <p className="text-xs text-slate-400 dark:text-slate-500">
        Data fetched at {new Date(overview.lastUpdated).toLocaleTimeString()}
      </p>

      {/* ── Anti-Phishing ──────────────────────────────────────────────────── */}
      <Section
        icon={<ShieldErrorRegular className="w-5 h-5" />}
        title="Anti-Phishing"
        subtitle="Spoof intelligence, impersonation protection &amp; safety tips"
        count={overview.antiPhish.totalCount}
        error={overview.antiPhish.error}
      >
        {overview.antiPhish.policies.map(p => (
          <PolicyCard
            key={p.identity}
            title={p.name}
            isDefault={p.isDefault}
            enabled={p.enabled}
            lastChanged={p.whenChanged}
          >
            <BoolRow label="Spoof intelligence" value={p.enableSpoofIntelligence} />
            <BoolRow label="Unauthenticated sender indicators" value={p.enableUnauthenticatedSender} />
            <BoolRow label="Via tag" value={p.enableViaTag} />
            <SettingRow label="Phish threshold level">
              <span className="text-xs font-medium text-slate-700 dark:text-slate-300">
                {friendlyPhishThreshold(p.phishThresholdLevel)}
              </span>
            </SettingRow>
            <BoolRow label="Mailbox intelligence" value={p.enableMailboxIntelligence} />
            <BoolRow label="Mailbox intelligence protection" value={p.enableMailboxIntelligenceProtection} />
            {p.enableMailboxIntelligenceProtection && (
              <SettingRow label="Intelligence protection action">
                <ActionBadge action={p.mailboxIntelligenceProtectionAction} />
              </SettingRow>
            )}
            <BoolRow label="Org domain impersonation protection" value={p.enableOrganizationDomainsProtection} />
            <BoolRow label="Targeted domain protection" value={p.enableTargetedDomainsProtection} />
            {p.enableTargetedDomainsProtection && (
              <SettingRow label="Domain protection action">
                <ActionBadge action={p.targetedDomainProtectionAction} />
              </SettingRow>
            )}
            <BoolRow label="Targeted user protection" value={p.enableTargetedUserProtection} />
            {p.enableTargetedUserProtection && (
              <>
                <SettingRow label="User protection action">
                  <ActionBadge action={p.targetedUserProtectionAction} />
                </SettingRow>
                <SettingRow label="Protected users">
                  <span className="text-xs text-slate-600 dark:text-slate-300">
                    {p.targetedUsersToProtect.length > 0
                      ? `${p.targetedUsersToProtect.length} user(s)`
                      : 'None configured'}
                  </span>
                </SettingRow>
              </>
            )}
            <BoolRow label="First contact safety tips" value={p.enableFirstContactSafetyTips} />
            <BoolRow label="Similar user safety tips" value={p.enableSimilarUsersSafetyTips} />
            <BoolRow label="Similar domain safety tips" value={p.enableSimilarDomainsSafetyTips} />
          </PolicyCard>
        ))}
      </Section>

      {/* ── Anti-Malware ───────────────────────────────────────────────────── */}
      <Section
        icon={<ShieldCheckmarkRegular className="w-5 h-5" />}
        title="Anti-Malware"
        subtitle="Malware filter policies with file type blocking and ZAP"
        count={overview.antiMalware.totalCount}
        error={overview.antiMalware.error}
      >
        {overview.antiMalware.policies.map(p => (
          <PolicyCard
            key={p.identity}
            title={p.name}
            isDefault={p.isDefault}
            enabled={p.enabled}
            lastChanged={p.whenChanged}
          >
            <SettingRow label="Action">
              <ActionBadge action={p.action} />
            </SettingRow>
            <BoolRow label="File type filter" value={p.enableFileFilter} />
            {p.enableFileFilter && p.fileTypes.length > 0 && (
              <SettingRow label="Blocked file types">
                <span className="text-xs text-slate-600 dark:text-slate-300">{p.fileTypes.length} types</span>
              </SettingRow>
            )}
            <BoolRow label="Zero-hour auto purge (ZAP)" value={p.zapEnabled} />
            <BoolRow label="Internal sender notifications" value={p.enableInternalSenderAdminNotifications} />
            {p.enableInternalSenderAdminNotifications && p.internalSenderAdminAddress && (
              <SettingRow label="Internal notify address">
                <span className="text-xs text-slate-600 dark:text-slate-300 truncate max-w-48">{p.internalSenderAdminAddress}</span>
              </SettingRow>
            )}
            <BoolRow label="External sender notifications" value={p.enableExternalSenderAdminNotifications} />
            {p.enableExternalSenderAdminNotifications && p.externalSenderAdminAddress && (
              <SettingRow label="External notify address">
                <span className="text-xs text-slate-600 dark:text-slate-300 truncate max-w-48">{p.externalSenderAdminAddress}</span>
              </SettingRow>
            )}
          </PolicyCard>
        ))}
      </Section>

      {/* ── Anti-Spam (inbound) ────────────────────────────────────────────── */}
      <Section
        icon={<FilterRegular className="w-5 h-5" />}
        title="Anti-Spam (Inbound)"
        subtitle="Spam, phish, and bulk mail filtering actions"
        count={overview.antiSpam.totalCount}
        error={overview.antiSpam.error}
      >
        {overview.antiSpam.policies.map(p => (
          <PolicyCard
            key={p.identity}
            title={p.name}
            isDefault={p.isDefault}
            enabled={p.enabled}
            lastChanged={p.whenChanged}
          >
            <SettingRow label="Spam action">
              <ActionBadge action={p.spamAction} />
            </SettingRow>
            <SettingRow label="High-confidence spam">
              <ActionBadge action={p.highConfidenceSpamAction} />
            </SettingRow>
            <SettingRow label="Phish action">
              <ActionBadge action={p.phishSpamAction} />
            </SettingRow>
            <SettingRow label="High-confidence phish">
              <ActionBadge action={p.highConfidencePhishAction} />
            </SettingRow>
            <SettingRow label="Bulk mail action">
              <ActionBadge action={p.bulkSpamAction} />
            </SettingRow>
            <SettingRow label="Bulk threshold (BCL)">
              <span className={`text-xs font-medium ${p.bulkThreshold <= 5
                ? 'text-green-600 dark:text-green-400'
                : p.bulkThreshold <= 7
                  ? 'text-amber-600 dark:text-amber-400'
                  : 'text-red-600 dark:text-red-400'}`}>
                {p.bulkThreshold} / 9
              </span>
            </SettingRow>
            <BoolRow label="Zero-hour auto purge (ZAP)" value={p.zapEnabled} />
            <BoolRow label="Spam ZAP" value={p.spamZapEnabled} />
            <BoolRow label="Phish ZAP" value={p.phishZapEnabled} />
            <BoolRow label="End-user spam notifications" value={p.enableEndUserSpamNotifications} />
            {p.enableEndUserSpamNotifications && (
              <SettingRow label="Notification frequency">
                <span className="text-xs text-slate-600 dark:text-slate-300">Every {p.endUserSpamNotificationFrequency} day(s)</span>
              </SettingRow>
            )}
            {p.allowedSenderDomainsPresent && (
              <SettingRow label="Allowed sender domains">
                <span className="inline-flex items-center gap-1 text-xs text-amber-600 dark:text-amber-400">
                  <WarningRegular className="w-3 h-3" /> Configured
                </span>
              </SettingRow>
            )}
            {p.blockedSenderDomainsPresent && (
              <SettingRow label="Blocked sender domains">
                <span className="inline-flex items-center gap-1 text-xs text-green-600 dark:text-green-400">
                  <CheckmarkCircleRegular className="w-3 h-3" /> Configured
                </span>
              </SettingRow>
            )}
          </PolicyCard>
        ))}
      </Section>

      {/* ── Outbound Spam ──────────────────────────────────────────────────── */}
      <Section
        icon={<MailRegular className="w-5 h-5" />}
        title="Outbound Spam"
        subtitle="Outbound spam filter and auto-forwarding controls"
        count={overview.outboundSpam.totalCount}
        error={overview.outboundSpam.error}
      >
        {overview.outboundSpam.policies.map(p => (
          <PolicyCard
            key={p.identity}
            title={p.name}
            isDefault={p.isDefault}
            enabled={p.enabled}
            lastChanged={p.whenChanged}
          >
            <SettingRow label="Action when threshold reached">
              <ActionBadge action={p.actionWhenThresholdReached} />
            </SettingRow>
            <SettingRow label="Auto-forwarding mode">
              <span className={`text-xs font-medium
                ${p.autoForwardingModeValue === 'Off'
                  ? 'text-green-600 dark:text-green-400'
                  : 'text-amber-600 dark:text-amber-400'}`}>
                {friendlyAutoForward(p.autoForwardingModeValue)}
              </span>
            </SettingRow>
            <SettingRow label="External recipient limit / hour">
              <span className="text-xs text-slate-600 dark:text-slate-300">
                {p.recipientLimitExternalPerHour === 0 ? 'Unlimited' : p.recipientLimitExternalPerHour}
              </span>
            </SettingRow>
            <SettingRow label="Internal recipient limit / hour">
              <span className="text-xs text-slate-600 dark:text-slate-300">
                {p.recipientLimitInternalPerHour === 0 ? 'Unlimited' : p.recipientLimitInternalPerHour}
              </span>
            </SettingRow>
            <SettingRow label="Recipient limit / day">
              <span className="text-xs text-slate-600 dark:text-slate-300">
                {p.recipientLimitPerDay === 0 ? 'Unlimited' : p.recipientLimitPerDay}
              </span>
            </SettingRow>
          </PolicyCard>
        ))}
      </Section>

      {/* ── Safe Attachments ───────────────────────────────────────────────── */}
      <Section
        icon={<DocumentRegular className="w-5 h-5" />}
        title="Safe Attachments"
        subtitle="Defender for Office 365 P1 — attachment sandboxing"
        count={overview.safeAttachments.totalCount}
        error={overview.safeAttachments.isLicensed ? overview.safeAttachments.error : null}
        unlicensed={!overview.safeAttachments.isLicensed}
      >
        {overview.safeAttachments.policies.length === 0 && overview.safeAttachments.isLicensed && (
          <div className="flex items-center gap-2 text-sm text-amber-600 dark:text-amber-400">
            <WarningRegular className="w-4 h-4" /> No Safe Attachments policies found — attachments are not scanned.
          </div>
        )}
        {overview.safeAttachments.policies.map(p => (
          <PolicyCard
            key={p.identity}
            title={p.name}
            isDefault={p.isDefault}
            enabled={p.enabled}
            lastChanged={p.whenChanged}
          >
            <SettingRow label="Action">
              <ActionBadge action={p.action} />
            </SettingRow>
            <BoolRow label="Action on detection error" value={p.actionOnError} />
            <BoolRow label="Redirect malicious attachments" value={p.redirect} />
            {p.redirect && p.redirectAddress && (
              <SettingRow label="Redirect address">
                <span className="text-xs text-slate-600 dark:text-slate-300 truncate max-w-48">{p.redirectAddress}</span>
              </SettingRow>
            )}
          </PolicyCard>
        ))}
      </Section>

      {/* ── Safe Links ─────────────────────────────────────────────────────── */}
      <Section
        icon={<LinkRegular className="w-5 h-5" />}
        title="Safe Links"
        subtitle="Defender for Office 365 P1 — URL rewriting and click protection"
        count={overview.safeLinks.totalCount}
        error={overview.safeLinks.isLicensed ? overview.safeLinks.error : null}
        unlicensed={!overview.safeLinks.isLicensed}
      >
        {overview.safeLinks.policies.length === 0 && overview.safeLinks.isLicensed && (
          <div className="flex items-center gap-2 text-sm text-amber-600 dark:text-amber-400">
            <WarningRegular className="w-4 h-4" /> No Safe Links policies found — URLs are not being scanned.
          </div>
        )}
        {overview.safeLinks.policies.map(p => (
          <PolicyCard
            key={p.identity}
            title={p.name}
            isDefault={p.isDefault}
            enabled={p.enabled}
            lastChanged={p.whenChanged}
          >
            <BoolRow label="Safe Links for email" value={p.enableSafeLinksForEmail} />
            <BoolRow label="Safe Links for Teams" value={p.enableSafeLinksForTeams} />
            <BoolRow label="Safe Links for Office apps" value={p.enableSafeLinksForOffice} />
            <BoolRow label="Scan URLs" value={p.scanUrls} />
            <BoolRow label="Track user clicks" value={p.trackClicks} />
            <BoolRow label="Allow click-through" value={p.allowClickThrough} goodWhenTrue={false} />
            <BoolRow label="URL rewrite disabled" value={p.disableUrlRewrite} goodWhenTrue={false} />
            <BoolRow label="Apply for internal senders" value={p.enableForInternalSenders} />
            {p.doNotRewriteUrls.length > 0 && (
              <SettingRow label="Do-not-rewrite URLs">
                <span className="text-xs text-amber-600 dark:text-amber-400">
                  {p.doNotRewriteUrls.length} exception(s)
                </span>
              </SettingRow>
            )}
          </PolicyCard>
        ))}
      </Section>



    </div>
  );
};

export default DefenderForOfficePage;
