import React, { useState, useEffect, useCallback, useRef } from 'react';
import {
  SettingsRegular,
  SaveRegular,
  ImageRegular,
  DeleteRegular,
  CheckmarkCircleFilled,
  DismissCircleFilled,
  ArrowSyncRegular,
  ColorRegular,
  DocumentRegular,
  ChatRegular,
  AddRegular,
  EditRegular,
} from '@fluentui/react-icons';
import { useAppContext } from '../contexts/AppContext';

interface ReportQuote {
  bigNumber: string;
  line1: string;
  line2: string;
  source: string;
  enabled: boolean;
}

interface ReportSettings {
  companyName: string;
  reportTitle: string;
  logoBase64: string | null;
  logoContentType: string | null;
  primaryColor: string;
  accentColor: string;
  showInfoGraphics: boolean;
  showQuotes: boolean;
  footerText: string | null;
  updatedAt: string;
  quotes: ReportQuote[];
  excludedDomains: string[];
}

const DEFAULT_QUOTES: ReportQuote[] = [
  { bigNumber: '99%',  line1: 'of breaches could be mitigated',       line2: 'with strong passwords and MFA',              source: 'Source: Microsoft Security Report',        enabled: true },
  { bigNumber: '84%',  line1: 'of businesses fell victim',             line2: 'to phishing attacks in 2024',                source: 'Source: Cyber Security Breaches Survey',   enabled: true },
  { bigNumber: '31%',  line1: 'of all breaches over the past',         line2: '10 years involved stolen credentials',      source: 'Source: Verizon DBIR',                     enabled: true },
  { bigNumber: '300%', line1: 'increase in reported cyber incidents',  line2: 'since the start of remote working',         source: 'Source: NCSC Annual Review',               enabled: true },
  { bigNumber: '4.5M', line1: 'average cost of a data breach in 2023', line2: 'a record high for the 13th consecutive year', source: 'Source: IBM Cost of a Data Breach Report', enabled: true },
  { bigNumber: '11s',  line1: 'a business falls victim to ransomware', line2: 'every 11 seconds globally',                 source: 'Source: Cybersecurity Ventures',           enabled: true },
  { bigNumber: '74%',  line1: 'of all breaches include',               line2: 'a human element',                          source: 'Source: Verizon DBIR 2023',                enabled: true },
  { bigNumber: '85%',  line1: 'of organisations have experienced',     line2: 'at least one cloud data breach',            source: 'Source: Thales Cloud Security Study',      enabled: true },
  { bigNumber: '50%',  line1: 'of SMBs have suffered a cyberattack',  line2: 'and 60% close within 6 months',             source: 'Source: SCORE.org',                        enabled: true },
  { bigNumber: '98%',  line1: 'of cyberattacks can be prevented',      line2: 'by implementing basic cyber hygiene',       source: 'Source: Microsoft Digital Defence Report', enabled: true },
];

// Small controlled input for adding a domain to the exclusion list
const ExcludedDomainInput: React.FC<{ onAdd: (domain: string) => void }> = ({ onAdd }) => {
  const [value, setValue] = React.useState('');
  const submit = () => {
    const trimmed = value.trim().toLowerCase();
    if (trimmed) { onAdd(trimmed); setValue(''); }
  };
  return (
    <div className="flex gap-2">
      <input
        type="text"
        value={value}
        onChange={e => setValue(e.target.value)}
        onKeyDown={e => e.key === 'Enter' && submit()}
        placeholder="e.g. legacy.example.com"
        className="flex-1 px-3 py-1.5 text-sm border border-slate-300 dark:border-slate-600 rounded bg-white dark:bg-slate-700 text-slate-900 dark:text-white placeholder-slate-400"
      />
      <button
        onClick={submit}
        disabled={!value.trim()}
        className="px-3 py-1.5 text-sm bg-blue-600 text-white rounded hover:bg-blue-700 disabled:opacity-40 flex items-center gap-1.5"
      >
        <AddRegular className="w-4 h-4" />
        Add
      </button>
    </div>
  );
};

const ReportSettingsPage: React.FC = () => {
  const { getAccessToken } = useAppContext();
  const [editingQuoteIndex, setEditingQuoteIndex] = useState<number | null>(null);
  const [editingQuote, setEditingQuote] = useState<ReportQuote | null>(null);

  const [settings, setSettings] = useState<ReportSettings>({
    companyName: 'M365 Dashboard',
    reportTitle: 'Microsoft 365 Security Assessment',
    logoBase64: null,
    logoContentType: null,
    primaryColor: '#0078d4',
    accentColor: '#e07c3a',
    showInfoGraphics: true,
    showQuotes: true,
    footerText: null,
    updatedAt: new Date().toISOString(),
    quotes: DEFAULT_QUOTES,
    excludedDomains: [],
  });
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);
  const [uploading, setUploading] = useState(false);
  const [message, setMessage] = useState<{ type: 'success' | 'error'; text: string } | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const loadSettings = useCallback(async () => {
    try {
      setLoading(true);
      const token = await getAccessToken();

      const response = await fetch('/api/settings/report', {
        headers: {
          'Authorization': `Bearer ${token}`,
        },
      });

      if (response.ok) {
        const data = await response.json();
        // Backfill quotes if not yet in DB (first load after upgrade)
        if (!data.quotes || data.quotes.length === 0) {
          data.quotes = DEFAULT_QUOTES;
        }
        if (!data.excludedDomains) {
          data.excludedDomains = [];
        }
        setSettings(data);
      }
    } catch (err) {
      console.error('Error loading settings:', err);
      setMessage({ type: 'error', text: 'Failed to load settings' });
    } finally {
      setLoading(false);
    }
  }, [getAccessToken]);

  const saveSettings = async () => {
    try {
      setSaving(true);
      setMessage(null);
      const token = await getAccessToken();

      const response = await fetch('/api/settings/report', {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${token}`,
          'Content-Type': 'application/json',
        },
        body: JSON.stringify(settings),
      });

      if (response.ok) {
        const data = await response.json();
        setSettings(data);
        setMessage({ type: 'success', text: 'Settings saved successfully' });
      } else {
        throw new Error('Failed to save settings');
      }
    } catch (err) {
      console.error('Error saving settings:', err);
      setMessage({ type: 'error', text: 'Failed to save settings' });
    } finally {
      setSaving(false);
    }
  };

  const uploadLogo = async (file: File) => {
    try {
      setUploading(true);
      setMessage(null);
      const token = await getAccessToken();

      const formData = new FormData();
      formData.append('file', file);

      const response = await fetch('/api/settings/report/logo', {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${token}`,
        },
        body: formData,
      });

      if (response.ok) {
        const data = await response.json();
        setSettings(prev => ({
          ...prev,
          logoBase64: data.logoBase64,
          logoContentType: data.contentType,
        }));
        setMessage({ type: 'success', text: 'Logo uploaded successfully' });
      } else {
        const error = await response.json();
        throw new Error(error.error || 'Failed to upload logo');
      }
    } catch (err) {
      console.error('Error uploading logo:', err);
      setMessage({ type: 'error', text: err instanceof Error ? err.message : 'Failed to upload logo' });
    } finally {
      setUploading(false);
    }
  };

  const removeLogo = async () => {
    try {
      setUploading(true);
      setMessage(null);
      const token = await getAccessToken();

      const response = await fetch('/api/settings/report/logo', {
        method: 'DELETE',
        headers: {
          'Authorization': `Bearer ${token}`,
        },
      });

      if (response.ok) {
        setSettings(prev => ({
          ...prev,
          logoBase64: null,
          logoContentType: null,
        }));
        setMessage({ type: 'success', text: 'Logo removed successfully' });
      } else {
        throw new Error('Failed to remove logo');
      }
    } catch (err) {
      console.error('Error removing logo:', err);
      setMessage({ type: 'error', text: 'Failed to remove logo' });
    } finally {
      setUploading(false);
    }
  };

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      uploadLogo(file);
    }
  };

  useEffect(() => {
    loadSettings();
  }, [loadSettings]);

  if (loading) {
    return (
      <div className="p-6">
        <div className="flex items-center justify-center h-64">
          <div className="text-center">
            <ArrowSyncRegular className="w-8 h-8 animate-spin text-blue-500 mx-auto mb-4" />
            <p className="text-slate-600 dark:text-slate-400">Loading settings...</p>
          </div>
        </div>
      </div>
    );
  }

  return (
    <div className="p-6 max-w-4xl mx-auto">
      {/* Header */}
      <div className="flex items-center gap-3 mb-6">
        <div className="p-2 bg-blue-100 dark:bg-blue-900/30 rounded-lg">
          <SettingsRegular className="w-6 h-6 text-blue-600 dark:text-blue-400" />
        </div>
        <div>
          <h1 className="text-2xl font-bold text-slate-900 dark:text-white">Report Settings</h1>
          <p className="text-slate-600 dark:text-slate-400">Customize the branding for your security reports</p>
        </div>
      </div>

      {/* Message */}
      {message && (
        <div className={`mb-6 p-4 rounded-lg flex items-center gap-3 ${
          message.type === 'success' 
            ? 'bg-green-50 dark:bg-green-900/20 text-green-700 dark:text-green-400' 
            : 'bg-red-50 dark:bg-red-900/20 text-red-700 dark:text-red-400'
        }`}>
          {message.type === 'success' ? (
            <CheckmarkCircleFilled className="w-5 h-5" />
          ) : (
            <DismissCircleFilled className="w-5 h-5" />
          )}
          {message.text}
        </div>
      )}

      <div className="space-y-6">
        {/* Logo Section */}
        <div className="bg-white dark:bg-slate-800 rounded-lg shadow-sm p-6">
          <h2 className="text-lg font-semibold text-slate-900 dark:text-white mb-4 flex items-center gap-2">
            <ImageRegular className="w-5 h-5" />
            Company Logo
          </h2>
          
          <div className="flex items-start gap-6">
            {/* Logo Preview */}
            <div className="w-48 h-32 border-2 border-dashed border-slate-300 dark:border-slate-600 rounded-lg flex items-center justify-center bg-slate-50 dark:bg-slate-700/50 overflow-hidden">
              {settings.logoBase64 ? (
                <img 
                  src={`data:${settings.logoContentType};base64,${settings.logoBase64}`}
                  alt="Company Logo"
                  className="max-w-full max-h-full object-contain"
                />
              ) : (
                <div className="text-center text-slate-400">
                  <ImageRegular className="w-8 h-8 mx-auto mb-2" />
                  <span className="text-sm">No logo uploaded</span>
                </div>
              )}
            </div>

            {/* Upload Controls */}
            <div className="flex-1">
              <p className="text-sm text-slate-600 dark:text-slate-400 mb-4">
                Upload your company logo to appear on the cover page and footer of generated reports.
                Recommended size: 300x100 pixels. Supported formats: PNG, JPEG, GIF, SVG.
              </p>
              
              <input
                type="file"
                ref={fileInputRef}
                onChange={handleFileChange}
                accept="image/png,image/jpeg,image/gif,image/svg+xml"
                className="hidden"
              />
              
              <div className="flex gap-3">
                <button
                  onClick={() => fileInputRef.current?.click()}
                  disabled={uploading}
                  className="px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 transition-colors disabled:opacity-50 flex items-center gap-2"
                >
                  <ImageRegular className="w-4 h-4" />
                  {uploading ? 'Uploading...' : 'Upload Logo'}
                </button>
                
                {settings.logoBase64 && (
                  <button
                    onClick={removeLogo}
                    disabled={uploading}
                    className="px-4 py-2 bg-red-100 dark:bg-red-900/30 text-red-700 dark:text-red-400 rounded-lg hover:bg-red-200 dark:hover:bg-red-900/50 transition-colors disabled:opacity-50 flex items-center gap-2"
                  >
                    <DeleteRegular className="w-4 h-4" />
                    Remove
                  </button>
                )}
              </div>
            </div>
          </div>
        </div>

        {/* Text Settings */}
        <div className="bg-white dark:bg-slate-800 rounded-lg shadow-sm p-6">
          <h2 className="text-lg font-semibold text-slate-900 dark:text-white mb-4 flex items-center gap-2">
            <DocumentRegular className="w-5 h-5" />
            Report Text
          </h2>
          
          <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
            <div>
              <label className="block text-sm font-medium text-slate-700 dark:text-slate-300 mb-2">
                Company Name
              </label>
              <input
                type="text"
                value={settings.companyName}
                onChange={(e) => setSettings(prev => ({ ...prev, companyName: e.target.value }))}
                placeholder="Your Company Name"
                className="w-full px-4 py-2 border border-slate-300 dark:border-slate-600 rounded-lg bg-white dark:bg-slate-700 text-slate-900 dark:text-white focus:ring-2 focus:ring-blue-500 focus:border-transparent"
              />
              <p className="mt-1 text-xs text-slate-500 dark:text-slate-400">
                Displayed in the cover page and footer
              </p>
            </div>
            
            <div>
              <label className="block text-sm font-medium text-slate-700 dark:text-slate-300 mb-2">
                Report Title
              </label>
              <input
                type="text"
                value={settings.reportTitle}
                onChange={(e) => setSettings(prev => ({ ...prev, reportTitle: e.target.value }))}
                placeholder="Security Assessment Report"
                className="w-full px-4 py-2 border border-slate-300 dark:border-slate-600 rounded-lg bg-white dark:bg-slate-700 text-slate-900 dark:text-white focus:ring-2 focus:ring-blue-500 focus:border-transparent"
              />
              <p className="mt-1 text-xs text-slate-500 dark:text-slate-400">
                Main title shown on the cover page
              </p>
            </div>
            
            <div className="md:col-span-2">
              <label className="block text-sm font-medium text-slate-700 dark:text-slate-300 mb-2">
                Footer Text (Optional)
              </label>
              <input
                type="text"
                value={settings.footerText || ''}
                onChange={(e) => setSettings(prev => ({ ...prev, footerText: e.target.value || null }))}
                placeholder="Confidential - For internal use only"
                className="w-full px-4 py-2 border border-slate-300 dark:border-slate-600 rounded-lg bg-white dark:bg-slate-700 text-slate-900 dark:text-white focus:ring-2 focus:ring-blue-500 focus:border-transparent"
              />
              <p className="mt-1 text-xs text-slate-500 dark:text-slate-400">
                Additional text to show in the footer (e.g., confidentiality notice)
              </p>
            </div>
          </div>
        </div>

        {/* Colors */}
        <div className="bg-white dark:bg-slate-800 rounded-lg shadow-sm p-6">
          <h2 className="text-lg font-semibold text-slate-900 dark:text-white mb-4 flex items-center gap-2">
            <ColorRegular className="w-5 h-5" />
            Colors
          </h2>
          
          <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
            <div>
              <label className="block text-sm font-medium text-slate-700 dark:text-slate-300 mb-2">
                Primary Color
              </label>
              <div className="flex items-center gap-3">
                <input
                  type="color"
                  value={settings.primaryColor}
                  onChange={(e) => setSettings(prev => ({ ...prev, primaryColor: e.target.value }))}
                  className="w-12 h-10 rounded cursor-pointer border border-slate-300 dark:border-slate-600"
                />
                <input
                  type="text"
                  value={settings.primaryColor}
                  onChange={(e) => setSettings(prev => ({ ...prev, primaryColor: e.target.value }))}
                  className="flex-1 px-4 py-2 border border-slate-300 dark:border-slate-600 rounded-lg bg-white dark:bg-slate-700 text-slate-900 dark:text-white font-mono text-sm"
                />
              </div>
              <p className="mt-1 text-xs text-slate-500 dark:text-slate-400">
                Used for headers and section backgrounds
              </p>
            </div>
            
            <div>
              <label className="block text-sm font-medium text-slate-700 dark:text-slate-300 mb-2">
                Accent Color
              </label>
              <div className="flex items-center gap-3">
                <input
                  type="color"
                  value={settings.accentColor}
                  onChange={(e) => setSettings(prev => ({ ...prev, accentColor: e.target.value }))}
                  className="w-12 h-10 rounded cursor-pointer border border-slate-300 dark:border-slate-600"
                />
                <input
                  type="text"
                  value={settings.accentColor}
                  onChange={(e) => setSettings(prev => ({ ...prev, accentColor: e.target.value }))}
                  className="flex-1 px-4 py-2 border border-slate-300 dark:border-slate-600 rounded-lg bg-white dark:bg-slate-700 text-slate-900 dark:text-white font-mono text-sm"
                />
              </div>
              <p className="mt-1 text-xs text-slate-500 dark:text-slate-400">
                Used for highlights and emphasis
              </p>
            </div>
          </div>
        </div>

        {/* Options */}
        <div className="bg-white dark:bg-slate-800 rounded-lg shadow-sm p-6">
          <h2 className="text-lg font-semibold text-slate-900 dark:text-white mb-4">Report Options</h2>
          <div className="space-y-4">
            <label className="flex items-center gap-3 cursor-pointer">
              <input type="checkbox" checked={settings.showInfoGraphics}
                onChange={(e) => setSettings(prev => ({ ...prev, showInfoGraphics: e.target.checked }))}
                className="w-5 h-5 rounded border-slate-300 dark:border-slate-600 text-blue-600 focus:ring-blue-500" />
              <div>
                <span className="text-slate-900 dark:text-white font-medium">Include Infographic Pages</span>
                <p className="text-sm text-slate-500 dark:text-slate-400">Show full-page statistic cards between report sections</p>
              </div>
            </label>
            {settings.showInfoGraphics && (
              <label className="flex items-center gap-3 cursor-pointer ml-8">
                <input type="checkbox" checked={settings.showQuotes}
                  onChange={(e) => setSettings(prev => ({ ...prev, showQuotes: e.target.checked }))}
                  className="w-5 h-5 rounded border-slate-300 dark:border-slate-600 text-blue-600 focus:ring-blue-500" />
                <div>
                  <span className="text-slate-900 dark:text-white font-medium">Include Quote Statistics</span>
                  <p className="text-sm text-slate-500 dark:text-slate-400">3 quotes are randomly selected from the pool below each time a report is generated</p>
                </div>
              </label>
            )}
          </div>
        </div>

        {/* Quotes Pool */}
        <div className="bg-white dark:bg-slate-800 rounded-lg shadow-sm p-6">
          <div className="flex items-center justify-between mb-4">
            <h2 className="text-lg font-semibold text-slate-900 dark:text-white flex items-center gap-2">
              <ChatRegular className="w-5 h-5" />
              Quote Pool
            </h2>
            <div className="flex items-center gap-3">
              <span className="text-xs text-slate-500 dark:text-slate-400">
                {settings.quotes.filter(q => q.enabled).length} enabled · 3 selected randomly per report
              </span>
              <button
                onClick={() => { setEditingQuoteIndex(-1); setEditingQuote({ bigNumber: '', line1: '', line2: '', source: '', enabled: true }); }}
                className="flex items-center gap-1.5 px-3 py-1.5 text-sm bg-blue-600 text-white rounded-lg hover:bg-blue-700 transition-colors"
              >
                <AddRegular className="w-4 h-4" /> Add Quote
              </button>
            </div>
          </div>

          <div className="space-y-2">
            {settings.quotes.map((quote, idx) => (
              <div key={idx} className={`border rounded-lg p-3 transition-colors ${
                quote.enabled
                  ? 'border-slate-200 dark:border-slate-700 bg-white dark:bg-slate-800'
                  : 'border-slate-100 dark:border-slate-800 bg-slate-50 dark:bg-slate-900/50 opacity-60'
              }`}>
                {editingQuoteIndex === idx ? (
                  <div className="space-y-2">
                    <div className="grid grid-cols-2 gap-2">
                      <div>
                        <label className="text-xs font-medium text-slate-600 dark:text-slate-400">Big Number / Stat</label>
                        <input type="text" value={editingQuote?.bigNumber ?? ''} placeholder="e.g. 99%"
                          onChange={e => setEditingQuote(q => q ? { ...q, bigNumber: e.target.value } : q)}
                          className="w-full mt-0.5 px-3 py-1.5 text-sm border border-slate-300 dark:border-slate-600 rounded bg-white dark:bg-slate-700 text-slate-900 dark:text-white" />
                      </div>
                      <div>
                        <label className="text-xs font-medium text-slate-600 dark:text-slate-400">Source</label>
                        <input type="text" value={editingQuote?.source ?? ''} placeholder="e.g. Source: Microsoft"
                          onChange={e => setEditingQuote(q => q ? { ...q, source: e.target.value } : q)}
                          className="w-full mt-0.5 px-3 py-1.5 text-sm border border-slate-300 dark:border-slate-600 rounded bg-white dark:bg-slate-700 text-slate-900 dark:text-white" />
                      </div>
                    </div>
                    <div>
                      <label className="text-xs font-medium text-slate-600 dark:text-slate-400">Line 1</label>
                      <input type="text" value={editingQuote?.line1 ?? ''} placeholder="e.g. of breaches could be mitigated"
                        onChange={e => setEditingQuote(q => q ? { ...q, line1: e.target.value } : q)}
                        className="w-full mt-0.5 px-3 py-1.5 text-sm border border-slate-300 dark:border-slate-600 rounded bg-white dark:bg-slate-700 text-slate-900 dark:text-white" />
                    </div>
                    <div>
                      <label className="text-xs font-medium text-slate-600 dark:text-slate-400">Line 2</label>
                      <input type="text" value={editingQuote?.line2 ?? ''} placeholder="e.g. with strong passwords and MFA"
                        onChange={e => setEditingQuote(q => q ? { ...q, line2: e.target.value } : q)}
                        className="w-full mt-0.5 px-3 py-1.5 text-sm border border-slate-300 dark:border-slate-600 rounded bg-white dark:bg-slate-700 text-slate-900 dark:text-white" />
                    </div>
                    <div className="flex gap-2 pt-1">
                      <button onClick={() => {
                        if (editingQuote) {
                          const updated = [...settings.quotes];
                          updated[idx] = editingQuote;
                          setSettings(prev => ({ ...prev, quotes: updated }));
                        }
                        setEditingQuoteIndex(null); setEditingQuote(null);
                      }} className="px-3 py-1.5 text-sm bg-blue-600 text-white rounded hover:bg-blue-700">Save</button>
                      <button onClick={() => { setEditingQuoteIndex(null); setEditingQuote(null); }}
                        className="px-3 py-1.5 text-sm bg-slate-100 dark:bg-slate-700 text-slate-700 dark:text-slate-300 rounded hover:bg-slate-200 dark:hover:bg-slate-600">Cancel</button>
                    </div>
                  </div>
                ) : (
                  <div className="flex items-start gap-3">
                    <input type="checkbox" checked={quote.enabled}
                      onChange={e => {
                        const updated = [...settings.quotes];
                        updated[idx] = { ...updated[idx], enabled: e.target.checked };
                        setSettings(prev => ({ ...prev, quotes: updated }));
                      }}
                      className="mt-1 w-4 h-4 rounded border-slate-300 text-blue-600 cursor-pointer" />
                    <div className="flex-1 min-w-0">
                      <div className="flex items-baseline gap-2 flex-wrap">
                        <span className="text-base font-bold text-blue-600 dark:text-blue-400 font-mono">{quote.bigNumber}</span>
                        <span className="text-sm text-slate-700 dark:text-slate-300">{quote.line1} {quote.line2}</span>
                      </div>
                      <p className="text-xs text-slate-400 dark:text-slate-500 italic mt-0.5">{quote.source}</p>
                    </div>
                    <div className="flex gap-1 flex-shrink-0">
                      <button
                        onClick={() => { setEditingQuoteIndex(idx); setEditingQuote({ ...quote }); }}
                        className="p-1.5 text-slate-400 hover:text-blue-600 dark:hover:text-blue-400 rounded transition-colors"
                        title="Edit"
                      >
                        <EditRegular className="w-4 h-4" />
                      </button>
                      <button
                        onClick={() => {
                          const updated = settings.quotes.filter((_, i) => i !== idx);
                          setSettings(prev => ({ ...prev, quotes: updated }));
                        }}
                        className="p-1.5 text-slate-400 hover:text-red-600 dark:hover:text-red-400 rounded transition-colors"
                        title="Delete"
                      >
                        <DeleteRegular className="w-4 h-4" />
                      </button>
                    </div>
                  </div>
                )}
              </div>
            ))}

            {/* New quote inline form */}
            {editingQuoteIndex === -1 && editingQuote && (
              <div className="border-2 border-blue-300 dark:border-blue-700 border-dashed rounded-lg p-3 space-y-2">
                <p className="text-sm font-medium text-blue-600 dark:text-blue-400">New Quote</p>
                <div className="grid grid-cols-2 gap-2">
                  <div>
                    <label className="text-xs font-medium text-slate-600 dark:text-slate-400">Big Number / Stat</label>
                    <input type="text" value={editingQuote.bigNumber} placeholder="e.g. 99%"
                      onChange={e => setEditingQuote(q => q ? { ...q, bigNumber: e.target.value } : q)}
                      className="w-full mt-0.5 px-3 py-1.5 text-sm border border-slate-300 dark:border-slate-600 rounded bg-white dark:bg-slate-700 text-slate-900 dark:text-white" />
                  </div>
                  <div>
                    <label className="text-xs font-medium text-slate-600 dark:text-slate-400">Source</label>
                    <input type="text" value={editingQuote.source} placeholder="e.g. Source: Microsoft"
                      onChange={e => setEditingQuote(q => q ? { ...q, source: e.target.value } : q)}
                      className="w-full mt-0.5 px-3 py-1.5 text-sm border border-slate-300 dark:border-slate-600 rounded bg-white dark:bg-slate-700 text-slate-900 dark:text-white" />
                  </div>
                </div>
                <div>
                  <label className="text-xs font-medium text-slate-600 dark:text-slate-400">Line 1</label>
                  <input type="text" value={editingQuote.line1} placeholder="e.g. of breaches could be mitigated"
                    onChange={e => setEditingQuote(q => q ? { ...q, line1: e.target.value } : q)}
                    className="w-full mt-0.5 px-3 py-1.5 text-sm border border-slate-300 dark:border-slate-600 rounded bg-white dark:bg-slate-700 text-slate-900 dark:text-white" />
                </div>
                <div>
                  <label className="text-xs font-medium text-slate-600 dark:text-slate-400">Line 2</label>
                  <input type="text" value={editingQuote.line2} placeholder="e.g. with strong passwords and MFA"
                    onChange={e => setEditingQuote(q => q ? { ...q, line2: e.target.value } : q)}
                    className="w-full mt-0.5 px-3 py-1.5 text-sm border border-slate-300 dark:border-slate-600 rounded bg-white dark:bg-slate-700 text-slate-900 dark:text-white" />
                </div>
                <div className="flex gap-2 pt-1">
                  <button onClick={() => {
                    if (editingQuote && editingQuote.bigNumber && editingQuote.line1) {
                      setSettings(prev => ({ ...prev, quotes: [...prev.quotes, { ...editingQuote, enabled: true }] }));
                    }
                    setEditingQuoteIndex(null); setEditingQuote(null);
                  }} className="px-3 py-1.5 text-sm bg-blue-600 text-white rounded hover:bg-blue-700">Add</button>
                  <button onClick={() => { setEditingQuoteIndex(null); setEditingQuote(null); }}
                    className="px-3 py-1.5 text-sm bg-slate-100 dark:bg-slate-700 text-slate-700 dark:text-slate-300 rounded hover:bg-slate-200 dark:hover:bg-slate-600">Cancel</button>
                </div>
              </div>
            )}
          </div>

          <div className="mt-4 flex items-start gap-2 text-xs text-slate-500 dark:text-slate-400 bg-slate-50 dark:bg-slate-700/50 rounded-lg px-3 py-2.5">
            <ChatRegular className="w-4 h-4 flex-shrink-0 mt-0.5" />
            <span>3 quotes are picked at random from the enabled pool each time a report is generated, so monthly reports won't always show the same statistics. Uncheck a quote to exclude it without deleting it.</span>
          </div>
        </div>

        {/* Excluded Domains */}
        <div className="bg-white dark:bg-slate-800 rounded-lg shadow-sm p-6">
          <h2 className="text-lg font-semibold text-slate-900 dark:text-white mb-1">Excluded Domains</h2>
          <p className="text-sm text-slate-500 dark:text-slate-400 mb-4">
            Domains listed here will be hidden from the Domain Security section of the report.
            Useful for internal or legacy domains where SPF/DMARC cannot be configured.
          </p>

          {/* Existing excluded domains */}
          <div className="space-y-2 mb-3">
            {(settings.excludedDomains ?? []).map((domain, i) => (
              <div key={i} className="flex items-center gap-2 px-3 py-2 bg-slate-50 dark:bg-slate-700/50 rounded-lg">
                <span className="flex-1 text-sm font-mono text-slate-700 dark:text-slate-300">{domain}</span>
                <button
                  onClick={() => setSettings(prev => ({
                    ...prev,
                    excludedDomains: (prev.excludedDomains ?? []).filter((_, idx) => idx !== i)
                  }))}
                  className="p-1 text-slate-400 hover:text-red-500 dark:hover:text-red-400 transition-colors"
                  title="Remove"
                >
                  <DeleteRegular className="w-4 h-4" />
                </button>
              </div>
            ))}
            {(settings.excludedDomains ?? []).length === 0 && (
              <p className="text-sm text-slate-400 dark:text-slate-500 italic">No domains excluded.</p>
            )}
          </div>

          {/* Add domain input */}
          <ExcludedDomainInput onAdd={domain => setSettings(prev => ({
            ...prev,
            excludedDomains: [...(prev.excludedDomains ?? []), domain]
          }))} />
        </div>

        {/* Preview */}
        <div className="bg-white dark:bg-slate-800 rounded-lg shadow-sm p-6">
          <h2 className="text-lg font-semibold text-slate-900 dark:text-white mb-4">
            Preview
          </h2>
          
          <div className="border border-slate-200 dark:border-slate-700 rounded-lg overflow-hidden">
            {/* Mini Cover Preview */}
            <div 
              className="h-48 p-6 flex flex-col justify-between text-white"
              style={{ background: `linear-gradient(135deg, ${settings.primaryColor} 0%, #2d3748 100%)` }}
            >
              <div>
                <div className="text-lg font-light tracking-wider uppercase">MICROSOFT 365</div>
                <div 
                  className="text-sm font-light uppercase tracking-widest"
                  style={{ color: settings.accentColor }}
                >
                  {settings.reportTitle.toUpperCase().replace('MICROSOFT 365 ', '')}
                </div>
              </div>
              <div className="flex items-end justify-between">
                <div className="text-sm opacity-80">Sample Tenant Name</div>
                <div className="flex items-center gap-3">
                  {settings.logoBase64 ? (
                    <img 
                      src={`data:${settings.logoContentType};base64,${settings.logoBase64}`}
                      alt="Logo Preview"
                      className="h-8 object-contain"
                    />
                  ) : (
                    <span 
                      className="font-semibold"
                      style={{ color: settings.accentColor }}
                    >
                      {settings.companyName}
                    </span>
                  )}
                </div>
              </div>
            </div>
          </div>
        </div>

        {/* Save Button */}
        <div className="flex justify-end">
          <button
            onClick={saveSettings}
            disabled={saving}
            className="px-6 py-3 bg-blue-600 text-white rounded-lg hover:bg-blue-700 transition-colors disabled:opacity-50 flex items-center gap-2 font-medium"
          >
            <SaveRegular className="w-5 h-5" />
            {saving ? 'Saving...' : 'Save Settings'}
          </button>
        </div>
      </div>
    </div>
  );
};

export default ReportSettingsPage;
