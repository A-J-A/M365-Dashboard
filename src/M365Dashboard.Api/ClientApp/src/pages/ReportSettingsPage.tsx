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
} from '@fluentui/react-icons';
import { useAppContext } from '../contexts/AppContext';

interface ReportSettings {
  companyName: string;
  reportTitle: string;
  logoBase64: string | null;
  logoContentType: string | null;
  primaryColor: string;
  accentColor: string;
  showInfoGraphics: boolean;
  footerText: string | null;
  updatedAt: string;
}

const ReportSettingsPage: React.FC = () => {
  const { getAccessToken } = useAppContext();
  const [settings, setSettings] = useState<ReportSettings>({
    companyName: 'M365 Dashboard',
    reportTitle: 'Microsoft 365 Security Assessment',
    logoBase64: null,
    logoContentType: null,
    primaryColor: '#0078d4',
    accentColor: '#e07c3a',
    showInfoGraphics: true,
    footerText: null,
    updatedAt: new Date().toISOString(),
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
          <h2 className="text-lg font-semibold text-slate-900 dark:text-white mb-4">
            Report Options
          </h2>
          
          <label className="flex items-center gap-3 cursor-pointer">
            <input
              type="checkbox"
              checked={settings.showInfoGraphics}
              onChange={(e) => setSettings(prev => ({ ...prev, showInfoGraphics: e.target.checked }))}
              className="w-5 h-5 rounded border-slate-300 dark:border-slate-600 text-blue-600 focus:ring-blue-500"
            />
            <div>
              <span className="text-slate-900 dark:text-white font-medium">Include Infographic Pages</span>
              <p className="text-sm text-slate-500 dark:text-slate-400">
                Show statistics and quotes between report sections (recommended for client-facing reports)
              </p>
            </div>
          </label>
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
