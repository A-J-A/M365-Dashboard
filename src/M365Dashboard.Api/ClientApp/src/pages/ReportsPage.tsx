import React, { useState, useEffect } from 'react';
import {
  DocumentRegular,
  CalendarRegular,
  ArrowDownloadRegular,
  AddRegular,
  DeleteRegular,
  EditRegular,
  PlayRegular,
  ClockRegular,
  CheckmarkCircleRegular,
  DismissCircleRegular,
  ChevronDownRegular,
  ChevronUpRegular,
  FilterRegular,
  MailRegular,
  ShieldCheckmarkRegular,
  LaptopRegular,
  PeopleRegular,
} from '@fluentui/react-icons';
import { useAppContext } from '../contexts/AppContext';

interface ReportDefinition {
  reportType: string;
  displayName: string;
  description: string;
  category: string;
  availableFormats: string[];
  supportsDateRange: boolean;
  supportsScheduling: boolean;
}

interface ReportSummary {
  totalRecords: number;
  highlights?: Record<string, unknown>;
}

interface ReportResult {
  reportType: string;
  displayName: string;
  generatedAt: string;
  dateRange: string;
  data: unknown;
  summary?: ReportSummary;
}

interface ScheduledReport {
  id: string;
  reportType: string;
  displayName: string;
  schedule: string;
  frequency: string;
  recipients: string[];
  dateRange?: string;
  isEnabled: boolean;
  lastRunAt?: string;
  nextRunAt?: string;
  createdAt: string;
}

interface ReportHistory {
  id: string;
  reportType: string;
  displayName: string;
  generatedAt: string;
  status: string;
  errorMessage?: string;
  recordCount?: number;
  wasScheduled: boolean;
}

const categoryIcons: Record<string, React.ReactNode> = {
  'Usage': <DocumentRegular className="w-5 h-5" />,
  'Security': <ShieldCheckmarkRegular className="w-5 h-5" />,
  'Devices': <LaptopRegular className="w-5 h-5" />,
  'Identity': <PeopleRegular className="w-5 h-5" />,
  'Executive': <DocumentRegular className="w-5 h-5" />,
};

const categoryColors: Record<string, string> = {
  'Usage': 'bg-blue-100 text-blue-600 dark:bg-blue-900/30 dark:text-blue-400',
  'Executive': 'bg-indigo-100 text-indigo-600 dark:bg-indigo-900/30 dark:text-indigo-400',
  'Security': 'bg-red-100 text-red-600 dark:bg-red-900/30 dark:text-red-400',
  'Devices': 'bg-green-100 text-green-600 dark:bg-green-900/30 dark:text-green-400',
  'Identity': 'bg-purple-100 text-purple-600 dark:bg-purple-900/30 dark:text-purple-400',
};

const ReportsPage: React.FC = () => {
  const { getAccessToken } = useAppContext();
  const [definitions, setDefinitions] = useState<ReportDefinition[]>([]);
  const [scheduledReports, setScheduledReports] = useState<ScheduledReport[]>([]);
  const [reportHistory, setReportHistory] = useState<ReportHistory[]>([]);
  const [loading, setLoading] = useState(true);
  const [generatingReport, setGeneratingReport] = useState<string | null>(null);
  const [selectedCategory, setSelectedCategory] = useState<string>('all');
  const [showScheduleModal, setShowScheduleModal] = useState(false);
  const [selectedReportForSchedule, setSelectedReportForSchedule] = useState<ReportDefinition | null>(null);
  const [expandedSections, setExpandedSections] = useState<Record<string, boolean>>({
    available: true,
    scheduled: true,
    history: false,
  });

  // Schedule form state
  const [scheduleForm, setScheduleForm] = useState({
    frequency: 'weekly',
    time: '08:00',
    dayOfWeek: 1,
    dayOfMonth: 1,
    recipients: '',
    dateRange: 'last30days',
  });

  useEffect(() => {
    fetchData();
  }, []);

  const fetchData = async () => {
    try {
      setLoading(true);
      const token = await getAccessToken();
      const headers = {
        'Authorization': `Bearer ${token}`,
        'Content-Type': 'application/json',
      };

      const [defsRes, schedulesRes, historyRes] = await Promise.all([
        fetch('/api/reports/definitions', { headers }),
        fetch('/api/reports/schedules', { headers }),
        fetch('/api/reports/history?take=10', { headers }),
      ]);

      if (defsRes.ok) {
        const defs = await defsRes.json();
        setDefinitions(defs);
      }
      if (schedulesRes.ok) {
        const schedules = await schedulesRes.json();
        setScheduledReports(schedules);
      }
      if (historyRes.ok) {
        const history = await historyRes.json();
        setReportHistory(history);
      }
    } catch (error) {
      console.error('Error fetching report data:', error);
    } finally {
      setLoading(false);
    }
  };

  const generateReport = async (reportType: string, format: string = 'json', dateRange: string = 'last30days') => {
    try {
      setGeneratingReport(reportType);
      const token = await getAccessToken();

      if (format === 'pdf') {
        const response = await fetch('/api/reports/download', {
          method: 'POST',
          headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' },
          body: JSON.stringify({ reportType, dateRange, format }),
        });
        if (response.ok) {
          const blob = await response.blob();
          const url = window.URL.createObjectURL(blob);
          const a = document.createElement('a');
          a.href = url;
          a.download = `Executive_Summary_${new Date().toISOString().split('T')[0]}.pdf`;
          document.body.appendChild(a); a.click();
          window.URL.revokeObjectURL(url); a.remove();
        }
        return;
      }
      
      if (format === 'csv' || format === 'html') {
        const response = await fetch('/api/reports/export', {
          method: 'POST',
          headers: {
            'Authorization': `Bearer ${token}`,
            'Content-Type': 'application/json',
          },
          body: JSON.stringify({ reportType, dateRange, format }),
        });

        if (response.ok) {
          const blob = await response.blob();
          const url = window.URL.createObjectURL(blob);
          
          if (format === 'html') {
            // Open HTML report in new tab for viewing
            window.open(url, '_blank');
          } else {
            // Download CSV file
            const a = document.createElement('a');
            a.href = url;
            a.download = `${reportType}_${new Date().toISOString().split('T')[0]}.${format}`;
            document.body.appendChild(a);
            a.click();
            window.URL.revokeObjectURL(url);
            a.remove();
          }
        }
      } else {
        const response = await fetch('/api/reports/generate', {
          method: 'POST',
          headers: {
            'Authorization': `Bearer ${token}`,
            'Content-Type': 'application/json',
          },
          body: JSON.stringify({ reportType, dateRange, format }),
        });

        if (response.ok) {
          const result: ReportResult = await response.json();
          // Open in new modal/tab to view
          const jsonStr = JSON.stringify(result, null, 2);
          const blob = new Blob([jsonStr], { type: 'application/json' });
          const url = window.URL.createObjectURL(blob);
          window.open(url, '_blank');
        }
      }
    } catch (error) {
      console.error('Error generating report:', error);
    } finally {
      setGeneratingReport(null);
    }
  };

  const createSchedule = async () => {
    if (!selectedReportForSchedule) return;

    try {
      const token = await getAccessToken();
      const response = await fetch('/api/reports/schedules', {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${token}`,
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          reportType: selectedReportForSchedule.reportType,
          frequency: scheduleForm.frequency,
          time: scheduleForm.time,
          dayOfWeek: scheduleForm.frequency === 'weekly' ? scheduleForm.dayOfWeek : null,
          dayOfMonth: scheduleForm.frequency === 'monthly' ? scheduleForm.dayOfMonth : null,
          recipients: scheduleForm.recipients.split(',').map(r => r.trim()).filter(r => r),
          dateRange: selectedReportForSchedule.supportsDateRange ? scheduleForm.dateRange : null,
        }),
      });

      if (response.ok) {
        const newSchedule = await response.json();
        setScheduledReports([...scheduledReports, newSchedule]);
        setShowScheduleModal(false);
        setSelectedReportForSchedule(null);
      }
    } catch (error) {
      console.error('Error creating schedule:', error);
    }
  };

  const deleteSchedule = async (scheduleId: string) => {
    if (!confirm('Are you sure you want to delete this scheduled report?')) return;

    try {
      const token = await getAccessToken();
      const response = await fetch(`/api/reports/schedules/${scheduleId}`, {
        method: 'DELETE',
        headers: {
          'Authorization': `Bearer ${token}`,
        },
      });

      if (response.ok) {
        setScheduledReports(scheduledReports.filter(s => s.id !== scheduleId));
      }
    } catch (error) {
      console.error('Error deleting schedule:', error);
    }
  };

  const toggleSection = (section: string) => {
    setExpandedSections(prev => ({ ...prev, [section]: !prev[section] }));
  };

  const categories = ['all', ...new Set(definitions.map(d => d.category))];
  const filteredDefinitions = selectedCategory === 'all' 
    ? definitions 
    : definitions.filter(d => d.category === selectedCategory);

  const groupedDefinitions = filteredDefinitions.reduce((acc, def) => {
    if (!acc[def.category]) acc[def.category] = [];
    acc[def.category].push(def);
    return acc;
  }, {} as Record<string, ReportDefinition[]>);

  if (loading) {
    return (
      <div className="p-4 flex items-center justify-center h-64">
        <div className="animate-spin rounded-full h-8 w-8 border-b-2 border-blue-600"></div>
      </div>
    );
  }

  return (
    <div className="p-4 space-y-6 max-w-7xl mx-auto">
      {/* Header */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4">
        <div>
          <h1 className="text-2xl font-bold text-slate-900 dark:text-white">Reports</h1>
          <p className="text-sm text-slate-500 dark:text-slate-400">
            Generate and schedule Microsoft 365 tenant reports
          </p>
        </div>
      </div>

      {/* Category Filter */}
      <div className="flex items-center gap-2 overflow-x-auto pb-2">
        <FilterRegular className="w-5 h-5 text-slate-400 flex-shrink-0" />
        {categories.map(cat => (
          <button
            key={cat}
            onClick={() => setSelectedCategory(cat)}
            className={`px-3 py-1.5 rounded-full text-sm font-medium whitespace-nowrap transition-colors ${
              selectedCategory === cat
                ? 'bg-blue-600 text-white'
                : 'bg-slate-100 text-slate-600 hover:bg-slate-200 dark:bg-slate-700 dark:text-slate-300 dark:hover:bg-slate-600'
            }`}
          >
            {cat === 'all' ? 'All Reports' : cat}
          </button>
        ))}
      </div>

      {/* Available Reports Section */}
      <section className="bg-white dark:bg-slate-800 rounded-xl border border-slate-200 dark:border-slate-700 overflow-hidden">
        <button
          onClick={() => toggleSection('available')}
          className="w-full px-6 py-4 flex items-center justify-between hover:bg-slate-50 dark:hover:bg-slate-700/50 transition-colors"
        >
          <div className="flex items-center gap-3">
            <DocumentRegular className="w-5 h-5 text-slate-600 dark:text-slate-400" />
            <div className="text-left">
              <h2 className="font-semibold text-slate-900 dark:text-white">Available Reports</h2>
              <p className="text-sm text-slate-500 dark:text-slate-400">
                {filteredDefinitions.length} reports available
              </p>
            </div>
          </div>
          {expandedSections.available ? (
            <ChevronUpRegular className="w-5 h-5 text-slate-400" />
          ) : (
            <ChevronDownRegular className="w-5 h-5 text-slate-400" />
          )}
        </button>

        {expandedSections.available && (
          <div className="border-t border-slate-200 dark:border-slate-700">
            {Object.entries(groupedDefinitions).map(([category, reports]) => (
              <div key={category} className="border-b border-slate-100 dark:border-slate-700 last:border-b-0">
                <div className="px-6 py-3 bg-slate-50 dark:bg-slate-700/50 flex items-center gap-2">
                  <span className={`p-1.5 rounded ${categoryColors[category] || 'bg-slate-100 text-slate-600'}`}>
                    {categoryIcons[category] || <DocumentRegular className="w-4 h-4" />}
                  </span>
                  <span className="font-medium text-slate-700 dark:text-slate-300">{category}</span>
                  <span className="text-xs text-slate-400">({reports.length})</span>
                </div>
                <div className="divide-y divide-slate-100 dark:divide-slate-700">
                  {reports.map(report => (
                    <ReportCard
                      key={report.reportType}
                      report={report}
                      isGenerating={generatingReport === report.reportType}
                      onGenerate={(format, dateRange) => generateReport(report.reportType, format, dateRange)}
                      onSchedule={() => {
                        setSelectedReportForSchedule(report);
                        setShowScheduleModal(true);
                      }}
                    />
                  ))}
                </div>
              </div>
            ))}
          </div>
        )}
      </section>

      {/* Scheduled Reports Section */}
      <section className="bg-white dark:bg-slate-800 rounded-xl border border-slate-200 dark:border-slate-700 overflow-hidden">
        <button
          onClick={() => toggleSection('scheduled')}
          className="w-full px-6 py-4 flex items-center justify-between hover:bg-slate-50 dark:hover:bg-slate-700/50 transition-colors"
        >
          <div className="flex items-center gap-3">
            <CalendarRegular className="w-5 h-5 text-slate-600 dark:text-slate-400" />
            <div className="text-left">
              <h2 className="font-semibold text-slate-900 dark:text-white">Scheduled Reports</h2>
              <p className="text-sm text-slate-500 dark:text-slate-400">
                {scheduledReports.length} scheduled
              </p>
            </div>
          </div>
          {expandedSections.scheduled ? (
            <ChevronUpRegular className="w-5 h-5 text-slate-400" />
          ) : (
            <ChevronDownRegular className="w-5 h-5 text-slate-400" />
          )}
        </button>

        {expandedSections.scheduled && (
          <div className="border-t border-slate-200 dark:border-slate-700">
            {scheduledReports.length === 0 ? (
              <div className="px-6 py-8 text-center">
                <CalendarRegular className="w-12 h-12 mx-auto text-slate-300 dark:text-slate-600 mb-3" />
                <p className="text-slate-500 dark:text-slate-400">No scheduled reports yet</p>
                <p className="text-sm text-slate-400 dark:text-slate-500 mt-1">
                  Click the calendar icon on any report to schedule it
                </p>
              </div>
            ) : (
              <div className="divide-y divide-slate-100 dark:divide-slate-700">
                {scheduledReports.map(schedule => (
                  <div key={schedule.id} className="px-6 py-4 flex items-center justify-between">
                    <div className="flex-1 min-w-0">
                      <div className="flex items-center gap-2">
                        <p className="font-medium text-slate-900 dark:text-white truncate">
                          {schedule.displayName}
                        </p>
                        <span className={`px-2 py-0.5 rounded text-xs ${
                          schedule.isEnabled 
                            ? 'bg-green-100 text-green-700 dark:bg-green-900/30 dark:text-green-400'
                            : 'bg-slate-100 text-slate-500 dark:bg-slate-700 dark:text-slate-400'
                        }`}>
                          {schedule.isEnabled ? 'Active' : 'Paused'}
                        </span>
                      </div>
                      <div className="flex items-center gap-4 mt-1 text-sm text-slate-500 dark:text-slate-400">
                        <span className="flex items-center gap-1">
                          <ClockRegular className="w-4 h-4" />
                          {schedule.schedule}
                        </span>
                        <span className="flex items-center gap-1">
                          <MailRegular className="w-4 h-4" />
                          {schedule.recipients.length} recipient{schedule.recipients.length !== 1 ? 's' : ''}
                        </span>
                      </div>
                      {schedule.nextRunAt && (
                        <p className="text-xs text-slate-400 mt-1">
                          Next run: {new Date(schedule.nextRunAt).toLocaleString()}
                        </p>
                      )}
                    </div>
                    <div className="flex items-center gap-2">
                      <button
                        onClick={() => deleteSchedule(schedule.id)}
                        className="p-2 text-slate-400 hover:text-red-600 hover:bg-red-50 dark:hover:bg-red-900/20 rounded-lg transition-colors"
                        title="Delete schedule"
                      >
                        <DeleteRegular className="w-5 h-5" />
                      </button>
                    </div>
                  </div>
                ))}
              </div>
            )}
          </div>
        )}
      </section>

      {/* Report History Section */}
      <section className="bg-white dark:bg-slate-800 rounded-xl border border-slate-200 dark:border-slate-700 overflow-hidden">
        <button
          onClick={() => toggleSection('history')}
          className="w-full px-6 py-4 flex items-center justify-between hover:bg-slate-50 dark:hover:bg-slate-700/50 transition-colors"
        >
          <div className="flex items-center gap-3">
            <ClockRegular className="w-5 h-5 text-slate-600 dark:text-slate-400" />
            <div className="text-left">
              <h2 className="font-semibold text-slate-900 dark:text-white">Report History</h2>
              <p className="text-sm text-slate-500 dark:text-slate-400">
                Recent report generations
              </p>
            </div>
          </div>
          {expandedSections.history ? (
            <ChevronUpRegular className="w-5 h-5 text-slate-400" />
          ) : (
            <ChevronDownRegular className="w-5 h-5 text-slate-400" />
          )}
        </button>

        {expandedSections.history && (
          <div className="border-t border-slate-200 dark:border-slate-700">
            {reportHistory.length === 0 ? (
              <div className="px-6 py-8 text-center">
                <ClockRegular className="w-12 h-12 mx-auto text-slate-300 dark:text-slate-600 mb-3" />
                <p className="text-slate-500 dark:text-slate-400">No report history yet</p>
              </div>
            ) : (
              <div className="divide-y divide-slate-100 dark:divide-slate-700">
                {reportHistory.map(history => (
                  <div key={history.id} className="px-6 py-3 flex items-center justify-between">
                    <div className="flex items-center gap-3">
                      {history.status === 'success' ? (
                        <CheckmarkCircleRegular className="w-5 h-5 text-green-500" />
                      ) : (
                        <DismissCircleRegular className="w-5 h-5 text-red-500" />
                      )}
                      <div>
                        <p className="font-medium text-slate-900 dark:text-white">
                          {history.displayName}
                        </p>
                        <p className="text-xs text-slate-500 dark:text-slate-400">
                          {new Date(history.generatedAt).toLocaleString()}
                          {history.recordCount && ` • ${history.recordCount} records`}
                          {history.wasScheduled && ' • Scheduled'}
                        </p>
                      </div>
                    </div>
                  </div>
                ))}
              </div>
            )}
          </div>
        )}
      </section>

      {/* Schedule Modal */}
      {showScheduleModal && selectedReportForSchedule && (
        <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50 p-4">
          <div className="bg-white dark:bg-slate-800 rounded-xl shadow-xl max-w-md w-full max-h-[90vh] overflow-y-auto">
            <div className="px-6 py-4 border-b border-slate-200 dark:border-slate-700">
              <h3 className="text-lg font-semibold text-slate-900 dark:text-white">
                Schedule Report
              </h3>
              <p className="text-sm text-slate-500 dark:text-slate-400">
                {selectedReportForSchedule.displayName}
              </p>
            </div>

            <div className="p-6 space-y-4">
              {/* Frequency */}
              <div>
                <label className="block text-sm font-medium text-slate-700 dark:text-slate-300 mb-1">
                  Frequency
                </label>
                <select
                  value={scheduleForm.frequency}
                  onChange={e => setScheduleForm({ ...scheduleForm, frequency: e.target.value })}
                  className="w-full px-3 py-2 border border-slate-300 dark:border-slate-600 rounded-lg bg-white dark:bg-slate-700 text-slate-900 dark:text-white"
                >
                  <option value="daily">Daily</option>
                  <option value="weekly">Weekly</option>
                  <option value="monthly">Monthly</option>
                </select>
              </div>

              {/* Time */}
              <div>
                <label className="block text-sm font-medium text-slate-700 dark:text-slate-300 mb-1">
                  Time (UTC)
                </label>
                <input
                  type="time"
                  value={scheduleForm.time}
                  onChange={e => setScheduleForm({ ...scheduleForm, time: e.target.value })}
                  className="w-full px-3 py-2 border border-slate-300 dark:border-slate-600 rounded-lg bg-white dark:bg-slate-700 text-slate-900 dark:text-white"
                />
              </div>

              {/* Day of Week (for weekly) */}
              {scheduleForm.frequency === 'weekly' && (
                <div>
                  <label className="block text-sm font-medium text-slate-700 dark:text-slate-300 mb-1">
                    Day of Week
                  </label>
                  <select
                    value={scheduleForm.dayOfWeek}
                    onChange={e => setScheduleForm({ ...scheduleForm, dayOfWeek: Number(e.target.value) })}
                    className="w-full px-3 py-2 border border-slate-300 dark:border-slate-600 rounded-lg bg-white dark:bg-slate-700 text-slate-900 dark:text-white"
                  >
                    <option value={0}>Sunday</option>
                    <option value={1}>Monday</option>
                    <option value={2}>Tuesday</option>
                    <option value={3}>Wednesday</option>
                    <option value={4}>Thursday</option>
                    <option value={5}>Friday</option>
                    <option value={6}>Saturday</option>
                  </select>
                </div>
              )}

              {/* Day of Month (for monthly) */}
              {scheduleForm.frequency === 'monthly' && (
                <div>
                  <label className="block text-sm font-medium text-slate-700 dark:text-slate-300 mb-1">
                    Day of Month
                  </label>
                  <select
                    value={scheduleForm.dayOfMonth}
                    onChange={e => setScheduleForm({ ...scheduleForm, dayOfMonth: Number(e.target.value) })}
                    className="w-full px-3 py-2 border border-slate-300 dark:border-slate-600 rounded-lg bg-white dark:bg-slate-700 text-slate-900 dark:text-white"
                  >
                    {Array.from({ length: 28 }, (_, i) => (
                      <option key={i + 1} value={i + 1}>{i + 1}</option>
                    ))}
                  </select>
                </div>
              )}

              {/* Date Range (if supported) */}
              {selectedReportForSchedule.supportsDateRange && (
                <div>
                  <label className="block text-sm font-medium text-slate-700 dark:text-slate-300 mb-1">
                    Date Range
                  </label>
                  <select
                    value={scheduleForm.dateRange}
                    onChange={e => setScheduleForm({ ...scheduleForm, dateRange: e.target.value })}
                    className="w-full px-3 py-2 border border-slate-300 dark:border-slate-600 rounded-lg bg-white dark:bg-slate-700 text-slate-900 dark:text-white"
                  >
                    <option value="last7days">Last 7 days</option>
                    <option value="last30days">Last 30 days</option>
                    <option value="last90days">Last 90 days</option>
                  </select>
                </div>
              )}

              {/* Recipients */}
              <div>
                <label className="block text-sm font-medium text-slate-700 dark:text-slate-300 mb-1">
                  Email Recipients
                </label>
                <input
                  type="text"
                  value={scheduleForm.recipients}
                  onChange={e => setScheduleForm({ ...scheduleForm, recipients: e.target.value })}
                  placeholder="email1@example.com, email2@example.com"
                  className="w-full px-3 py-2 border border-slate-300 dark:border-slate-600 rounded-lg bg-white dark:bg-slate-700 text-slate-900 dark:text-white placeholder-slate-400"
                />
                <p className="text-xs text-slate-500 dark:text-slate-400 mt-1">
                  Separate multiple emails with commas. Your email will be added automatically.
                </p>
              </div>
            </div>

            <div className="px-6 py-4 border-t border-slate-200 dark:border-slate-700 flex justify-end gap-3">
              <button
                onClick={() => {
                  setShowScheduleModal(false);
                  setSelectedReportForSchedule(null);
                }}
                className="px-4 py-2 text-slate-600 dark:text-slate-400 hover:bg-slate-100 dark:hover:bg-slate-700 rounded-lg transition-colors"
              >
                Cancel
              </button>
              <button
                onClick={createSchedule}
                className="px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 transition-colors flex items-center gap-2"
              >
                <AddRegular className="w-4 h-4" />
                Create Schedule
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

// Report Card Component
interface ReportCardProps {
  report: ReportDefinition;
  isGenerating: boolean;
  onGenerate: (format: string, dateRange: string) => void;
  onSchedule: () => void;
}

const ReportCard: React.FC<ReportCardProps> = ({ report, isGenerating, onGenerate, onSchedule }) => {
  const [dateRange, setDateRange] = useState('last30days');
  const supportsHtml = report.availableFormats.includes('html');
  const supportsPdf  = report.availableFormats.includes('pdf');
  const pdfOnly      = supportsPdf && report.availableFormats.length === 1;

  return (
    <div className="px-6 py-4 hover:bg-slate-50 dark:hover:bg-slate-700/30 transition-colors">
      <div className="flex items-start justify-between gap-4">
        <div className="flex-1 min-w-0">
          <h3 className="font-medium text-slate-900 dark:text-white">{report.displayName}</h3>
          <p className="text-sm text-slate-500 dark:text-slate-400 mt-0.5">{report.description}</p>
        </div>
        <div className="flex items-center gap-2 flex-shrink-0">
          {report.supportsDateRange && (
            <select
              value={dateRange}
              onChange={e => setDateRange(e.target.value)}
              className="px-2 py-1.5 text-sm border border-slate-300 dark:border-slate-600 rounded bg-white dark:bg-slate-700 text-slate-700 dark:text-slate-300"
            >
              <option value="last7days">7 days</option>
              <option value="last30days">30 days</option>
              <option value="last90days">90 days</option>
            </select>
          )}
          {pdfOnly ? (
            // PDF-only report: single prominent download button
            <button
              onClick={() => onGenerate('pdf', dateRange)}
              disabled={isGenerating}
              className="flex items-center gap-1.5 px-3 py-1.5 text-sm font-medium text-white bg-blue-600 hover:bg-blue-700 rounded-lg transition-colors disabled:opacity-50"
              title="Download PDF"
            >
              {isGenerating ? (
                <div className="w-4 h-4 border-2 border-white border-t-transparent rounded-full animate-spin" />
              ) : (
                <ArrowDownloadRegular className="w-4 h-4" />
              )}
              Download PDF
            </button>
          ) : (
            <>
              {supportsHtml && (
                <button
                  onClick={() => onGenerate('html', dateRange)}
                  disabled={isGenerating}
                  className="p-2 text-slate-600 hover:text-orange-600 hover:bg-orange-50 dark:text-slate-400 dark:hover:text-orange-400 dark:hover:bg-orange-900/20 rounded-lg transition-colors disabled:opacity-50"
                  title="View HTML Report"
                >
                  <DocumentRegular className="w-5 h-5" />
                </button>
              )}
              <button
                onClick={() => onGenerate('csv', dateRange)}
                disabled={isGenerating}
                className="p-2 text-slate-600 hover:text-blue-600 hover:bg-blue-50 dark:text-slate-400 dark:hover:text-blue-400 dark:hover:bg-blue-900/20 rounded-lg transition-colors disabled:opacity-50"
                title="Download CSV"
              >
                <ArrowDownloadRegular className="w-5 h-5" />
              </button>
              <button
                onClick={() => onGenerate('json', dateRange)}
                disabled={isGenerating}
                className="p-2 text-slate-600 hover:text-green-600 hover:bg-green-50 dark:text-slate-400 dark:hover:text-green-400 dark:hover:bg-green-900/20 rounded-lg transition-colors disabled:opacity-50"
                title="Generate Report"
              >
                {isGenerating ? (
                  <div className="w-5 h-5 border-2 border-current border-t-transparent rounded-full animate-spin" />
                ) : (
                  <PlayRegular className="w-5 h-5" />
                )}
              </button>
            </>
          )}
          {report.supportsScheduling && (
            <button
              onClick={onSchedule}
              className="p-2 text-slate-600 hover:text-purple-600 hover:bg-purple-50 dark:text-slate-400 dark:hover:text-purple-400 dark:hover:bg-purple-900/20 rounded-lg transition-colors"
              title="Schedule Report"
            >
              <CalendarRegular className="w-5 h-5" />
            </button>
          )}
        </div>
      </div>
    </div>
  );
};

export default ReportsPage;
