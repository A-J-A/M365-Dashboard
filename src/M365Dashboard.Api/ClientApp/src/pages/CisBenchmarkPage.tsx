import React, { useState, useEffect, useCallback } from 'react';
import {
  ShieldCheckmarkRegular,
  ArrowDownloadRegular,
  ArrowSyncRegular,
  CheckmarkCircleFilled,
  DismissCircleFilled,
  WarningRegular,
  InfoRegular,
  OpenRegular,
  FilterRegular,
  ChevronDownRegular,
  ChevronRightRegular,
} from '@fluentui/react-icons';
import { useAppContext } from '../contexts/AppContext';

// Types
interface CisControlResult {
  controlId: string;
  title: string;
  description: string;
  rationale: string;
  category: string;
  subCategory: string;
  level: 'L1' | 'L2';
  profile: 'E3' | 'E5';
  status: 'Pass' | 'Fail' | 'Manual' | 'NotApplicable' | 'Error' | 'Unknown';
  statusReason: string;
  currentValue: string;
  expectedValue: string;
  remediation: string;
  impact: string;
  reference: string;
  isAutomated: boolean;
  affectedItems: string[];
}

interface CisCategoryResult {
  categoryId: string;
  categoryName: string;
  totalControls: number;
  passedControls: number;
  failedControls: number;
  manualControls: number;
  compliancePercentage: number;
}

interface CisBenchmarkResult {
  reportTitle: string;
  benchmarkVersion: string;
  generatedAt: string;
  tenantId: string;
  tenantName: string;
  totalControls: number;
  passedControls: number;
  failedControls: number;
  manualControls: number;
  notApplicableControls: number;
  errorControls: number;
  compliancePercentage: number;
  level1Total: number;
  level1Passed: number;
  level2Total: number;
  level2Passed: number;
  categories: CisCategoryResult[];
  controls: CisControlResult[];
}

const CisBenchmarkPage: React.FC = () => {
  const { getAccessToken } = useAppContext();
  const [benchmarkData, setBenchmarkData] = useState<CisBenchmarkResult | null>(null);
  const [loading, setLoading] = useState(false);
  const [downloading, setDownloading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [includeLevel2, setIncludeLevel2] = useState(true);
  const [includeE5, setIncludeE5] = useState(true);
  const [statusFilter, setStatusFilter] = useState<string>('all');
  const [categoryFilter, setCategoryFilter] = useState<string>('all');
  const [expandedControls, setExpandedControls] = useState<Set<string>>(new Set());

  const runBenchmark = useCallback(async () => {
    try {
      setLoading(true);
      setError(null);
      const token = await getAccessToken();
      
      const response = await fetch('/api/cisbenchmark/run', {
        method: 'GET',
        headers: {
          'Authorization': `Bearer ${token}`,
          'Content-Type': 'application/json',
        },
      });

      if (!response.ok) {
        throw new Error('Failed to run benchmark assessment');
      }

      const data: CisBenchmarkResult = await response.json();
      setBenchmarkData(data);
    } catch (err) {
      setError(err instanceof Error ? err.message : 'An error occurred');
    } finally {
      setLoading(false);
    }
  }, [getAccessToken]);

  const downloadReport = async (format: 'docx' | 'html') => {
    try {
      setDownloading(true);
      const token = await getAccessToken();
      
      const endpoint = format === 'docx' ? '/api/cisbenchmark/download' : '/api/cisbenchmark/html';
      const response = await fetch(`${endpoint}?includeLevel2=${includeLevel2}&includeE5=${includeE5}`, {
        headers: {
          'Authorization': `Bearer ${token}`,
        },
      });

      if (!response.ok) {
        throw new Error('Failed to download report');
      }

      if (format === 'html') {
        const html = await response.text();
        const blob = new Blob([html], { type: 'text/html' });
        const url = window.URL.createObjectURL(blob);
        window.open(url, '_blank');
      } else {
        const blob = await response.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `CIS_M365_Benchmark_Report_${new Date().toISOString().split('T')[0]}.docx`;
        document.body.appendChild(a);
        a.click();
        window.URL.revokeObjectURL(url);
        document.body.removeChild(a);
      }
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Failed to download report');
    } finally {
      setDownloading(false);
    }
  };

  useEffect(() => {
    runBenchmark();
  }, []);

  const toggleControl = (controlId: string) => {
    setExpandedControls(prev => {
      const next = new Set(prev);
      if (next.has(controlId)) {
        next.delete(controlId);
      } else {
        next.add(controlId);
      }
      return next;
    });
  };

  const getStatusIcon = (status: string) => {
    switch (status) {
      case 'Pass':
        return <CheckmarkCircleFilled className="w-5 h-5 text-green-500" />;
      case 'Fail':
        return <DismissCircleFilled className="w-5 h-5 text-red-500" />;
      case 'Manual':
        return <WarningRegular className="w-5 h-5 text-amber-500" />;
      default:
        return <InfoRegular className="w-5 h-5 text-slate-400" />;
    }
  };

  const getStatusBadgeClass = (status: string) => {
    switch (status) {
      case 'Pass':
        return 'bg-green-100 text-green-800 dark:bg-green-900/30 dark:text-green-400';
      case 'Fail':
        return 'bg-red-100 text-red-800 dark:bg-red-900/30 dark:text-red-400';
      case 'Manual':
        return 'bg-amber-100 text-amber-800 dark:bg-amber-900/30 dark:text-amber-400';
      default:
        return 'bg-slate-100 text-slate-600 dark:bg-slate-700 dark:text-slate-400';
    }
  };

  const getLevelBadgeClass = (level: string) => {
    return level === 'L1' 
      ? 'bg-blue-100 text-blue-800 dark:bg-blue-900/30 dark:text-blue-400'
      : 'bg-purple-100 text-purple-800 dark:bg-purple-900/30 dark:text-purple-400';
  };

  const getScoreColor = (score: number) => {
    if (score >= 80) return 'text-green-600 dark:text-green-400';
    if (score >= 60) return 'text-amber-600 dark:text-amber-400';
    return 'text-red-600 dark:text-red-400';
  };

  const filteredControls = benchmarkData?.controls.filter(c => {
    if (statusFilter !== 'all' && c.status !== statusFilter) return false;
    if (categoryFilter !== 'all' && !c.category.startsWith(categoryFilter)) return false;
    return true;
  }) || [];

  // Compliance Score Ring
  const ComplianceRing: React.FC<{ percentage: number }> = ({ percentage }) => {
    const radius = 70;
    const strokeWidth = 12;
    const normalizedRadius = radius - strokeWidth / 2;
    const circumference = normalizedRadius * 2 * Math.PI;
    const strokeDashoffset = circumference - (percentage / 100) * circumference;

    const getColor = () => {
      if (percentage >= 80) return '#107C10';
      if (percentage >= 60) return '#FFB900';
      return '#D13438';
    };

    return (
      <div className="relative inline-flex items-center justify-center">
        <svg height={radius * 2} width={radius * 2} className="transform -rotate-90">
          <circle
            stroke="currentColor"
            className="text-slate-200 dark:text-slate-700"
            fill="transparent"
            strokeWidth={strokeWidth}
            r={normalizedRadius}
            cx={radius}
            cy={radius}
          />
          <circle
            stroke={getColor()}
            fill="transparent"
            strokeWidth={strokeWidth}
            strokeDasharray={circumference + ' ' + circumference}
            style={{ strokeDashoffset, transition: 'stroke-dashoffset 0.5s ease' }}
            strokeLinecap="round"
            r={normalizedRadius}
            cx={radius}
            cy={radius}
          />
        </svg>
        <div className="absolute flex flex-col items-center justify-center">
          <span className={`text-3xl font-bold ${getScoreColor(percentage)}`}>
            {percentage}%
          </span>
          <span className="text-xs text-slate-500 dark:text-slate-400">Compliance</span>
        </div>
      </div>
    );
  };

  // Summary Card
  const SummaryCard: React.FC<{
    title: string;
    value: number;
    color: 'green' | 'red' | 'amber' | 'blue' | 'default';
    icon: React.ReactNode;
  }> = ({ title, value, color, icon }) => {
    const colorClasses = {
      green: 'border-l-green-500 text-green-600 dark:text-green-400',
      red: 'border-l-red-500 text-red-600 dark:text-red-400',
      amber: 'border-l-amber-500 text-amber-600 dark:text-amber-400',
      blue: 'border-l-blue-500 text-blue-600 dark:text-blue-400',
      default: 'border-l-slate-500 text-slate-600 dark:text-slate-400',
    };

    return (
      <div className={`bg-white dark:bg-slate-800 rounded-lg border border-slate-200 dark:border-slate-700 border-l-4 ${colorClasses[color].split(' ')[0]} p-4`}>
        <div className="flex items-center gap-3">
          {icon}
          <div>
            <p className="text-sm text-slate-500 dark:text-slate-400">{title}</p>
            <p className={`text-2xl font-bold ${colorClasses[color].split(' ').slice(1).join(' ')}`}>{value}</p>
          </div>
        </div>
      </div>
    );
  };

  // Control Card
  const ControlCard: React.FC<{ control: CisControlResult }> = ({ control }) => {
    const isExpanded = expandedControls.has(control.controlId);

    return (
      <div className={`bg-white dark:bg-slate-800 rounded-lg border ${
        control.status === 'Fail' ? 'border-l-4 border-l-red-500' :
        control.status === 'Pass' ? 'border-l-4 border-l-green-500' :
        control.status === 'Manual' ? 'border-l-4 border-l-amber-500' :
        ''
      } border-slate-200 dark:border-slate-700 overflow-hidden`}>
        <button
          onClick={() => toggleControl(control.controlId)}
          className="w-full px-4 py-3 flex items-center justify-between hover:bg-slate-50 dark:hover:bg-slate-700/50 transition-colors"
        >
          <div className="flex items-center gap-3">
            {getStatusIcon(control.status)}
            <div className="text-left">
              <span className="text-blue-600 dark:text-blue-400 font-mono text-sm">{control.controlId}</span>
              <span className="mx-2 text-slate-300 dark:text-slate-600">|</span>
              <span className="font-medium text-slate-900 dark:text-white">{control.title}</span>
            </div>
          </div>
          <div className="flex items-center gap-2">
            <span className={`px-2 py-0.5 rounded text-xs font-medium ${getLevelBadgeClass(control.level)}`}>
              {control.level}
            </span>
            <span className={`px-2 py-0.5 rounded text-xs font-medium ${getStatusBadgeClass(control.status)}`}>
              {control.status}
            </span>
            {isExpanded ? (
              <ChevronDownRegular className="w-5 h-5 text-slate-400" />
            ) : (
              <ChevronRightRegular className="w-5 h-5 text-slate-400" />
            )}
          </div>
        </button>

        {isExpanded && (
          <div className="px-4 py-4 border-t border-slate-200 dark:border-slate-700 bg-slate-50 dark:bg-slate-900/50 space-y-4">
            {/* Description */}
            <div>
              <dt className="font-medium text-slate-700 dark:text-slate-300 text-sm mb-1">Description</dt>
              <dd className="text-sm text-slate-600 dark:text-slate-400">{control.description}</dd>
            </div>

            {/* Rationale */}
            {control.rationale && (
              <div className="p-3 bg-slate-100 dark:bg-slate-800 rounded-lg">
                <dt className="font-medium text-slate-700 dark:text-slate-300 text-sm mb-1">Rationale</dt>
                <dd className="text-sm text-slate-600 dark:text-slate-400">{control.rationale}</dd>
              </div>
            )}

            {/* Impact */}
            {control.impact && (
              <div className="p-3 bg-amber-50 dark:bg-amber-900/20 rounded-lg">
                <dt className="font-medium text-amber-800 dark:text-amber-300 text-sm mb-1">Impact</dt>
                <dd className="text-sm text-amber-700 dark:text-amber-400">{control.impact}</dd>
              </div>
            )}

            {/* Assessment Result */}
            <div className="p-3 bg-white dark:bg-slate-800 rounded-lg border border-slate-200 dark:border-slate-700">
              <dt className="font-medium text-slate-700 dark:text-slate-300 text-sm mb-2">Assessment Result</dt>
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4 text-sm">
                <div>
                  <dt className="text-xs uppercase tracking-wide text-slate-500 dark:text-slate-400 mb-1">Current Value</dt>
                  <dd className="text-slate-700 dark:text-slate-300 font-medium">{control.currentValue || 'N/A'}</dd>
                </div>
                <div>
                  <dt className="text-xs uppercase tracking-wide text-slate-500 dark:text-slate-400 mb-1">Expected Value</dt>
                  <dd className="text-slate-700 dark:text-slate-300 font-medium">{control.expectedValue || 'N/A'}</dd>
                </div>
              </div>
              {control.statusReason && (
                <div className="mt-3 pt-3 border-t border-slate-200 dark:border-slate-700">
                  <dt className="text-xs uppercase tracking-wide text-slate-500 dark:text-slate-400 mb-1">Status Reason</dt>
                  <dd className="text-slate-600 dark:text-slate-400 text-sm">{control.statusReason}</dd>
                </div>
              )}
            </div>

            {/* Checked Items - Show for all statuses to allow verification */}
            {control.affectedItems && control.affectedItems.length > 0 && (
              <div className={`p-3 rounded-lg ${
                control.status === 'Fail' 
                  ? 'bg-red-50 dark:bg-red-900/20' 
                  : control.status === 'Pass'
                  ? 'bg-green-50 dark:bg-green-900/20'
                  : 'bg-slate-100 dark:bg-slate-800'
              }`}>
                <dt className={`font-medium text-sm mb-2 ${
                  control.status === 'Fail' 
                    ? 'text-red-800 dark:text-red-300' 
                    : control.status === 'Pass'
                    ? 'text-green-800 dark:text-green-300'
                    : 'text-slate-700 dark:text-slate-300'
                }`}>
                  {control.status === 'Fail' ? 'Affected Items' : 'Checked Items'} ({control.affectedItems.length})
                </dt>
                <dd className={`text-sm ${
                  control.status === 'Fail' 
                    ? 'text-red-700 dark:text-red-400' 
                    : control.status === 'Pass'
                    ? 'text-green-700 dark:text-green-400'
                    : 'text-slate-600 dark:text-slate-400'
                }`}>
                  <ul className="list-disc list-inside space-y-1">
                    {control.affectedItems.slice(0, 10).map((item, idx) => (
                      <li key={idx}>{item}</li>
                    ))}
                    {control.affectedItems.length > 10 && (
                      <li className="text-slate-500">...and {control.affectedItems.length - 10} more</li>
                    )}
                  </ul>
                </dd>
              </div>
            )}

            {/* Remediation */}
            {(control.status === 'Fail' || control.status === 'Manual') && control.remediation && (
              <div className="p-3 bg-blue-50 dark:bg-blue-900/20 rounded-lg">
                <dt className="font-medium text-blue-800 dark:text-blue-300 text-sm">Remediation</dt>
                <dd className="text-blue-700 dark:text-blue-400 text-sm mt-1">{control.remediation}</dd>
              </div>
            )}

            {/* Reference Link */}
            {control.reference && (
              <div>
                <a
                  href={control.reference}
                  target="_blank"
                  rel="noopener noreferrer"
                  className="text-blue-600 hover:text-blue-700 text-sm flex items-center gap-1"
                >
                  Learn more <OpenRegular className="w-3 h-3" />
                </a>
              </div>
            )}
          </div>
        )}
      </div>
    );
  };

  return (
    <div className="p-4 space-y-4 w-full max-w-full overflow-hidden">
      {/* Header */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
        <div>
          <h1 className="text-xl font-semibold text-slate-900 dark:text-white flex items-center gap-2">
            <ShieldCheckmarkRegular className="w-6 h-6 text-blue-600" />
            CIS Microsoft 365 Benchmark
          </h1>
          <p className="text-sm text-slate-500 dark:text-slate-400">
            Security compliance assessment based on CIS Foundations Benchmark v6.0.0
          </p>
        </div>
        <div className="flex items-center gap-2">
          <button
            onClick={runBenchmark}
            disabled={loading}
            className="px-3 py-2 bg-slate-100 dark:bg-slate-700 text-slate-700 dark:text-slate-300 rounded-lg hover:bg-slate-200 dark:hover:bg-slate-600 transition-colors disabled:opacity-50 flex items-center gap-2"
          >
            <ArrowSyncRegular className={`w-4 h-4 ${loading ? 'animate-spin' : ''}`} />
            {loading ? 'Running...' : 'Run Assessment'}
          </button>
          <button
            onClick={() => downloadReport('docx')}
            disabled={downloading || !benchmarkData}
            className="inline-flex items-center gap-2 px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 transition-colors disabled:opacity-50"
          >
            <ArrowDownloadRegular className="w-4 h-4" />
            <span>Download Report</span>
          </button>
        </div>
      </div>

      {/* Loading State */}
      {loading && !benchmarkData && (
        <div className="flex items-center justify-center h-64">
          <div className="text-center">
            <div className="animate-spin rounded-full h-8 w-8 border-b-2 border-blue-600 mx-auto mb-4"></div>
            <p className="text-slate-500 dark:text-slate-400">Running CIS benchmark assessment...</p>
            <p className="text-sm text-slate-400 dark:text-slate-500 mt-1">This may take a few moments</p>
          </div>
        </div>
      )}

      {/* Error State */}
      {error && (
        <div className="bg-red-50 dark:bg-red-900/20 border border-red-200 dark:border-red-800 rounded-lg p-4">
          <p className="text-red-700 dark:text-red-400">{error}</p>
          <button onClick={runBenchmark} className="mt-2 text-sm text-red-600 dark:text-red-400 underline">
            Retry
          </button>
        </div>
      )}

      {/* Results */}
      {benchmarkData && (
        <>
          {/* Report Header */}
          <div className="bg-gradient-to-r from-blue-600 to-indigo-600 rounded-lg p-6 text-white">
            <div className="flex items-center justify-between">
              <div>
                <h2 className="text-2xl font-bold">CIS Microsoft 365 Foundations Benchmark</h2>
                <p className="text-blue-100 mt-1">Version {benchmarkData.benchmarkVersion}</p>
                <p className="text-sm text-blue-200 mt-2">
                  Tenant: {benchmarkData.tenantName || benchmarkData.tenantId || 'Connected Tenant'}
                </p>
                <p className="text-sm text-blue-200">
                  Generated: {new Date(benchmarkData.generatedAt).toLocaleString()}
                </p>
              </div>
              <ComplianceRing percentage={benchmarkData.compliancePercentage} />
            </div>
          </div>

          {/* Summary Cards */}
          <div className="grid grid-cols-2 md:grid-cols-5 gap-3">
            <SummaryCard
              title="Total Controls"
              value={benchmarkData.totalControls}
              color="blue"
              icon={<ShieldCheckmarkRegular className="w-6 h-6 text-blue-600" />}
            />
            <SummaryCard
              title="Passed"
              value={benchmarkData.passedControls}
              color="green"
              icon={<CheckmarkCircleFilled className="w-6 h-6 text-green-500" />}
            />
            <SummaryCard
              title="Failed"
              value={benchmarkData.failedControls}
              color="red"
              icon={<DismissCircleFilled className="w-6 h-6 text-red-500" />}
            />
            <SummaryCard
              title="Manual Review"
              value={benchmarkData.manualControls}
              color="amber"
              icon={<WarningRegular className="w-6 h-6 text-amber-500" />}
            />
            <SummaryCard
              title="Not Applicable"
              value={benchmarkData.notApplicableControls}
              color="default"
              icon={<InfoRegular className="w-6 h-6 text-slate-400" />}
            />
          </div>

          {/* Level Breakdown */}
          <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
            <div className="bg-white dark:bg-slate-800 rounded-lg border border-slate-200 dark:border-slate-700 p-4">
              <h3 className="font-semibold text-slate-900 dark:text-white mb-3">Level 1 (Essential Security)</h3>
              <div className="flex items-center justify-between mb-2">
                <span className="text-sm text-slate-600 dark:text-slate-400">
                  {benchmarkData.level1Passed} / {benchmarkData.level1Total} passed
                </span>
                <span className={`font-bold ${getScoreColor(benchmarkData.level1Total > 0 ? Math.round(benchmarkData.level1Passed / benchmarkData.level1Total * 100) : 0)}`}>
                  {benchmarkData.level1Total > 0 ? Math.round(benchmarkData.level1Passed / benchmarkData.level1Total * 100) : 0}%
                </span>
              </div>
              <div className="w-full bg-slate-200 dark:bg-slate-700 rounded-full h-2">
                <div
                  className="h-2 rounded-full bg-blue-500 transition-all"
                  style={{ width: `${benchmarkData.level1Total > 0 ? (benchmarkData.level1Passed / benchmarkData.level1Total * 100) : 0}%` }}
                />
              </div>
            </div>
            <div className="bg-white dark:bg-slate-800 rounded-lg border border-slate-200 dark:border-slate-700 p-4">
              <h3 className="font-semibold text-slate-900 dark:text-white mb-3">Level 2 (Defense in Depth)</h3>
              <div className="flex items-center justify-between mb-2">
                <span className="text-sm text-slate-600 dark:text-slate-400">
                  {benchmarkData.level2Passed} / {benchmarkData.level2Total} passed
                </span>
                <span className={`font-bold ${getScoreColor(benchmarkData.level2Total > 0 ? Math.round(benchmarkData.level2Passed / benchmarkData.level2Total * 100) : 0)}`}>
                  {benchmarkData.level2Total > 0 ? Math.round(benchmarkData.level2Passed / benchmarkData.level2Total * 100) : 0}%
                </span>
              </div>
              <div className="w-full bg-slate-200 dark:bg-slate-700 rounded-full h-2">
                <div
                  className="h-2 rounded-full bg-purple-500 transition-all"
                  style={{ width: `${benchmarkData.level2Total > 0 ? (benchmarkData.level2Passed / benchmarkData.level2Total * 100) : 0}%` }}
                />
              </div>
            </div>
          </div>

          {/* Category Breakdown */}
          <div className="bg-white dark:bg-slate-800 rounded-lg border border-slate-200 dark:border-slate-700 overflow-hidden">
            <div className="px-4 py-3 border-b border-slate-200 dark:border-slate-700">
              <h3 className="font-semibold text-slate-900 dark:text-white">Compliance by Category</h3>
            </div>
            <div className="overflow-x-auto">
              <table className="w-full text-sm">
                <thead className="bg-slate-50 dark:bg-slate-900">
                  <tr>
                    <th className="px-4 py-3 text-left font-medium text-slate-600 dark:text-slate-400">Category</th>
                    <th className="px-4 py-3 text-center font-medium text-slate-600 dark:text-slate-400">Passed</th>
                    <th className="px-4 py-3 text-center font-medium text-slate-600 dark:text-slate-400">Failed</th>
                    <th className="px-4 py-3 text-center font-medium text-slate-600 dark:text-slate-400">Manual</th>
                    <th className="px-4 py-3 text-right font-medium text-slate-600 dark:text-slate-400">Score</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-200 dark:divide-slate-700">
                  {benchmarkData.categories.map((cat) => (
                    <tr key={cat.categoryId} className="hover:bg-slate-50 dark:hover:bg-slate-700/50">
                      <td className="px-4 py-3 text-slate-700 dark:text-slate-300">{cat.categoryName}</td>
                      <td className="px-4 py-3 text-center text-green-600 dark:text-green-400 font-medium">{cat.passedControls}</td>
                      <td className="px-4 py-3 text-center text-red-600 dark:text-red-400 font-medium">{cat.failedControls}</td>
                      <td className="px-4 py-3 text-center text-amber-600 dark:text-amber-400 font-medium">{cat.manualControls}</td>
                      <td className="px-4 py-3 text-right">
                        <span className={`font-bold ${getScoreColor(cat.compliancePercentage)}`}>
                          {cat.compliancePercentage}%
                        </span>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>

          {/* Filters */}
          <div className="flex flex-wrap items-center gap-3 bg-white dark:bg-slate-800 rounded-lg border border-slate-200 dark:border-slate-700 p-3">
            <FilterRegular className="w-5 h-5 text-slate-400" />
            <select
              value={statusFilter}
              onChange={(e) => setStatusFilter(e.target.value)}
              className="bg-slate-100 dark:bg-slate-700 border-0 rounded-lg px-3 py-2 text-sm text-slate-700 dark:text-slate-300"
            >
              <option value="all">All Statuses</option>
              <option value="Pass">Passed</option>
              <option value="Fail">Failed</option>
              <option value="Manual">Manual Review</option>
            </select>
            <select
              value={categoryFilter}
              onChange={(e) => setCategoryFilter(e.target.value)}
              className="bg-slate-100 dark:bg-slate-700 border-0 rounded-lg px-3 py-2 text-sm text-slate-700 dark:text-slate-300"
            >
              <option value="all">All Categories</option>
              {benchmarkData.categories.map((cat) => (
                <option key={cat.categoryId} value={cat.categoryId}>
                  {cat.categoryName}
                </option>
              ))}
            </select>
            <span className="text-sm text-slate-500 dark:text-slate-400 ml-auto">
              Showing {filteredControls.length} of {benchmarkData.controls.length} controls
            </span>
          </div>

          {/* Controls List */}
          <div className="space-y-2">
            {filteredControls.map((control) => (
              <ControlCard key={control.controlId} control={control} />
            ))}
          </div>

          {/* Footer */}
          <div className="bg-slate-50 dark:bg-slate-800/50 border border-slate-200 dark:border-slate-700 rounded-lg p-4">
            <div className="flex items-start gap-3">
              <InfoRegular className="w-5 h-5 text-slate-500 dark:text-slate-400 flex-shrink-0 mt-0.5" />
              <div className="text-sm text-slate-600 dark:text-slate-400">
                <p>This assessment is based on the CIS Microsoft 365 Foundations Benchmark v6.0.0.</p>
                <p className="mt-1">Some controls require manual verification in the respective admin centers.</p>
                <a
                  href="https://www.cisecurity.org/benchmark/microsoft_365"
                  target="_blank"
                  rel="noopener noreferrer"
                  className="text-blue-600 hover:text-blue-700 flex items-center gap-1 mt-2"
                >
                  View full CIS Benchmark documentation <OpenRegular className="w-3 h-3" />
                </a>
              </div>
            </div>
          </div>
        </>
      )}
    </div>
  );
};

export default CisBenchmarkPage;
