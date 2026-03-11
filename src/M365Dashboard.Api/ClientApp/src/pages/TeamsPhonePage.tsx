import React, { useState, useEffect, useCallback } from 'react';
import {
    Call24Regular,
    CallInbound24Regular,
    CallOutbound24Regular,
    CallMissed24Regular,
    Clock24Regular,
    Person24Regular,
    ArrowSync24Regular,
    Filter24Regular,
    ChartMultiple24Regular,
    CallForward24Regular,
    CallEnd24Regular,
    Checkmark24Regular,
    Dismiss24Regular,
    Warning24Regular,
    Info24Regular,
    ArrowLeft24Regular,
    ArrowRight24Regular,
    Beaker24Regular,
    Money24Regular,
    Globe24Regular,
    Router24Regular,
    Chat24Regular,
} from '@fluentui/react-icons';
import {
    Card,
    Spinner,
    Badge,
    ProgressBar,
    Button,
} from '@fluentui/react-components';
import { BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, LineChart, Line, Legend } from 'recharts';
import { useAppContext } from '../contexts/AppContext';

interface PstnSummary {
    totalCalls: number;
    inboundCalls: number;
    outboundCalls: number;
    answeredCalls: number;
    missedCalls: number;
    totalMinutes: number;
    answerRate: number;
}

interface DirectRoutingSummary {
    totalCalls: number;
    successfulCalls: number;
    failedCalls: number;
    totalMinutes: number;
    successRate: number;
}

interface DailyTrend {
    date: string;
    pstnCalls: number;
    directRoutingCalls: number;
    totalCalls: number;
}

interface TopUser {
    userPrincipalName: string;
    displayName: string;
    callCount: number;
    totalMinutes: number;
}

interface CallsByHour {
    hour: number;
    count: number;
}

interface DashboardData {
    pstn: PstnSummary;
    directRouting: DirectRoutingSummary;
    combined: {
        totalCalls: number;
        totalMinutes: number;
    };
    dailyTrend: DailyTrend[];
    topUsers: TopUser[];
    callsByHour: CallsByHour[];
    period: {
        fromDate: string;
        toDate: string;
    };
    lastUpdated: string;
    error?: string;
}

interface PstnCallDetail {
    id: string;
    userPrincipalName: string;
    userDisplayName: string;
    startDateTime: string;
    endDateTime: string;
    duration: number;
    callType: string;
    calleeNumber: string;
    callerNumber: string;
    destinationContext: string;
    destinationName: string;
    isAnswered: boolean;
    charge: number;
    currency: string;
    connectionCharge: number;
    licenseCapability: string;
    inventoryType: string;
}

interface PstnCallsResponse {
    summary: PstnSummary;
    dailyTrend: any[];
    topCallers: TopUser[];
    callsByHour: CallsByHour[];
    calls?: PstnCallDetail[];
    period: { fromDate: string; toDate: string };
    lastUpdated: string;
    error?: string;
}

type CallFilter = 'all' | 'inbound' | 'outbound' | 'answered' | 'missed';
type ViewTab = 'standard' | 'beta';

const COLORS = ['#3b82f6', '#22c55e', '#f59e0b', '#ef4444', '#8b5cf6', '#ec4899'];

const TeamsPhonePage: React.FC = () => {
    const { getAccessToken } = useAppContext();
    const [dashboardData, setDashboardData] = useState<DashboardData | null>(null);
    const [detailedCalls, setDetailedCalls] = useState<PstnCallDetail[]>([]);
    const [loading, setLoading] = useState(true);
    const [detailLoading, setDetailLoading] = useState(false);
    const [error, setError] = useState<string | null>(null);
    const [days, setDays] = useState(30);
    const [activeFilter, setActiveFilter] = useState<CallFilter | null>(null);
    const [showDetailView, setShowDetailView] = useState(false);
    const [activeTab, setActiveTab] = useState<ViewTab>('standard');

    const fetchDashboardData = useCallback(async () => {
        try {
            setLoading(true);
            setError(null);
            const token = await getAccessToken();
            
            const response = await fetch(`/api/teamsphone/dashboard?days=${days}`, {
                headers: {
                    'Authorization': `Bearer ${token}`,
                    'Content-Type': 'application/json',
                },
            });

            if (!response.ok) {
                throw new Error('Failed to fetch Teams Phone data');
            }

            const data = await response.json();
            setDashboardData(data);
        } catch (err) {
            setError(err instanceof Error ? err.message : 'An error occurred');
        } finally {
            setLoading(false);
        }
    }, [getAccessToken, days]);

    const fetchDetailedCalls = useCallback(async (filter: CallFilter) => {
        try {
            setDetailLoading(true);
            const token = await getAccessToken();
            
            const response = await fetch(`/api/teamsphone/pstn-calls?days=${days}`, {
                headers: {
                    'Authorization': `Bearer ${token}`,
                    'Content-Type': 'application/json',
                },
            });

            if (!response.ok) {
                throw new Error('Failed to fetch call details');
            }

            const data: PstnCallsResponse = await response.json();
            
            // The API returns aggregated data, so we'll simulate call details from the summary
            // In a real implementation, you'd have an endpoint that returns individual call records
            setDetailedCalls(data.calls || []);
            setActiveFilter(filter);
            setShowDetailView(true);
        } catch (err) {
            console.error('Failed to fetch call details:', err);
        } finally {
            setDetailLoading(false);
        }
    }, [getAccessToken, days]);

    useEffect(() => {
        fetchDashboardData();
    }, [fetchDashboardData]);

    const formatDuration = (minutes: number) => {
        if (minutes < 60) return `${Math.round(minutes)} min`;
        const hours = Math.floor(minutes / 60);
        const mins = Math.round(minutes % 60);
        return `${hours}h ${mins}m`;
    };

    const formatSeconds = (seconds: number) => {
        if (seconds < 60) return `${Math.round(seconds)}s`;
        const mins = Math.floor(seconds / 60);
        const secs = Math.round(seconds % 60);
        return `${mins}m ${secs}s`;
    };

    const formatHour = (hour: number) => {
        if (hour === 0) return '12 AM';
        if (hour === 12) return '12 PM';
        if (hour < 12) return `${hour} AM`;
        return `${hour - 12} PM`;
    };

    const handleMetricClick = (filter: CallFilter) => {
        setActiveFilter(filter);
        setShowDetailView(true);
    };

    const handleBackToDashboard = () => {
        setShowDetailView(false);
        setActiveFilter(null);
    };

    if (loading) {
        return (
            <div className="flex items-center justify-center min-h-[400px]">
                <Spinner size="large" label="Loading Teams Phone data..." />
            </div>
        );
    }

    const hasData = dashboardData && (dashboardData.pstn.totalCalls > 0 || dashboardData.directRouting.totalCalls > 0);

    // Detail View
    if (showDetailView && dashboardData) {
        return (
            <CallDetailView
                filter={activeFilter}
                dashboardData={dashboardData}
                days={days}
                onBack={handleBackToDashboard}
                getAccessToken={getAccessToken}
            />
        );
    }

    // Beta View
    if (activeTab === 'beta') {
        return (
            <BetaApiView
                days={days}
                setDays={setDays}
                activeTab={activeTab}
                setActiveTab={setActiveTab}
                getAccessToken={getAccessToken}
            />
        );
    }

    return (
        <div className="p-6 space-y-6">
            {/* Header */}
            <div className="flex items-start justify-between">
                <div>
                    <h1 className="text-2xl font-bold text-gray-900 dark:text-white flex items-center gap-2">
                        <Call24Regular className="w-7 h-7" />
                        Teams Phone System
                    </h1>
                    <p className="mt-1 text-sm text-gray-500 dark:text-gray-400">
                        PSTN and Direct Routing call analytics
                    </p>
                </div>
                <div className="flex items-center gap-3">
                    {/* Tab Selector */}
                    <div className="flex items-center bg-gray-100 dark:bg-gray-700 rounded-lg p-1">
                        <button
                            onClick={() => setActiveTab('standard')}
                            className={`flex items-center gap-1.5 px-3 py-1.5 rounded-md text-sm font-medium transition-colors ${
                                activeTab === 'standard'
                                    ? 'bg-white dark:bg-gray-600 text-gray-900 dark:text-white shadow-sm'
                                    : 'text-gray-600 dark:text-gray-400 hover:text-gray-900 dark:hover:text-white'
                            }`}
                        >
                            <Call24Regular className="w-4 h-4" />
                            Standard
                        </button>
                        <button
                            onClick={() => setActiveTab('beta')}
                            className="flex items-center gap-1.5 px-3 py-1.5 rounded-md text-sm font-medium transition-colors text-gray-600 dark:text-gray-400 hover:text-gray-900 dark:hover:text-white"
                        >
                            <Beaker24Regular className="w-4 h-4" />
                            Beta API
                        </button>
                    </div>

                    <div className="flex items-center gap-2">
                        <Filter24Regular className="w-4 h-4 text-gray-400" />
                        <select
                            value={days}
                            onChange={(e) => setDays(Number(e.target.value))}
                            className="px-3 py-1.5 border border-gray-300 dark:border-gray-600 rounded-lg bg-white dark:bg-gray-700 text-gray-900 dark:text-white text-sm"
                        >
                            <option value={7}>Last 7 days</option>
                            <option value={14}>Last 14 days</option>
                            <option value={30}>Last 30 days</option>
                            <option value={60}>Last 60 days</option>
                            <option value={90}>Last 90 days</option>
                        </select>
                    </div>
                    <button
                        onClick={fetchDashboardData}
                        disabled={loading}
                        className="p-2 text-gray-600 hover:text-blue-600 hover:bg-blue-50 dark:text-gray-400 dark:hover:text-blue-400 dark:hover:bg-blue-900/20 rounded-lg transition-colors disabled:opacity-50"
                        title="Refresh"
                    >
                        <ArrowSync24Regular className={`w-5 h-5 ${loading ? 'animate-spin' : ''}`} />
                    </button>
                </div>
            </div>

            {error && (
                <div className="bg-red-50 dark:bg-red-900/20 border border-red-200 dark:border-red-800 rounded-lg p-4">
                    <p className="text-red-700 dark:text-red-300">{error}</p>
                </div>
            )}

            {dashboardData?.error && (
                <div className="bg-amber-50 dark:bg-amber-900/20 border border-amber-200 dark:border-amber-800 rounded-lg p-4 flex items-start gap-3">
                    <Warning24Regular className="w-5 h-5 text-amber-600 flex-shrink-0 mt-0.5" />
                    <div>
                        <p className="text-amber-700 dark:text-amber-300 font-medium">Limited Data Available</p>
                        <p className="text-sm text-amber-600 dark:text-amber-400 mt-1">{dashboardData.error}</p>
                    </div>
                </div>
            )}

            {!hasData && !error && (
                <div className="bg-blue-50 dark:bg-blue-900/20 border border-blue-200 dark:border-blue-800 rounded-lg p-6 text-center">
                    <Info24Regular className="w-12 h-12 text-blue-500 mx-auto mb-3" />
                    <h3 className="text-lg font-semibold text-blue-800 dark:text-blue-200">No Call Data Available</h3>
                    <p className="text-sm text-blue-600 dark:text-blue-400 mt-2 max-w-md mx-auto">
                        Teams Phone System data requires Microsoft Teams Phone licenses and Call Records API permissions. 
                        Ensure your organization has Teams Phone System configured and the app has the required permissions.
                    </p>
                </div>
            )}

            {hasData && dashboardData && (
                <>
                    {/* Summary Cards - Clickable */}
                    <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
                        <ClickableMetricCard
                            title="Total Calls"
                            value={dashboardData.combined.totalCalls}
                            icon={Call24Regular}
                            color="blue"
                            subtitle={`${formatDuration(dashboardData.combined.totalMinutes)} total duration`}
                            onClick={() => handleMetricClick('all')}
                        />
                        <ClickableMetricCard
                            title="Inbound Calls"
                            value={dashboardData.pstn.inboundCalls}
                            icon={CallInbound24Regular}
                            color="green"
                            subtitle={`${Math.round((dashboardData.pstn.inboundCalls / (dashboardData.pstn.totalCalls || 1)) * 100)}% of PSTN calls`}
                            onClick={() => handleMetricClick('inbound')}
                        />
                        <ClickableMetricCard
                            title="Outbound Calls"
                            value={dashboardData.pstn.outboundCalls}
                            icon={CallOutbound24Regular}
                            color="purple"
                            subtitle={`${Math.round((dashboardData.pstn.outboundCalls / (dashboardData.pstn.totalCalls || 1)) * 100)}% of PSTN calls`}
                            onClick={() => handleMetricClick('outbound')}
                        />
                        <ClickableMetricCard
                            title="Missed Calls"
                            value={dashboardData.pstn.missedCalls}
                            icon={CallMissed24Regular}
                            color={dashboardData.pstn.missedCalls > 0 ? 'red' : 'green'}
                            subtitle={`${(100 - dashboardData.pstn.answerRate).toFixed(1)}% miss rate`}
                            onClick={() => handleMetricClick('missed')}
                            highlight={dashboardData.pstn.missedCalls > 0}
                        />
                    </div>

                    {/* PSTN vs Direct Routing */}
                    <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
                        {/* PSTN Summary */}
                        <Card className="p-5">
                            <div className="flex items-center gap-2 mb-4">
                                <div className="p-2 bg-blue-50 dark:bg-blue-900/30 rounded-lg">
                                    <Call24Regular className="w-5 h-5 text-blue-600 dark:text-blue-400" />
                                </div>
                                <h3 className="font-semibold text-gray-900 dark:text-white">PSTN Calls</h3>
                            </div>
                            <div className="space-y-4">
                                <div className="grid grid-cols-2 gap-4">
                                    <div>
                                        <p className="text-2xl font-bold text-gray-900 dark:text-white">
                                            {dashboardData.pstn.totalCalls.toLocaleString()}
                                        </p>
                                        <p className="text-sm text-gray-500 dark:text-gray-400">Total Calls</p>
                                    </div>
                                    <div>
                                        <p className="text-2xl font-bold text-gray-900 dark:text-white">
                                            {formatDuration(dashboardData.pstn.totalMinutes)}
                                        </p>
                                        <p className="text-sm text-gray-500 dark:text-gray-400">Total Duration</p>
                                    </div>
                                </div>
                                <div>
                                    <div className="flex justify-between text-sm mb-1">
                                        <span className="text-gray-500 dark:text-gray-400">Answer Rate</span>
                                        <span className="font-medium text-gray-900 dark:text-white">{dashboardData.pstn.answerRate}%</span>
                                    </div>
                                    <ProgressBar 
                                        value={dashboardData.pstn.answerRate / 100} 
                                        color={dashboardData.pstn.answerRate >= 80 ? 'success' : dashboardData.pstn.answerRate >= 60 ? 'warning' : 'error'}
                                    />
                                </div>
                                <div className="grid grid-cols-3 gap-2 text-center">
                                    <button 
                                        onClick={() => handleMetricClick('answered')}
                                        className="p-2 bg-green-50 dark:bg-green-900/20 rounded-lg hover:bg-green-100 dark:hover:bg-green-900/40 transition-colors cursor-pointer"
                                    >
                                        <p className="text-lg font-semibold text-green-600 dark:text-green-400">{dashboardData.pstn.answeredCalls}</p>
                                        <p className="text-xs text-gray-500 dark:text-gray-400">Answered</p>
                                    </button>
                                    <button 
                                        onClick={() => handleMetricClick('missed')}
                                        className="p-2 bg-red-50 dark:bg-red-900/20 rounded-lg hover:bg-red-100 dark:hover:bg-red-900/40 transition-colors cursor-pointer"
                                    >
                                        <p className="text-lg font-semibold text-red-600 dark:text-red-400">{dashboardData.pstn.missedCalls}</p>
                                        <p className="text-xs text-gray-500 dark:text-gray-400">Missed</p>
                                    </button>
                                    <button 
                                        onClick={() => handleMetricClick('inbound')}
                                        className="p-2 bg-blue-50 dark:bg-blue-900/20 rounded-lg hover:bg-blue-100 dark:hover:bg-blue-900/40 transition-colors cursor-pointer"
                                    >
                                        <p className="text-lg font-semibold text-blue-600 dark:text-blue-400">{dashboardData.pstn.inboundCalls}</p>
                                        <p className="text-xs text-gray-500 dark:text-gray-400">Inbound</p>
                                    </button>
                                </div>
                            </div>
                        </Card>

                        {/* Direct Routing Summary */}
                        <Card className="p-5">
                            <div className="flex items-center gap-2 mb-4">
                                <div className="p-2 bg-purple-50 dark:bg-purple-900/30 rounded-lg">
                                    <CallForward24Regular className="w-5 h-5 text-purple-600 dark:text-purple-400" />
                                </div>
                                <h3 className="font-semibold text-gray-900 dark:text-white">Direct Routing</h3>
                            </div>
                            {dashboardData.directRouting.totalCalls > 0 ? (
                                <div className="space-y-4">
                                    <div className="grid grid-cols-2 gap-4">
                                        <div>
                                            <p className="text-2xl font-bold text-gray-900 dark:text-white">
                                                {dashboardData.directRouting.totalCalls.toLocaleString()}
                                            </p>
                                            <p className="text-sm text-gray-500 dark:text-gray-400">Total Calls</p>
                                        </div>
                                        <div>
                                            <p className="text-2xl font-bold text-gray-900 dark:text-white">
                                                {formatDuration(dashboardData.directRouting.totalMinutes)}
                                            </p>
                                            <p className="text-sm text-gray-500 dark:text-gray-400">Total Duration</p>
                                        </div>
                                    </div>
                                    <div>
                                        <div className="flex justify-between text-sm mb-1">
                                            <span className="text-gray-500 dark:text-gray-400">Success Rate</span>
                                            <span className="font-medium text-gray-900 dark:text-white">{dashboardData.directRouting.successRate}%</span>
                                        </div>
                                        <ProgressBar 
                                            value={dashboardData.directRouting.successRate / 100} 
                                            color={dashboardData.directRouting.successRate >= 95 ? 'success' : dashboardData.directRouting.successRate >= 80 ? 'warning' : 'error'}
                                        />
                                    </div>
                                    <div className="grid grid-cols-2 gap-2 text-center">
                                        <div className="p-2 bg-green-50 dark:bg-green-900/20 rounded-lg">
                                            <p className="text-lg font-semibold text-green-600 dark:text-green-400">{dashboardData.directRouting.successfulCalls}</p>
                                            <p className="text-xs text-gray-500 dark:text-gray-400">Successful</p>
                                        </div>
                                        <div className="p-2 bg-red-50 dark:bg-red-900/20 rounded-lg">
                                            <p className="text-lg font-semibold text-red-600 dark:text-red-400">{dashboardData.directRouting.failedCalls}</p>
                                            <p className="text-xs text-gray-500 dark:text-gray-400">Failed</p>
                                        </div>
                                    </div>
                                </div>
                            ) : (
                                <div className="flex flex-col items-center justify-center py-8 text-gray-500 dark:text-gray-400">
                                    <Info24Regular className="w-8 h-8 mb-2" />
                                    <p className="text-sm">No Direct Routing calls found</p>
                                    <p className="text-xs mt-1">Direct Routing may not be configured</p>
                                </div>
                            )}
                        </Card>
                    </div>

                    {/* Charts Row */}
                    <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
                        {/* Daily Trend Chart */}
                        <Card className="p-5">
                            <div className="flex items-center gap-2 mb-4">
                                <ChartMultiple24Regular className="w-5 h-5 text-gray-500" />
                                <h3 className="font-semibold text-gray-900 dark:text-white">Call Volume Trend</h3>
                            </div>
                            {dashboardData.dailyTrend.length > 0 ? (
                                <div className="h-64">
                                    <ResponsiveContainer width="100%" height="100%">
                                        <LineChart data={dashboardData.dailyTrend}>
                                            <XAxis 
                                                dataKey="date" 
                                                tickFormatter={(value) => new Date(value).toLocaleDateString('en-GB', { day: 'numeric', month: 'short' })}
                                                tick={{ fontSize: 12 }}
                                            />
                                            <YAxis tick={{ fontSize: 12 }} />
                                            <Tooltip 
                                                labelFormatter={(value) => new Date(value).toLocaleDateString('en-GB', { weekday: 'short', day: 'numeric', month: 'short' })}
                                            />
                                            <Legend />
                                            <Line type="monotone" dataKey="pstnCalls" name="PSTN" stroke="#3b82f6" strokeWidth={2} dot={false} />
                                            <Line type="monotone" dataKey="directRoutingCalls" name="Direct Routing" stroke="#8b5cf6" strokeWidth={2} dot={false} />
                                        </LineChart>
                                    </ResponsiveContainer>
                                </div>
                            ) : (
                                <div className="h-64 flex items-center justify-center text-gray-500 dark:text-gray-400">
                                    <p>No trend data available</p>
                                </div>
                            )}
                        </Card>

                        {/* Calls by Hour */}
                        <Card className="p-5">
                            <div className="flex items-center gap-2 mb-4">
                                <Clock24Regular className="w-5 h-5 text-gray-500" />
                                <h3 className="font-semibold text-gray-900 dark:text-white">Calls by Hour of Day</h3>
                            </div>
                            {dashboardData.callsByHour.length > 0 ? (
                                <div className="h-64">
                                    <ResponsiveContainer width="100%" height="100%">
                                        <BarChart data={dashboardData.callsByHour}>
                                            <XAxis 
                                                dataKey="hour" 
                                                tickFormatter={formatHour}
                                                tick={{ fontSize: 10 }}
                                                interval={2}
                                            />
                                            <YAxis tick={{ fontSize: 12 }} />
                                            <Tooltip 
                                                labelFormatter={(value) => formatHour(value as number)}
                                                formatter={(value: number) => [value, 'Calls']}
                                            />
                                            <Bar dataKey="count" fill="#3b82f6" radius={[4, 4, 0, 0]} />
                                        </BarChart>
                                    </ResponsiveContainer>
                                </div>
                            ) : (
                                <div className="h-64 flex items-center justify-center text-gray-500 dark:text-gray-400">
                                    <p>No hourly data available</p>
                                </div>
                            )}
                        </Card>
                    </div>

                    {/* Top Users */}
                    {dashboardData.topUsers.length > 0 && (
                        <Card className="p-5">
                            <div className="flex items-center gap-2 mb-4">
                                <Person24Regular className="w-5 h-5 text-gray-500" />
                                <h3 className="font-semibold text-gray-900 dark:text-white">Top Callers</h3>
                            </div>
                            <div className="overflow-x-auto">
                                <table className="w-full">
                                    <thead>
                                        <tr className="border-b border-gray-200 dark:border-gray-700">
                                            <th className="text-left py-2 px-3 text-sm font-medium text-gray-500 dark:text-gray-400">User</th>
                                            <th className="text-right py-2 px-3 text-sm font-medium text-gray-500 dark:text-gray-400">Calls</th>
                                            <th className="text-right py-2 px-3 text-sm font-medium text-gray-500 dark:text-gray-400">Duration</th>
                                            <th className="text-right py-2 px-3 text-sm font-medium text-gray-500 dark:text-gray-400">Avg/Call</th>
                                        </tr>
                                    </thead>
                                    <tbody className="divide-y divide-gray-100 dark:divide-gray-700">
                                        {dashboardData.topUsers.map((user, index) => (
                                            <tr key={user.userPrincipalName} className="hover:bg-gray-50 dark:hover:bg-gray-700/50">
                                                <td className="py-2 px-3">
                                                    <div className="flex items-center gap-2">
                                                        <div className="w-6 h-6 rounded-full bg-blue-100 dark:bg-blue-900/30 flex items-center justify-center text-xs font-medium text-blue-600 dark:text-blue-400">
                                                            {index + 1}
                                                        </div>
                                                        <div>
                                                            <p className="text-sm font-medium text-gray-900 dark:text-white">{user.displayName}</p>
                                                            <p className="text-xs text-gray-500 dark:text-gray-400">{user.userPrincipalName}</p>
                                                        </div>
                                                    </div>
                                                </td>
                                                <td className="py-2 px-3 text-right">
                                                    <span className="text-sm font-medium text-gray-900 dark:text-white">{user.callCount}</span>
                                                </td>
                                                <td className="py-2 px-3 text-right">
                                                    <span className="text-sm text-gray-700 dark:text-gray-300">{formatDuration(user.totalMinutes)}</span>
                                                </td>
                                                <td className="py-2 px-3 text-right">
                                                    <span className="text-sm text-gray-500 dark:text-gray-400">
                                                        {user.callCount > 0 ? formatDuration(user.totalMinutes / user.callCount) : '-'}
                                                    </span>
                                                </td>
                                            </tr>
                                        ))}
                                    </tbody>
                                </table>
                            </div>
                        </Card>
                    )}

                    {/* Last Updated */}
                    <div className="flex items-center justify-center gap-2 text-xs text-gray-400 dark:text-gray-500">
                        <Clock24Regular className="w-4 h-4" />
                        Last updated: {dashboardData.lastUpdated ? new Date(dashboardData.lastUpdated).toLocaleString() : 'Unknown'}
                        <span className="mx-2">•</span>
                        Period: {new Date(dashboardData.period.fromDate).toLocaleDateString()} - {new Date(dashboardData.period.toDate).toLocaleDateString()}
                    </div>
                </>
            )}
        </div>
    );
};

// Clickable Metric Card Component
interface ClickableMetricCardProps {
    title: string;
    value: number;
    icon: React.ComponentType<{ className?: string }>;
    color: 'blue' | 'green' | 'purple' | 'red' | 'orange';
    subtitle?: string;
    onClick: () => void;
    highlight?: boolean;
}

const colorClasses = {
    blue: { bg: 'bg-blue-50 dark:bg-blue-900/30', icon: 'text-blue-600 dark:text-blue-400', hover: 'hover:border-blue-300 dark:hover:border-blue-700' },
    green: { bg: 'bg-green-50 dark:bg-green-900/30', icon: 'text-green-600 dark:text-green-400', hover: 'hover:border-green-300 dark:hover:border-green-700' },
    purple: { bg: 'bg-purple-50 dark:bg-purple-900/30', icon: 'text-purple-600 dark:text-purple-400', hover: 'hover:border-purple-300 dark:hover:border-purple-700' },
    red: { bg: 'bg-red-50 dark:bg-red-900/30', icon: 'text-red-600 dark:text-red-400', hover: 'hover:border-red-300 dark:hover:border-red-700' },
    orange: { bg: 'bg-orange-50 dark:bg-orange-900/30', icon: 'text-orange-600 dark:text-orange-400', hover: 'hover:border-orange-300 dark:hover:border-orange-700' },
};

const ClickableMetricCard: React.FC<ClickableMetricCardProps> = ({ title, value, icon: Icon, color, subtitle, onClick, highlight }) => {
    const colors = colorClasses[color];
    
    return (
        <button
            onClick={onClick}
            className={`w-full text-left p-5 bg-white dark:bg-gray-800 rounded-xl border-2 border-transparent ${colors.hover} transition-all cursor-pointer hover:shadow-md ${highlight ? 'ring-2 ring-red-200 dark:ring-red-800' : ''}`}
        >
            <div className="flex items-start justify-between">
                <div>
                    <p className="text-sm font-medium text-gray-500 dark:text-gray-400">{title}</p>
                    <p className="text-2xl font-bold text-gray-900 dark:text-white mt-1">
                        {value.toLocaleString()}
                    </p>
                    {subtitle && (
                        <p className="text-xs text-gray-400 dark:text-gray-500 mt-1">{subtitle}</p>
                    )}
                </div>
                <div className={`p-2.5 rounded-lg ${colors.bg}`}>
                    <Icon className={`w-5 h-5 ${colors.icon}`} />
                </div>
            </div>
            <div className="mt-3 flex items-center text-xs text-blue-600 dark:text-blue-400">
                <span>View details</span>
                <ArrowRight24Regular className="w-4 h-4 ml-1" />
            </div>
        </button>
    );
};

// Call Detail View Component
interface CallDetailViewProps {
    filter: CallFilter | null;
    dashboardData: DashboardData;
    days: number;
    onBack: () => void;
    getAccessToken: () => Promise<string>;
}

const CallDetailView: React.FC<CallDetailViewProps> = ({ filter, dashboardData, days, onBack, getAccessToken }) => {
    const [calls, setCalls] = useState<PstnCallDetail[]>([]);
    const [loading, setLoading] = useState(true);
    const [error, setError] = useState<string | null>(null);

    const filterLabels: Record<CallFilter, string> = {
        all: 'All Calls',
        inbound: 'Inbound Calls',
        outbound: 'Outbound Calls',
        answered: 'Answered Calls',
        missed: 'Missed Calls',
    };

    const filterCounts: Record<CallFilter, number> = {
        all: dashboardData.pstn.totalCalls,
        inbound: dashboardData.pstn.inboundCalls,
        outbound: dashboardData.pstn.outboundCalls,
        answered: dashboardData.pstn.answeredCalls,
        missed: dashboardData.pstn.missedCalls,
    };

    useEffect(() => {
        const fetchCalls = async () => {
            try {
                setLoading(true);
                const token = await getAccessToken();
                const response = await fetch(`/api/teamsphone/pstn-calls?days=${days}`, {
                    headers: {
                        'Authorization': `Bearer ${token}`,
                        'Content-Type': 'application/json',
                    },
                });

                if (response.ok) {
                    const data = await response.json();
                    // Use the actual call records from the API
                    setCalls(data.calls || []);
                } else {
                    setError('Failed to load call details');
                }
            } catch (err) {
                setError('Failed to load call details');
            } finally {
                setLoading(false);
            }
        };

        fetchCalls();
    }, [getAccessToken, days]);

    // Filter calls based on selected filter
    const filteredCalls = calls.filter(call => {
        if (!filter || filter === 'all') return true;
        if (filter === 'inbound') return call.callType?.toLowerCase().includes('inbound');
        if (filter === 'outbound') return call.callType?.toLowerCase().includes('outbound');
        if (filter === 'answered') return call.isAnswered;
        if (filter === 'missed') return !call.isAnswered;
        return true;
    });

    const formatDuration = (seconds: number) => {
        if (!seconds || seconds === 0) return '0s';
        if (seconds < 60) return `${seconds}s`;
        const mins = Math.floor(seconds / 60);
        const secs = seconds % 60;
        return secs > 0 ? `${mins}m ${secs}s` : `${mins}m`;
    };

    return (
        <div className="p-6 space-y-6">
            {/* Header */}
            <div className="flex items-center gap-4">
                <Button
                    appearance="subtle"
                    icon={<ArrowLeft24Regular />}
                    onClick={onBack}
                >
                    Back
                </Button>
                <div>
                    <h1 className="text-2xl font-bold text-gray-900 dark:text-white">
                        {filter ? filterLabels[filter] : 'Call Details'}
                    </h1>
                    <p className="text-sm text-gray-500 dark:text-gray-400">
                        {filter ? filterCounts[filter].toLocaleString() : 0} calls in the last {days} days
                    </p>
                </div>
            </div>

            {/* Stats Summary */}
            <div className="grid grid-cols-2 sm:grid-cols-4 gap-4">
                <div className={`p-4 rounded-lg ${filter === 'all' ? 'bg-blue-50 dark:bg-blue-900/20 ring-2 ring-blue-300' : 'bg-gray-50 dark:bg-gray-800'}`}>
                    <p className="text-2xl font-bold text-gray-900 dark:text-white">{dashboardData.pstn.totalCalls}</p>
                    <p className="text-sm text-gray-500 dark:text-gray-400">Total</p>
                </div>
                <div className={`p-4 rounded-lg ${filter === 'answered' ? 'bg-green-50 dark:bg-green-900/20 ring-2 ring-green-300' : 'bg-gray-50 dark:bg-gray-800'}`}>
                    <p className="text-2xl font-bold text-green-600 dark:text-green-400">{dashboardData.pstn.answeredCalls}</p>
                    <p className="text-sm text-gray-500 dark:text-gray-400">Answered</p>
                </div>
                <div className={`p-4 rounded-lg ${filter === 'missed' ? 'bg-red-50 dark:bg-red-900/20 ring-2 ring-red-300' : 'bg-gray-50 dark:bg-gray-800'}`}>
                    <p className="text-2xl font-bold text-red-600 dark:text-red-400">{dashboardData.pstn.missedCalls}</p>
                    <p className="text-sm text-gray-500 dark:text-gray-400">Missed</p>
                </div>
                <div className="p-4 rounded-lg bg-gray-50 dark:bg-gray-800">
                    <p className="text-2xl font-bold text-gray-900 dark:text-white">{dashboardData.pstn.answerRate}%</p>
                    <p className="text-sm text-gray-500 dark:text-gray-400">Answer Rate</p>
                </div>
            </div>

            {/* Call List */}
            <Card className="overflow-hidden">
                <div className="px-4 py-3 bg-gray-50 dark:bg-gray-700/50 border-b border-gray-200 dark:border-gray-700">
                    <p className="text-sm font-medium text-gray-700 dark:text-gray-300">
                        {filteredCalls.length} {filter ? filterLabels[filter].toLowerCase() : 'calls'} shown
                    </p>
                </div>
                
                {loading ? (
                    <div className="flex items-center justify-center py-12">
                        <Spinner size="medium" label="Loading calls..." />
                    </div>
                ) : error ? (
                    <div className="flex items-center justify-center py-12 text-red-500">
                        <p>{error}</p>
                    </div>
                ) : filteredCalls.length === 0 ? (
                    <div className="flex flex-col items-center justify-center py-12 text-gray-500 dark:text-gray-400">
                        <Info24Regular className="w-12 h-12 mb-3" />
                        <p>No calls found for this filter</p>
                        <p className="text-sm mt-1">Try selecting a different time range</p>
                    </div>
                ) : (
                    <div className="divide-y divide-gray-100 dark:divide-gray-700">
                        {filteredCalls.slice(0, 100).map((call, index) => (
                            <div key={call.id || index} className="px-4 py-3 flex items-center gap-4 hover:bg-gray-50 dark:hover:bg-gray-700/50">
                                {/* Status Icon */}
                                <div className="flex-shrink-0">
                                    {call.isAnswered ? (
                                        call.callType?.toLowerCase().includes('inbound') ? (
                                            <CallInbound24Regular className="w-6 h-6 text-green-500" />
                                        ) : (
                                            <CallOutbound24Regular className="w-6 h-6 text-blue-500" />
                                        )
                                    ) : (
                                        <CallMissed24Regular className="w-6 h-6 text-red-500" />
                                    )}
                                </div>

                                {/* User Info */}
                                <div className="flex-1 min-w-0">
                                    <p className="font-medium text-gray-900 dark:text-white text-sm truncate">
                                        {call.userDisplayName || call.userPrincipalName || 'Unknown'}
                                    </p>
                                    <p className="text-xs text-gray-500 dark:text-gray-400 truncate">
                                        {call.userPrincipalName || '-'}
                                    </p>
                                </div>

                                {/* Phone Numbers */}
                                <div className="hidden lg:block w-40">
                                    <p className="text-xs text-gray-700 dark:text-gray-300 truncate" title={call.callerNumber}>
                                        From: {call.callerNumber || '-'}
                                    </p>
                                    <p className="text-xs text-gray-500 dark:text-gray-400 truncate" title={call.calleeNumber}>
                                        To: {call.calleeNumber || '-'}
                                    </p>
                                </div>

                                {/* Call Type */}
                                <div className="hidden sm:block">
                                    <Badge 
                                        appearance="tint" 
                                        color={call.callType?.toLowerCase().includes('inbound') ? 'success' : 'brand'}
                                        size="small"
                                    >
                                        {call.callType || 'Unknown'}
                                    </Badge>
                                </div>

                                {/* Duration */}
                                <div className="w-20 text-right">
                                    <p className="text-sm font-medium text-gray-900 dark:text-white">
                                        {call.isAnswered ? formatDuration(call.duration) : '-'}
                                    </p>
                                    <p className="text-xs text-gray-500 dark:text-gray-400">
                                        {call.isAnswered ? 'Duration' : 'Missed'}
                                    </p>
                                </div>

                                {/* Time */}
                                <div className="w-32 text-right hidden md:block">
                                    <p className="text-sm text-gray-700 dark:text-gray-300">
                                        {call.startDateTime ? new Date(call.startDateTime).toLocaleDateString() : '-'}
                                    </p>
                                    <p className="text-xs text-gray-500 dark:text-gray-400">
                                        {call.startDateTime ? new Date(call.startDateTime).toLocaleTimeString() : '-'}
                                    </p>
                                </div>

                                {/* Destination */}
                                <div className="hidden xl:block w-32 text-right">
                                    <p className="text-xs text-gray-700 dark:text-gray-300 truncate" title={call.destinationName}>
                                        {call.destinationName || '-'}
                                    </p>
                                    <p className="text-xs text-gray-500 dark:text-gray-400 truncate">
                                        {call.destinationContext || '-'}
                                    </p>
                                </div>
                            </div>
                        ))}
                    </div>
                )}

                {filteredCalls.length > 100 && (
                    <div className="px-4 py-3 bg-gray-50 dark:bg-gray-700/50 border-t border-gray-200 dark:border-gray-700 text-center">
                        <p className="text-sm text-gray-500 dark:text-gray-400">
                            Showing first 100 of {filteredCalls.length} calls
                        </p>
                    </div>
                )}
            </Card>

            {/* Info Note */}
            <div className="flex items-start gap-2 p-4 bg-blue-50 dark:bg-blue-900/20 rounded-lg">
                <Info24Regular className="w-5 h-5 text-blue-500 flex-shrink-0 mt-0.5" />
                <div className="text-sm text-blue-700 dark:text-blue-300">
                    <p className="font-medium">Note about call details</p>
                    <p className="mt-1 text-blue-600 dark:text-blue-400">
                        Individual call records are aggregated from the Microsoft Graph Call Records API. 
                        For more detailed call analytics including call quality metrics, use the Teams Admin Center 
                        or Call Quality Dashboard.
                    </p>
                </div>
            </div>
        </div>
    );
};

// Beta API View Component
interface BetaApiViewProps {
    days: number;
    setDays: (days: number) => void;
    activeTab: ViewTab;
    setActiveTab: (tab: ViewTab) => void;
    getAccessToken: () => Promise<string>;
}

interface BetaCallRecord {
    id: string;
    callId?: string;
    userPrincipalName: string;
    userDisplayName: string;
    startDateTime: string;
    endDateTime: string;
    duration: number;
    charge: number;
    callType: string;
    currency: string;
    calleeNumber: string;
    callerNumber: string;
    destinationContext: string;
    destinationName: string;
    licenseCapability: string;
    inventoryType: string;
    tenantCountryCode?: string;
    usageCountryCode?: string;
    connectionCharge?: number;
    otherPartyCountryCode?: string;
    clientLocalIpV4Address?: string;
    clientPublicIpV4Address?: string;
    isAnswered: boolean;
}

interface BetaSbcStats {
    sbc: string;
    totalCalls: number;
    successfulCalls: number;
    failedCalls: number;
    totalMinutes: number;
}

interface BetaSipFailure {
    sipCode: number;
    sipCodePhrase: string;
    count: number;
}

interface BetaDirectRoutingCall {
    id: string;
    correlationId?: string;
    userPrincipalName: string;
    userDisplayName: string;
    startDateTime: string;
    endDateTime: string;
    inviteDateTime?: string;
    failureDateTime?: string;
    duration: number;
    callType: string;
    calleeNumber: string;
    callerNumber: string;
    trunkFullyQualifiedDomainName?: string;
    mediaBypassEnabled?: boolean;
    finalSipCode: number;
    finalSipCodePhrase?: string;
    mediaPathLocation?: string;
    signalingLocation?: string;
    isSuccess: boolean;
}

const BetaApiView: React.FC<BetaApiViewProps> = ({ days, setDays, activeTab, setActiveTab, getAccessToken }) => {
    const [loading, setLoading] = useState(true);
    const [error, setError] = useState<string | null>(null);
    const [pstnCalls, setPstnCalls] = useState<BetaCallRecord[]>([]);
    const [pstnSummary, setPstnSummary] = useState<any>(null);
    const [drCalls, setDrCalls] = useState<BetaDirectRoutingCall[]>([]);
    const [drSummary, setDrSummary] = useState<any>(null);
    const [sbcStats, setSbcStats] = useState<BetaSbcStats[]>([]);
    const [sipFailures, setSipFailures] = useState<BetaSipFailure[]>([]);
    const [activeSection, setActiveSection] = useState<'pstn' | 'dr' | 'sms'>('pstn');

    const fetchBetaData = useCallback(async () => {
        try {
            setLoading(true);
            setError(null);
            const token = await getAccessToken();

            // Fetch enhanced PSTN calls
            const pstnResponse = await fetch(`/api/teamsphone/beta/pstn-calls-enhanced?days=${days}`, {
                headers: {
                    'Authorization': `Bearer ${token}`,
                    'Content-Type': 'application/json',
                },
            });
            if (pstnResponse.ok) {
                const data = await pstnResponse.json();
                setPstnCalls(data.calls || []);
                setPstnSummary(data.summary);
            }

            // Fetch enhanced Direct Routing calls
            const drResponse = await fetch(`/api/teamsphone/beta/direct-routing-enhanced?days=${days}`, {
                headers: {
                    'Authorization': `Bearer ${token}`,
                    'Content-Type': 'application/json',
                },
            });
            if (drResponse.ok) {
                const data = await drResponse.json();
                setDrCalls(data.calls || []);
                setDrSummary(data.summary);
                setSbcStats(data.sbcStats || []);
                setSipFailures(data.sipFailures || []);
            }
        } catch (err) {
            setError(err instanceof Error ? err.message : 'Failed to load beta data');
        } finally {
            setLoading(false);
        }
    }, [getAccessToken, days]);

    useEffect(() => {
        fetchBetaData();
    }, [fetchBetaData]);

    const formatDuration = (seconds: number) => {
        if (!seconds || seconds === 0) return '0s';
        if (seconds < 60) return `${seconds}s`;
        const mins = Math.floor(seconds / 60);
        const secs = seconds % 60;
        return secs > 0 ? `${mins}m ${secs}s` : `${mins}m`;
    };

    const formatCurrency = (amount: number, currency: string) => {
        if (!amount) return '-';
        return new Intl.NumberFormat('en-US', {
            style: 'currency',
            currency: currency || 'USD',
            minimumFractionDigits: 2,
        }).format(amount);
    };

    return (
        <div className="p-6 space-y-6">
            {/* Header */}
            <div className="flex items-start justify-between">
                <div>
                    <h1 className="text-2xl font-bold text-gray-900 dark:text-white flex items-center gap-2">
                        <Beaker24Regular className="w-7 h-7" />
                        Teams Phone - Beta API
                    </h1>
                    <p className="mt-1 text-sm text-gray-500 dark:text-gray-400">
                        Enhanced call analytics with additional data fields from Microsoft Graph Beta API
                    </p>
                </div>
                <div className="flex items-center gap-3">
                    {/* Tab Selector */}
                    <div className="flex items-center bg-gray-100 dark:bg-gray-700 rounded-lg p-1">
                        <button
                            onClick={() => setActiveTab('standard')}
                            className={`flex items-center gap-1.5 px-3 py-1.5 rounded-md text-sm font-medium transition-colors ${
                                activeTab === 'standard'
                                    ? 'bg-white dark:bg-gray-600 text-gray-900 dark:text-white shadow-sm'
                                    : 'text-gray-600 dark:text-gray-400 hover:text-gray-900 dark:hover:text-white'
                            }`}
                        >
                            <Call24Regular className="w-4 h-4" />
                            Standard
                        </button>
                        <button
                            onClick={() => setActiveTab('beta')}
                            className={`flex items-center gap-1.5 px-3 py-1.5 rounded-md text-sm font-medium transition-colors ${
                                activeTab === 'beta'
                                    ? 'bg-white dark:bg-gray-600 text-gray-900 dark:text-white shadow-sm'
                                    : 'text-gray-600 dark:text-gray-400 hover:text-gray-900 dark:hover:text-white'
                            }`}
                        >
                            <Beaker24Regular className="w-4 h-4" />
                            Beta API
                        </button>
                    </div>

                    <div className="flex items-center gap-2">
                        <Filter24Regular className="w-4 h-4 text-gray-400" />
                        <select
                            value={days}
                            onChange={(e) => setDays(Number(e.target.value))}
                            className="px-3 py-1.5 border border-gray-300 dark:border-gray-600 rounded-lg bg-white dark:bg-gray-700 text-gray-900 dark:text-white text-sm"
                        >
                            <option value={7}>Last 7 days</option>
                            <option value={14}>Last 14 days</option>
                            <option value={30}>Last 30 days</option>
                            <option value={60}>Last 60 days</option>
                            <option value={90}>Last 90 days</option>
                        </select>
                    </div>
                    <button
                        onClick={fetchBetaData}
                        disabled={loading}
                        className="p-2 text-gray-600 hover:text-blue-600 hover:bg-blue-50 dark:text-gray-400 dark:hover:text-blue-400 dark:hover:bg-blue-900/20 rounded-lg transition-colors disabled:opacity-50"
                        title="Refresh"
                    >
                        <ArrowSync24Regular className={`w-5 h-5 ${loading ? 'animate-spin' : ''}`} />
                    </button>
                </div>
            </div>

            {/* Beta Notice */}
            <div className="flex items-start gap-2 p-4 bg-purple-50 dark:bg-purple-900/20 border border-purple-200 dark:border-purple-800 rounded-lg">
                <Beaker24Regular className="w-5 h-5 text-purple-500 flex-shrink-0 mt-0.5" />
                <div className="text-sm text-purple-700 dark:text-purple-300">
                    <p className="font-medium">Beta API Features</p>
                    <p className="mt-1 text-purple-600 dark:text-purple-400">
                        This view uses Microsoft Graph Beta endpoints which provide additional data including:
                        call charges, country codes, IP addresses, SBC details, SIP response codes, and more.
                        Beta APIs may change without notice.
                    </p>
                </div>
            </div>

            {loading ? (
                <div className="flex items-center justify-center min-h-[400px]">
                    <Spinner size="large" label="Loading Beta API data..." />
                </div>
            ) : error ? (
                <div className="bg-red-50 dark:bg-red-900/20 border border-red-200 dark:border-red-800 rounded-lg p-4">
                    <p className="text-red-700 dark:text-red-300">{error}</p>
                </div>
            ) : (
                <>
                    {/* Section Tabs */}
                    <div className="flex items-center gap-2 border-b border-gray-200 dark:border-gray-700">
                        <button
                            onClick={() => setActiveSection('pstn')}
                            className={`flex items-center gap-2 px-4 py-2 text-sm font-medium border-b-2 transition-colors ${
                                activeSection === 'pstn'
                                    ? 'border-blue-500 text-blue-600 dark:text-blue-400'
                                    : 'border-transparent text-gray-500 hover:text-gray-700 dark:text-gray-400 dark:hover:text-gray-300'
                            }`}
                        >
                            <Call24Regular className="w-4 h-4" />
                            PSTN Calls ({pstnCalls.length})
                        </button>
                        <button
                            onClick={() => setActiveSection('dr')}
                            className={`flex items-center gap-2 px-4 py-2 text-sm font-medium border-b-2 transition-colors ${
                                activeSection === 'dr'
                                    ? 'border-blue-500 text-blue-600 dark:text-blue-400'
                                    : 'border-transparent text-gray-500 hover:text-gray-700 dark:text-gray-400 dark:hover:text-gray-300'
                            }`}
                        >
                            <Router24Regular className="w-4 h-4" />
                            Direct Routing ({drCalls.length})
                        </button>
                    </div>

                    {/* PSTN Section */}
                    {activeSection === 'pstn' && (
                        <div className="space-y-6">
                            {/* Summary Cards */}
                            {pstnSummary && (
                                <div className="grid grid-cols-2 sm:grid-cols-4 lg:grid-cols-6 gap-4">
                                    <Card className="p-4">
                                        <p className="text-2xl font-bold text-gray-900 dark:text-white">{pstnSummary.totalCalls}</p>
                                        <p className="text-sm text-gray-500 dark:text-gray-400">Total Calls</p>
                                    </Card>
                                    <Card className="p-4">
                                        <p className="text-2xl font-bold text-green-600 dark:text-green-400">{pstnSummary.answeredCalls}</p>
                                        <p className="text-sm text-gray-500 dark:text-gray-400">Answered</p>
                                    </Card>
                                    <Card className="p-4">
                                        <p className="text-2xl font-bold text-red-600 dark:text-red-400">{pstnSummary.missedCalls}</p>
                                        <p className="text-sm text-gray-500 dark:text-gray-400">Missed</p>
                                    </Card>
                                    <Card className="p-4">
                                        <p className="text-2xl font-bold text-gray-900 dark:text-white">{pstnSummary.totalDurationMinutes}m</p>
                                        <p className="text-sm text-gray-500 dark:text-gray-400">Duration</p>
                                    </Card>
                                    <Card className="p-4">
                                        <p className="text-2xl font-bold text-gray-900 dark:text-white">{pstnSummary.answerRate}%</p>
                                        <p className="text-sm text-gray-500 dark:text-gray-400">Answer Rate</p>
                                    </Card>
                                    <Card className="p-4">
                                        <p className="text-2xl font-bold text-amber-600 dark:text-amber-400">
                                            {formatCurrency(pstnSummary.totalCharge, 'USD')}
                                        </p>
                                        <p className="text-sm text-gray-500 dark:text-gray-400">Total Charges</p>
                                    </Card>
                                </div>
                            )}

                            {/* Call Records Table */}
                            <Card className="overflow-hidden">
                                <div className="px-4 py-3 bg-gray-50 dark:bg-gray-700/50 border-b border-gray-200 dark:border-gray-700">
                                    <h3 className="font-semibold text-gray-900 dark:text-white flex items-center gap-2">
                                        <Money24Regular className="w-5 h-5" />
                                        PSTN Call Records with Charges
                                    </h3>
                                </div>
                                <div className="overflow-x-auto">
                                    <table className="w-full text-sm">
                                        <thead className="bg-gray-50 dark:bg-gray-800">
                                            <tr>
                                                <th className="text-left py-2 px-3 font-medium text-gray-500 dark:text-gray-400">User</th>
                                                <th className="text-left py-2 px-3 font-medium text-gray-500 dark:text-gray-400">Type</th>
                                                <th className="text-left py-2 px-3 font-medium text-gray-500 dark:text-gray-400">From/To</th>
                                                <th className="text-left py-2 px-3 font-medium text-gray-500 dark:text-gray-400">Destination</th>
                                                <th className="text-right py-2 px-3 font-medium text-gray-500 dark:text-gray-400">Duration</th>
                                                <th className="text-right py-2 px-3 font-medium text-gray-500 dark:text-gray-400">Charge</th>
                                                <th className="text-left py-2 px-3 font-medium text-gray-500 dark:text-gray-400">Country</th>
                                                <th className="text-left py-2 px-3 font-medium text-gray-500 dark:text-gray-400">Date/Time</th>
                                            </tr>
                                        </thead>
                                        <tbody className="divide-y divide-gray-100 dark:divide-gray-700">
                                            {pstnCalls.slice(0, 50).map((call, index) => (
                                                <tr key={call.id || index} className="hover:bg-gray-50 dark:hover:bg-gray-700/50">
                                                    <td className="py-2 px-3">
                                                        <p className="font-medium text-gray-900 dark:text-white truncate max-w-[150px]">
                                                            {call.userDisplayName || 'Unknown'}
                                                        </p>
                                                        <p className="text-xs text-gray-500 truncate max-w-[150px]">
                                                            {call.userPrincipalName}
                                                        </p>
                                                    </td>
                                                    <td className="py-2 px-3">
                                                        <Badge 
                                                            appearance="tint" 
                                                            color={call.isAnswered ? 'success' : 'danger'}
                                                            size="small"
                                                        >
                                                            {call.callType || 'Unknown'}
                                                        </Badge>
                                                    </td>
                                                    <td className="py-2 px-3">
                                                        <p className="text-xs text-gray-700 dark:text-gray-300">{call.callerNumber || '-'}</p>
                                                        <p className="text-xs text-gray-500">→ {call.calleeNumber || '-'}</p>
                                                    </td>
                                                    <td className="py-2 px-3">
                                                        <p className="text-xs text-gray-700 dark:text-gray-300 truncate max-w-[120px]">
                                                            {call.destinationName || '-'}
                                                        </p>
                                                    </td>
                                                    <td className="py-2 px-3 text-right">
                                                        <span className="text-gray-900 dark:text-white">
                                                            {call.isAnswered ? formatDuration(call.duration) : '-'}
                                                        </span>
                                                    </td>
                                                    <td className="py-2 px-3 text-right">
                                                        <span className={`font-medium ${call.charge > 0 ? 'text-amber-600 dark:text-amber-400' : 'text-gray-500'}`}>
                                                            {formatCurrency(call.charge, call.currency)}
                                                        </span>
                                                    </td>
                                                    <td className="py-2 px-3">
                                                        <div className="flex items-center gap-1">
                                                            <Globe24Regular className="w-3 h-3 text-gray-400" />
                                                            <span className="text-xs text-gray-600 dark:text-gray-400">
                                                                {call.usageCountryCode || call.tenantCountryCode || '-'}
                                                            </span>
                                                        </div>
                                                    </td>
                                                    <td className="py-2 px-3">
                                                        <p className="text-xs text-gray-700 dark:text-gray-300">
                                                            {call.startDateTime ? new Date(call.startDateTime).toLocaleDateString() : '-'}
                                                        </p>
                                                        <p className="text-xs text-gray-500">
                                                            {call.startDateTime ? new Date(call.startDateTime).toLocaleTimeString() : '-'}
                                                        </p>
                                                    </td>
                                                </tr>
                                            ))}
                                        </tbody>
                                    </table>
                                </div>
                                {pstnCalls.length > 50 && (
                                    <div className="px-4 py-2 bg-gray-50 dark:bg-gray-700/50 border-t border-gray-200 dark:border-gray-700 text-center">
                                        <p className="text-xs text-gray-500">Showing first 50 of {pstnCalls.length} calls</p>
                                    </div>
                                )}
                            </Card>
                        </div>
                    )}

                    {/* Direct Routing Section */}
                    {activeSection === 'dr' && (
                        <div className="space-y-6">
                            {/* Summary Cards */}
                            {drSummary && (
                                <div className="grid grid-cols-2 sm:grid-cols-4 lg:grid-cols-5 gap-4">
                                    <Card className="p-4">
                                        <p className="text-2xl font-bold text-gray-900 dark:text-white">{drSummary.totalCalls}</p>
                                        <p className="text-sm text-gray-500 dark:text-gray-400">Total Calls</p>
                                    </Card>
                                    <Card className="p-4">
                                        <p className="text-2xl font-bold text-green-600 dark:text-green-400">{drSummary.successfulCalls}</p>
                                        <p className="text-sm text-gray-500 dark:text-gray-400">Successful</p>
                                    </Card>
                                    <Card className="p-4">
                                        <p className="text-2xl font-bold text-red-600 dark:text-red-400">{drSummary.failedCalls}</p>
                                        <p className="text-sm text-gray-500 dark:text-gray-400">Failed</p>
                                    </Card>
                                    <Card className="p-4">
                                        <p className="text-2xl font-bold text-gray-900 dark:text-white">{drSummary.totalDurationMinutes}m</p>
                                        <p className="text-sm text-gray-500 dark:text-gray-400">Duration</p>
                                    </Card>
                                    <Card className="p-4">
                                        <p className="text-2xl font-bold text-gray-900 dark:text-white">{drSummary.successRate}%</p>
                                        <p className="text-sm text-gray-500 dark:text-gray-400">Success Rate</p>
                                    </Card>
                                </div>
                            )}

                            {/* SBC Stats and SIP Failures */}
                            <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
                                {/* SBC Stats */}
                                {sbcStats.length > 0 && (
                                    <Card className="p-5">
                                        <div className="flex items-center gap-2 mb-4">
                                            <Router24Regular className="w-5 h-5 text-gray-500" />
                                            <h3 className="font-semibold text-gray-900 dark:text-white">SBC Statistics</h3>
                                        </div>
                                        <div className="space-y-3">
                                            {sbcStats.map((sbc, index) => (
                                                <div key={index} className="p-3 bg-gray-50 dark:bg-gray-700/50 rounded-lg">
                                                    <div className="flex items-center justify-between mb-2">
                                                        <p className="font-medium text-gray-900 dark:text-white text-sm truncate">
                                                            {sbc.sbc}
                                                        </p>
                                                        <Badge appearance="tint" color="brand" size="small">
                                                            {sbc.totalCalls} calls
                                                        </Badge>
                                                    </div>
                                                    <div className="grid grid-cols-3 gap-2 text-xs">
                                                        <div>
                                                            <p className="text-green-600 dark:text-green-400 font-medium">{sbc.successfulCalls}</p>
                                                            <p className="text-gray-500">Success</p>
                                                        </div>
                                                        <div>
                                                            <p className="text-red-600 dark:text-red-400 font-medium">{sbc.failedCalls}</p>
                                                            <p className="text-gray-500">Failed</p>
                                                        </div>
                                                        <div>
                                                            <p className="text-gray-900 dark:text-white font-medium">{sbc.totalMinutes}m</p>
                                                            <p className="text-gray-500">Duration</p>
                                                        </div>
                                                    </div>
                                                </div>
                                            ))}
                                        </div>
                                    </Card>
                                )}

                                {/* SIP Failures */}
                                {sipFailures.length > 0 && (
                                    <Card className="p-5">
                                        <div className="flex items-center gap-2 mb-4">
                                            <Warning24Regular className="w-5 h-5 text-red-500" />
                                            <h3 className="font-semibold text-gray-900 dark:text-white">SIP Failure Codes</h3>
                                        </div>
                                        <div className="space-y-2">
                                            {sipFailures.map((failure, index) => (
                                                <div key={index} className="flex items-center justify-between p-2 bg-red-50 dark:bg-red-900/20 rounded-lg">
                                                    <div className="flex items-center gap-2">
                                                        <Badge appearance="filled" color="danger" size="small">
                                                            {failure.sipCode}
                                                        </Badge>
                                                        <span className="text-sm text-gray-700 dark:text-gray-300">
                                                            {failure.sipCodePhrase}
                                                        </span>
                                                    </div>
                                                    <span className="text-sm font-medium text-red-600 dark:text-red-400">
                                                        {failure.count} calls
                                                    </span>
                                                </div>
                                            ))}
                                        </div>
                                    </Card>
                                )}
                            </div>

                            {/* Direct Routing Call Records */}
                            <Card className="overflow-hidden">
                                <div className="px-4 py-3 bg-gray-50 dark:bg-gray-700/50 border-b border-gray-200 dark:border-gray-700">
                                    <h3 className="font-semibold text-gray-900 dark:text-white flex items-center gap-2">
                                        <Router24Regular className="w-5 h-5" />
                                        Direct Routing Call Records
                                    </h3>
                                </div>
                                <div className="overflow-x-auto">
                                    <table className="w-full text-sm">
                                        <thead className="bg-gray-50 dark:bg-gray-800">
                                            <tr>
                                                <th className="text-left py-2 px-3 font-medium text-gray-500 dark:text-gray-400">User</th>
                                                <th className="text-left py-2 px-3 font-medium text-gray-500 dark:text-gray-400">SBC</th>
                                                <th className="text-left py-2 px-3 font-medium text-gray-500 dark:text-gray-400">From/To</th>
                                                <th className="text-center py-2 px-3 font-medium text-gray-500 dark:text-gray-400">SIP Code</th>
                                                <th className="text-right py-2 px-3 font-medium text-gray-500 dark:text-gray-400">Duration</th>
                                                <th className="text-left py-2 px-3 font-medium text-gray-500 dark:text-gray-400">Media Path</th>
                                                <th className="text-left py-2 px-3 font-medium text-gray-500 dark:text-gray-400">Date/Time</th>
                                            </tr>
                                        </thead>
                                        <tbody className="divide-y divide-gray-100 dark:divide-gray-700">
                                            {drCalls.slice(0, 50).map((call, index) => (
                                                <tr key={call.id || index} className="hover:bg-gray-50 dark:hover:bg-gray-700/50">
                                                    <td className="py-2 px-3">
                                                        <p className="font-medium text-gray-900 dark:text-white truncate max-w-[150px]">
                                                            {call.userDisplayName || 'Unknown'}
                                                        </p>
                                                    </td>
                                                    <td className="py-2 px-3">
                                                        <p className="text-xs text-gray-700 dark:text-gray-300 truncate max-w-[150px]">
                                                            {call.trunkFullyQualifiedDomainName || '-'}
                                                        </p>
                                                        {call.mediaBypassEnabled && (
                                                            <Badge appearance="tint" color="success" size="small">Bypass</Badge>
                                                        )}
                                                    </td>
                                                    <td className="py-2 px-3">
                                                        <p className="text-xs text-gray-700 dark:text-gray-300">{call.callerNumber || '-'}</p>
                                                        <p className="text-xs text-gray-500">→ {call.calleeNumber || '-'}</p>
                                                    </td>
                                                    <td className="py-2 px-3 text-center">
                                                        <Badge 
                                                            appearance="filled" 
                                                            color={call.isSuccess ? 'success' : 'danger'}
                                                            size="small"
                                                        >
                                                            {call.finalSipCode}
                                                        </Badge>
                                                        <p className="text-xs text-gray-500 mt-1 truncate max-w-[100px]">
                                                            {call.finalSipCodePhrase}
                                                        </p>
                                                    </td>
                                                    <td className="py-2 px-3 text-right">
                                                        <span className="text-gray-900 dark:text-white">
                                                            {call.isSuccess ? formatDuration(call.duration) : '-'}
                                                        </span>
                                                    </td>
                                                    <td className="py-2 px-3">
                                                        <p className="text-xs text-gray-600 dark:text-gray-400">
                                                            {call.mediaPathLocation || '-'}
                                                        </p>
                                                        <p className="text-xs text-gray-500">
                                                            {call.signalingLocation || '-'}
                                                        </p>
                                                    </td>
                                                    <td className="py-2 px-3">
                                                        <p className="text-xs text-gray-700 dark:text-gray-300">
                                                            {call.startDateTime ? new Date(call.startDateTime).toLocaleDateString() : '-'}
                                                        </p>
                                                        <p className="text-xs text-gray-500">
                                                            {call.startDateTime ? new Date(call.startDateTime).toLocaleTimeString() : '-'}
                                                        </p>
                                                    </td>
                                                </tr>
                                            ))}
                                        </tbody>
                                    </table>
                                </div>
                                {drCalls.length === 0 && (
                                    <div className="flex flex-col items-center justify-center py-12 text-gray-500 dark:text-gray-400">
                                        <Info24Regular className="w-12 h-12 mb-3" />
                                        <p>No Direct Routing calls found</p>
                                        <p className="text-sm mt-1">Direct Routing may not be configured</p>
                                    </div>
                                )}
                                {drCalls.length > 50 && (
                                    <div className="px-4 py-2 bg-gray-50 dark:bg-gray-700/50 border-t border-gray-200 dark:border-gray-700 text-center">
                                        <p className="text-xs text-gray-500">Showing first 50 of {drCalls.length} calls</p>
                                    </div>
                                )}
                            </Card>
                        </div>
                    )}
                </>
            )}
        </div>
    );
};

export default TeamsPhonePage;
