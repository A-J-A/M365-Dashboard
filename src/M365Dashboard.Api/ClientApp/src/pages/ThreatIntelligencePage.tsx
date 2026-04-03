import React, { useState, useEffect, useCallback } from 'react';
import {
    ShieldError24Regular,
    Warning24Regular,
    Person24Regular,
    ArrowSync24Regular,
    Bug24Regular,
    Mail24Regular,
    Shield24Regular,
    Eye24Regular,
} from '@fluentui/react-icons';
import { Card, Spinner, Badge, ProgressBar } from '@fluentui/react-components';
import { BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, LineChart, Line, PieChart, Pie, Cell, Legend } from 'recharts';
import { useAppContext } from '../contexts/AppContext';

const COLORS = ['#ef4444', '#f59e0b', '#22c55e', '#3b82f6', '#8b5cf6'];

const ThreatIntelligencePage: React.FC = () => {
    const { getAccessToken } = useAppContext();
    const [alerts, setAlerts] = useState<any[]>([]);
    const [alertSummary, setAlertSummary] = useState<any>(null);
    const [riskyUsers, setRiskyUsers] = useState<any[]>([]);
    const [riskDetections, setRiskDetections] = useState<any[]>([]);
    const [secureScore, setSecureScore] = useState<any>(null);
    const [loading, setLoading] = useState(true);
    const [error, setError] = useState<string | null>(null);
    const [days, setDays] = useState(30);

    const fetchData = useCallback(async () => {
        try {
            setLoading(true);
            const token = await getAccessToken();
            const headers = { 'Authorization': `Bearer ${token}` };

            const [alertsRes, riskyUsersRes, detectionsRes, scoreRes] = await Promise.all([
                fetch(`/api/threatintelligence/alerts?days=${days}`, { headers }),
                fetch('/api/threatintelligence/risky-users', { headers }),
                fetch(`/api/threatintelligence/risk-detections?days=${days}`, { headers }),
                fetch('/api/threatintelligence/secure-score', { headers })
            ]);

            if (alertsRes.ok) {
                const data = await alertsRes.json();
                setAlerts(data.alerts || []);
                setAlertSummary(data.summary);
            }
            if (riskyUsersRes.ok) {
                const data = await riskyUsersRes.json();
                setRiskyUsers(data.users || []);
            }
            if (detectionsRes.ok) {
                const data = await detectionsRes.json();
                setRiskDetections(data.detections || []);
            }
            if (scoreRes.ok) {
                const data = await scoreRes.json();
                setSecureScore(data.score);
            }
        } catch (err) {
            setError(err instanceof Error ? err.message : 'Failed to load data');
        } finally {
            setLoading(false);
        }
    }, [getAccessToken, days]);

    useEffect(() => { fetchData(); }, [fetchData]);

    if (loading) {
        return <div className="flex items-center justify-center min-h-[400px]"><Spinner size="large" label="Loading..." /></div>;
    }

    const severityData = alertSummary ? [
        { name: 'High', value: alertSummary.highSeverity, color: '#ef4444' },
        { name: 'Medium', value: alertSummary.mediumSeverity, color: '#f59e0b' },
        { name: 'Low', value: alertSummary.lowSeverity, color: '#22c55e' },
    ].filter(d => d.value > 0) : [];

    return (
        <div className="p-6 space-y-6">
            <div className="flex items-start justify-between">
                <div>
                    <h1 className="text-2xl font-bold text-gray-900 dark:text-white flex items-center gap-2">
                        <ShieldError24Regular className="w-7 h-7" />
                        Threat Intelligence
                    </h1>
                    <p className="mt-1 text-sm text-gray-500 dark:text-gray-400">Security alerts, risky users, and risk detections</p>
                </div>
                <div className="flex items-center gap-3">
                    <select value={days} onChange={(e) => setDays(Number(e.target.value))}
                        className="px-3 py-1.5 border border-gray-300 dark:border-gray-600 rounded-lg bg-white dark:bg-gray-700 text-sm">
                        <option value={7}>Last 7 days</option>
                        <option value={30}>Last 30 days</option>
                        <option value={90}>Last 90 days</option>
                    </select>
                    <button onClick={fetchData} className="p-2 hover:bg-gray-100 dark:hover:bg-gray-700 rounded-lg">
                        <ArrowSync24Regular className="w-5 h-5" />
                    </button>
                </div>
            </div>

            {error && <div className="bg-red-50 dark:bg-red-900/20 border border-red-200 p-4 rounded-lg text-red-700">{error}</div>}

            {/* Summary Cards */}
            <div className="grid grid-cols-2 md:grid-cols-5 gap-4">
                <Card className="p-4">
                    <p className="text-3xl font-bold text-gray-900 dark:text-white">{alertSummary?.totalAlerts || 0}</p>
                    <p className="text-sm text-gray-500">Security Alerts</p>
                </Card>
                <Card className="p-4">
                    <p className="text-3xl font-bold text-red-600">{alertSummary?.highSeverity || 0}</p>
                    <p className="text-sm text-gray-500">High Severity</p>
                </Card>
                <Card className="p-4">
                    <p className="text-3xl font-bold text-amber-600">{riskyUsers.length}</p>
                    <p className="text-sm text-gray-500">Risky Users</p>
                </Card>
                <Card className="p-4">
                    <p className="text-3xl font-bold text-purple-600">{riskDetections.length}</p>
                    <p className="text-sm text-gray-500">Risk Detections</p>
                </Card>
                <Card className="p-4">
                    <p className="text-3xl font-bold text-blue-600">{secureScore?.percentage || 0}%</p>
                    <p className="text-sm text-gray-500">Secure Score</p>
                </Card>
            </div>

            {/* Secure Score */}
            {secureScore && (
                <Card className="p-5">
                    <div className="flex items-center justify-between mb-4">
                        <h3 className="font-semibold flex items-center gap-2">
                            <Shield24Regular className="w-5 h-5" />
                            Microsoft Secure Score
                        </h3>
                        <span className="text-2xl font-bold">{secureScore.currentScore}/{secureScore.maxScore}</span>
                    </div>
                    <ProgressBar value={secureScore.percentage / 100} color={secureScore.percentage >= 70 ? 'success' : secureScore.percentage >= 40 ? 'warning' : 'error'} />
                    <p className="text-sm text-gray-500 mt-2">Active Users: {secureScore.activeUserCount?.toLocaleString()}</p>
                </Card>
            )}

            <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
                {/* Alert Severity */}
                <Card className="p-5">
                    <h3 className="font-semibold mb-4">Alerts by Severity</h3>
                    <div className="h-64">
                        <ResponsiveContainer width="100%" height="100%">
                            <PieChart>
                                <Pie data={severityData} dataKey="value" nameKey="name" cx="50%" cy="50%" outerRadius={80} label>
                                    {severityData.map((entry, index) => (
                                        <Cell key={index} fill={entry.color} />
                                    ))}
                                </Pie>
                                <Legend />
                                <Tooltip />
                            </PieChart>
                        </ResponsiveContainer>
                    </div>
                </Card>

                {/* Risky Users */}
                <Card className="p-5">
                    <h3 className="font-semibold mb-4 flex items-center gap-2">
                        <Warning24Regular className="w-5 h-5 text-amber-500" />
                        Risky Users ({riskyUsers.length})
                    </h3>
                    {riskyUsers.length === 0 ? (
                        <p className="text-gray-500 text-center py-8">No risky users detected</p>
                    ) : (
                        <div className="space-y-2 max-h-64 overflow-y-auto">
                            {riskyUsers.slice(0, 10).map((user: any) => (
                                <div key={user.id} className="flex items-center justify-between p-2 bg-gray-50 dark:bg-gray-700/50 rounded-lg">
                                    <div className="flex items-center gap-2">
                                        <Person24Regular className="w-5 h-5" />
                                        <div>
                                            <p className="font-medium text-sm">{user.userDisplayName}</p>
                                            <p className="text-xs text-gray-500">{user.userPrincipalName}</p>
                                        </div>
                                    </div>
                                    <Badge appearance="filled" color={user.riskLevel === 'high' ? 'danger' : user.riskLevel === 'medium' ? 'warning' : 'success'} size="small">
                                        {user.riskLevel}
                                    </Badge>
                                </div>
                            ))}
                        </div>
                    )}
                </Card>
            </div>

            {/* Security Alerts */}
            {alerts.length > 0 && (
                <Card className="p-5">
                    <h3 className="font-semibold mb-4 flex items-center gap-2">
                        <Bug24Regular className="w-5 h-5 text-red-500" />
                        Recent Security Alerts
                    </h3>
                    <div className="overflow-x-auto">
                        <table className="w-full text-sm">
                            <thead>
                                <tr className="border-b dark:border-gray-700">
                                    <th className="text-left py-2 px-3">Alert</th>
                                    <th className="text-left py-2 px-3">Severity</th>
                                    <th className="text-left py-2 px-3">Status</th>
                                    <th className="text-left py-2 px-3">Category</th>
                                    <th className="text-left py-2 px-3">Time</th>
                                </tr>
                            </thead>
                            <tbody className="divide-y dark:divide-gray-700">
                                {alerts.slice(0, 20).map((alert: any) => (
                                    <tr key={alert.id} className="hover:bg-gray-50 dark:hover:bg-gray-700/50">
                                        <td className="py-2 px-3">
                                            <p className="font-medium">{alert.title}</p>
                                            <p className="text-xs text-gray-500 truncate max-w-md">{alert.description}</p>
                                        </td>
                                        <td className="py-2 px-3">
                                            <Badge appearance="filled" color={alert.severity === 'High' ? 'danger' : alert.severity === 'Medium' ? 'warning' : 'success'} size="small">
                                                {alert.severity}
                                            </Badge>
                                        </td>
                                        <td className="py-2 px-3">
                                            <Badge appearance="tint" color={alert.status === 'Resolved' ? 'success' : alert.status === 'InProgress' ? 'warning' : 'brand'} size="small">
                                                {alert.status}
                                            </Badge>
                                        </td>
                                        <td className="py-2 px-3 text-gray-600">{alert.category || '-'}</td>
                                        <td className="py-2 px-3 text-gray-500">
                                            {alert.createdDateTime ? new Date(alert.createdDateTime).toLocaleString() : '-'}
                                        </td>
                                    </tr>
                                ))}
                            </tbody>
                        </table>
                    </div>
                </Card>
            )}

            {/* Risk Detections */}
            {riskDetections.length > 0 && (
                <Card className="p-5">
                    <h3 className="font-semibold mb-4 flex items-center gap-2">
                        <Eye24Regular className="w-5 h-5" />
                        Recent Risk Detections
                    </h3>
                    <div className="overflow-x-auto">
                        <table className="w-full text-sm">
                            <thead>
                                <tr className="border-b dark:border-gray-700">
                                    <th className="text-left py-2 px-3">User</th>
                                    <th className="text-left py-2 px-3">Risk Type</th>
                                    <th className="text-left py-2 px-3">Level</th>
                                    <th className="text-left py-2 px-3">Location</th>
                                    <th className="text-left py-2 px-3">Detected</th>
                                </tr>
                            </thead>
                            <tbody className="divide-y dark:divide-gray-700">
                                {riskDetections.slice(0, 15).map((detection: any) => (
                                    <tr key={detection.id} className="hover:bg-gray-50 dark:hover:bg-gray-700/50">
                                        <td className="py-2 px-3">
                                            <p className="font-medium">{detection.userDisplayName}</p>
                                            <p className="text-xs text-gray-500">{detection.userPrincipalName}</p>
                                        </td>
                                        <td className="py-2 px-3">{detection.riskType || '-'}</td>
                                        <td className="py-2 px-3">
                                            <Badge appearance="filled" color={detection.riskLevel === 'high' ? 'danger' : detection.riskLevel === 'medium' ? 'warning' : 'success'} size="small">
                                                {detection.riskLevel}
                                            </Badge>
                                        </td>
                                        <td className="py-2 px-3 text-gray-600">
                                            {detection.location ? `${detection.location.city || ''}, ${detection.location.country || ''}` : detection.ipAddress || '-'}
                                        </td>
                                        <td className="py-2 px-3 text-gray-500">
                                            {detection.detectedDateTime ? new Date(detection.detectedDateTime).toLocaleString() : '-'}
                                        </td>
                                    </tr>
                                ))}
                            </tbody>
                        </table>
                    </div>
                </Card>
            )}
        </div>
    );
};

export default ThreatIntelligencePage;
