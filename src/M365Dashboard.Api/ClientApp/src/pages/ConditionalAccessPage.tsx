import React, { useState, useEffect, useCallback } from 'react';
import {
    Shield24Regular,
    Checkmark24Regular,
    Dismiss24Regular,
    Warning24Regular,
    Info24Regular,
    ArrowSync24Regular,
    LockClosed24Regular,
    Person24Regular,
    Apps24Regular,
} from '@fluentui/react-icons';
import { Card, Spinner, Badge, ProgressBar } from '@fluentui/react-components';
import { PieChart, Pie, Cell, ResponsiveContainer, BarChart, Bar, XAxis, YAxis, Tooltip, Legend } from 'recharts';
import { useAppContext } from '../contexts/AppContext';

const COLORS = ['#22c55e', '#ef4444', '#6b7280', '#f59e0b', '#3b82f6'];

const ConditionalAccessPage: React.FC = () => {
    const { getAccessToken } = useAppContext();
    const [policies, setPolicies] = useState<any[]>([]);
    const [summary, setSummary] = useState<any>(null);
    const [insights, setInsights] = useState<any>(null);
    const [loading, setLoading] = useState(true);
    const [error, setError] = useState<string | null>(null);
    const [days, setDays] = useState(7);

    const fetchData = useCallback(async () => {
        try {
            setLoading(true);
            const token = await getAccessToken();
            const headers = { 'Authorization': `Bearer ${token}` };

            const [policiesRes, insightsRes] = await Promise.all([
                fetch('/api/conditionalaccess/policies', { headers }),
                fetch(`/api/conditionalaccess/sign-in-insights?days=${days}`, { headers })
            ]);

            if (policiesRes.ok) {
                const data = await policiesRes.json();
                setPolicies(data.policies || []);
                setSummary(data.summary);
            }
            if (insightsRes.ok) {
                const data = await insightsRes.json();
                setInsights(data);
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

    const policyStateData = summary ? [
        { name: 'Enabled', value: summary.enabledPolicies, color: '#22c55e' },
        { name: 'Report Only', value: summary.reportOnlyPolicies, color: '#f59e0b' },
        { name: 'Disabled', value: summary.disabledPolicies, color: '#6b7280' },
    ] : [];

    const caResultsData = insights?.summary?.caResults ? [
        { name: 'Success', value: insights.summary.caResults.success },
        { name: 'Blocked', value: insights.summary.caResults.failure },
        { name: 'Not Applied', value: insights.summary.caResults.notApplied },
    ] : [];

    return (
        <div className="p-6 space-y-6">
            <div className="flex items-start justify-between">
                <div>
                    <h1 className="text-2xl font-bold text-gray-900 dark:text-white flex items-center gap-2">
                        <Shield24Regular className="w-7 h-7" />
                        Conditional Access Monitor
                    </h1>
                    <p className="mt-1 text-sm text-gray-500 dark:text-gray-400">CA policies and sign-in impact analysis</p>
                </div>
                <div className="flex items-center gap-3">
                    <select value={days} onChange={(e) => setDays(Number(e.target.value))}
                        className="px-3 py-1.5 border border-gray-300 dark:border-gray-600 rounded-lg bg-white dark:bg-gray-700 text-sm">
                        <option value={1}>Last 24 hours</option>
                        <option value={7}>Last 7 days</option>
                        <option value={30}>Last 30 days</option>
                    </select>
                    <button onClick={fetchData} className="p-2 hover:bg-gray-100 dark:hover:bg-gray-700 rounded-lg">
                        <ArrowSync24Regular className="w-5 h-5" />
                    </button>
                </div>
            </div>

            {error && <div className="bg-red-50 dark:bg-red-900/20 border border-red-200 dark:border-red-800 rounded-lg p-4 text-red-700 dark:text-red-300">{error}</div>}

            {/* Summary Cards */}
            <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
                <Card className="p-4">
                    <p className="text-3xl font-bold text-gray-900 dark:text-white">{summary?.totalPolicies || 0}</p>
                    <p className="text-sm text-gray-500">Total Policies</p>
                </Card>
                <Card className="p-4">
                    <p className="text-3xl font-bold text-green-600">{summary?.enabledPolicies || 0}</p>
                    <p className="text-sm text-gray-500">Enabled</p>
                </Card>
                <Card className="p-4">
                    <p className="text-3xl font-bold text-amber-600">{summary?.reportOnlyPolicies || 0}</p>
                    <p className="text-sm text-gray-500">Report Only</p>
                </Card>
                <Card className="p-4">
                    <p className="text-3xl font-bold text-red-600">{insights?.blockedSignIns?.length || 0}</p>
                    <p className="text-sm text-gray-500">Blocked Sign-ins</p>
                </Card>
            </div>

            {/* Charts Row */}
            <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
                <Card className="p-5">
                    <h3 className="font-semibold mb-4">Policy States</h3>
                    <div className="h-64">
                        <ResponsiveContainer width="100%" height="100%">
                            <PieChart>
                                <Pie data={policyStateData} dataKey="value" nameKey="name" cx="50%" cy="50%" outerRadius={80} label>
                                    {policyStateData.map((entry, index) => (
                                        <Cell key={index} fill={entry.color} />
                                    ))}
                                </Pie>
                                <Legend />
                                <Tooltip />
                            </PieChart>
                        </ResponsiveContainer>
                    </div>
                </Card>

                <Card className="p-5">
                    <h3 className="font-semibold mb-4">CA Evaluation Results</h3>
                    <div className="h-64">
                        <ResponsiveContainer width="100%" height="100%">
                            <BarChart data={caResultsData}>
                                <XAxis dataKey="name" />
                                <YAxis />
                                <Tooltip />
                                <Bar dataKey="value" fill="#3b82f6" radius={[4, 4, 0, 0]} />
                            </BarChart>
                        </ResponsiveContainer>
                    </div>
                </Card>
            </div>

            {/* Policy Hits */}
            {insights?.policyHits?.length > 0 && (
                <Card className="p-5">
                    <h3 className="font-semibold mb-4">Top Policy Activity</h3>
                    <div className="space-y-3">
                        {insights.policyHits.slice(0, 10).map((hit: any, index: number) => (
                            <div key={index} className="flex items-center justify-between p-3 bg-gray-50 dark:bg-gray-700/50 rounded-lg">
                                <div className="flex-1 min-w-0">
                                    <p className="font-medium text-sm truncate">{hit.policyName}</p>
                                </div>
                                <div className="flex items-center gap-4 text-sm">
                                    <span className="text-green-600">{hit.successCount} passed</span>
                                    <span className="text-red-600">{hit.failureCount} blocked</span>
                                </div>
                            </div>
                        ))}
                    </div>
                </Card>
            )}

            {/* Blocked Sign-ins */}
            {insights?.blockedSignIns?.length > 0 && (
                <Card className="p-5">
                    <h3 className="font-semibold mb-4 flex items-center gap-2">
                        <Dismiss24Regular className="w-5 h-5 text-red-500" />
                        Recent Blocked Sign-ins
                    </h3>
                    <div className="overflow-x-auto">
                        <table className="w-full text-sm">
                            <thead>
                                <tr className="border-b dark:border-gray-700">
                                    <th className="text-left py-2 px-3">User</th>
                                    <th className="text-left py-2 px-3">App</th>
                                    <th className="text-left py-2 px-3">Location</th>
                                    <th className="text-left py-2 px-3">Policy</th>
                                    <th className="text-left py-2 px-3">Time</th>
                                </tr>
                            </thead>
                            <tbody className="divide-y dark:divide-gray-700">
                                {insights.blockedSignIns.slice(0, 20).map((signIn: any) => (
                                    <tr key={signIn.id} className="hover:bg-gray-50 dark:hover:bg-gray-700/50">
                                        <td className="py-2 px-3">
                                            <p className="font-medium">{signIn.userDisplayName}</p>
                                            <p className="text-xs text-gray-500">{signIn.userPrincipalName}</p>
                                        </td>
                                        <td className="py-2 px-3">{signIn.appDisplayName || '-'}</td>
                                        <td className="py-2 px-3">{signIn.location || '-'}</td>
                                        <td className="py-2 px-3">
                                            {signIn.blockedByPolicies?.map((p: any, i: number) => (
                                                <Badge key={i} appearance="tint" color="danger" size="small" className="mr-1">
                                                    {p.displayName}
                                                </Badge>
                                            ))}
                                        </td>
                                        <td className="py-2 px-3 text-gray-500">
                                            {signIn.createdDateTime ? new Date(signIn.createdDateTime).toLocaleString() : '-'}
                                        </td>
                                    </tr>
                                ))}
                            </tbody>
                        </table>
                    </div>
                </Card>
            )}

            {/* All Policies */}
            <Card className="p-5">
                <h3 className="font-semibold mb-4">All Conditional Access Policies</h3>
                <div className="overflow-x-auto">
                    <table className="w-full text-sm">
                        <thead>
                            <tr className="border-b dark:border-gray-700">
                                <th className="text-left py-2 px-3">Policy Name</th>
                                <th className="text-left py-2 px-3">State</th>
                                <th className="text-left py-2 px-3">Grant Controls</th>
                                <th className="text-left py-2 px-3">Modified</th>
                            </tr>
                        </thead>
                        <tbody className="divide-y dark:divide-gray-700">
                            {policies.map((policy: any) => (
                                <tr key={policy.id} className="hover:bg-gray-50 dark:hover:bg-gray-700/50">
                                    <td className="py-2 px-3 font-medium">{policy.displayName}</td>
                                    <td className="py-2 px-3">
                                        <Badge appearance="filled" 
                                            color={policy.state === 'Enabled' ? 'success' : policy.state === 'EnabledForReportingButNotEnforced' ? 'warning' : 'subtle'}>
                                            {policy.state === 'EnabledForReportingButNotEnforced' ? 'Report Only' : policy.state}
                                        </Badge>
                                    </td>
                                    <td className="py-2 px-3">
                                        <div className="flex flex-wrap gap-1">
                                            {policy.grantControls?.builtInControls?.map((ctrl: string, i: number) => (
                                                <Badge key={i} appearance="tint" size="small">{ctrl}</Badge>
                                            ))}
                                        </div>
                                    </td>
                                    <td className="py-2 px-3 text-gray-500">
                                        {policy.modifiedDateTime ? new Date(policy.modifiedDateTime).toLocaleDateString() : '-'}
                                    </td>
                                </tr>
                            ))}
                        </tbody>
                    </table>
                </div>
            </Card>
        </div>
    );
};

export default ConditionalAccessPage;
