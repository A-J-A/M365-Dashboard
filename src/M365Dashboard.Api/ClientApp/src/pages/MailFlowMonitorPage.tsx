import React, { useState, useEffect, useCallback } from 'react';
import {
    Mail24Regular,
    Send24Regular,
    MailRead24Regular,
    ArrowSync24Regular,
    Warning24Regular,
    ArrowForward24Regular,
    Person24Regular,
    Link24Regular,
} from '@fluentui/react-icons';
import { Card, Spinner, Badge } from '@fluentui/react-components';
import { BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, LineChart, Line, Legend } from 'recharts';
import { useAppContext } from '../contexts/AppContext';

const MailFlowMonitorPage: React.FC = () => {
    const { getAccessToken } = useAppContext();
    const [trafficData, setTrafficData] = useState<any>(null);
    const [topSenders, setTopSenders] = useState<any[]>([]);
    const [topReceivers, setTopReceivers] = useState<any[]>([]);
    const [forwardingConfigs, setForwardingConfigs] = useState<any[]>([]);
    const [threatSummary, setThreatSummary] = useState<any>(null);
    const [loading, setLoading] = useState(true);
    const [error, setError] = useState<string | null>(null);
    const [days, setDays] = useState(7);

    const fetchData = useCallback(async () => {
        try {
            setLoading(true);
            const token = await getAccessToken();
            const headers = { 'Authorization': `Bearer ${token}` };

            const [trafficRes, sendersRes, forwardingRes, threatRes] = await Promise.all([
                fetch(`/api/mailflow/traffic-summary?days=${days}`, { headers }),
                fetch(`/api/mailflow/top-senders?days=${days}`, { headers }),
                fetch('/api/mailflow/forwarding-configs?top=50', { headers }),
                fetch(`/api/mailflow/threat-summary?days=${days}`, { headers })
            ]);

            if (trafficRes.ok) {
                const data = await trafficRes.json();
                setTrafficData(data);
            }
            if (sendersRes.ok) {
                const data = await sendersRes.json();
                setTopSenders(data.topSenders || []);
                setTopReceivers(data.topReceivers || []);
            }
            if (forwardingRes.ok) {
                const data = await forwardingRes.json();
                setForwardingConfigs(data.users || []);
            }
            if (threatRes.ok) {
                const data = await threatRes.json();
                setThreatSummary(data);
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

    return (
        <div className="p-6 space-y-6">
            <div className="flex items-start justify-between">
                <div>
                    <h1 className="text-2xl font-bold text-gray-900 dark:text-white flex items-center gap-2">
                        <Mail24Regular className="w-7 h-7" />
                        Mail Flow Monitor
                    </h1>
                    <p className="mt-1 text-sm text-gray-500 dark:text-gray-400">Email traffic, forwarding configurations, and mail-related security</p>
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
            <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
                <Card className="p-4">
                    <div className="flex items-center gap-2 mb-2">
                        <Send24Regular className="w-5 h-5 text-blue-500" />
                        <span className="text-sm text-gray-500">Sent</span>
                    </div>
                    <p className="text-3xl font-bold text-blue-600">{trafficData?.summary?.totalSent?.toLocaleString() || 0}</p>
                </Card>
                <Card className="p-4">
                    <div className="flex items-center gap-2 mb-2">
                        <MailRead24Regular className="w-5 h-5 text-green-500" />
                        <span className="text-sm text-gray-500">Received</span>
                    </div>
                    <p className="text-3xl font-bold text-green-600">{trafficData?.summary?.totalReceived?.toLocaleString() || 0}</p>
                </Card>
                <Card className="p-4">
                    <div className="flex items-center gap-2 mb-2">
                        <ArrowForward24Regular className="w-5 h-5 text-amber-500" />
                        <span className="text-sm text-gray-500">Forwarding Rules</span>
                    </div>
                    <p className="text-3xl font-bold text-amber-600">{forwardingConfigs.length}</p>
                </Card>
                <Card className="p-4">
                    <div className="flex items-center gap-2 mb-2">
                        <Warning24Regular className="w-5 h-5 text-red-500" />
                        <span className="text-sm text-gray-500">Mail Alerts</span>
                    </div>
                    <p className="text-3xl font-bold text-red-600">{threatSummary?.summary?.totalMailAlerts || 0}</p>
                </Card>
            </div>

            {/* Traffic Trend */}
            {trafficData?.dailyData?.length > 0 && (
                <Card className="p-5">
                    <h3 className="font-semibold mb-4">Email Traffic Trend</h3>
                    <div className="h-64">
                        <ResponsiveContainer width="100%" height="100%">
                            <LineChart data={trafficData.dailyData}>
                                <XAxis dataKey="reportDate" tickFormatter={(v) => new Date(v).toLocaleDateString('en-GB', { day: 'numeric', month: 'short' })} />
                                <YAxis />
                                <Tooltip labelFormatter={(v) => new Date(v).toLocaleDateString()} />
                                <Legend />
                                <Line type="monotone" dataKey="send" name="Sent" stroke="#3b82f6" strokeWidth={2} dot={false} />
                                <Line type="monotone" dataKey="receive" name="Received" stroke="#22c55e" strokeWidth={2} dot={false} />
                            </LineChart>
                        </ResponsiveContainer>
                    </div>
                </Card>
            )}

            <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
                {/* Top Senders */}
                <Card className="p-5">
                    <h3 className="font-semibold mb-4 flex items-center gap-2">
                        <Send24Regular className="w-5 h-5" />
                        Top Email Senders
                    </h3>
                    {topSenders.length === 0 ? (
                        <p className="text-gray-500 text-center py-8">No data available</p>
                    ) : (
                        <div className="space-y-2">
                            {topSenders.slice(0, 10).map((sender: any, index: number) => (
                                <div key={index} className="flex items-center justify-between p-2 bg-gray-50 dark:bg-gray-700/50 rounded-lg">
                                    <div className="flex items-center gap-2 min-w-0">
                                        <span className="text-sm font-medium text-gray-500 w-5">{index + 1}</span>
                                        <Person24Regular className="w-5 h-5 flex-shrink-0" />
                                        <div className="min-w-0">
                                            <p className="font-medium text-sm truncate">{sender.displayName || sender.userPrincipalName}</p>
                                        </div>
                                    </div>
                                    <Badge appearance="filled" color="brand">{sender.sendCount?.toLocaleString()}</Badge>
                                </div>
                            ))}
                        </div>
                    )}
                </Card>

                {/* Top Receivers */}
                <Card className="p-5">
                    <h3 className="font-semibold mb-4 flex items-center gap-2">
                        <MailRead24Regular className="w-5 h-5" />
                        Top Email Receivers
                    </h3>
                    {topReceivers.length === 0 ? (
                        <p className="text-gray-500 text-center py-8">No data available</p>
                    ) : (
                        <div className="space-y-2">
                            {topReceivers.slice(0, 10).map((receiver: any, index: number) => (
                                <div key={index} className="flex items-center justify-between p-2 bg-gray-50 dark:bg-gray-700/50 rounded-lg">
                                    <div className="flex items-center gap-2 min-w-0">
                                        <span className="text-sm font-medium text-gray-500 w-5">{index + 1}</span>
                                        <Person24Regular className="w-5 h-5 flex-shrink-0" />
                                        <div className="min-w-0">
                                            <p className="font-medium text-sm truncate">{receiver.displayName || receiver.userPrincipalName}</p>
                                        </div>
                                    </div>
                                    <Badge appearance="filled" color="success">{receiver.receiveCount?.toLocaleString()}</Badge>
                                </div>
                            ))}
                        </div>
                    )}
                </Card>
            </div>

            {/* Forwarding Configurations */}
            {forwardingConfigs.length > 0 && (
                <Card className="p-5">
                    <h3 className="font-semibold mb-4 flex items-center gap-2">
                        <ArrowForward24Regular className="w-5 h-5 text-amber-500" />
                        Email Forwarding Rules
                        <Badge appearance="filled" color="warning">{forwardingConfigs.length}</Badge>
                    </h3>
                    <div className="overflow-x-auto">
                        <table className="w-full text-sm">
                            <thead>
                                <tr className="border-b dark:border-gray-700">
                                    <th className="text-left py-2 px-3">User</th>
                                    <th className="text-left py-2 px-3">Rule</th>
                                    <th className="text-left py-2 px-3">Forward To</th>
                                    <th className="text-center py-2 px-3">Status</th>
                                </tr>
                            </thead>
                            <tbody className="divide-y dark:divide-gray-700">
                                {forwardingConfigs.map((config: any, index: number) => (
                                    config.forwardingRules?.map((rule: any, ruleIndex: number) => (
                                        <tr key={`${index}-${ruleIndex}`} className="hover:bg-gray-50 dark:hover:bg-gray-700/50">
                                            <td className="py-2 px-3">
                                                <p className="font-medium">{config.userDisplayName}</p>
                                                <p className="text-xs text-gray-500">{config.userPrincipalName}</p>
                                            </td>
                                            <td className="py-2 px-3">{rule.ruleName || 'Unnamed Rule'}</td>
                                            <td className="py-2 px-3">
                                                <div className="space-y-1">
                                                    {rule.forwardTo?.map((email: string, i: number) => (
                                                        <Badge key={i} appearance="tint" size="small">{email}</Badge>
                                                    ))}
                                                    {rule.redirectTo?.map((email: string, i: number) => (
                                                        <Badge key={i} appearance="tint" color="warning" size="small">Redirect: {email}</Badge>
                                                    ))}
                                                </div>
                                            </td>
                                            <td className="py-2 px-3 text-center">
                                                <Badge appearance="filled" color={rule.isEnabled ? 'success' : 'subtle'} size="small">
                                                    {rule.isEnabled ? 'Active' : 'Disabled'}
                                                </Badge>
                                            </td>
                                        </tr>
                                    ))
                                ))}
                            </tbody>
                        </table>
                    </div>
                </Card>
            )}

            {/* Quick Links */}
            <Card className="p-5">
                <h3 className="font-semibold mb-4">Additional Mail Flow Tools</h3>
                <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                    <a href="https://security.microsoft.com/messagetrace" target="_blank" rel="noopener noreferrer"
                        className="flex items-center gap-3 p-4 bg-blue-50 dark:bg-blue-900/20 rounded-lg hover:bg-blue-100 dark:hover:bg-blue-900/40 transition-colors">
                        <Link24Regular className="w-6 h-6 text-blue-600" />
                        <div>
                            <p className="font-medium">Message Trace</p>
                            <p className="text-xs text-gray-500">Detailed email tracking</p>
                        </div>
                    </a>
                    <a href="https://security.microsoft.com/quarantine" target="_blank" rel="noopener noreferrer"
                        className="flex items-center gap-3 p-4 bg-amber-50 dark:bg-amber-900/20 rounded-lg hover:bg-amber-100 dark:hover:bg-amber-900/40 transition-colors">
                        <Link24Regular className="w-6 h-6 text-amber-600" />
                        <div>
                            <p className="font-medium">Quarantine</p>
                            <p className="text-xs text-gray-500">Blocked messages</p>
                        </div>
                    </a>
                    <a href="https://admin.exchange.microsoft.com" target="_blank" rel="noopener noreferrer"
                        className="flex items-center gap-3 p-4 bg-purple-50 dark:bg-purple-900/20 rounded-lg hover:bg-purple-100 dark:hover:bg-purple-900/40 transition-colors">
                        <Link24Regular className="w-6 h-6 text-purple-600" />
                        <div>
                            <p className="font-medium">Exchange Admin</p>
                            <p className="text-xs text-gray-500">Full mail management</p>
                        </div>
                    </a>
                </div>
            </Card>
        </div>
    );
};

export default MailFlowMonitorPage;
