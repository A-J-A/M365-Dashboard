import React, { useState, useEffect, useCallback } from 'react';
import {
    PersonKey24Regular,
    Person24Regular,
    Key24Regular,
    Warning24Regular,
    ArrowSync24Regular,
    Crown24Regular,
    History24Regular,
    Apps24Regular,
} from '@fluentui/react-icons';
import { Card, Spinner, Badge } from '@fluentui/react-components';
import { PieChart, Pie, Cell, ResponsiveContainer, Tooltip, Legend } from 'recharts';
import { useAppContext } from '../contexts/AppContext';

const COLORS = ['#ef4444', '#f59e0b', '#22c55e', '#3b82f6', '#8b5cf6', '#ec4899'];

const PrivilegedAccessPage: React.FC = () => {
    const { getAccessToken } = useAppContext();
    const [roles, setRoles] = useState<any[]>([]);
    const [roleChanges, setRoleChanges] = useState<any[]>([]);
    const [privilegedSignIns, setPrivilegedSignIns] = useState<any>(null);
    const [summary, setSummary] = useState<any>(null);
    const [loading, setLoading] = useState(true);
    const [error, setError] = useState<string | null>(null);
    const [days, setDays] = useState(30);

    const fetchData = useCallback(async () => {
        try {
            setLoading(true);
            const token = await getAccessToken();
            const headers = { 'Authorization': `Bearer ${token}` };

            const [rolesRes, changesRes, signInsRes] = await Promise.all([
                fetch('/api/privilegedaccess/directory-roles', { headers }),
                fetch(`/api/privilegedaccess/role-changes?days=${days}`, { headers }),
                fetch(`/api/privilegedaccess/privileged-sign-ins?days=${Math.min(days, 7)}`, { headers })
            ]);

            if (rolesRes.ok) {
                const data = await rolesRes.json();
                setRoles(data.roles || []);
                setSummary(data.summary);
            }
            if (changesRes.ok) {
                const data = await changesRes.json();
                setRoleChanges(data.changes || []);
            }
            if (signInsRes.ok) {
                const data = await signInsRes.json();
                setPrivilegedSignIns(data);
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

    const privilegedRoles = roles.filter(r => r.isPrivileged);
    const globalAdmins = roles.find(r => r.isGlobalAdmin);

    const topRolesData = privilegedRoles
        .filter(r => r.memberCount > 0)
        .slice(0, 6)
        .map((r, i) => ({ name: r.displayName?.replace(' Administrator', '').replace('Global ', 'GA'), value: r.memberCount, color: COLORS[i % COLORS.length] }));

    return (
        <div className="p-6 space-y-6">
            <div className="flex items-start justify-between">
                <div>
                    <h1 className="text-2xl font-bold text-gray-900 dark:text-white flex items-center gap-2">
                        <PersonKey24Regular className="w-7 h-7" />
                        Privileged Access Monitor
                    </h1>
                    <p className="mt-1 text-sm text-gray-500 dark:text-gray-400">Monitor admin roles, assignments, and privileged sign-ins</p>
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

            {error && <div className="bg-red-50 dark:bg-red-900/20 border border-red-200 dark:border-red-800 rounded-lg p-4 text-red-700 dark:text-red-300">{error}</div>}

            {/* Summary Cards */}
            <div className="grid grid-cols-2 md:grid-cols-5 gap-4">
                <Card className="p-4">
                    <div className="flex items-center gap-2 mb-2">
                        <Crown24Regular className="w-5 h-5 text-amber-500" />
                        <span className="text-sm text-gray-500">Global Admins</span>
                    </div>
                    <p className="text-3xl font-bold text-amber-600">{globalAdmins?.memberCount || 0}</p>
                </Card>
                <Card className="p-4">
                    <p className="text-3xl font-bold text-gray-900 dark:text-white">{summary?.totalRoles || 0}</p>
                    <p className="text-sm text-gray-500">Active Roles</p>
                </Card>
                <Card className="p-4">
                    <p className="text-3xl font-bold text-blue-600">{summary?.totalAssignments || 0}</p>
                    <p className="text-sm text-gray-500">Total Assignments</p>
                </Card>
                <Card className="p-4">
                    <p className="text-3xl font-bold text-purple-600">{roleChanges.length}</p>
                    <p className="text-sm text-gray-500">Recent Changes</p>
                </Card>
                <Card className="p-4">
                    <p className="text-3xl font-bold text-red-600">{privilegedSignIns?.summary?.riskySignIns || 0}</p>
                    <p className="text-sm text-gray-500">Risky Sign-ins</p>
                </Card>
            </div>

            {/* Global Admin Alert */}
            {globalAdmins?.memberCount > 5 && (
                <div className="bg-amber-50 dark:bg-amber-900/20 border border-amber-200 dark:border-amber-800 rounded-lg p-4 flex items-start gap-3">
                    <Warning24Regular className="w-5 h-5 text-amber-600 flex-shrink-0 mt-0.5" />
                    <div>
                        <p className="font-medium text-amber-800 dark:text-amber-200">High Number of Global Administrators</p>
                        <p className="text-sm text-amber-700 dark:text-amber-300">
                            You have {globalAdmins.memberCount} Global Administrators. Microsoft recommends keeping this number below 5.
                        </p>
                    </div>
                </div>
            )}

            <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
                {/* Top Privileged Roles Chart */}
                <Card className="p-5">
                    <h3 className="font-semibold mb-4">Top Privileged Role Assignments</h3>
                    <div className="h-64">
                        <ResponsiveContainer width="100%" height="100%">
                            <PieChart>
                                <Pie data={topRolesData} dataKey="value" nameKey="name" cx="50%" cy="50%" outerRadius={80} label>
                                    {topRolesData.map((entry, index) => (
                                        <Cell key={index} fill={entry.color} />
                                    ))}
                                </Pie>
                                <Legend />
                                <Tooltip />
                            </PieChart>
                        </ResponsiveContainer>
                    </div>
                </Card>

                {/* Recent Role Changes */}
                <Card className="p-5">
                    <h3 className="font-semibold mb-4 flex items-center gap-2">
                        <History24Regular className="w-5 h-5" />
                        Recent Role Changes
                    </h3>
                    {roleChanges.length === 0 ? (
                        <p className="text-gray-500 text-center py-8">No role changes in this period</p>
                    ) : (
                        <div className="space-y-2 max-h-64 overflow-y-auto">
                            {roleChanges.slice(0, 10).map((change: any, index: number) => (
                                <div key={index} className="p-2 bg-gray-50 dark:bg-gray-700/50 rounded-lg text-sm">
                                    <div className="flex items-center justify-between">
                                        <Badge appearance="tint" color={change.activityDisplayName?.includes('Add') ? 'success' : 'danger'} size="small">
                                            {change.activityDisplayName?.includes('Add') ? 'Added' : 'Removed'}
                                        </Badge>
                                        <span className="text-xs text-gray-500">
                                            {change.activityDateTime ? new Date(change.activityDateTime).toLocaleDateString() : ''}
                                        </span>
                                    </div>
                                    <p className="mt-1 text-gray-700 dark:text-gray-300">
                                        {change.initiatedBy?.displayName || 'System'} → {change.targetResources?.[0]?.displayName || 'Unknown'}
                                    </p>
                                </div>
                            ))}
                        </div>
                    )}
                </Card>
            </div>

            {/* Global Administrators */}
            {globalAdmins?.members?.length > 0 && (
                <Card className="p-5">
                    <h3 className="font-semibold mb-4 flex items-center gap-2">
                        <Crown24Regular className="w-5 h-5 text-amber-500" />
                        Global Administrators
                    </h3>
                    <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-3">
                        {globalAdmins.members.map((admin: any) => (
                            <div key={admin.id} className="flex items-center gap-3 p-3 bg-amber-50 dark:bg-amber-900/20 rounded-lg">
                                <div className="w-10 h-10 rounded-full bg-amber-100 dark:bg-amber-800 flex items-center justify-center">
                                    <Person24Regular className="w-5 h-5 text-amber-600" />
                                </div>
                                <div className="min-w-0">
                                    <p className="font-medium text-sm truncate">{admin.displayName}</p>
                                    <p className="text-xs text-gray-500 truncate">{admin.userPrincipalName}</p>
                                </div>
                                {admin.accountEnabled === false && (
                                    <Badge appearance="filled" color="danger" size="small">Disabled</Badge>
                                )}
                            </div>
                        ))}
                    </div>
                </Card>
            )}

            {/* Privileged Sign-ins */}
            {privilegedSignIns?.signIns?.length > 0 && (
                <Card className="p-5">
                    <h3 className="font-semibold mb-4 flex items-center gap-2">
                        <Key24Regular className="w-5 h-5" />
                        Recent Privileged Account Sign-ins
                    </h3>
                    <div className="overflow-x-auto">
                        <table className="w-full text-sm">
                            <thead>
                                <tr className="border-b dark:border-gray-700">
                                    <th className="text-left py-2 px-3">User</th>
                                    <th className="text-left py-2 px-3">App</th>
                                    <th className="text-left py-2 px-3">Location</th>
                                    <th className="text-left py-2 px-3">Risk</th>
                                    <th className="text-left py-2 px-3">Time</th>
                                </tr>
                            </thead>
                            <tbody className="divide-y dark:divide-gray-700">
                                {privilegedSignIns.signIns.slice(0, 20).map((signIn: any) => (
                                    <tr key={signIn.id} className="hover:bg-gray-50 dark:hover:bg-gray-700/50">
                                        <td className="py-2 px-3">
                                            <p className="font-medium">{signIn.userDisplayName}</p>
                                        </td>
                                        <td className="py-2 px-3">{signIn.appDisplayName || '-'}</td>
                                        <td className="py-2 px-3">
                                            {signIn.location ? `${signIn.location.city || ''}, ${signIn.location.country || ''}` : signIn.ipAddress || '-'}
                                        </td>
                                        <td className="py-2 px-3">
                                            {signIn.riskLevel && signIn.riskLevel !== 'none' && signIn.riskLevel !== 'None' ? (
                                                <Badge appearance="filled" color={signIn.riskLevel === 'high' ? 'danger' : 'warning'} size="small">
                                                    {signIn.riskLevel}
                                                </Badge>
                                            ) : '-'}
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

            {/* All Privileged Roles */}
            <Card className="p-5">
                <h3 className="font-semibold mb-4">All Privileged Roles</h3>
                <div className="overflow-x-auto">
                    <table className="w-full text-sm">
                        <thead>
                            <tr className="border-b dark:border-gray-700">
                                <th className="text-left py-2 px-3">Role</th>
                                <th className="text-center py-2 px-3">Members</th>
                                <th className="text-left py-2 px-3">Type</th>
                            </tr>
                        </thead>
                        <tbody className="divide-y dark:divide-gray-700">
                            {privilegedRoles.map((role: any) => (
                                <tr key={role.id} className="hover:bg-gray-50 dark:hover:bg-gray-700/50">
                                    <td className="py-2 px-3">
                                        <div className="flex items-center gap-2">
                                            {role.isGlobalAdmin && <Crown24Regular className="w-4 h-4 text-amber-500" />}
                                            <span className="font-medium">{role.displayName}</span>
                                        </div>
                                    </td>
                                    <td className="py-2 px-3 text-center">
                                        <Badge appearance={role.memberCount > 0 ? 'filled' : 'outline'} color={role.memberCount > 10 ? 'warning' : 'brand'}>
                                            {role.memberCount}
                                        </Badge>
                                    </td>
                                    <td className="py-2 px-3">
                                        {role.isGlobalAdmin ? (
                                            <Badge appearance="filled" color="danger" size="small">Critical</Badge>
                                        ) : (
                                            <Badge appearance="tint" color="warning" size="small">Privileged</Badge>
                                        )}
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

export default PrivilegedAccessPage;
