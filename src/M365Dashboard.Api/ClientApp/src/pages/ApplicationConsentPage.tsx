import React, { useState, useEffect, useCallback } from 'react';
import {
    Apps24Regular,
    Warning24Regular,
    ArrowSync24Regular,
    ShieldCheckmark24Regular,
    Key24Regular,
    Person24Regular,
    Building24Regular,
    Clock24Regular,
} from '@fluentui/react-icons';
import { Card, Spinner, Badge, ProgressBar } from '@fluentui/react-components';
import { PieChart, Pie, Cell, ResponsiveContainer, Tooltip, Legend, BarChart, Bar, XAxis, YAxis } from 'recharts';
import { useAppContext } from '../contexts/AppContext';

const COLORS = ['#ef4444', '#f59e0b', '#22c55e', '#3b82f6', '#8b5cf6'];

const ApplicationConsentPage: React.FC = () => {
    const { getAccessToken } = useAppContext();
    const [oauth2Grants, setOauth2Grants] = useState<any>(null);
    const [enterpriseApps, setEnterpriseApps] = useState<any>(null);
    const [appRegistrations, setAppRegistrations] = useState<any>(null);
    const [riskyConsents, setRiskyConsents] = useState<any>(null);
    const [loading, setLoading] = useState(true);
    const [error, setError] = useState<string | null>(null);
    const [activeTab, setActiveTab] = useState<'overview' | 'grants' | 'apps' | 'risky'>('overview');

    const fetchData = useCallback(async () => {
        try {
            setLoading(true);
            const token = await getAccessToken();
            const headers = { 'Authorization': `Bearer ${token}` };

            const [grantsRes, appsRes, registrationsRes, riskyRes] = await Promise.all([
                fetch('/api/applicationconsent/oauth2-grants', { headers }),
                fetch('/api/applicationconsent/enterprise-apps', { headers }),
                fetch('/api/applicationconsent/app-registrations', { headers }),
                fetch('/api/applicationconsent/risky-consents', { headers })
            ]);

            if (grantsRes.ok) setOauth2Grants(await grantsRes.json());
            if (appsRes.ok) setEnterpriseApps(await appsRes.json());
            if (registrationsRes.ok) setAppRegistrations(await registrationsRes.json());
            if (riskyRes.ok) setRiskyConsents(await riskyRes.json());
        } catch (err) {
            setError(err instanceof Error ? err.message : 'Failed to load data');
        } finally {
            setLoading(false);
        }
    }, [getAccessToken]);

    useEffect(() => { fetchData(); }, [fetchData]);

    if (loading) {
        return <div className="flex items-center justify-center min-h-[400px]"><Spinner size="large" label="Loading..." /></div>;
    }

    const riskLevelData = riskyConsents?.summary ? [
        { name: 'High', value: riskyConsents.summary.highRisk, color: '#ef4444' },
        { name: 'Medium', value: riskyConsents.summary.mediumRisk, color: '#f59e0b' },
        { name: 'Low', value: riskyConsents.summary.lowRisk, color: '#22c55e' },
    ].filter(d => d.value > 0) : [];

    return (
        <div className="p-6 space-y-6">
            <div className="flex items-start justify-between">
                <div>
                    <h1 className="text-2xl font-bold text-gray-900 dark:text-white flex items-center gap-2">
                        <Apps24Regular className="w-7 h-7" />
                        Application Consent & Permissions
                    </h1>
                    <p className="mt-1 text-sm text-gray-500 dark:text-gray-400">OAuth consents, enterprise apps, and permission analysis</p>
                </div>
                <button onClick={fetchData} className="p-2 hover:bg-gray-100 dark:hover:bg-gray-700 rounded-lg">
                    <ArrowSync24Regular className="w-5 h-5" />
                </button>
            </div>

            {error && <div className="bg-red-50 dark:bg-red-900/20 border border-red-200 p-4 rounded-lg text-red-700">{error}</div>}

            {/* Summary Cards */}
            <div className="grid grid-cols-2 md:grid-cols-5 gap-4">
                <Card className="p-4">
                    <p className="text-3xl font-bold text-gray-900 dark:text-white">{oauth2Grants?.summary?.totalGrants || 0}</p>
                    <p className="text-sm text-gray-500">OAuth Grants</p>
                </Card>
                <Card className="p-4">
                    <p className="text-3xl font-bold text-blue-600">{enterpriseApps?.summary?.totalApps || 0}</p>
                    <p className="text-sm text-gray-500">Enterprise Apps</p>
                </Card>
                <Card className="p-4">
                    <p className="text-3xl font-bold text-purple-600">{appRegistrations?.summary?.totalApps || 0}</p>
                    <p className="text-sm text-gray-500">App Registrations</p>
                </Card>
                <Card className="p-4">
                    <p className="text-3xl font-bold text-red-600">{riskyConsents?.summary?.totalRiskyConsents || 0}</p>
                    <p className="text-sm text-gray-500">Risky Consents</p>
                </Card>
                <Card className="p-4">
                    <p className="text-3xl font-bold text-amber-600">{appRegistrations?.summary?.appsWithExpiringSoonCredentials || 0}</p>
                    <p className="text-sm text-gray-500">Expiring Creds</p>
                </Card>
            </div>

            {/* Alerts */}
            {(riskyConsents?.summary?.highRisk > 0 || appRegistrations?.summary?.appsWithExpiredCredentials > 0) && (
                <div className="space-y-2">
                    {riskyConsents?.summary?.highRisk > 0 && (
                        <div className="bg-red-50 dark:bg-red-900/20 border border-red-200 dark:border-red-800 rounded-lg p-4 flex items-start gap-3">
                            <Warning24Regular className="w-5 h-5 text-red-600 flex-shrink-0 mt-0.5" />
                            <div>
                                <p className="font-medium text-red-800 dark:text-red-200">High-Risk App Consents Detected</p>
                                <p className="text-sm text-red-700 dark:text-red-300">
                                    {riskyConsents.summary.highRisk} applications have been granted high-risk permissions. Review these immediately.
                                </p>
                            </div>
                        </div>
                    )}
                    {appRegistrations?.summary?.appsWithExpiredCredentials > 0 && (
                        <div className="bg-amber-50 dark:bg-amber-900/20 border border-amber-200 dark:border-amber-800 rounded-lg p-4 flex items-start gap-3">
                            <Clock24Regular className="w-5 h-5 text-amber-600 flex-shrink-0 mt-0.5" />
                            <div>
                                <p className="font-medium text-amber-800 dark:text-amber-200">Expired Application Credentials</p>
                                <p className="text-sm text-amber-700 dark:text-amber-300">
                                    {appRegistrations.summary.appsWithExpiredCredentials} applications have expired credentials that need renewal.
                                </p>
                            </div>
                        </div>
                    )}
                </div>
            )}

            {/* Tabs */}
            <div className="flex gap-2 border-b dark:border-gray-700">
                {['overview', 'risky', 'grants', 'apps'].map((tab) => (
                    <button key={tab} onClick={() => setActiveTab(tab as any)}
                        className={`px-4 py-2 text-sm font-medium border-b-2 transition-colors ${
                            activeTab === tab 
                                ? 'border-blue-500 text-blue-600' 
                                : 'border-transparent text-gray-500 hover:text-gray-700'
                        }`}>
                        {tab === 'overview' && 'Overview'}
                        {tab === 'risky' && `Risky (${riskyConsents?.summary?.totalRiskyConsents || 0})`}
                        {tab === 'grants' && 'OAuth Grants'}
                        {tab === 'apps' && 'Enterprise Apps'}
                    </button>
                ))}
            </div>

            {/* Overview Tab */}
            {activeTab === 'overview' && (
                <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
                    <Card className="p-5">
                        <h3 className="font-semibold mb-4">Risky Consents by Level</h3>
                        <div className="h-64">
                            <ResponsiveContainer width="100%" height="100%">
                                <PieChart>
                                    <Pie data={riskLevelData} dataKey="value" nameKey="name" cx="50%" cy="50%" outerRadius={80} label>
                                        {riskLevelData.map((entry, index) => <Cell key={index} fill={entry.color} />)}
                                    </Pie>
                                    <Legend /><Tooltip />
                                </PieChart>
                            </ResponsiveContainer>
                        </div>
                    </Card>

                    <Card className="p-5">
                        <h3 className="font-semibold mb-4">Top Apps by Permissions</h3>
                        <div className="space-y-2">
                            {oauth2Grants?.topAppsByPermissions?.slice(0, 8).map((app: any, index: number) => (
                                <div key={index} className="flex items-center justify-between p-2 bg-gray-50 dark:bg-gray-700/50 rounded-lg">
                                    <div className="flex items-center gap-2 min-w-0">
                                        <Apps24Regular className="w-5 h-5 flex-shrink-0" />
                                        <span className="font-medium text-sm truncate">{app.appName}</span>
                                    </div>
                                    <div className="flex items-center gap-2">
                                        <Badge appearance="tint" size="small">{app.totalScopes} scopes</Badge>
                                        {app.hasHighRiskPermissions && (
                                            <Warning24Regular className="w-4 h-4 text-red-500" />
                                        )}
                                    </div>
                                </div>
                            ))}
                        </div>
                    </Card>
                </div>
            )}

            {/* Risky Tab */}
            {activeTab === 'risky' && (
                <Card className="p-5">
                    <h3 className="font-semibold mb-4 flex items-center gap-2">
                        <Warning24Regular className="w-5 h-5 text-red-500" />
                        Risky Application Consents
                    </h3>
                    {riskyConsents?.riskyConsents?.length === 0 ? (
                        <p className="text-gray-500 text-center py-8">No risky consents detected</p>
                    ) : (
                        <div className="overflow-x-auto">
                            <table className="w-full text-sm">
                                <thead>
                                    <tr className="border-b dark:border-gray-700">
                                        <th className="text-left py-2 px-3">Application</th>
                                        <th className="text-left py-2 px-3">Risk Level</th>
                                        <th className="text-left py-2 px-3">Risk Factors</th>
                                        <th className="text-left py-2 px-3">High-Risk Scopes</th>
                                        <th className="text-left py-2 px-3">Consent Type</th>
                                    </tr>
                                </thead>
                                <tbody className="divide-y dark:divide-gray-700">
                                    {riskyConsents?.riskyConsents?.map((consent: any, index: number) => (
                                        <tr key={index} className="hover:bg-gray-50 dark:hover:bg-gray-700/50">
                                            <td className="py-2 px-3">
                                                <div className="flex items-center gap-2">
                                                    {consent.isVerified ? (
                                                        <ShieldCheckmark24Regular className="w-5 h-5 text-green-500" />
                                                    ) : (
                                                        <Warning24Regular className="w-5 h-5 text-amber-500" />
                                                    )}
                                                    <div>
                                                        <p className="font-medium">{consent.appName}</p>
                                                        <p className="text-xs text-gray-500">{consent.publisherName || 'Unknown Publisher'}</p>
                                                    </div>
                                                </div>
                                            </td>
                                            <td className="py-2 px-3">
                                                <Badge appearance="filled" 
                                                    color={consent.riskLevel === 'High' ? 'danger' : consent.riskLevel === 'Medium' ? 'warning' : 'success'}>
                                                    {consent.riskLevel}
                                                </Badge>
                                            </td>
                                            <td className="py-2 px-3">
                                                <div className="flex flex-wrap gap-1">
                                                    {consent.riskFactors?.map((factor: string, i: number) => (
                                                        <Badge key={i} appearance="tint" size="small">{factor}</Badge>
                                                    ))}
                                                </div>
                                            </td>
                                            <td className="py-2 px-3">
                                                <div className="flex flex-wrap gap-1 max-w-xs">
                                                    {consent.riskyScopes?.slice(0, 3).map((scope: string, i: number) => (
                                                        <Badge key={i} appearance="outline" color="danger" size="small">{scope}</Badge>
                                                    ))}
                                                    {consent.riskyScopes?.length > 3 && (
                                                        <Badge appearance="tint" size="small">+{consent.riskyScopes.length - 3}</Badge>
                                                    )}
                                                </div>
                                            </td>
                                            <td className="py-2 px-3">
                                                <Badge appearance="tint" color={consent.consentType === 'AllPrincipals' ? 'warning' : 'brand'} size="small">
                                                    {consent.consentType === 'AllPrincipals' ? 'Admin (All Users)' : 'User'}
                                                </Badge>
                                            </td>
                                        </tr>
                                    ))}
                                </tbody>
                            </table>
                        </div>
                    )}
                </Card>
            )}

            {/* OAuth Grants Tab */}
            {activeTab === 'grants' && (
                <Card className="p-5">
                    <h3 className="font-semibold mb-4">OAuth2 Permission Grants</h3>
                    <div className="overflow-x-auto">
                        <table className="w-full text-sm">
                            <thead>
                                <tr className="border-b dark:border-gray-700">
                                    <th className="text-left py-2 px-3">Application</th>
                                    <th className="text-left py-2 px-3">Consent Type</th>
                                    <th className="text-left py-2 px-3">Scopes</th>
                                </tr>
                            </thead>
                            <tbody className="divide-y dark:divide-gray-700">
                                {oauth2Grants?.grants?.slice(0, 50).map((grant: any, index: number) => (
                                    <tr key={index} className="hover:bg-gray-50 dark:hover:bg-gray-700/50">
                                        <td className="py-2 px-3 font-medium">{grant.clientDisplayName}</td>
                                        <td className="py-2 px-3">
                                            <Badge appearance="tint" color={grant.consentType === 'AllPrincipals' ? 'warning' : 'brand'} size="small">
                                                {grant.consentType === 'AllPrincipals' ? 'Admin' : 'User'}
                                            </Badge>
                                        </td>
                                        <td className="py-2 px-3">
                                            <div className="flex flex-wrap gap-1 max-w-md">
                                                {grant.scopes?.slice(0, 5).map((scope: string, i: number) => (
                                                    <Badge key={i} appearance="outline" size="small">{scope}</Badge>
                                                ))}
                                                {grant.scopes?.length > 5 && <Badge appearance="tint" size="small">+{grant.scopes.length - 5}</Badge>}
                                            </div>
                                        </td>
                                    </tr>
                                ))}
                            </tbody>
                        </table>
                    </div>
                </Card>
            )}

            {/* Enterprise Apps Tab */}
            {activeTab === 'apps' && (
                <Card className="p-5">
                    <h3 className="font-semibold mb-4">Enterprise Applications</h3>
                    <div className="overflow-x-auto">
                        <table className="w-full text-sm">
                            <thead>
                                <tr className="border-b dark:border-gray-700">
                                    <th className="text-left py-2 px-3">Application</th>
                                    <th className="text-left py-2 px-3">Publisher</th>
                                    <th className="text-center py-2 px-3">Verified</th>
                                    <th className="text-center py-2 px-3">App Roles</th>
                                    <th className="text-center py-2 px-3">Delegated</th>
                                    <th className="text-center py-2 px-3">Status</th>
                                </tr>
                            </thead>
                            <tbody className="divide-y dark:divide-gray-700">
                                {enterpriseApps?.apps?.slice(0, 50).map((app: any, index: number) => (
                                    <tr key={index} className="hover:bg-gray-50 dark:hover:bg-gray-700/50">
                                        <td className="py-2 px-3">
                                            <p className="font-medium">{app.displayName}</p>
                                            <p className="text-xs text-gray-500">{app.appId}</p>
                                        </td>
                                        <td className="py-2 px-3 text-gray-600">{app.publisherName || '-'}</td>
                                        <td className="py-2 px-3 text-center">
                                            {app.isVerified ? (
                                                <ShieldCheckmark24Regular className="w-5 h-5 text-green-500 mx-auto" />
                                            ) : (
                                                <span className="text-gray-400">-</span>
                                            )}
                                        </td>
                                        <td className="py-2 px-3 text-center">
                                            <Badge appearance={app.appRoleAssignmentCount > 0 ? 'filled' : 'outline'} color="brand" size="small">
                                                {app.appRoleAssignmentCount || 0}
                                            </Badge>
                                        </td>
                                        <td className="py-2 px-3 text-center">
                                            <Badge appearance={app.oauth2GrantCount > 0 ? 'filled' : 'outline'} color="success" size="small">
                                                {app.oauth2GrantCount || 0}
                                            </Badge>
                                        </td>
                                        <td className="py-2 px-3 text-center">
                                            <Badge appearance="tint" color={app.accountEnabled ? 'success' : 'danger'} size="small">
                                                {app.accountEnabled ? 'Enabled' : 'Disabled'}
                                            </Badge>
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

export default ApplicationConsentPage;
