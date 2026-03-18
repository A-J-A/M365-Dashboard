import React, { useState, useEffect, useCallback } from 'react';
import { 
    Spinner, 
    Text, 
    Card, 
    Badge,
    Button,
    ProgressBar,
    tokens,
} from '@fluentui/react-components';
import { 
    ArrowSync24Regular,
    People24Regular,
    ShieldCheckmark24Regular,
    Laptop24Regular,
    ArrowUp24Regular,
    ArrowDown24Regular,
    ArrowRight24Regular,
    Warning24Regular,
    Checkmark24Regular,
    Clock24Regular,
    Globe24Regular,
    ShieldError24Regular,
    PersonWarning24Regular,
    LockClosed24Regular,
    Eye24Regular,
    PersonAvailable24Regular,
    Shield24Regular,
    Certificate24Regular,
} from '@fluentui/react-icons';
import { ResponsiveContainer, AreaChart, Area, BarChart, Bar, XAxis, YAxis, Tooltip } from 'recharts';
import { useSettings, useTheme, useUser, useAppContext } from '../../contexts/AppContext';
import { useDashboardSummary, useSignInAnalytics, useDeviceCompliance } from '../../hooks/useWidgetData';
import { Link } from 'react-router-dom';

interface UserStats {
    totalUsers: number;
    enabledUsers: number;
    disabledUsers: number;
    memberUsers: number;
    guestUsers: number;
    licensedUsers: number;
    unlicensedUsers: number;
    usersSignedInLast30Days: number;
    usersNeverSignedIn: number;
    lastUpdated: string;
}

interface SecurityStats {
    totalRiskyUsers: number;
    highRiskUsers: number;
    mediumRiskUsers: number;
    lowRiskUsers: number;
    usersAtRisk: number;
    riskySignInsLast24Hours: number;
    mfaRegisteredUsers: number;
    mfaNotRegisteredUsers: number;
    mfaRegistrationPercentage: number;
    lastUpdated: string;
}

interface RiskyUser {
    id: string;
    userPrincipalName: string;
    displayName: string;
    riskLevel: string;
    riskState: string;
    riskDetail: string;
    riskLastUpdatedDateTime: string;
}

interface RiskySignIn {
    id: string;
    userPrincipalName: string;
    displayName: string;
    createdDateTime: string;
    ipAddress: string;
    location: string;
    riskLevel: string;
    riskState: string;
    clientAppUsed: string;
}

interface VpnSignIn {
    id: string;
    userPrincipalName: string;
    displayName: string | null;
    createdDateTime: string | null;
    ipAddress: string | null;
    city: string | null;
    countryOrRegion: string | null;
    isSuccess: boolean;
    clientAppUsed: string | null;
    riskLevel: string | null;
}

interface ApplePushCertificate {
    isConfigured: boolean;
    appleIdentifier?: string;
    expirationDateTime?: string;
    daysUntilExpiry?: number;
    status?: 'Healthy' | 'Warning' | 'Critical' | 'Expired' | 'Unknown';
    error?: string;
    permissionRequired?: boolean;
}

export function Dashboard() {
    const { isLoading: settingsLoading } = useSettings();
    const { data: summary, isLoading: summaryLoading, refresh: refreshSummary } = useDashboardSummary();
    const { data: signIns, isLoading: signInsLoading } = useSignInAnalytics();
    const { data: devices, isLoading: devicesLoading, error: devicesError } = useDeviceCompliance();
    const { profile } = useUser();
    const { resolvedTheme } = useTheme();
    const { getAccessToken } = useAppContext();
    
    const [currentTime, setCurrentTime] = useState(new Date());
    const [recentAlerts, setRecentAlerts] = useState<Alert[]>([]);
    const [userStats, setUserStats] = useState<UserStats | null>(null);
    const [userStatsLoading, setUserStatsLoading] = useState(true);
    const [securityStats, setSecurityStats] = useState<SecurityStats | null>(null);
    const [securityStatsLoading, setSecurityStatsLoading] = useState(true);
    const [securityStatsError, setSecurityStatsError] = useState<string | null>(null);
    const [riskyUsers, setRiskyUsers] = useState<RiskyUser[]>([]);
    const [riskySignIns, setRiskySignIns] = useState<RiskySignIn[]>([]);
    const [applePushCert, setApplePushCert] = useState<ApplePushCertificate | null>(null);
    const [vpnSignIns, setVpnSignIns] = useState<VpnSignIn[]>([]);
    const [vpnSignInsLoading, setVpnSignInsLoading] = useState(true);

    useEffect(() => {
        const timer = setInterval(() => setCurrentTime(new Date()), 60000);
        return () => clearInterval(timer);
    }, []);

    // Fetch user stats
    const fetchUserStats = useCallback(async () => {
        try {
            setUserStatsLoading(true);
            const token = await getAccessToken();
            const response = await fetch('/api/users/stats', {
                headers: { Authorization: `Bearer ${token}` },
            });
            if (response.ok) {
                const data = await response.json();
                setUserStats(data);
            }
        } catch (error) {
            console.error('Failed to fetch user stats:', error);
        } finally {
            setUserStatsLoading(false);
        }
    }, [getAccessToken]);

    // Fetch security stats
    const fetchSecurityStats = useCallback(async () => {
        try {
            setSecurityStatsLoading(true);
            setSecurityStatsError(null);
            const token = await getAccessToken();
            const response = await fetch('/api/security/stats', {
                headers: { Authorization: `Bearer ${token}` },
            });
            if (response.ok) {
                const data = await response.json();
                setSecurityStats(data);
            } else {
                const err = await response.json().catch(() => ({ error: `HTTP ${response.status}` }));
                setSecurityStatsError(err.error || `HTTP ${response.status}`);
                console.error('Security stats error:', err);
            }
        } catch (error) {
            setSecurityStatsError(error instanceof Error ? error.message : 'Unknown error');
            console.error('Failed to fetch security stats:', error);
        } finally {
            setSecurityStatsLoading(false);
        }
    }, [getAccessToken]);

    // Fetch risky users
    const fetchRiskyUsers = useCallback(async () => {
        try {
            const token = await getAccessToken();
            const response = await fetch('/api/security/risky-users', {
                headers: { Authorization: `Bearer ${token}` },
            });
            if (response.ok) {
                const data = await response.json();
                setRiskyUsers(data.slice(0, 5));
            }
        } catch (error) {
            console.error('Failed to fetch risky users:', error);
        }
    }, [getAccessToken]);

    // Fetch risky sign-ins
    const fetchRiskySignIns = useCallback(async () => {
        try {
            const token = await getAccessToken();
            const response = await fetch('/api/security/risky-signins?hours=24', {
                headers: { Authorization: `Bearer ${token}` },
            });
            if (response.ok) {
                const data = await response.json();
                setRiskySignIns(data.slice(0, 5));
            }
        } catch (error) {
            console.error('Failed to fetch risky sign-ins:', error);
        }
    }, [getAccessToken]);

    // Fetch VPN/proxy sign-ins
    const fetchVpnSignIns = useCallback(async () => {
        try {
            setVpnSignInsLoading(true);
            const token = await getAccessToken();
            const response = await fetch('/api/signins/vpn-proxy?hours=24&take=10', {
                headers: { Authorization: `Bearer ${token}` },
            });
            if (response.ok) {
                const data = await response.json();
                setVpnSignIns(data.signIns ?? []);
            }
        } catch (error) {
            console.error('Failed to fetch VPN/proxy sign-ins:', error);
        } finally {
            setVpnSignInsLoading(false);
        }
    }, [getAccessToken]);

    // Fetch Apple Push certificate status
    const fetchApplePushCert = useCallback(async () => {
        try {
            const token = await getAccessToken();
            const response = await fetch('/api/devices/apple-push-certificate', {
                headers: { Authorization: `Bearer ${token}` },
            });
            if (response.ok) {
                const data = await response.json();
                setApplePushCert(data);
            }
        } catch (error) {
            console.error('Failed to fetch Apple Push certificate:', error);
        }
    }, [getAccessToken]);

    useEffect(() => {
        fetchUserStats();
        fetchSecurityStats();
        fetchRiskyUsers();
        fetchRiskySignIns();
        fetchApplePushCert();
        fetchVpnSignIns();
    }, [fetchUserStats, fetchSecurityStats, fetchRiskyUsers, fetchRiskySignIns, fetchApplePushCert, fetchVpnSignIns]);

    // Generate alerts based on data
    useEffect(() => {
        const alerts: Alert[] = [];
        
        // Apple Push Certificate expiry warning
        if (applePushCert?.isConfigured && applePushCert.daysUntilExpiry !== undefined) {
            if (applePushCert.status === 'Expired') {
                alerts.push({
                    type: 'danger',
                    title: 'Apple Push Certificate Expired',
                    message: `Certificate expired ${Math.abs(applePushCert.daysUntilExpiry)} days ago. iOS/macOS devices cannot be managed.`,
                    link: '/devices',
                    linkParams: '?tab=certificates',
                    icon: 'certificate'
                });
            } else if (applePushCert.status === 'Critical') {
                alerts.push({
                    type: 'danger',
                    title: 'Apple Push Certificate Expiring Soon',
                    message: `Certificate expires in ${applePushCert.daysUntilExpiry} days. Renew immediately to avoid service disruption.`,
                    link: '/devices',
                    linkParams: '?tab=certificates',
                    icon: 'certificate'
                });
            } else if (applePushCert.status === 'Warning') {
                alerts.push({
                    type: 'warning',
                    title: 'Apple Push Certificate Expiring',
                    message: `Certificate expires in ${applePushCert.daysUntilExpiry} days. Plan renewal soon.`,
                    link: '/devices',
                    linkParams: '?tab=certificates',
                    icon: 'certificate'
                });
            }
        }
        
        if (securityStats && securityStats.highRiskUsers > 0) {
            alerts.push({
                type: 'danger',
                title: 'High Risk Users Detected',
                message: `${securityStats.highRiskUsers} user(s) identified as high risk`,
                link: '/security'
            });
        }
        
        if (securityStats && securityStats.riskySignInsLast24Hours > 0) {
            alerts.push({
                type: 'danger',
                title: 'Risky Sign-ins in Last 24 Hours',
                message: `${securityStats.riskySignInsLast24Hours} risky sign-in(s) detected`,
                link: '/signins'
            });
        }

        if (signIns && signIns.riskySignIns > 0) {
            alerts.push({
                type: 'warning',
                title: 'Risky Sign-ins (7 days)',
                message: `${signIns.riskySignIns} risky sign-ins in the last 7 days`,
                link: '/security'
            });
        }

        setRecentAlerts(alerts);
    }, [securityStats, devices, signIns, applePushCert]);

    const getGreeting = () => {
        const hour = currentTime.getHours();
        if (hour < 12) return 'Good morning';
        if (hour < 18) return 'Good afternoon';
        return 'Good evening';
    };

    const refreshAll = () => {
        refreshSummary();
        fetchUserStats();
        fetchSecurityStats();
        fetchRiskyUsers();
        fetchRiskySignIns();
        fetchVpnSignIns();
    };

    if (settingsLoading) {
        return (
            <div className="flex items-center justify-center min-h-[400px]">
                <Spinner size="large" label="Loading dashboard..." />
            </div>
        );
    }

    return (
        <div className="space-y-6">
            {/* Welcome Header */}
            <div className="flex items-start justify-between">
                <div>
                    <h1 className="text-2xl font-bold text-gray-900 dark:text-white">
                        {getGreeting()}, {profile?.displayName?.split(' ')[0] || 'Analyst'}
                    </h1>
                </div>
                <div className="flex items-center gap-3">
                    <span className="text-sm text-gray-500 dark:text-gray-400">
                        {currentTime.toLocaleDateString('en-GB', { 
                            weekday: 'long', 
                            day: 'numeric', 
                            month: 'long',
                            year: 'numeric'
                        })}
                    </span>
                    <Button
                        appearance="subtle"
                        icon={<ArrowSync24Regular className={summaryLoading ? 'animate-spin' : ''} />}
                        onClick={refreshAll}
                    >
                        Refresh
                    </Button>
                </div>
            </div>

            {/* Security Alerts Section */}
            {recentAlerts.length > 0 && (
                <div className="space-y-2">
                    {recentAlerts.map((alert, index) => (
                        <AlertBanner key={index} alert={alert} />
                    ))}
                </div>
            )}

            {/* Key Security Metrics Grid */}
            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
                <MetricCard
                    title="Risky Users"
                    value={securityStats?.totalRiskyUsers ?? 0}
                    subtitle="Users at risk"
                    icon={PersonWarning24Regular}
                    color={securityStats && securityStats.totalRiskyUsers > 0 ? 'red' : 'green'}
                    isLoading={securityStatsLoading}
                    extraInfo={securityStats ? `${securityStats.highRiskUsers} high, ${securityStats.mediumRiskUsers} medium` : undefined}
                    link="/security"
                />
                <MetricCard
                    title="MFA Coverage"
                    value={securityStats?.mfaRegistrationPercentage ?? 0}
                    suffix="%"
                    subtitle={securityStats ? `${securityStats.mfaRegisteredUsers.toLocaleString()} of ${(securityStats.mfaRegisteredUsers + securityStats.mfaNotRegisteredUsers).toLocaleString()} users` : (securityStatsError ? `Error: ${securityStatsError}` : 'Users with MFA')}
                    icon={LockClosed24Regular}
                    color={securityStats && securityStats.mfaRegistrationPercentage < 90 ? 'orange' : 'green'}
                    isLoading={securityStatsLoading}
                    showProgressBar
                    link="/security/mfa"
                />
                <MetricCard
                    title="Sign-in Success"
                    value={signIns?.successRate ?? summary?.signInSuccessRate ?? 0}
                    suffix="%"
                    subtitle="Last 7 days"
                    icon={ShieldCheckmark24Regular}
                    color="green"
                    isLoading={signInsLoading}
                    extraInfo={signIns ? `${signIns.failedSignIns.toLocaleString()} failed, ${signIns.riskySignIns} risky` : undefined}
                    link="/signins"
                />
                <MetricCard
                    title="Device Compliance"
                    value={devices?.complianceRate ?? summary?.deviceComplianceRate ?? 0}
                    suffix="%"
                    subtitle={devices ? `${devices.compliantDevices.toLocaleString()} of ${devices.totalDevices.toLocaleString()} devices` : (devicesError ? `Error: ${devicesError}` : 'Compliant devices')}
                    icon={Laptop24Regular}
                    color={devices && devices.complianceRate < 80 ? 'orange' : 'green'}
                    isLoading={devicesLoading}
                    showProgressBar
                    link="/devices"
                />
            </div>

            {/* Secondary Row - Users and Sign-in Activity */}
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                {/* User Overview */}
                <Card className="p-5">
                    <div className="flex items-center justify-between mb-4">
                        <div className="flex items-center gap-2">
                            <div className="p-2 bg-blue-50 dark:bg-blue-900/30 rounded-lg">
                                <People24Regular className="w-5 h-5 text-blue-600 dark:text-blue-400" />
                            </div>
                            <Text weight="semibold">User Overview</Text>
                        </div>
                        <Link to="/users" className="text-sm text-blue-600 dark:text-blue-400 hover:underline flex items-center gap-1">
                            View All <ArrowRight24Regular className="w-4 h-4" />
                        </Link>
                    </div>
                    {userStatsLoading ? (
                        <div className="flex justify-center py-4">
                            <Spinner size="small" />
                        </div>
                    ) : userStats ? (
                        <div className="grid grid-cols-2 gap-4">
                            <div className="p-3 bg-gray-50 dark:bg-gray-800/50 rounded-lg">
                                <p className="text-2xl font-bold text-gray-900 dark:text-white">{userStats.enabledUsers.toLocaleString()}</p>
                                <p className="text-xs text-gray-500 dark:text-gray-400">Enabled Users</p>
                            </div>
                            <div className="p-3 bg-gray-50 dark:bg-gray-800/50 rounded-lg">
                                <p className="text-2xl font-bold text-gray-900 dark:text-white">{userStats.disabledUsers.toLocaleString()}</p>
                                <p className="text-xs text-gray-500 dark:text-gray-400">Disabled Users</p>
                            </div>
                            <div className="p-3 bg-gray-50 dark:bg-gray-800/50 rounded-lg">
                                <p className="text-2xl font-bold text-gray-900 dark:text-white">{userStats.guestUsers.toLocaleString()}</p>
                                <p className="text-xs text-gray-500 dark:text-gray-400">Guest Users</p>
                            </div>
                            <div className="p-3 bg-gray-50 dark:bg-gray-800/50 rounded-lg">
                                <p className="text-2xl font-bold text-amber-600 dark:text-amber-400">{userStats.usersNeverSignedIn.toLocaleString()}</p>
                                <p className="text-xs text-gray-500 dark:text-gray-400">Never Signed In</p>
                            </div>
                        </div>
                    ) : null}
                </Card>

                {/* Sign-in Activity Card */}
                <Card className="p-5">
                    <div className="flex items-center justify-between mb-4">
                        <div className="flex items-center gap-2">
                            <div className="p-2 bg-green-50 dark:bg-green-900/30 rounded-lg">
                                <Globe24Regular className="w-5 h-5 text-green-600 dark:text-green-400" />
                            </div>
                            <Text weight="semibold">Sign-in Activity (7 days)</Text>
                        </div>
                        <Link to="/signins" className="text-sm text-blue-600 dark:text-blue-400 hover:underline flex items-center gap-1">
                            View All <ArrowRight24Regular className="w-4 h-4" />
                        </Link>
                    </div>
                    {signInsLoading ? (
                        <div className="flex justify-center py-4">
                            <Spinner size="small" />
                        </div>
                    ) : signIns ? (
                        <div className="space-y-3">
                            <div className="flex justify-between items-center">
                                <span className="text-sm text-gray-500 dark:text-gray-400">Total Sign-ins</span>
                                <span className="font-semibold text-gray-900 dark:text-white">{signIns.totalSignIns.toLocaleString()}</span>
                            </div>
                            <div className="flex justify-between items-center">
                                <span className="text-sm text-gray-500 dark:text-gray-400 flex items-center gap-1">
                                    <Checkmark24Regular className="w-4 h-4 text-green-500" /> Successful
                                </span>
                                <span className="font-semibold text-green-600 dark:text-green-400">{signIns.successfulSignIns.toLocaleString()}</span>
                            </div>
                            <Link to="/signins?status=failure" className="flex justify-between items-center hover:bg-gray-100 dark:hover:bg-gray-700 -mx-2 px-2 py-1 rounded transition-colors">
                                <span className="text-sm text-gray-500 dark:text-gray-400 flex items-center gap-1">
                                    <ShieldError24Regular className="w-4 h-4 text-red-500" /> Failed
                                </span>
                                <span className="font-semibold text-red-600 dark:text-red-400 flex items-center gap-1">
                                    {signIns.failedSignIns.toLocaleString()}
                                    <ArrowRight24Regular className="w-4 h-4" />
                                </span>
                            </Link>
                            {signIns.riskySignIns > 0 && (
                                <div className="flex justify-between items-center pt-2 border-t border-gray-200 dark:border-gray-700">
                                    <span className="text-sm text-amber-600 dark:text-amber-400 flex items-center gap-1">
                                        <Warning24Regular className="w-4 h-4" /> Risky
                                    </span>
                                    <span className="font-semibold text-amber-600 dark:text-amber-400">{signIns.riskySignIns.toLocaleString()}</span>
                                </div>
                            )}
                        </div>
                    ) : null}
                </Card>
            </div>

            {/* Risky Users, Sign-ins and VPN */}
            <div className="grid grid-cols-1 lg:grid-cols-3 gap-4">
                {/* Risky Users */}
                <Card className="p-5">
                    <div className="flex items-center justify-between mb-4">
                        <div className="flex items-center gap-2">
                            <div className="p-2 bg-red-50 dark:bg-red-900/30 rounded-lg">
                                <PersonWarning24Regular className="w-5 h-5 text-red-600 dark:text-red-400" />
                            </div>
                            <Text weight="semibold">Risky Users</Text>
                        </div>
                        <Link to="/security" className="text-sm text-blue-600 dark:text-blue-400 hover:underline flex items-center gap-1">
                            View All <ArrowRight24Regular className="w-4 h-4" />
                        </Link>
                    </div>
                    {riskyUsers.length > 0 ? (
                        <div className="space-y-2">
                            {riskyUsers.map((riskyUser) => (
                                <div key={riskyUser.id} className="flex items-center justify-between p-3 bg-gray-50 dark:bg-gray-800/50 rounded-lg">
                                    <div className="min-w-0 flex-1">
                                        <p className="text-sm font-medium text-gray-900 dark:text-white truncate">
                                            {riskyUser.displayName || riskyUser.userPrincipalName}
                                        </p>
                                        <p className="text-xs text-gray-500 dark:text-gray-400 truncate">
                                            {riskyUser.userPrincipalName}
                                        </p>
                                    </div>
                                    <Badge 
                                        appearance="tint" 
                                        color={riskyUser.riskLevel === 'high' ? 'danger' : riskyUser.riskLevel === 'medium' ? 'warning' : 'important'}
                                        size="small"
                                    >
                                        {riskyUser.riskLevel}
                                    </Badge>
                                </div>
                            ))}
                        </div>
                    ) : (
                        <div className="flex flex-col items-center justify-center py-8 text-gray-500 dark:text-gray-400">
                            <Checkmark24Regular className="w-8 h-8 text-green-500 mb-2" />
                            <p className="text-sm">No risky users detected</p>
                        </div>
                    )}
                </Card>

                {/* Risky Sign-ins (Last 24 hours) */}
                <Card className="p-5">
                    <div className="flex items-center justify-between mb-4">
                        <div className="flex items-center gap-2">
                            <div className="p-2 bg-amber-50 dark:bg-amber-900/30 rounded-lg">
                                <Warning24Regular className="w-5 h-5 text-amber-600 dark:text-amber-400" />
                            </div>
                            <Text weight="semibold">Risky Sign-ins (24h)</Text>
                        </div>
                        <Link to="/signins" className="text-sm text-blue-600 dark:text-blue-400 hover:underline flex items-center gap-1">
                            View All <ArrowRight24Regular className="w-4 h-4" />
                        </Link>
                    </div>
                    {riskySignIns.length > 0 ? (
                        <div className="space-y-2">
                            {riskySignIns.map((signIn) => (
                                <div key={signIn.id} className="flex items-center justify-between p-3 bg-gray-50 dark:bg-gray-800/50 rounded-lg">
                                    <div className="min-w-0 flex-1">
                                        <p className="text-sm font-medium text-gray-900 dark:text-white truncate">
                                            {signIn.displayName || signIn.userPrincipalName}
                                        </p>
                                        <p className="text-xs text-gray-500 dark:text-gray-400">
                                            {signIn.location || signIn.ipAddress} • {signIn.createdDateTime ? new Date(signIn.createdDateTime).toLocaleTimeString() : 'Unknown time'}
                                        </p>
                                    </div>
                                    <Badge 
                                        appearance="tint" 
                                        color={signIn.riskLevel === 'high' ? 'danger' : signIn.riskLevel === 'medium' ? 'warning' : 'important'}
                                        size="small"
                                    >
                                        {signIn.riskLevel}
                                    </Badge>
                                </div>
                            ))}
                        </div>
                    ) : (
                        <div className="flex flex-col items-center justify-center py-8 text-gray-500 dark:text-gray-400">
                            <Checkmark24Regular className="w-8 h-8 text-green-500 mb-2" />
                            <p className="text-sm">No risky sign-ins in the last 24 hours</p>
                        </div>
                    )}
                </Card>

                {/* VPN / Proxy Sign-ins */}
                <Card className="p-5">
                    <div className="flex items-center justify-between mb-4">
                        <div className="flex items-center gap-2">
                            <div className="p-2 bg-purple-50 dark:bg-purple-900/30 rounded-lg">
                                <Eye24Regular className="w-5 h-5 text-purple-600 dark:text-purple-400" />
                            </div>
                            <Text weight="semibold">VPN / Proxy Sign-ins (24h)</Text>
                        </div>
                        <Link to="/signins" className="text-sm text-blue-600 dark:text-blue-400 hover:underline flex items-center gap-1">
                            View All <ArrowRight24Regular className="w-4 h-4" />
                        </Link>
                    </div>
                    {vpnSignInsLoading ? (
                        <div className="flex justify-center py-4">
                            <Spinner size="small" />
                        </div>
                    ) : vpnSignIns.length > 0 ? (
                        <div className="space-y-2">
                            {vpnSignIns.map((signIn) => (
                                <div key={signIn.id} className="flex items-center justify-between p-3 bg-purple-50 dark:bg-purple-900/20 rounded-lg border border-purple-200 dark:border-purple-800">
                                    <div className="min-w-0 flex-1">
                                        <p className="text-sm font-medium text-gray-900 dark:text-white truncate">
                                            {signIn.displayName || signIn.userPrincipalName}
                                        </p>
                                        <p className="text-xs text-gray-500 dark:text-gray-400 truncate">
                                            {[signIn.city, signIn.countryOrRegion].filter(Boolean).join(', ') || signIn.ipAddress || 'Unknown location'}
                                            {signIn.createdDateTime ? ` • ${new Date(signIn.createdDateTime).toLocaleTimeString()}` : ''}
                                        </p>
                                    </div>
                                    <Badge
                                        appearance="tint"
                                        color={signIn.isSuccess ? 'warning' : 'danger'}
                                        size="small"
                                    >
                                        {signIn.isSuccess ? 'Success' : 'Failed'}
                                    </Badge>
                                </div>
                            ))}
                        </div>
                    ) : (
                        <div className="flex flex-col items-center justify-center py-8 text-gray-500 dark:text-gray-400">
                            <Checkmark24Regular className="w-8 h-8 text-green-500 mb-2" />
                            <p className="text-sm">No VPN/proxy sign-ins detected</p>
                            <p className="text-xs text-gray-400 dark:text-gray-500 mt-1">Requires Entra ID P2</p>
                        </div>
                    )}
                </Card>
            </div>

            {/* Device Compliance and Quick Actions */}
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                {/* Device Status Card */}
                <Card className="p-5">
                    <div className="flex items-center justify-between mb-4">
                        <div className="flex items-center gap-2">
                            <div className="p-2 bg-purple-50 dark:bg-purple-900/30 rounded-lg">
                                <Laptop24Regular className="w-5 h-5 text-purple-600 dark:text-purple-400" />
                            </div>
                            <Text weight="semibold">Device Compliance</Text>
                        </div>
                        <Link to="/devices" className="text-sm text-blue-600 dark:text-blue-400 hover:underline flex items-center gap-1">
                            View All <ArrowRight24Regular className="w-4 h-4" />
                        </Link>
                    </div>
                    {devicesLoading ? (
                        <div className="flex justify-center py-4">
                            <Spinner size="small" />
                        </div>
                    ) : devices ? (
                        <div className="space-y-3">
                            <div className="flex justify-between items-center">
                                <span className="text-sm text-gray-500 dark:text-gray-400">Total Devices</span>
                                <span className="font-semibold text-gray-900 dark:text-white">{devices.totalDevices.toLocaleString()}</span>
                            </div>
                            <div className="flex justify-between items-center">
                                <span className="text-sm text-gray-500 dark:text-gray-400 flex items-center gap-1">
                                    <Checkmark24Regular className="w-4 h-4 text-green-500" /> Compliant
                                </span>
                                <span className="font-semibold text-green-600 dark:text-green-400">{devices.compliantDevices.toLocaleString()}</span>
                            </div>
                            <div className="flex justify-between items-center">
                                <span className="text-sm text-gray-500 dark:text-gray-400 flex items-center gap-1">
                                    <Warning24Regular className="w-4 h-4 text-red-500" /> Non-Compliant
                                </span>
                                <span className="font-semibold text-red-600 dark:text-red-400">{devices.nonCompliantDevices.toLocaleString()}</span>
                            </div>
                            {devices.byPlatform && devices.byPlatform.length > 0 && (
                                <div className="pt-2 border-t border-gray-200 dark:border-gray-700">
                                    <p className="text-xs text-gray-500 dark:text-gray-400 mb-2">By Platform</p>
                                    <div className="flex flex-wrap gap-2">
                                        {devices.byPlatform.slice(0, 4).map((platform, idx) => (
                                            <Badge key={idx} appearance="outline" size="small">
                                                {platform.platform}: {platform.total}
                                            </Badge>
                                        ))}
                                    </div>
                                </div>
                            )}
                        </div>
                    ) : null}
                </Card>

                {/* Quick Actions Card */}
                <Card className="p-5">
                    <div className="flex items-center gap-2 mb-4">
                        <div className="p-2 bg-blue-50 dark:bg-blue-900/30 rounded-lg">
                            <Shield24Regular className="w-5 h-5 text-blue-600 dark:text-blue-400" />
                        </div>
                        <Text weight="semibold">Security Quick Actions</Text>
                    </div>
                    <div className="space-y-2">
                        <QuickLink 
                            icon={PersonWarning24Regular} 
                            label="Review Risky Users" 
                            to="/security" 
                        />
                        <QuickLink 
                            icon={Eye24Regular} 
                            label="View Sign-in Logs" 
                            to="/signins" 
                        />
                        <QuickLink 
                            icon={LockClosed24Regular} 
                            label="MFA Registration Status" 
                            to="/security/mfa" 
                        />
                        <QuickLink 
                            icon={Laptop24Regular} 
                            label="Non-Compliant Devices" 
                            to="/devices" 
                        />
                    </div>
                </Card>
            </div>

            {/* Last Updated */}
            <div className="flex items-center justify-center gap-2 text-xs text-gray-400 dark:text-gray-500">
                <Clock24Regular className="w-4 h-4" />
                Last updated: {summary?.lastUpdated ? new Date(summary.lastUpdated).toLocaleString() : 'Unknown'}
            </div>
        </div>
    );
}

// Helper Components
interface Alert {
    type: 'info' | 'warning' | 'danger';
    title: string;
    message: string;
    link?: string;
    linkParams?: string;
    icon?: string;
}

function AlertBanner({ alert }: { alert: Alert }) {
    const bgColor = alert.type === 'danger' 
        ? 'bg-red-50 dark:bg-red-900/20 border-red-200 dark:border-red-800'
        : alert.type === 'warning'
        ? 'bg-amber-50 dark:bg-amber-900/20 border-amber-200 dark:border-amber-800'
        : 'bg-blue-50 dark:bg-blue-900/20 border-blue-200 dark:border-blue-800';
    
    const textColor = alert.type === 'danger'
        ? 'text-red-800 dark:text-red-200'
        : alert.type === 'warning'
        ? 'text-amber-800 dark:text-amber-200'
        : 'text-blue-800 dark:text-blue-200';

    return (
        <div className={`p-3 rounded-lg border ${bgColor} flex items-center justify-between`}>
            <div className="flex items-center gap-3">
                <Warning24Regular className={`w-5 h-5 ${textColor}`} />
                <div>
                    <p className={`font-medium ${textColor}`}>{alert.title}</p>
                    <p className={`text-sm ${textColor} opacity-80`}>{alert.message}</p>
                </div>
            </div>
            {alert.link && (
                <Link 
                    to={alert.link} 
                    className={`text-sm font-medium ${textColor} hover:underline flex items-center gap-1`}
                >
                    View <ArrowRight24Regular className="w-4 h-4" />
                </Link>
            )}
        </div>
    );
}

interface MetricCardProps {
    title: string;
    value: number;
    suffix?: string;
    subtitle: string;
    icon: React.ComponentType<{ className?: string }>;
    color: 'blue' | 'green' | 'purple' | 'orange' | 'red';
    isLoading: boolean;
    trend?: { value: number; isPositive: boolean };
    sparklineData?: number[];
    extraInfo?: string;
    showProgressBar?: boolean;
    link?: string;
}

const colorClasses = {
    blue: { bg: 'bg-blue-50 dark:bg-blue-900/30', icon: 'text-blue-600 dark:text-blue-400', bar: '#3b82f6' },
    green: { bg: 'bg-green-50 dark:bg-green-900/30', icon: 'text-green-600 dark:text-green-400', bar: '#22c55e' },
    purple: { bg: 'bg-purple-50 dark:bg-purple-900/30', icon: 'text-purple-600 dark:text-purple-400', bar: '#a855f7' },
    orange: { bg: 'bg-orange-50 dark:bg-orange-900/30', icon: 'text-orange-600 dark:text-orange-400', bar: '#f97316' },
    red: { bg: 'bg-red-50 dark:bg-red-900/30', icon: 'text-red-600 dark:text-red-400', bar: '#ef4444' },
};

function MetricCard({ title, value, suffix, subtitle, icon: Icon, color, isLoading, trend, sparklineData, extraInfo, showProgressBar, link }: MetricCardProps) {
    const colors = colorClasses[color];
    
    const content = (
        <Card className="p-5 h-full transition-shadow hover:shadow-md cursor-pointer">
            <div className="flex items-start justify-between">
                <div className="flex-1 min-w-0">
                    <p className="text-sm font-medium text-gray-500 dark:text-gray-400 truncate">
                        {title}
                    </p>
                    {isLoading ? (
                        <div className="mt-2 h-8 w-24 bg-gray-200 dark:bg-gray-700 rounded animate-pulse" />
                    ) : (
                        <div className="flex items-baseline gap-2 mt-1">
                            <p className="text-2xl font-bold text-gray-900 dark:text-white">
                                {typeof value === 'number' ? value.toLocaleString(undefined, { maximumFractionDigits: 1 }) : value}
                                {suffix && <span className="text-lg">{suffix}</span>}
                            </p>
                            {trend && (
                                <span className={`flex items-center text-xs font-medium ${trend.isPositive ? 'text-green-600' : 'text-red-600'}`}>
                                    {trend.isPositive ? <ArrowUp24Regular className="w-4 h-4" /> : <ArrowDown24Regular className="w-4 h-4" />}
                                    {Math.abs(trend.value).toFixed(1)}%
                                </span>
                            )}
                        </div>
                    )}
                    <p className="mt-1 text-xs text-gray-400 dark:text-gray-500">
                        {subtitle}
                    </p>
                    {extraInfo && (
                        <p className="mt-1 text-xs text-gray-500 dark:text-gray-400">
                            {extraInfo}
                        </p>
                    )}
                </div>
                <div className="flex flex-col items-end gap-2">
                    <div className={`p-2.5 rounded-lg ${colors.bg}`}>
                        <Icon className={`w-5 h-5 ${colors.icon}`} />
                    </div>
                    {sparklineData && sparklineData.length > 0 && !isLoading && (
                        <div className="w-16 h-8">
                            <ResponsiveContainer width="100%" height="100%">
                                <AreaChart data={sparklineData.map((v, i) => ({ v }))}>
                                    <Area 
                                        type="monotone" 
                                        dataKey="v" 
                                        stroke={colors.bar}
                                        fill={colors.bar}
                                        fillOpacity={0.2}
                                        strokeWidth={1.5}
                                    />
                                </AreaChart>
                            </ResponsiveContainer>
                        </div>
                    )}
                </div>
            </div>
            {showProgressBar && !isLoading && (
                <div className="mt-3">
                    <ProgressBar 
                        value={Math.min(value / 100, 1)} 
                        color={value > 90 ? 'success' : value > 75 ? 'warning' : 'error'}
                        thickness="medium"
                    />
                </div>
            )}
        </Card>
    );

    return link ? <Link to={link} className="block">{content}</Link> : content;
}

function QuickLink({ icon: Icon, label, to }: { icon: React.ComponentType<{ className?: string }>; label: string; to: string }) {
    return (
        <Link 
            to={to}
            className="flex items-center gap-3 p-2 rounded-lg hover:bg-gray-100 dark:hover:bg-gray-700 transition-colors"
        >
            <Icon className="w-5 h-5 text-gray-500 dark:text-gray-400" />
            <span className="text-sm text-gray-700 dark:text-gray-300">{label}</span>
            <ArrowRight24Regular className="w-4 h-4 text-gray-400 ml-auto" />
        </Link>
    );
}
