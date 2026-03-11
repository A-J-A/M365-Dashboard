import React, { useState, useEffect, useCallback } from 'react';
import {
    makeStyles,
    tokens,
    Card,
    CardHeader,
    Text,
    Title3,
    Badge,
    Spinner,
    Button,
    Input,
    Table,
    TableHeader,
    TableHeaderCell,
    TableBody,
    TableRow,
    TableCell,
    ProgressBar,
    Tab,
    TabList,
    Switch,
} from '@fluentui/react-components';
import {
    Key24Regular,
    Certificate24Regular,
    People24Regular,
    CheckmarkCircle24Regular,
    Warning24Regular,
    ArrowClockwise24Regular,
    Search24Regular,
    ChevronUp24Regular,
    ChevronDown24Regular,
    ArrowDownload24Regular,
    Info24Regular,
} from '@fluentui/react-icons';
import { useAppContext } from '../../contexts/AppContext';

const useStyles = makeStyles({
    container: {
        padding: tokens.spacingHorizontalL,
        display: 'flex',
        flexDirection: 'column',
        gap: tokens.spacingVerticalL,
        maxWidth: '1600px',
        margin: '0 auto',
    },
    header: {
        display: 'flex',
        justifyContent: 'space-between',
        alignItems: 'center',
        flexWrap: 'wrap',
        gap: tokens.spacingHorizontalM,
    },
    headerTitle: {
        display: 'flex',
        alignItems: 'center',
        gap: tokens.spacingHorizontalS,
    },
    statsGrid: {
        display: 'grid',
        gridTemplateColumns: 'repeat(auto-fit, minmax(180px, 1fr))',
        gap: tokens.spacingHorizontalM,
    },
    statCard: {
        padding: tokens.spacingVerticalM,
        display: 'flex',
        flexDirection: 'column',
        gap: tokens.spacingVerticalXS,
    },
    statValue: {
        fontSize: tokens.fontSizeHero800,
        fontWeight: tokens.fontWeightSemibold,
        lineHeight: '1',
    },
    statLabel: {
        color: tokens.colorNeutralForeground3,
        fontSize: tokens.fontSizeBase200,
    },
    utilizationCard: {
        padding: tokens.spacingVerticalL,
    },
    utilizationHeader: {
        display: 'flex',
        justifyContent: 'space-between',
        alignItems: 'center',
        marginBottom: tokens.spacingVerticalM,
    },
    utilizationInfo: {
        display: 'flex',
        gap: tokens.spacingHorizontalXL,
        marginBottom: tokens.spacingVerticalM,
    },
    utilizationItem: {
        display: 'flex',
        flexDirection: 'column',
        gap: tokens.spacingVerticalXXS,
    },
    contentGrid: {
        display: 'grid',
        gridTemplateColumns: 'repeat(auto-fit, minmax(400px, 1fr))',
        gap: tokens.spacingHorizontalL,
    },
    listCard: {
        minHeight: '280px',
    },
    filterBar: {
        display: 'flex',
        gap: tokens.spacingHorizontalM,
        alignItems: 'center',
        flexWrap: 'wrap',
        marginBottom: tokens.spacingVerticalM,
    },
    searchInput: {
        minWidth: '250px',
    },
    tableContainer: {
        overflowX: 'auto',
    },
    sortableHeader: {
        cursor: 'pointer',
        display: 'flex',
        alignItems: 'center',
        gap: tokens.spacingHorizontalXS,
        ':hover': {
            color: tokens.colorBrandForeground1,
        },
    },
    loadingContainer: {
        display: 'flex',
        justifyContent: 'center',
        alignItems: 'center',
        padding: tokens.spacingVerticalXXL,
    },
    emptyState: {
        display: 'flex',
        flexDirection: 'column',
        alignItems: 'center',
        justifyContent: 'center',
        padding: tokens.spacingVerticalXXL,
        color: tokens.colorNeutralForeground3,
    },
    progressCell: {
        display: 'flex',
        alignItems: 'center',
        gap: tokens.spacingHorizontalS,
        minWidth: '150px',
    },
    progressBar: {
        flex: 1,
        minWidth: '80px',
    },
});

interface License {
    skuId: string;
    skuPartNumber: string;
    displayName: string;
    totalUnits: number;
    consumedUnits: number;
    availableUnits: number;
    warningUnits: number;
    suspendedUnits: number;
    utilizationPercentage: number;
    status: string;
    appliesTo: string;
    isTrial: boolean;
    servicePlanCount: number;
}

interface LicenseStats {
    totalSubscriptions: number;
    totalLicenses: number;
    assignedLicenses: number;
    availableLicenses: number;
    utilizationPercentage: number;
    subscriptionsWithWarnings: number;
    trialSubscriptions: number;
}

interface LicenseOverview {
    stats: LicenseStats;
    topUtilized: License[];
    lowUtilization: License[];
    recentlyAdded: License[];
    lastUpdated: string;
}

const getUtilizationColor = (percentage: number): 'success' | 'warning' | 'danger' => {
    if (percentage >= 95) return 'danger';
    if (percentage >= 80) return 'warning';
    return 'success';
};

const getUtilizationBarColor = (percentage: number): 'success' | 'warning' | 'error' | 'brand' => {
    if (percentage >= 95) return 'error';
    if (percentage >= 80) return 'warning';
    if (percentage >= 50) return 'brand';
    return 'success';
};

export const LicensesPage: React.FC = () => {
    const styles = useStyles();
    const { getAccessToken } = useAppContext();
    
    const [overview, setOverview] = useState<LicenseOverview | null>(null);
    const [licenses, setLicenses] = useState<License[]>([]);
    const [loading, setLoading] = useState(true);
    const [licensesLoading, setLicensesLoading] = useState(false);
    const [error, setError] = useState<string | null>(null);
    
    const [selectedTab, setSelectedTab] = useState<string>('overview');
    const [searchQuery, setSearchQuery] = useState('');
    const [sortBy, setSortBy] = useState<string>('displayName');
    const [sortAscending, setSortAscending] = useState(true);
    const [filterType, setFilterType] = useState<string>('excludeFreeTrial');
    const [excludeFreeTrialOverview, setExcludeFreeTrialOverview] = useState(true);

    const fetchOverview = useCallback(async (excludeFreeTrial: boolean) => {
        try {
            setLoading(true);
            setError(null);
            
            const token = await getAccessToken();
            const response = await fetch(`/api/licenses/overview?excludeFreeTrial=${excludeFreeTrial}`, {
                headers: { Authorization: `Bearer ${token}` },
            });

            if (!response.ok) {
                throw new Error('Failed to fetch license overview');
            }

            const data = await response.json();
            setOverview(data);
        } catch (err) {
            setError(err instanceof Error ? err.message : 'An error occurred');
        } finally {
            setLoading(false);
        }
    }, [getAccessToken]);

    const fetchLicenses = useCallback(async () => {
        try {
            setLicensesLoading(true);
            
            const token = await getAccessToken();
            const response = await fetch('/api/licenses', {
                headers: { Authorization: `Bearer ${token}` },
            });

            if (!response.ok) {
                throw new Error('Failed to fetch licenses');
            }

            const data = await response.json();
            setLicenses(data.licenses);
        } catch (err) {
            console.error('Error fetching licenses:', err);
        } finally {
            setLicensesLoading(false);
        }
    }, [getAccessToken]);

    useEffect(() => {
        fetchOverview(excludeFreeTrialOverview);
    }, [fetchOverview, excludeFreeTrialOverview]);

    useEffect(() => {
        if (selectedTab === 'licenses') {
            fetchLicenses();
        }
    }, [selectedTab, fetchLicenses]);

    const handleSort = (column: string) => {
        if (sortBy === column) {
            setSortAscending(!sortAscending);
        } else {
            setSortBy(column);
            setSortAscending(true);
        }
    };

    const renderSortIcon = (column: string) => {
        if (sortBy !== column) return null;
        return sortAscending ? <ChevronUp24Regular /> : <ChevronDown24Regular />;
    };

    const filteredLicenses = licenses
        .filter(license => {
            if (searchQuery) {
                const query = searchQuery.toLowerCase();
                if (!license.displayName.toLowerCase().includes(query) &&
                    !license.skuPartNumber.toLowerCase().includes(query)) {
                    return false;
                }
            }
            if (filterType === 'trial' && !license.isTrial) return false;
            if (filterType === 'paid' && license.isTrial) return false;
            if (filterType === 'excludeFreeTrial' && license.isTrial) return false;
            if (filterType === 'warning' && license.status !== 'Warning') return false;
            return true;
        })
        .sort((a, b) => {
            let comparison = 0;
            switch (sortBy) {
                case 'totalUnits':
                    comparison = a.totalUnits - b.totalUnits;
                    break;
                case 'consumedUnits':
                    comparison = a.consumedUnits - b.consumedUnits;
                    break;
                case 'availableUnits':
                    comparison = a.availableUnits - b.availableUnits;
                    break;
                case 'utilization':
                    comparison = a.utilizationPercentage - b.utilizationPercentage;
                    break;
                default:
                    comparison = a.displayName.localeCompare(b.displayName);
            }
            return sortAscending ? comparison : -comparison;
        });

    const handleExport = () => {
        if (filteredLicenses.length === 0) return;
        
        const headers = ['License Name', 'SKU', 'Total', 'Assigned', 'Available', 'Utilization %', 'Status', 'Trial'];
        const rows = filteredLicenses.map(license => [
            license.displayName,
            license.skuPartNumber,
            license.totalUnits.toString(),
            license.consumedUnits.toString(),
            license.availableUnits.toString(),
            license.utilizationPercentage.toFixed(1),
            license.status,
            license.isTrial ? 'Yes' : 'No'
        ]);
        
        const csvContent = [
            headers.join(','),
            ...rows.map(row => row.map(cell => `"${(cell || '').replace(/"/g, '""')}"`).join(','))
        ].join('\n');
        
        const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = `M365_Licenses_${new Date().toISOString().split('T')[0]}.csv`;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    };

    if (loading) {
        return (
            <div className={styles.container}>
                <div className={styles.loadingContainer}>
                    <Spinner size="large" label="Loading license data..." />
                </div>
            </div>
        );
    }

    if (error) {
        return (
            <div className={styles.container}>
                <Card>
                    <div className={styles.emptyState}>
                        <Warning24Regular />
                        <Text>{error}</Text>
                        <Button appearance="primary" onClick={() => fetchOverview(true)} style={{ marginTop: '16px' }}>
                            Retry
                        </Button>
                    </div>
                </Card>
            </div>
        );
    }

    return (
        <div className={styles.container}>
            {/* Header */}
            <div className={styles.header}>
                <div className={styles.headerTitle}>
                    <Key24Regular />
                    <Title3>Licenses</Title3>
                </div>
                <Button
                    appearance="subtle"
                    icon={<ArrowClockwise24Regular />}
                    onClick={() => {
                        fetchOverview(excludeFreeTrialOverview);
                        if (selectedTab === 'licenses') fetchLicenses();
                    }}
                >
                    Refresh
                </Button>
            </div>

            {/* Tabs */}
            <TabList selectedValue={selectedTab} onTabSelect={(_, data) => setSelectedTab(data.value as string)}>
                <Tab value="overview">Overview</Tab>
                <Tab value="licenses">All Licenses</Tab>
            </TabList>

            {selectedTab === 'overview' && overview && (
                <>
                    {/* Filter Toggle */}
                    <div style={{ display: 'flex', alignItems: 'center', gap: tokens.spacingHorizontalM }}>
                        <Switch
                            checked={excludeFreeTrialOverview}
                            onChange={(_, data) => setExcludeFreeTrialOverview(data.checked)}
                            label="Exclude Free/Trial licenses"
                        />
                    </div>

                    {/* Stats Cards */}
                    <div className={styles.statsGrid}>
                        <Card className={styles.statCard}>
                            <Certificate24Regular />
                            <Text className={styles.statValue}>{overview.stats.totalSubscriptions}</Text>
                            <Text className={styles.statLabel}>Subscriptions</Text>
                        </Card>
                        <Card className={styles.statCard}>
                            <Key24Regular />
                            <Text className={styles.statValue}>{overview.stats.totalLicenses.toLocaleString()}</Text>
                            <Text className={styles.statLabel}>Total Licenses</Text>
                        </Card>
                        <Card className={styles.statCard}>
                            <People24Regular />
                            <Text className={styles.statValue}>{overview.stats.assignedLicenses.toLocaleString()}</Text>
                            <Text className={styles.statLabel}>Assigned</Text>
                        </Card>
                        <Card className={styles.statCard}>
                            <CheckmarkCircle24Regular />
                            <Text className={styles.statValue}>{overview.stats.availableLicenses.toLocaleString()}</Text>
                            <Text className={styles.statLabel}>Available</Text>
                        </Card>
                        <Card className={styles.statCard}>
                            <Info24Regular />
                            <Text className={styles.statValue}>{overview.stats.trialSubscriptions}</Text>
                            <Text className={styles.statLabel}>Trial/Free</Text>
                        </Card>
                        <Card className={styles.statCard}>
                            <Warning24Regular />
                            <Text className={styles.statValue}>{overview.stats.subscriptionsWithWarnings}</Text>
                            <Text className={styles.statLabel}>Warnings</Text>
                        </Card>
                    </div>

                    {/* Utilization Overview */}
                    <Card className={styles.utilizationCard}>
                        <div className={styles.utilizationHeader}>
                            <Title3>Overall Utilization</Title3>
                            <Badge color={getUtilizationColor(overview.stats.utilizationPercentage)}>
                                {overview.stats.utilizationPercentage.toFixed(1)}% Used
                            </Badge>
                        </div>
                        <div className={styles.utilizationInfo}>
                            <div className={styles.utilizationItem}>
                                <Text weight="semibold">{overview.stats.assignedLicenses.toLocaleString()}</Text>
                                <Text size={200} style={{ color: tokens.colorNeutralForeground3 }}>Assigned</Text>
                            </div>
                            <div className={styles.utilizationItem}>
                                <Text weight="semibold">{overview.stats.totalLicenses.toLocaleString()}</Text>
                                <Text size={200} style={{ color: tokens.colorNeutralForeground3 }}>Total</Text>
                            </div>
                            <div className={styles.utilizationItem}>
                                <Text weight="semibold">{overview.stats.availableLicenses.toLocaleString()}</Text>
                                <Text size={200} style={{ color: tokens.colorNeutralForeground3 }}>Available</Text>
                            </div>
                        </div>
                        <ProgressBar
                            value={overview.stats.utilizationPercentage / 100}
                            color={getUtilizationBarColor(overview.stats.utilizationPercentage)}
                            thickness="large"
                        />
                    </Card>

                    {/* Content Grid */}
                    <div className={styles.contentGrid}>
                        {/* Top Utilized */}
                        <Card className={styles.listCard}>
                            <CardHeader
                                header={<Text weight="semibold">Highest Utilization</Text>}
                                description="Licenses with highest usage"
                            />
                            <div className={styles.tableContainer}>
                                <Table size="small">
                                    <TableBody>
                                        {overview.topUtilized.map((license) => (
                                            <TableRow key={license.skuId}>
                                                <TableCell>
                                                    <Text title={license.skuPartNumber}>{license.displayName}</Text>
                                                </TableCell>
                                                <TableCell style={{ textAlign: 'right', whiteSpace: 'nowrap' }}>
                                                    <Text size={200} style={{ color: tokens.colorNeutralForeground3, marginRight: '8px' }}>
                                                        {license.consumedUnits} / {license.totalUnits}
                                                    </Text>
                                                    <Badge color={getUtilizationColor(license.utilizationPercentage)}>
                                                        {license.utilizationPercentage.toFixed(0)}%
                                                    </Badge>
                                                </TableCell>
                                            </TableRow>
                                        ))}
                                    </TableBody>
                                </Table>
                            </div>
                        </Card>

                        {/* Low Utilization */}
                        <Card className={styles.listCard}>
                            <CardHeader
                                header={<Text weight="semibold">Low Utilization</Text>}
                                description="Licenses with less than 50% usage"
                            />
                            <div className={styles.tableContainer}>
                                {overview.lowUtilization.length > 0 ? (
                                    <Table size="small">
                                        <TableBody>
                                            {overview.lowUtilization.map((license) => (
                                                <TableRow key={license.skuId}>
                                                    <TableCell>
                                                        <Text title={license.skuPartNumber}>{license.displayName}</Text>
                                                    </TableCell>
                                                    <TableCell style={{ textAlign: 'right' }}>
                                                        <Text size={200}>
                                                            {license.consumedUnits} / {license.totalUnits}
                                                        </Text>
                                                    </TableCell>
                                                </TableRow>
                                            ))}
                                        </TableBody>
                                    </Table>
                                ) : (
                                    <div className={styles.emptyState}>
                                        <CheckmarkCircle24Regular />
                                        <Text>All licenses well utilized</Text>
                                    </div>
                                )}
                            </div>
                        </Card>
                    </div>
                </>
            )}

            {selectedTab === 'licenses' && (
                <>
                    {/* Filters */}
                    <div className={styles.filterBar}>
                        <Input
                            className={styles.searchInput}
                            contentBefore={<Search24Regular />}
                            placeholder="Search licenses..."
                            value={searchQuery}
                            onChange={(_, data) => setSearchQuery(data.value)}
                        />
                        <select
                            value={filterType}
                            onChange={(e) => setFilterType(e.target.value)}
                            style={{
                                padding: '6px 12px',
                                borderRadius: '4px',
                                border: `1px solid ${tokens.colorNeutralStroke1}`,
                                backgroundColor: tokens.colorNeutralBackground1,
                                color: tokens.colorNeutralForeground1,
                                fontSize: tokens.fontSizeBase300,
                                minWidth: '150px',
                            }}
                        >
                            <option value="all">All Licenses</option>
                            <option value="paid">Paid Only</option>
                            <option value="excludeFreeTrial">Exclude Free/Trial</option>
                            <option value="trial">Trial/Free Only</option>
                            <option value="warning">With Warnings</option>
                        </select>
                        <Button
                            appearance="subtle"
                            icon={<ArrowDownload24Regular />}
                            onClick={handleExport}
                            disabled={filteredLicenses.length === 0}
                        >
                            Export
                        </Button>
                    </div>

                    {/* Licenses Table */}
                    <Card>
                        {licensesLoading ? (
                            <div className={styles.loadingContainer}>
                                <Spinner size="medium" label="Loading licenses..." />
                            </div>
                        ) : (
                            <div className={styles.tableContainer}>
                                <Table>
                                    <TableHeader>
                                        <TableRow>
                                            <TableHeaderCell>
                                                <span
                                                    className={styles.sortableHeader}
                                                    onClick={() => handleSort('displayName')}
                                                >
                                                    License Name {renderSortIcon('displayName')}
                                                </span>
                                            </TableHeaderCell>
                                            <TableHeaderCell>
                                                <span
                                                    className={styles.sortableHeader}
                                                    onClick={() => handleSort('totalUnits')}
                                                >
                                                    Total {renderSortIcon('totalUnits')}
                                                </span>
                                            </TableHeaderCell>
                                            <TableHeaderCell>
                                                <span
                                                    className={styles.sortableHeader}
                                                    onClick={() => handleSort('consumedUnits')}
                                                >
                                                    Assigned {renderSortIcon('consumedUnits')}
                                                </span>
                                            </TableHeaderCell>
                                            <TableHeaderCell>
                                                <span
                                                    className={styles.sortableHeader}
                                                    onClick={() => handleSort('availableUnits')}
                                                >
                                                    Available {renderSortIcon('availableUnits')}
                                                </span>
                                            </TableHeaderCell>
                                            <TableHeaderCell>
                                                <span
                                                    className={styles.sortableHeader}
                                                    onClick={() => handleSort('utilization')}
                                                >
                                                    Utilization {renderSortIcon('utilization')}
                                                </span>
                                            </TableHeaderCell>
                                            <TableHeaderCell>Status</TableHeaderCell>
                                        </TableRow>
                                    </TableHeader>
                                    <TableBody>
                                        {filteredLicenses.map((license) => (
                                            <TableRow key={license.skuId}>
                                                <TableCell>
                                                    <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                                                        <Text title={license.skuPartNumber}>{license.displayName}</Text>
                                                        {license.isTrial && (
                                                            <Badge size="small" appearance="outline" color="informative">
                                                                Trial
                                                            </Badge>
                                                        )}
                                                    </div>
                                                </TableCell>
                                                <TableCell>
                                                    <Text>{license.totalUnits.toLocaleString()}</Text>
                                                </TableCell>
                                                <TableCell>
                                                    <Text>{license.consumedUnits.toLocaleString()}</Text>
                                                </TableCell>
                                                <TableCell>
                                                    <Text>{license.availableUnits.toLocaleString()}</Text>
                                                </TableCell>
                                                <TableCell>
                                                    <div className={styles.progressCell}>
                                                        <div className={styles.progressBar}>
                                                            <ProgressBar
                                                                value={license.utilizationPercentage / 100}
                                                                color={getUtilizationBarColor(license.utilizationPercentage)}
                                                                thickness="medium"
                                                            />
                                                        </div>
                                                        <Text size={200}>{license.utilizationPercentage.toFixed(0)}%</Text>
                                                    </div>
                                                </TableCell>
                                                <TableCell>
                                                    <Badge 
                                                        color={license.status === 'Enabled' ? 'success' : 
                                                               license.status === 'Warning' ? 'warning' : 'danger'}
                                                        appearance="tint"
                                                    >
                                                        {license.status}
                                                    </Badge>
                                                </TableCell>
                                            </TableRow>
                                        ))}
                                    </TableBody>
                                </Table>
                                {filteredLicenses.length === 0 && (
                                    <div className={styles.emptyState}>
                                        <Text>No licenses found</Text>
                                    </div>
                                )}
                            </div>
                        )}
                    </Card>
                </>
            )}
        </div>
    );
};

export default LicensesPage;
