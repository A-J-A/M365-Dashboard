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
    Dropdown,
    Option,
    Table,
    TableHeader,
    TableHeaderCell,
    TableBody,
    TableRow,
    TableCell,
    ProgressBar,
    Tooltip,
    Tab,
    TabList,
} from '@fluentui/react-components';
import {
    Database24Regular,
    Storage24Regular,
    People24Regular,
    Globe24Regular,
    ArrowClockwise24Regular,
    Search24Regular,
    ChevronUp24Regular,
    ChevronDown24Regular,
    Warning24Regular,
    Checkmark24Regular,
    Open24Regular,
    ArrowDownload24Regular,
    DocumentBulletList24Regular,
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
        gridTemplateColumns: 'repeat(auto-fit, minmax(200px, 1fr))',
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
    storageCard: {
        padding: tokens.spacingVerticalL,
    },
    storageHeader: {
        display: 'flex',
        justifyContent: 'space-between',
        alignItems: 'center',
        marginBottom: tokens.spacingVerticalM,
    },
    storageInfo: {
        display: 'flex',
        gap: tokens.spacingHorizontalXL,
        marginBottom: tokens.spacingVerticalM,
    },
    storageItem: {
        display: 'flex',
        flexDirection: 'column',
        gap: tokens.spacingVerticalXXS,
    },
    progressContainer: {
        marginTop: tokens.spacingVerticalS,
    },
    contentGrid: {
        display: 'grid',
        gridTemplateColumns: 'repeat(auto-fit, minmax(400px, 1fr))',
        gap: tokens.spacingHorizontalL,
    },
    listCard: {
        minHeight: '300px',
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
    siteLink: {
        color: tokens.colorBrandForeground1,
        textDecoration: 'none',
        display: 'flex',
        alignItems: 'center',
        gap: tokens.spacingHorizontalXS,
        ':hover': {
            textDecoration: 'underline',
        },
    },
    storageCell: {
        display: 'flex',
        flexDirection: 'column',
        gap: '2px',
    },
    miniProgress: {
        width: '80px',
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
});

interface SharePointSite {
    id: string;
    name: string;
    displayName: string;
    description?: string;
    webUrl: string;
    siteTemplate?: string;
    createdDateTime?: string;
    lastModifiedDateTime?: string;
    storageUsedBytes: number;
    storageAllocatedBytes: number;
    storageUsedPercentage: number;
    ownerDisplayName?: string;
    ownerEmail?: string;
    isPersonalSite: boolean;
    itemCount?: number;
    status?: string;
}

interface SharePointStats {
    totalSites: number;
    teamSites: number;
    communicationSites: number;
    personalSites: number;
    otherSites: number;
    totalStorageUsedBytes: number;
    totalStorageAllocatedBytes: number;
    overallStorageUsedPercentage: number;
    sitesNearQuota: number;
    activeSitesLast30Days: number;
    inactiveSitesLast30Days: number;
    lastUpdated: string;
}

interface SharePointOverview {
    stats: SharePointStats;
    largestSites: SharePointSite[];
    recentlyCreatedSites: SharePointSite[];
    sitesNearStorageLimit: SharePointSite[];
    lastUpdated: string;
}

const formatBytes = (bytes: number): string => {
    if (bytes === 0) return '0 B';
    const k = 1024;
    const sizes = ['B', 'KB', 'MB', 'GB', 'TB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
};

const formatDate = (dateString?: string): string => {
    if (!dateString) return 'N/A';
    return new Date(dateString).toLocaleDateString('en-GB', {
        day: '2-digit',
        month: 'short',
        year: 'numeric',
    });
};

export const SharePointPage: React.FC = () => {
    const styles = useStyles();
    const { getAccessToken } = useAppContext();
    
    const [overview, setOverview] = useState<SharePointOverview | null>(null);
    const [sites, setSites] = useState<SharePointSite[]>([]);
    const [loading, setLoading] = useState(true);
    const [sitesLoading, setSitesLoading] = useState(false);
    const [error, setError] = useState<string | null>(null);
    
    const [selectedTab, setSelectedTab] = useState<string>('overview');
    const [searchQuery, setSearchQuery] = useState('');
    const [siteTypeFilter, setSiteTypeFilter] = useState<string>('all');
    const [sortBy, setSortBy] = useState<string>('name');
    const [sortAscending, setSortAscending] = useState(true);

    const fetchOverview = useCallback(async () => {
        try {
            setLoading(true);
            setError(null);
            
            const token = await getAccessToken();
            const response = await fetch('/api/sharepoint/overview', {
                headers: { Authorization: `Bearer ${token}` },
            });

            if (!response.ok) {
                throw new Error('Failed to fetch SharePoint overview');
            }

            const data = await response.json();
            setOverview(data);
        } catch (err) {
            setError(err instanceof Error ? err.message : 'An error occurred');
        } finally {
            setLoading(false);
        }
    }, [getAccessToken]);

    const fetchSites = useCallback(async () => {
        try {
            setSitesLoading(true);
            
            const token = await getAccessToken();
            const params = new URLSearchParams({
                orderBy: sortBy,
                ascending: sortAscending.toString(),
                take: '100',
            });
            
            if (searchQuery) params.append('search', searchQuery);
            if (siteTypeFilter !== 'all') params.append('siteType', siteTypeFilter);

            const response = await fetch(`/api/sharepoint/sites?${params}`, {
                headers: { Authorization: `Bearer ${token}` },
            });

            if (!response.ok) {
                throw new Error('Failed to fetch sites');
            }

            const data = await response.json();
            setSites(data.sites);
        } catch (err) {
            console.error('Error fetching sites:', err);
        } finally {
            setSitesLoading(false);
        }
    }, [getAccessToken, searchQuery, siteTypeFilter, sortBy, sortAscending]);

    useEffect(() => {
        fetchOverview();
    }, [fetchOverview]);

    useEffect(() => {
        if (selectedTab === 'sites') {
            fetchSites();
        }
    }, [selectedTab, fetchSites]);

    const handleSort = (column: string) => {
        if (sortBy === column) {
            setSortAscending(!sortAscending);
        } else {
            setSortBy(column);
            setSortAscending(true);
        }
    };

    const handleExport = () => {
        if (sites.length === 0) return;
        
        // Create CSV content
        const headers = ['Site Name', 'URL', 'Storage Used', 'Last Modified', 'Created', 'Site Type'];
        const rows = sites.map(site => [
            site.displayName,
            site.webUrl,
            formatBytes(site.storageUsedBytes),
            site.lastModifiedDateTime ? new Date(site.lastModifiedDateTime).toISOString().split('T')[0] : '',
            site.createdDateTime ? new Date(site.createdDateTime).toISOString().split('T')[0] : '',
            site.siteTemplate === 'GROUP#0' ? 'Team Site' : 
                site.siteTemplate === 'SITEPAGEPUBLISHING#0' ? 'Communication Site' :
                site.isPersonalSite ? 'Personal Site' : 'Other'
        ]);
        
        const csvContent = [
            headers.join(','),
            ...rows.map(row => row.map(cell => `"${(cell || '').replace(/"/g, '""')}"`).join(','))
        ].join('\n');
        
        // Create and download file
        const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = `SharePoint_Sites_${new Date().toISOString().split('T')[0]}.csv`;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    };

    const renderSortIcon = (column: string) => {
        if (sortBy !== column) return null;
        return sortAscending ? <ChevronUp24Regular /> : <ChevronDown24Regular />;
    };

    const getStorageColor = (percentage: number): 'success' | 'warning' | 'danger' => {
        if (percentage >= 90) return 'danger';
        if (percentage >= 80) return 'warning';
        return 'success';
    };

    if (loading) {
        return (
            <div className={styles.container}>
                <div className={styles.loadingContainer}>
                    <Spinner size="large" label="Loading SharePoint data..." />
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
                        <Button appearance="primary" onClick={fetchOverview} style={{ marginTop: '16px' }}>
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
                    <DocumentBulletList24Regular />
                    <Title3>SharePoint</Title3>
                </div>
                <Button
                    appearance="subtle"
                    icon={<ArrowClockwise24Regular />}
                    onClick={() => {
                        fetchOverview();
                        if (selectedTab === 'sites') fetchSites();
                    }}
                >
                    Refresh
                </Button>
            </div>

            {/* Tabs */}
            <TabList selectedValue={selectedTab} onTabSelect={(_, data) => setSelectedTab(data.value as string)}>
                <Tab value="overview">Overview</Tab>
                <Tab value="sites">All Sites</Tab>
            </TabList>

            {selectedTab === 'overview' && overview && (
                <>
                    {/* Stats Cards */}
                    <div className={styles.statsGrid}>
                        <Card className={styles.statCard}>
                            <Database24Regular />
                            <Text className={styles.statValue}>{overview.stats.totalSites}</Text>
                            <Text className={styles.statLabel}>Total Sites</Text>
                        </Card>
                        <Card className={styles.statCard}>
                            <People24Regular />
                            <Text className={styles.statValue}>{overview.stats.teamSites}</Text>
                            <Text className={styles.statLabel}>Team Sites</Text>
                        </Card>
                        <Card className={styles.statCard}>
                            <Globe24Regular />
                            <Text className={styles.statValue}>{overview.stats.communicationSites}</Text>
                            <Text className={styles.statLabel}>Communication Sites</Text>
                        </Card>
                        <Card className={styles.statCard}>
                            <Storage24Regular />
                            <Text className={styles.statValue}>{overview.stats.personalSites}</Text>
                            <Text className={styles.statLabel}>Personal Sites (OneDrive)</Text>
                        </Card>
                        <Card className={styles.statCard}>
                            <Checkmark24Regular />
                            <Text className={styles.statValue}>{overview.stats.activeSitesLast30Days}</Text>
                            <Text className={styles.statLabel}>Active (Last 30 Days)</Text>
                        </Card>
                        <Card className={styles.statCard}>
                            <Warning24Regular />
                            <Text className={styles.statValue}>{overview.stats.sitesNearQuota}</Text>
                            <Text className={styles.statLabel}>Near Storage Limit</Text>
                        </Card>
                    </div>

                    {/* Storage Overview */}
                    <Card className={styles.storageCard}>
                        <div className={styles.storageHeader}>
                            <Title3>Storage Usage</Title3>
                        </div>
                        <div className={styles.storageInfo}>
                            <div className={styles.storageItem}>
                                <Text weight="semibold">{formatBytes(overview.stats.totalStorageUsedBytes)}</Text>
                                <Text size={200} style={{ color: tokens.colorNeutralForeground3 }}>Total Used</Text>
                            </div>
                        </div>
                        <Text size={200} style={{ color: tokens.colorNeutralForeground3 }}>
                            Note: Site discovery uses Microsoft Graph API search. Some sites may not appear if they haven't been indexed by SharePoint search yet.
                        </Text>
                    </Card>

                    {/* Content Grid */}
                    <div className={styles.contentGrid}>
                        {/* Largest Sites */}
                        <Card className={styles.listCard}>
                            <CardHeader
                                header={<Text weight="semibold">Largest Sites</Text>}
                                description="Sites using the most storage"
                            />
                            <div className={styles.tableContainer}>
                                <Table size="small">
                                    <TableBody>
                                        {overview.largestSites.slice(0, 5).map((site) => (
                                            <TableRow key={site.id}>
                                                <TableCell>
                                                    <a
                                                        href={site.webUrl}
                                                        target="_blank"
                                                        rel="noopener noreferrer"
                                                        className={styles.siteLink}
                                                    >
                                                        {site.displayName}
                                                        <Open24Regular style={{ width: 12, height: 12 }} />
                                                    </a>
                                                </TableCell>
                                                <TableCell style={{ textAlign: 'right' }}>
                                                    <Text weight="semibold">{formatBytes(site.storageUsedBytes)}</Text>
                                                </TableCell>
                                            </TableRow>
                                        ))}
                                    </TableBody>
                                </Table>
                            </div>
                        </Card>

                        {/* Sites Near Storage Limit */}
                        <Card className={styles.listCard}>
                            <CardHeader
                                header={<Text weight="semibold">Sites Near Storage Limit</Text>}
                                description="Sites with 80%+ storage used"
                            />
                            <div className={styles.tableContainer}>
                                {overview.sitesNearStorageLimit.length > 0 ? (
                                    <Table size="small">
                                        <TableBody>
                                            {overview.sitesNearStorageLimit.slice(0, 5).map((site) => (
                                                <TableRow key={site.id}>
                                                    <TableCell>
                                                        <a
                                                            href={site.webUrl}
                                                            target="_blank"
                                                            rel="noopener noreferrer"
                                                            className={styles.siteLink}
                                                        >
                                                            {site.displayName}
                                                            <Open24Regular style={{ width: 12, height: 12 }} />
                                                        </a>
                                                    </TableCell>
                                                    <TableCell style={{ textAlign: 'right' }}>
                                                        <Badge color={getStorageColor(site.storageUsedPercentage)}>
                                                            {site.storageUsedPercentage.toFixed(0)}%
                                                        </Badge>
                                                    </TableCell>
                                                </TableRow>
                                            ))}
                                        </TableBody>
                                    </Table>
                                ) : (
                                    <div className={styles.emptyState}>
                                        <Checkmark24Regular />
                                        <Text>No sites near storage limit</Text>
                                    </div>
                                )}
                            </div>
                        </Card>

                        {/* Recently Created Sites */}
                        <Card className={styles.listCard}>
                            <CardHeader
                                header={<Text weight="semibold">Recently Created Sites</Text>}
                                description="Newest sites in your tenant"
                            />
                            <div className={styles.tableContainer}>
                                <Table size="small">
                                    <TableBody>
                                        {overview.recentlyCreatedSites.slice(0, 5).map((site) => (
                                            <TableRow key={site.id}>
                                                <TableCell>
                                                    <a
                                                        href={site.webUrl}
                                                        target="_blank"
                                                        rel="noopener noreferrer"
                                                        className={styles.siteLink}
                                                    >
                                                        {site.displayName}
                                                        <Open24Regular style={{ width: 12, height: 12 }} />
                                                    </a>
                                                </TableCell>
                                                <TableCell style={{ textAlign: 'right' }}>
                                                    <Text size={200}>{formatDate(site.createdDateTime)}</Text>
                                                </TableCell>
                                            </TableRow>
                                        ))}
                                    </TableBody>
                                </Table>
                            </div>
                        </Card>
                    </div>
                </>
            )}

            {selectedTab === 'sites' && (
                <>
                    {/* Filters */}
                    <div className={styles.filterBar}>
                        <Input
                            className={styles.searchInput}
                            contentBefore={<Search24Regular />}
                            placeholder="Search sites..."
                            value={searchQuery}
                            onChange={(_, data) => setSearchQuery(data.value)}
                        />
                        <select
                            value={siteTypeFilter}
                            onChange={(e) => setSiteTypeFilter(e.target.value)}
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
                            <option value="all">All Types</option>
                            <option value="team">Team Sites</option>
                            <option value="communication">Communication Sites</option>
                            <option value="personal">Personal Sites</option>
                        </select>
                        <Button
                            appearance="subtle"
                            icon={<ArrowDownload24Regular />}
                            onClick={handleExport}
                            disabled={sites.length === 0}
                        >
                            Export
                        </Button>
                    </div>

                    {/* Sites Table */}
                    <Card>
                        {sitesLoading ? (
                            <div className={styles.loadingContainer}>
                                <Spinner size="medium" label="Loading sites..." />
                            </div>
                        ) : (
                            <div className={styles.tableContainer}>
                                <Table>
                                    <TableHeader>
                                        <TableRow>
                                            <TableHeaderCell>
                                                <span
                                                    className={styles.sortableHeader}
                                                    onClick={() => handleSort('name')}
                                                >
                                                    Site Name {renderSortIcon('name')}
                                                </span>
                                            </TableHeaderCell>
                                            <TableHeaderCell>URL</TableHeaderCell>
                                            <TableHeaderCell>
                                                <span
                                                    className={styles.sortableHeader}
                                                    onClick={() => handleSort('storage')}
                                                >
                                                    Storage {renderSortIcon('storage')}
                                                </span>
                                            </TableHeaderCell>
                                            <TableHeaderCell>
                                                <span
                                                    className={styles.sortableHeader}
                                                    onClick={() => handleSort('modified')}
                                                >
                                                    Last Modified {renderSortIcon('modified')}
                                                </span>
                                            </TableHeaderCell>
                                            <TableHeaderCell>
                                                <span
                                                    className={styles.sortableHeader}
                                                    onClick={() => handleSort('created')}
                                                >
                                                    Created {renderSortIcon('created')}
                                                </span>
                                            </TableHeaderCell>
                                        </TableRow>
                                    </TableHeader>
                                    <TableBody>
                                        {sites.map((site) => (
                                            <TableRow key={site.id}>
                                                <TableCell>
                                                    <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                                                        {site.displayName}
                                                        {site.isPersonalSite && (
                                                            <Badge size="small" appearance="outline">OneDrive</Badge>
                                                        )}
                                                    </div>
                                                </TableCell>
                                                <TableCell>
                                                    <a
                                                        href={site.webUrl}
                                                        target="_blank"
                                                        rel="noopener noreferrer"
                                                        className={styles.siteLink}
                                                    >
                                                        <Open24Regular style={{ width: 14, height: 14 }} />
                                                        Open
                                                    </a>
                                                </TableCell>
                                                <TableCell>
                                                    <Text>{formatBytes(site.storageUsedBytes)}</Text>
                                                </TableCell>
                                                <TableCell>{formatDate(site.lastModifiedDateTime)}</TableCell>
                                                <TableCell>{formatDate(site.createdDateTime)}</TableCell>
                                            </TableRow>
                                        ))}
                                    </TableBody>
                                </Table>
                                {sites.length === 0 && (
                                    <div className={styles.emptyState}>
                                        <Text>No sites found</Text>
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

export default SharePointPage;
