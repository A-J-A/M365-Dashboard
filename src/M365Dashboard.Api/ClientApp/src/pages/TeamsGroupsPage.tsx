import React, { useState, useEffect, useMemo } from 'react';
import {
  PeopleTeamRegular,
  PeopleTeamFilled,
  ShieldRegular,
  MailRegular,
  SearchRegular,
  OpenRegular,
  ChevronUpRegular,
  ChevronDownRegular,
  PeopleRegular,
  PersonRegular,
  LockClosedRegular,
  GlobeRegular,
  CheckmarkCircleFilled,
  DismissRegular,
  ArrowDownloadRegular,
  StarRegular,
  PeopleCommunityRegular,
  InfoRegular,
  ArrowSyncRegular,
} from '@fluentui/react-icons';
import { useAppContext } from '../contexts/AppContext';

interface TenantGroup {
  id: string;
  displayName: string;
  description: string | null;
  mail: string | null;
  groupType: string;
  mailEnabled: boolean | null;
  securityEnabled: boolean | null;
  visibility: string | null;
  createdDateTime: string | null;
  renewedDateTime: string | null;
  memberCount: number;
  ownerCount: number;
  isTeam: boolean;
  teamWebUrl: string | null;
  resourceProvisioningOptions: string[] | null;
  // Exchange-specific fields
  isExchangeDL?: boolean;
  primarySmtpAddress?: string | null;
  alias?: string | null;
}

interface GroupStats {
  totalGroups: number;
  microsoft365Groups: number;
  securityGroups: number;
  distributionGroups: number;
  teamsEnabled: number;
  publicGroups: number;
  privateGroups: number;
  groupsWithNoOwner: number;
  groupsWithNoMembers: number;
  lastUpdated: string;
}

interface GroupListResult {
  groups: TenantGroup[];
  totalCount: number;
  filteredCount: number;
  nextLink: string | null;
}

interface GroupMember {
  id: string;
  displayName: string;
  userPrincipalName: string | null;
  mail: string | null;
  memberType: string;
}

interface GroupDetail {
  id: string;
  displayName: string;
  description: string | null;
  mail: string | null;
  groupType: string;
  mailEnabled: boolean | null;
  securityEnabled: boolean | null;
  visibility: string | null;
  createdDateTime: string | null;
  renewedDateTime: string | null;
  expirationDateTime: string | null;
  members: GroupMember[] | null;
  owners: GroupMember[] | null;
  isTeam: boolean;
  teamWebUrl: string | null;
  isArchived: boolean | null;
  resourceProvisioningOptions: string[] | null;
}

// Exchange DDL interfaces
interface ExchangeDistributionList {
  id: string;
  displayName: string;
  primarySmtpAddress: string | null;
  alias: string | null;
  groupType: string;
  recipientType: string;
  memberCount: number;
  whenCreated: string | null;
  hiddenFromAddressListsEnabled: boolean;
}

interface ExchangeDLResult {
  distributionLists: ExchangeDistributionList[];
  totalCount: number;
}

interface ExchangeDLMember {
  id: string;
  displayName: string;
  primarySmtpAddress: string | null;
  recipientType: string;
}

interface ExchangeDLDetail {
  id: string;
  displayName: string;
  primarySmtpAddress: string | null;
  alias: string | null;
  description: string | null;
  managedBy: string[];
  groupType: string;
  recipientType: string;
  memberCount: number;
  members: ExchangeDLMember[];
  whenCreated: string | null;
  whenChanged: string | null;
  hiddenFromAddressListsEnabled: boolean;
  requireSenderAuthenticationEnabled: boolean;
  acceptMessagesOnlyFromSendersOrMembers: string[];
  emailAddresses: string[];
  isDynamic?: boolean;
  recipientFilter?: string | null;
}

type FilterType = 'all' | 'm365' | 'security' | 'distribution' | 'teams' | 'public' | 'private';
type SortField = 'displayName' | 'groupType' | 'memberCount' | 'createdDateTime' | 'isTeam';
type SortDirection = 'asc' | 'desc';

const TeamsGroupsPage: React.FC = () => {
  const { getAccessToken } = useAppContext();
  const [groups, setGroups] = useState<TenantGroup[]>([]);
  const [stats, setStats] = useState<GroupStats | null>(null);
  const [loading, setLoading] = useState(true);
  const [statsLoading, setStatsLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [searchTerm, setSearchTerm] = useState('');
  const [filterType, setFilterType] = useState<FilterType>('all');
  const [sortField, setSortField] = useState<SortField>('displayName');
  const [sortDirection, setSortDirection] = useState<SortDirection>('asc');
  
  // Exchange DDL state
  const [exchangeDLs, setExchangeDLs] = useState<TenantGroup[]>([]);
  const [exchangeDLsLoading, setExchangeDLsLoading] = useState(false);
  const [exchangeDLsError, setExchangeDLsError] = useState<string | null>(null);
  const [exchangeDLsLoaded, setExchangeDLsLoaded] = useState(false);
  
  // Member panel state
  const [selectedGroup, setSelectedGroup] = useState<GroupDetail | null>(null);
  const [selectedExchangeDL, setSelectedExchangeDL] = useState<ExchangeDLDetail | null>(null);
  const [memberPanelOpen, setMemberPanelOpen] = useState(false);
  const [memberLoading, setMemberLoading] = useState(false);
  const [memberSearchTerm, setMemberSearchTerm] = useState('');

  useEffect(() => {
    fetchGroups();
    fetchStats();
    fetchExchangeDLs(); // Fetch Exchange DDLs on load for accurate count
  }, []);

  // Fetch Exchange DDLs when distribution filter is selected
  useEffect(() => {
    if (filterType === 'distribution' && !exchangeDLsLoaded && !exchangeDLsLoading) {
      fetchExchangeDLs();
    }
  }, [filterType]);

  const fetchGroups = async () => {
    try {
      setLoading(true);
      const token = await getAccessToken();
      const response = await fetch('/api/groups?take=200', {
        headers: {
          'Authorization': `Bearer ${token}`,
          'Content-Type': 'application/json',
        },
      });

      if (!response.ok) {
        throw new Error('Failed to fetch groups');
      }

      const data: GroupListResult = await response.json();
      setGroups(data.groups);
      setError(null);
    } catch (err) {
      setError(err instanceof Error ? err.message : 'An error occurred');
    } finally {
      setLoading(false);
    }
  };

  const fetchStats = async () => {
    try {
      setStatsLoading(true);
      const token = await getAccessToken();
      const response = await fetch('/api/groups/stats', {
        headers: {
          'Authorization': `Bearer ${token}`,
          'Content-Type': 'application/json',
        },
      });

      if (!response.ok) {
        throw new Error('Failed to fetch group stats');
      }

      const data: GroupStats = await response.json();
      setStats(data);
    } catch (err) {
      console.error('Failed to fetch stats:', err);
    } finally {
      setStatsLoading(false);
    }
  };

  const fetchExchangeDLs = async () => {
    try {
      setExchangeDLsLoading(true);
      setExchangeDLsError(null);
      const token = await getAccessToken();
      const response = await fetch('/api/exchange/distribution-lists?take=200', {
        headers: {
          'Authorization': `Bearer ${token}`,
          'Content-Type': 'application/json',
        },
      });

      if (!response.ok) {
        const errorData = await response.json().catch(() => ({}));
        throw new Error(errorData.message || 'Failed to fetch Exchange distribution lists');
      }

      const data: ExchangeDLResult = await response.json();
      
      // Convert Exchange DDLs to TenantGroup format
      const convertedDLs: TenantGroup[] = data.distributionLists.map(dl => ({
        id: dl.id,
        displayName: dl.displayName,
        description: null,
        mail: dl.primarySmtpAddress,
        groupType: dl.groupType || 'Distribution',
        mailEnabled: true,
        securityEnabled: false,
        visibility: null,
        createdDateTime: dl.whenCreated,
        renewedDateTime: null,
        memberCount: dl.memberCount,
        ownerCount: 0,
        isTeam: false,
        teamWebUrl: null,
        resourceProvisioningOptions: null,
        isExchangeDL: true,
        primarySmtpAddress: dl.primarySmtpAddress,
        alias: dl.alias,
      }));
      
      setExchangeDLs(convertedDLs);
      setExchangeDLsLoaded(true);
      
      // Update stats with Exchange DDL count
      if (stats) {
        setStats({
          ...stats,
          distributionGroups: (stats.distributionGroups || 0) + data.totalCount,
        });
      }
    } catch (err) {
      setExchangeDLsError(err instanceof Error ? err.message : 'Failed to fetch Exchange DDLs');
      console.error('Failed to fetch Exchange DDLs:', err);
    } finally {
      setExchangeDLsLoading(false);
    }
  };

  const fetchGroupDetails = async (groupId: string, isExchangeDL: boolean = false, emailAddress?: string | null) => {
    try {
      setMemberLoading(true);
      setMemberPanelOpen(true);
      setSelectedGroup(null);
      setSelectedExchangeDL(null);
      const token = await getAccessToken();
      
      if (isExchangeDL) {
        // Use email address for Exchange DDLs as it's more reliable than GUID
        const identity = emailAddress || groupId;
        const response = await fetch(`/api/exchange/distribution-lists/${encodeURIComponent(identity)}`, {
          headers: {
            'Authorization': `Bearer ${token}`,
            'Content-Type': 'application/json',
          },
        });

        if (!response.ok) {
          throw new Error('Failed to fetch distribution list details');
        }

        const data: ExchangeDLDetail = await response.json();
        setSelectedExchangeDL(data);
      } else {
        const response = await fetch(`/api/groups/${groupId}`, {
          headers: {
            'Authorization': `Bearer ${token}`,
            'Content-Type': 'application/json',
          },
        });

        if (!response.ok) {
          throw new Error('Failed to fetch group details');
        }

        const data: GroupDetail = await response.json();
        setSelectedGroup(data);
      }
      setMemberSearchTerm('');
    } catch (err) {
      console.error('Failed to fetch group details:', err);
    } finally {
      setMemberLoading(false);
    }
  };

  const closeMemberPanel = () => {
    setMemberPanelOpen(false);
    setSelectedGroup(null);
    setSelectedExchangeDL(null);
    setMemberSearchTerm('');
  };

  const filteredMembers = useMemo(() => {
    if (selectedExchangeDL) {
      if (!memberSearchTerm) return selectedExchangeDL.members;
      const term = memberSearchTerm.toLowerCase();
      return selectedExchangeDL.members.filter(
        (member) =>
          member.displayName.toLowerCase().includes(term) ||
          (member.primarySmtpAddress?.toLowerCase().includes(term) ?? false)
      );
    }
    
    if (!selectedGroup?.members) return [];
    if (!memberSearchTerm) return selectedGroup.members;
    
    const term = memberSearchTerm.toLowerCase();
    return selectedGroup.members.filter(
      (member) =>
        member.displayName.toLowerCase().includes(term) ||
        (member.userPrincipalName?.toLowerCase().includes(term) ?? false) ||
        (member.mail?.toLowerCase().includes(term) ?? false)
    );
  }, [selectedGroup?.members, selectedExchangeDL?.members, memberSearchTerm]);

  const exportMembersToCsv = () => {
    const members = selectedExchangeDL ? selectedExchangeDL.members : selectedGroup?.members;
    if (!members) return;

    const headers = ['Display Name', 'Email', 'Type'];
    const rows = selectedExchangeDL 
      ? filteredMembers.map(m => {
          const member = m as ExchangeDLMember;
          return [member.displayName, member.primarySmtpAddress || '', member.recipientType];
        })
      : filteredMembers.map(m => {
          const member = m as GroupMember;
          return [member.displayName, member.mail || member.userPrincipalName || '', member.memberType];
        });

    const csvContent = [headers, ...rows].map(row => row.map(cell => `"${cell}"`).join(',')).join('\n');
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    const displayName = selectedExchangeDL?.displayName || selectedGroup?.displayName || 'group';
    const safeName = displayName.replace(/[^a-z0-9]/gi, '-').toLowerCase();
    link.download = `${safeName}-members-${new Date().toISOString().split('T')[0]}.csv`;
    link.click();
  };

  // Combine groups with Exchange DDLs when showing distribution filter
  const allGroups = useMemo(() => {
    if (filterType === 'distribution') {
      // When distribution filter is active, show both Graph DDLs and Exchange DDLs
      const graphDDLs = groups.filter(g => g.groupType === 'Distribution');
      // Dedupe by email address
      const graphEmails = new Set(graphDDLs.map(g => g.mail?.toLowerCase()).filter(Boolean));
      const uniqueExchangeDLs = exchangeDLs.filter(dl => !graphEmails.has(dl.mail?.toLowerCase()));
      return [...graphDDLs, ...uniqueExchangeDLs];
    }
    return groups;
  }, [groups, exchangeDLs, filterType]);

  const filteredAndSortedGroups = useMemo(() => {
    let result = [...allGroups];

    // Search filter
    if (searchTerm) {
      const term = searchTerm.toLowerCase();
      result = result.filter(
        (group) =>
          group.displayName.toLowerCase().includes(term) ||
          (group.description?.toLowerCase().includes(term) ?? false) ||
          (group.mail?.toLowerCase().includes(term) ?? false)
      );
    }

    // Type filter (already applied in allGroups for distribution)
    if (filterType !== 'distribution') {
      switch (filterType) {
        case 'm365':
          result = result.filter((g) => g.groupType === 'Microsoft 365');
          break;
        case 'security':
          result = result.filter((g) => g.groupType === 'Security' || g.groupType === 'Mail-enabled Security');
          break;
        case 'teams':
          result = result.filter((g) => g.isTeam);
          break;
        case 'public':
          result = result.filter((g) => g.visibility?.toLowerCase() === 'public');
          break;
        case 'private':
          result = result.filter((g) => g.visibility?.toLowerCase() === 'private');
          break;
      }
    }

    // Sort
    result.sort((a, b) => {
      let aValue: string | number | boolean | null = null;
      let bValue: string | number | boolean | null = null;

      switch (sortField) {
        case 'displayName':
          aValue = a.displayName;
          bValue = b.displayName;
          break;
        case 'groupType':
          aValue = a.groupType;
          bValue = b.groupType;
          break;
        case 'memberCount':
          aValue = a.memberCount;
          bValue = b.memberCount;
          break;
        case 'createdDateTime':
          aValue = a.createdDateTime;
          bValue = b.createdDateTime;
          break;
        case 'isTeam':
          aValue = a.isTeam;
          bValue = b.isTeam;
          break;
      }

      if (aValue === null && bValue === null) return 0;
      if (aValue === null) return 1;
      if (bValue === null) return -1;

      if (typeof aValue === 'boolean') {
        return sortDirection === 'asc'
          ? (aValue === bValue ? 0 : aValue ? -1 : 1)
          : (aValue === bValue ? 0 : aValue ? 1 : -1);
      }

      if (typeof aValue === 'number') {
        return sortDirection === 'asc' ? aValue - (bValue as number) : (bValue as number) - aValue;
      }

      const comparison = String(aValue).localeCompare(String(bValue));
      return sortDirection === 'asc' ? comparison : -comparison;
    });

    return result;
  }, [allGroups, searchTerm, filterType, sortField, sortDirection]);

  const handleSort = (field: SortField) => {
    if (sortField === field) {
      setSortDirection(sortDirection === 'asc' ? 'desc' : 'asc');
    } else {
      setSortField(field);
      setSortDirection('asc');
    }
  };

  const formatDate = (dateString: string | null) => {
    if (!dateString) return 'Never';
    const date = new Date(dateString);
    return date.toLocaleDateString();
  };

  const openInEntraPortal = (groupId: string) => {
    const url = `https://entra.microsoft.com/#view/Microsoft_AAD_IAM/GroupDetailsMenuBlade/~/Overview/groupId/${groupId}/menuId/`;
    window.open(url, '_blank');
  };

  const openInExchangeAdmin = (email: string) => {
    const url = `https://admin.exchange.microsoft.com/#/groups`;
    window.open(url, '_blank');
  };

  const openInTeams = (teamWebUrl: string) => {
    window.open(teamWebUrl, '_blank');
  };

  const getFilterLabel = (filter: FilterType): string => {
    switch (filter) {
      case 'all': return 'All Groups';
      case 'm365': return 'Microsoft 365';
      case 'security': return 'Security';
      case 'distribution': return 'Distribution';
      case 'teams': return 'Teams';
      case 'public': return 'Public';
      case 'private': return 'Private';
      default: return 'All Groups';
    }
  };

  const getGroupTypeIcon = (groupType: string, isTeam: boolean, isExchangeDL?: boolean) => {
    if (isTeam) {
      return <PeopleTeamFilled className="w-4 h-4 text-purple-600" />;
    }
    if (isExchangeDL) {
      return <MailRegular className="w-4 h-4 text-orange-600" />;
    }
    switch (groupType) {
      case 'Microsoft 365':
        return <PeopleTeamRegular className="w-4 h-4 text-blue-600" />;
      case 'Security':
      case 'Mail-enabled Security':
        return <ShieldRegular className="w-4 h-4 text-green-600" />;
      case 'Distribution':
        return <MailRegular className="w-4 h-4 text-orange-600" />;
      default:
        return <PeopleRegular className="w-4 h-4 text-slate-600" />;
    }
  };

  const getGroupTypeBadgeColor = (groupType: string, isExchangeDL?: boolean) => {
    if (isExchangeDL) {
      return 'bg-amber-100 text-amber-700 dark:bg-amber-900/30 dark:text-amber-400';
    }
    switch (groupType) {
      case 'Microsoft 365':
        return 'bg-blue-100 text-blue-700 dark:bg-blue-900/30 dark:text-blue-400';
      case 'Security':
        return 'bg-green-100 text-green-700 dark:bg-green-900/30 dark:text-green-400';
      case 'Mail-enabled Security':
        return 'bg-teal-100 text-teal-700 dark:bg-teal-900/30 dark:text-teal-400';
      case 'Distribution':
        return 'bg-orange-100 text-orange-700 dark:bg-orange-900/30 dark:text-orange-400';
      default:
        return 'bg-slate-100 text-slate-700 dark:bg-slate-900/30 dark:text-slate-400';
    }
  };

  const SortIcon: React.FC<{ field: SortField }> = ({ field }) => {
    if (sortField !== field) return null;
    return sortDirection === 'asc' ? (
      <ChevronUpRegular className="w-4 h-4 ml-1" />
    ) : (
      <ChevronDownRegular className="w-4 h-4 ml-1" />
    );
  };

  // Get distribution list count including Exchange DDLs
  const distributionCount = useMemo(() => {
    const graphCount = groups.filter(g => g.groupType === 'Distribution').length;
    return graphCount + exchangeDLs.length;
  }, [groups, exchangeDLs]);

  // Stat card component
  const StatCard: React.FC<{
    title: string;
    value: number;
    icon: React.ReactNode;
    color: string;
    isActive: boolean;
    onClick: () => void;
    loading?: boolean;
  }> = ({ title, value, icon, color, isActive, onClick, loading }) => (
    <button 
      type="button"
      className={`w-full bg-white dark:bg-slate-800 rounded-lg border-2 transition-all cursor-pointer hover:shadow-md p-2 text-left ${
        isActive 
          ? 'border-blue-500 ring-1 ring-blue-200 dark:ring-blue-800 shadow-md' 
          : 'border-slate-200 dark:border-slate-700 hover:border-slate-300 dark:hover:border-slate-600'
      }`}
      onClick={onClick}
    >
      <div className="flex items-center gap-1.5">
        <div className={`p-1 rounded ${color.replace('text-', 'bg-').replace('-600', '-100')} dark:bg-opacity-20 flex-shrink-0`}>
          {icon}
        </div>
        <div className="min-w-0 flex-1">
          <p className="text-[10px] text-slate-500 dark:text-slate-400 leading-tight truncate">{title}</p>
          <p className={`text-sm font-semibold leading-tight ${color}`}>
            {loading ? '...' : value}
          </p>
        </div>
        {isActive && (
          <CheckmarkCircleFilled className="w-3 h-3 text-blue-500 flex-shrink-0" />
        )}
      </div>
    </button>
  );

  return (
    <div className="p-4 space-y-4 w-full max-w-full overflow-hidden">
      {/* Header */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
        <div>
          <h1 className="text-xl font-semibold text-slate-900 dark:text-white">Teams & Groups</h1>
          <p className="text-sm text-slate-500 dark:text-slate-400 hidden sm:block">
            Manage Microsoft 365 groups, security groups, and Teams
          </p>
        </div>
        <a
          href="https://entra.microsoft.com/#view/Microsoft_AAD_IAM/GroupsManagementMenuBlade/~/AllGroups/menuId/"
          target="_blank"
          rel="noopener noreferrer"
          className="inline-flex items-center justify-center gap-2 px-3 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 transition-colors text-sm whitespace-nowrap"
        >
          <OpenRegular className="w-4 h-4" />
          <span>Entra ID</span>
        </a>
      </div>

      {/* Stats Cards - Row 1 */}
      <div className="grid grid-cols-2 sm:grid-cols-4 gap-2">
        <StatCard
          title="Total"
          value={stats?.totalGroups ?? 0}
          icon={<PeopleRegular className="w-4 h-4 text-blue-600" />}
          color="text-blue-600"
          isActive={filterType === 'all'}
          onClick={() => setFilterType('all')}
        />
        <StatCard
          title="M365 Groups"
          value={stats?.microsoft365Groups ?? 0}
          icon={<PeopleTeamRegular className="w-4 h-4 text-indigo-600" />}
          color="text-indigo-600"
          isActive={filterType === 'm365'}
          onClick={() => setFilterType('m365')}
        />
        <StatCard
          title="Security"
          value={stats?.securityGroups ?? 0}
          icon={<ShieldRegular className="w-4 h-4 text-green-600" />}
          color="text-green-600"
          isActive={filterType === 'security'}
          onClick={() => setFilterType('security')}
        />
        <StatCard
          title="Distribution"
          value={distributionCount}
          icon={<MailRegular className="w-4 h-4 text-orange-600" />}
          color="text-orange-600"
          isActive={filterType === 'distribution'}
          onClick={() => setFilterType('distribution')}
          loading={exchangeDLsLoading}
        />
      </div>

      {/* Stats Cards - Row 2 */}
      <div className="grid grid-cols-3 gap-2">
        <StatCard
          title="Teams"
          value={stats?.teamsEnabled ?? 0}
          icon={<PeopleTeamFilled className="w-4 h-4 text-purple-600" />}
          color="text-purple-600"
          isActive={filterType === 'teams'}
          onClick={() => setFilterType('teams')}
        />
        <StatCard
          title="Public"
          value={stats?.publicGroups ?? 0}
          icon={<GlobeRegular className="w-4 h-4 text-cyan-600" />}
          color="text-cyan-600"
          isActive={filterType === 'public'}
          onClick={() => setFilterType('public')}
        />
        <StatCard
          title="Private"
          value={stats?.privateGroups ?? 0}
          icon={<LockClosedRegular className="w-4 h-4 text-amber-600" />}
          color="text-amber-600"
          isActive={filterType === 'private'}
          onClick={() => setFilterType('private')}
        />
      </div>

      {/* Exchange DDL Info Banner */}
      {filterType === 'distribution' && (
        <div className="bg-blue-50 dark:bg-blue-900/20 border border-blue-200 dark:border-blue-700 rounded-lg p-3">
          <div className="flex items-start gap-2">
            <InfoRegular className="w-5 h-5 text-blue-600 dark:text-blue-400 flex-shrink-0 mt-0.5" />
            <div className="flex-1">
              <p className="text-sm text-blue-800 dark:text-blue-200">
                Showing distribution lists from both Entra ID (Azure AD groups) and Exchange Online.
                {exchangeDLsLoading && ' Loading Exchange DDLs...'}
                {exchangeDLsError && (
                  <span className="text-red-600 dark:text-red-400"> Error: {exchangeDLsError}</span>
                )}
              </p>
              {!exchangeDLsLoading && exchangeDLsLoaded && (
                <p className="text-xs text-blue-600 dark:text-blue-400 mt-1">
                  Found {exchangeDLs.length} Exchange distribution list(s)
                </p>
              )}
            </div>
            {exchangeDLsError && (
              <button
                onClick={fetchExchangeDLs}
                className="p-1.5 text-blue-600 hover:bg-blue-100 dark:hover:bg-blue-800 rounded"
                title="Retry"
              >
                <ArrowSyncRegular className="w-4 h-4" />
              </button>
            )}
          </div>
        </div>
      )}

      {/* Filters and Search */}
      <div className="bg-white dark:bg-slate-800 rounded-lg border border-slate-200 dark:border-slate-700 p-3">
        <div className="flex flex-col sm:flex-row gap-2">
          {/* Search */}
          <div className="relative flex-1">
            <SearchRegular className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-400" />
            <input
              type="text"
              placeholder="Search groups..."
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
              className="w-full pl-9 pr-3 py-2 border border-slate-200 dark:border-slate-600 rounded-lg bg-white dark:bg-slate-700 text-slate-900 dark:text-white focus:outline-none focus:ring-2 focus:ring-blue-500 text-sm"
            />
          </div>

          {/* Filter Dropdown */}
          <select
            value={filterType}
            onChange={(e) => setFilterType(e.target.value as FilterType)}
            className="px-3 py-2 border border-slate-200 dark:border-slate-600 rounded-lg bg-white dark:bg-slate-700 text-slate-900 dark:text-white focus:outline-none focus:ring-2 focus:ring-blue-500 text-sm"
          >
            <option value="all">All Groups</option>
            <option value="m365">Microsoft 365</option>
            <option value="security">Security</option>
            <option value="distribution">Distribution</option>
            <option value="teams">Teams</option>
            <option value="public">Public</option>
            <option value="private">Private</option>
          </select>
        </div>

        <div className="mt-2 text-xs text-slate-500 dark:text-slate-400">
          {filteredAndSortedGroups.length} of {filterType === 'distribution' ? distributionCount : groups.length} groups
          {filterType !== 'all' && (
            <span className="text-blue-600 dark:text-blue-400 ml-1">
              • {getFilterLabel(filterType)}
            </span>
          )}
        </div>
      </div>

      {/* Groups Table */}
      <div className="bg-white dark:bg-slate-800 rounded-lg border border-slate-200 dark:border-slate-700 overflow-hidden">
        {loading || (filterType === 'distribution' && exchangeDLsLoading && !exchangeDLsLoaded) ? (
          <div className="flex items-center justify-center h-64">
            <div className="animate-spin rounded-full h-8 w-8 border-b-2 border-blue-600"></div>
          </div>
        ) : error ? (
          <div className="flex items-center justify-center h-64 text-red-500">
            <p>{error}</p>
          </div>
        ) : (
          <>
            {/* Mobile Card View */}
            <div className="block lg:hidden divide-y divide-slate-200 dark:divide-slate-700">
              {filteredAndSortedGroups.map((group) => (
                <div key={group.id} className="p-3 hover:bg-slate-50 dark:hover:bg-slate-700/50">
                  <div className="flex items-start gap-3">
                    <div className={`w-10 h-10 rounded-full flex items-center justify-center flex-shrink-0 ${
                      group.isTeam ? 'bg-purple-100 dark:bg-purple-900' : 
                      group.isExchangeDL ? 'bg-amber-100 dark:bg-amber-900' :
                      'bg-blue-100 dark:bg-blue-900'
                    }`}>
                      {getGroupTypeIcon(group.groupType, group.isTeam, group.isExchangeDL)}
                    </div>
                    <div className="flex-1 min-w-0">
                      <div className="flex items-start justify-between gap-2">
                        <div className="min-w-0">
                          <p className="font-medium text-slate-900 dark:text-white truncate">
                            {group.displayName}
                          </p>
                          <p className="text-xs text-slate-500 dark:text-slate-400 truncate">
                            {group.mail || 'No email'}
                          </p>
                        </div>
                        <div className="flex items-center gap-1 flex-shrink-0">
                          <button
                            onClick={() => fetchGroupDetails(group.id, group.isExchangeDL, group.mail)}
                            className="p-1.5 text-slate-400 hover:text-green-600 hover:bg-green-50 dark:hover:bg-green-900/30 rounded-lg transition-colors"
                            title="View Members"
                          >
                            <PeopleRegular className="w-4 h-4" />
                          </button>
                          <button
                            onClick={() => group.isExchangeDL ? openInExchangeAdmin(group.mail || '') : openInEntraPortal(group.id)}
                            className="p-1.5 text-slate-400 hover:text-blue-600 hover:bg-blue-50 dark:hover:bg-blue-900/30 rounded-lg transition-colors"
                            title={group.isExchangeDL ? "Open in Exchange Admin" : "Open in Entra ID"}
                          >
                            <OpenRegular className="w-4 h-4" />
                          </button>
                        </div>
                      </div>
                      
                      <div className="flex flex-wrap gap-1.5 mt-2">
                        <span className={`inline-flex items-center gap-1 px-1.5 py-0.5 rounded text-xs font-medium ${getGroupTypeBadgeColor(group.groupType, group.isExchangeDL)}`}>
                          {group.isExchangeDL ? 'Exchange DDL' : group.groupType}
                        </span>
                        {group.isTeam && (
                          <span className="inline-flex items-center gap-1 px-1.5 py-0.5 rounded text-xs font-medium bg-purple-100 text-purple-700 dark:bg-purple-900/30 dark:text-purple-400">
                            <PeopleTeamFilled className="w-3 h-3" />
                            Team
                          </span>
                        )}
                        {group.visibility && (
                          <span className={`inline-flex items-center gap-1 px-1.5 py-0.5 rounded text-xs font-medium ${
                            group.visibility.toLowerCase() === 'public' 
                              ? 'bg-cyan-100 text-cyan-700 dark:bg-cyan-900/30 dark:text-cyan-400'
                              : 'bg-amber-100 text-amber-700 dark:bg-amber-900/30 dark:text-amber-400'
                          }`}>
                            {group.visibility.toLowerCase() === 'public' ? <GlobeRegular className="w-3 h-3" /> : <LockClosedRegular className="w-3 h-3" />}
                            {group.visibility}
                          </span>
                        )}
                      </div>

                      <div className="flex flex-wrap gap-x-3 gap-y-1 mt-2 text-xs text-slate-500 dark:text-slate-400">
                        <span className="flex items-center gap-1">
                          <PersonRegular className="w-3 h-3" />
                          {group.memberCount} members
                        </span>
                        {!group.isExchangeDL && (
                          <span className="flex items-center gap-1">
                            <PersonRegular className="w-3 h-3" />
                            {group.ownerCount} owners
                          </span>
                        )}
                      </div>
                    </div>
                  </div>
                </div>
              ))}
            </div>

            {/* Desktop Table View */}
            <div className="hidden lg:block overflow-x-auto">
              <table className="w-full">
                <thead className="bg-slate-50 dark:bg-slate-700">
                  <tr>
                    <th
                      className="px-4 py-3 text-left text-sm font-medium text-slate-600 dark:text-slate-300 cursor-pointer hover:bg-slate-100 dark:hover:bg-slate-600"
                      onClick={() => handleSort('displayName')}
                    >
                      <div className="flex items-center">
                        Group
                        <SortIcon field="displayName" />
                      </div>
                    </th>
                    <th
                      className="px-4 py-3 text-left text-sm font-medium text-slate-600 dark:text-slate-300 cursor-pointer hover:bg-slate-100 dark:hover:bg-slate-600"
                      onClick={() => handleSort('groupType')}
                    >
                      <div className="flex items-center">
                        Type
                        <SortIcon field="groupType" />
                      </div>
                    </th>
                    <th
                      className="px-4 py-3 text-left text-sm font-medium text-slate-600 dark:text-slate-300 cursor-pointer hover:bg-slate-100 dark:hover:bg-slate-600"
                      onClick={() => handleSort('isTeam')}
                    >
                      <div className="flex items-center">
                        Teams
                        <SortIcon field="isTeam" />
                      </div>
                    </th>
                    <th className="px-4 py-3 text-left text-sm font-medium text-slate-600 dark:text-slate-300">
                      Visibility
                    </th>
                    <th
                      className="px-4 py-3 text-left text-sm font-medium text-slate-600 dark:text-slate-300 cursor-pointer hover:bg-slate-100 dark:hover:bg-slate-600"
                      onClick={() => handleSort('memberCount')}
                    >
                      <div className="flex items-center">
                        Members
                        <SortIcon field="memberCount" />
                      </div>
                    </th>
                    <th
                      className="px-4 py-3 text-left text-sm font-medium text-slate-600 dark:text-slate-300 cursor-pointer hover:bg-slate-100 dark:hover:bg-slate-600"
                      onClick={() => handleSort('createdDateTime')}
                    >
                      <div className="flex items-center">
                        Created
                        <SortIcon field="createdDateTime" />
                      </div>
                    </th>
                    <th className="px-4 py-3 text-right text-sm font-medium text-slate-600 dark:text-slate-300">
                      Actions
                    </th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-200 dark:divide-slate-700">
                  {filteredAndSortedGroups.map((group) => (
                    <tr
                      key={group.id}
                      className="hover:bg-slate-50 dark:hover:bg-slate-700/50 transition-colors"
                    >
                      <td className="px-4 py-3">
                        <div className="flex items-center gap-3">
                          <div className={`w-10 h-10 rounded-full flex items-center justify-center flex-shrink-0 ${
                            group.isTeam ? 'bg-purple-100 dark:bg-purple-900' : 
                            group.isExchangeDL ? 'bg-amber-100 dark:bg-amber-900' :
                            'bg-blue-100 dark:bg-blue-900'
                          }`}>
                            {getGroupTypeIcon(group.groupType, group.isTeam, group.isExchangeDL)}
                          </div>
                          <div>
                            <p className="font-medium text-slate-900 dark:text-white">
                              {group.displayName}
                            </p>
                            <p className="text-sm text-slate-500 dark:text-slate-400 truncate max-w-xs">
                              {group.mail || group.description || 'No description'}
                            </p>
                          </div>
                        </div>
                      </td>
                      <td className="px-4 py-3">
                        <span className={`inline-flex items-center gap-1 px-2 py-1 rounded-full text-xs font-medium ${getGroupTypeBadgeColor(group.groupType, group.isExchangeDL)}`}>
                          {group.isExchangeDL ? 'Exchange DDL' : group.groupType}
                        </span>
                      </td>
                      <td className="px-4 py-3">
                        {group.isTeam ? (
                          <span className="inline-flex items-center gap-1 px-2 py-1 rounded-full text-xs font-medium bg-purple-100 text-purple-700 dark:bg-purple-900/30 dark:text-purple-400">
                            <PeopleTeamFilled className="w-3 h-3" />
                            Yes
                          </span>
                        ) : (
                          <span className="text-sm text-slate-400">-</span>
                        )}
                      </td>
                      <td className="px-4 py-3">
                        {group.visibility ? (
                          <span className={`inline-flex items-center gap-1 px-2 py-1 rounded-full text-xs font-medium ${
                            group.visibility.toLowerCase() === 'public' 
                              ? 'bg-cyan-100 text-cyan-700 dark:bg-cyan-900/30 dark:text-cyan-400'
                              : 'bg-amber-100 text-amber-700 dark:bg-amber-900/30 dark:text-amber-400'
                          }`}>
                            {group.visibility.toLowerCase() === 'public' ? <GlobeRegular className="w-3 h-3" /> : <LockClosedRegular className="w-3 h-3" />}
                            {group.visibility}
                          </span>
                        ) : (
                          <span className="text-sm text-slate-400">-</span>
                        )}
                      </td>
                      <td className="px-4 py-3">
                        <div className="text-sm">
                          {group.isExchangeDL && group.groupType === 'Dynamic Distribution' ? (
                            <span className="text-slate-400 italic">Dynamic</span>
                          ) : (
                            <>
                              <span className="text-slate-900 dark:text-white">{group.memberCount}</span>
                              {!group.isExchangeDL && (
                                <span className="text-slate-400 ml-1">/ {group.ownerCount} owners</span>
                              )}
                            </>
                          )}
                        </div>
                      </td>
                      <td className="px-4 py-3">
                        <span className="text-sm text-slate-600 dark:text-slate-300">
                          {formatDate(group.createdDateTime)}
                        </span>
                      </td>
                      <td className="px-4 py-3 text-right">
                        <div className="flex items-center justify-end gap-1">
                          {/* View Members Button */}
                          <button
                            onClick={() => fetchGroupDetails(group.id, group.isExchangeDL, group.mail)}
                            className="p-2 text-slate-400 hover:text-green-600 hover:bg-green-50 dark:hover:bg-green-900/30 rounded-lg transition-colors"
                            title="View Members"
                          >
                            <PeopleRegular className="w-5 h-5" />
                          </button>
                          {group.isTeam && group.teamWebUrl && (
                            <button
                              onClick={() => openInTeams(group.teamWebUrl!)}
                              className="p-2 text-purple-500 hover:text-purple-600 hover:bg-purple-50 dark:hover:bg-purple-900/30 rounded-lg transition-colors"
                              title="Open in Teams"
                            >
                              <PeopleTeamFilled className="w-5 h-5" />
                            </button>
                          )}
                          <button
                            onClick={() => group.isExchangeDL ? openInExchangeAdmin(group.mail || '') : openInEntraPortal(group.id)}
                            className="p-2 text-slate-400 hover:text-blue-600 hover:bg-blue-50 dark:hover:bg-blue-900/30 rounded-lg transition-colors"
                            title={group.isExchangeDL ? "Open in Exchange Admin" : "Open in Entra ID"}
                          >
                            <OpenRegular className="w-5 h-5" />
                          </button>
                        </div>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </>
        )}
      </div>

      {/* Member Panel Slide-out */}
      {memberPanelOpen && (
        <div className="fixed inset-0 z-50 overflow-hidden">
          {/* Backdrop */}
          <div 
            className="absolute inset-0 bg-black/50 transition-opacity"
            onClick={closeMemberPanel}
          />
          
          {/* Panel */}
          <div className="absolute inset-y-0 right-0 flex max-w-full">
            <div className="w-screen max-w-lg">
              <div className="flex h-full flex-col bg-white dark:bg-slate-800 shadow-xl">
                {/* Header */}
                <div className="px-4 py-4 border-b border-slate-200 dark:border-slate-700">
                  <div className="flex items-start justify-between">
                    <div className="min-w-0 flex-1">
                      <h2 className="text-lg font-semibold text-slate-900 dark:text-white truncate">
                        {selectedExchangeDL?.displayName || selectedGroup?.displayName}
                      </h2>
                      <p className="text-sm text-slate-500 dark:text-slate-400 truncate">
                        {selectedExchangeDL?.primarySmtpAddress || selectedGroup?.mail || 'No email address'}
                      </p>
                      <div className="flex flex-wrap gap-2 mt-2">
                        <span className={`inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-xs font-medium ${
                          selectedExchangeDL 
                            ? 'bg-amber-100 text-amber-700 dark:bg-amber-900/30 dark:text-amber-400'
                            : selectedGroup?.groupType === 'Distribution' 
                            ? 'bg-orange-100 text-orange-700 dark:bg-orange-900/30 dark:text-orange-400'
                            : selectedGroup?.groupType === 'Security'
                            ? 'bg-green-100 text-green-700 dark:bg-green-900/30 dark:text-green-400'
                            : 'bg-blue-100 text-blue-700 dark:bg-blue-900/30 dark:text-blue-400'
                        }`}>
                          {selectedExchangeDL ? 'Exchange DDL' : selectedGroup?.groupType}
                        </span>
                        <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-xs font-medium bg-slate-100 text-slate-700 dark:bg-slate-700 dark:text-slate-300">
                          <PeopleRegular className="w-3 h-3" />
                          {selectedExchangeDL?.members?.length ?? selectedExchangeDL?.memberCount ?? selectedGroup?.members?.length ?? 0} members
                        </span>
                        {selectedGroup?.owners && selectedGroup.owners.length > 0 && (
                          <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-xs font-medium bg-purple-100 text-purple-700 dark:bg-purple-900/30 dark:text-purple-400">
                            <StarRegular className="w-3 h-3" />
                            {selectedGroup.owners.length} owners
                          </span>
                        )}
                        {selectedExchangeDL?.managedBy && selectedExchangeDL.managedBy.length > 0 && (
                          <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-xs font-medium bg-purple-100 text-purple-700 dark:bg-purple-900/30 dark:text-purple-400">
                            <StarRegular className="w-3 h-3" />
                            {selectedExchangeDL.managedBy.length} managed by
                          </span>
                        )}
                      </div>
                    </div>
                    <button
                      onClick={closeMemberPanel}
                      className="p-2 text-slate-400 hover:text-slate-600 dark:hover:text-slate-300 rounded-lg hover:bg-slate-100 dark:hover:bg-slate-700"
                    >
                      <DismissRegular className="w-5 h-5" />
                    </button>
                  </div>
                </div>

                {/* Search and Export */}
                <div className="px-4 py-3 border-b border-slate-200 dark:border-slate-700">
                  <div className="flex gap-2">
                    <div className="relative flex-1">
                      <SearchRegular className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-400" />
                      <input
                        type="text"
                        placeholder="Search members..."
                        value={memberSearchTerm}
                        onChange={(e) => setMemberSearchTerm(e.target.value)}
                        className="w-full pl-9 pr-3 py-2 border border-slate-200 dark:border-slate-600 rounded-lg bg-white dark:bg-slate-700 text-slate-900 dark:text-white focus:outline-none focus:ring-2 focus:ring-blue-500 text-sm"
                      />
                    </div>
                    <button
                      onClick={exportMembersToCsv}
                      disabled={!filteredMembers.length}
                      className="inline-flex items-center gap-1.5 px-3 py-2 text-sm bg-slate-100 dark:bg-slate-700 text-slate-700 dark:text-slate-300 rounded-lg hover:bg-slate-200 dark:hover:bg-slate-600 disabled:opacity-50 disabled:cursor-not-allowed"
                      title="Export to CSV"
                    >
                      <ArrowDownloadRegular className="w-4 h-4" />
                      CSV
                    </button>
                  </div>
                  <p className="text-xs text-slate-500 dark:text-slate-400 mt-2">
                    {filteredMembers.length} of {selectedExchangeDL?.members?.length ?? selectedExchangeDL?.memberCount ?? selectedGroup?.members?.length ?? 0} members
                    {selectedExchangeDL?.isDynamic && ' (calculated from filter)'}
                  </p>
                </div>

                {/* Owners Section (if any) */}
                {selectedGroup?.owners && selectedGroup.owners.length > 0 && (
                  <div className="px-4 py-3 bg-purple-50 dark:bg-purple-900/20 border-b border-slate-200 dark:border-slate-700">
                    <h3 className="text-sm font-medium text-purple-700 dark:text-purple-400 mb-2 flex items-center gap-1">
                      <StarRegular className="w-4 h-4" />
                      Owners
                    </h3>
                    <div className="flex flex-wrap gap-2">
                      {selectedGroup.owners.map((owner) => (
                        <span 
                          key={owner.id}
                          className="inline-flex items-center gap-1.5 px-2 py-1 rounded-full text-xs font-medium bg-white dark:bg-slate-700 text-slate-700 dark:text-slate-300 border border-purple-200 dark:border-purple-700"
                          title={owner.userPrincipalName || owner.mail || ''}
                        >
                          <PersonRegular className="w-3 h-3 text-purple-600 dark:text-purple-400" />
                          {owner.displayName}
                        </span>
                      ))}
                    </div>
                  </div>
                )}

                {/* Managed By Section for Exchange DDLs */}
                {selectedExchangeDL?.managedBy && selectedExchangeDL.managedBy.length > 0 && (
                  <div className="px-4 py-3 bg-purple-50 dark:bg-purple-900/20 border-b border-slate-200 dark:border-slate-700">
                    <h3 className="text-sm font-medium text-purple-700 dark:text-purple-400 mb-2 flex items-center gap-1">
                      <StarRegular className="w-4 h-4" />
                      Managed By
                    </h3>
                    <div className="flex flex-wrap gap-2">
                      {selectedExchangeDL.managedBy.map((manager, idx) => (
                        <span 
                          key={idx}
                          className="inline-flex items-center gap-1.5 px-2 py-1 rounded-full text-xs font-medium bg-white dark:bg-slate-700 text-slate-700 dark:text-slate-300 border border-purple-200 dark:border-purple-700"
                        >
                          <PersonRegular className="w-3 h-3 text-purple-600 dark:text-purple-400" />
                          {manager}
                        </span>
                      ))}
                    </div>
                  </div>
                )}

                {/* Members List */}
                <div className="flex-1 overflow-y-auto">
                  {memberLoading ? (
                    <div className="flex items-center justify-center h-48">
                      <div className="animate-spin rounded-full h-8 w-8 border-b-2 border-blue-600"></div>
                    </div>
                  ) : filteredMembers.length > 0 ? (
                    <div className="divide-y divide-slate-200 dark:divide-slate-700">
                      {filteredMembers.map((member) => {
                        const isExchangeMember = selectedExchangeDL !== null;
                        const displayName = member.displayName;
                        const email = isExchangeMember 
                          ? (member as ExchangeDLMember).primarySmtpAddress 
                          : (member as GroupMember).mail || (member as GroupMember).userPrincipalName;
                        const memberType = isExchangeMember 
                          ? (member as ExchangeDLMember).recipientType 
                          : (member as GroupMember).memberType;
                        
                        return (
                          <div 
                            key={member.id} 
                            className="px-4 py-3 hover:bg-slate-50 dark:hover:bg-slate-700/50"
                          >
                            <div className="flex items-center gap-3">
                              <div className={`w-10 h-10 rounded-full flex items-center justify-center flex-shrink-0 ${
                                memberType === 'User' || memberType === 'UserMailbox' || memberType === 'MailUser'
                                  ? 'bg-blue-100 dark:bg-blue-900/30'
                                  : 'bg-slate-100 dark:bg-slate-700'
                              }`}>
                                {memberType === 'User' || memberType === 'UserMailbox' || memberType === 'MailUser' ? (
                                  <PersonRegular className="w-5 h-5 text-blue-600 dark:text-blue-400" />
                                ) : (
                                  <PeopleCommunityRegular className="w-5 h-5 text-slate-600 dark:text-slate-400" />
                                )}
                              </div>
                              <div className="min-w-0 flex-1">
                                <p className="font-medium text-slate-900 dark:text-white truncate">
                                  {displayName}
                                </p>
                                <p className="text-sm text-slate-500 dark:text-slate-400 truncate">
                                  {email || 'No email'}
                                </p>
                              </div>
                              <span className={`flex-shrink-0 inline-flex items-center px-2 py-0.5 rounded text-xs font-medium ${
                                memberType === 'User' || memberType === 'UserMailbox' || memberType === 'MailUser'
                                  ? 'bg-blue-100 text-blue-700 dark:bg-blue-900/30 dark:text-blue-400'
                                  : 'bg-slate-100 text-slate-600 dark:bg-slate-700 dark:text-slate-400'
                              }`}>
                                {memberType}
                              </span>
                            </div>
                          </div>
                        );
                      })}
                    </div>
                  ) : (
                    <div className="flex flex-col items-center justify-center h-48 text-slate-500">
                      <PeopleRegular className="w-12 h-12 text-slate-300 dark:text-slate-600 mb-2" />
                      <p className="text-sm">
                        {memberSearchTerm ? 'No members match your search' : 'No members in this group'}
                      </p>
                    </div>
                  )}
                </div>

                {/* Footer */}
                <div className="px-4 py-3 border-t border-slate-200 dark:border-slate-700 bg-slate-50 dark:bg-slate-800">
                  <button
                    onClick={() => {
                      if (selectedExchangeDL) {
                        openInExchangeAdmin(selectedExchangeDL.primarySmtpAddress || '');
                      } else if (selectedGroup) {
                        openInEntraPortal(selectedGroup.id);
                      }
                    }}
                    className="w-full inline-flex items-center justify-center gap-2 px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 transition-colors text-sm"
                  >
                    <OpenRegular className="w-4 h-4" />
                    {selectedExchangeDL ? 'Manage in Exchange Admin' : 'Manage in Entra ID'}
                  </button>
                </div>
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

export default TeamsGroupsPage;
