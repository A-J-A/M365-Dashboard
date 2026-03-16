import React, { useState, useEffect, useMemo } from 'react';
import {
  PersonRegular,
  PersonFilled,
  CheckmarkCircleFilled,
  DismissCircleFilled,
  GuestRegular,
  PersonAccountsRegular,
  SearchRegular,
  OpenRegular,
  ChevronUpRegular,
  ChevronDownRegular,
  CalendarRegular,
  ShieldCheckmarkRegular,
  LockClosedRegular,
  ShieldErrorRegular,
  LocationRegular,
  WarningRegular,
} from '@fluentui/react-icons';
import { useAppContext } from '../contexts/AppContext';

interface TenantUser {
  id: string;
  displayName: string;
  userPrincipalName: string;
  mail: string | null;
  userType: string;
  accountEnabled: boolean;
  createdDateTime: string | null;
  lastSignInDateTime: string | null;
  lastNonInteractiveSignInDateTime: string | null;
  jobTitle: string | null;
  department: string | null;
  officeLocation: string | null;
  city: string | null;
  country: string | null;
  mobilePhone: string | null;
  businessPhones: string | null;
  assignedLicenses: string[] | null;
  hasMailbox: boolean;
  managerDisplayName: string | null;
  profilePhoto: string | null;
  isMfaRegistered: boolean;
  isMfaCapable: boolean;
  defaultMfaMethod: string | null;
  mfaMethods: string[] | null;
  usageLocation: string | null;
}

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

interface UserListResult {
  users: TenantUser[];
  totalCount: number;
  filteredCount: number;
  nextLink: string | null;
}

type FilterType = 'all' | 'members' | 'guests' | 'enabled' | 'disabled' | 'licensed' | 'unlicensed' | 'active' | 'mfaEnabled' | 'mfaDisabled' | 'noUsageLocation';
type SortField = 'displayName' | 'userPrincipalName' | 'userType' | 'accountEnabled' | 'lastSignInDateTime' | 'createdDateTime' | 'isMfaRegistered';
type SortDirection = 'asc' | 'desc';

const UsersPage: React.FC = () => {
  const { getAccessToken } = useAppContext();
  const [users, setUsers] = useState<TenantUser[]>([]);
  const [stats, setStats] = useState<UserStats | null>(null);
  const [loading, setLoading] = useState(true);
  const [statsLoading, setStatsLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [searchTerm, setSearchTerm] = useState('');
  const [filterType, setFilterType] = useState<FilterType>('all');
  const [sortField, setSortField] = useState<SortField>('displayName');
  const [sortDirection, setSortDirection] = useState<SortDirection>('asc');

  const tenantId = import.meta.env.VITE_AZURE_TENANT_ID;

  useEffect(() => {
    fetchUsers();
    fetchStats();
  }, []);

  const fetchUsers = async () => {
    try {
      setLoading(true);
      const token = await getAccessToken();
      const response = await fetch('/api/users?take=500', {
        headers: {
          'Authorization': `Bearer ${token}`,
          'Content-Type': 'application/json',
        },
      });

      if (!response.ok) {
        throw new Error('Failed to fetch users');
      }

      const data: UserListResult = await response.json();
      setUsers(data.users);
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
      const response = await fetch('/api/users/stats', {
        headers: {
          'Authorization': `Bearer ${token}`,
          'Content-Type': 'application/json',
        },
      });

      if (!response.ok) {
        throw new Error('Failed to fetch user stats');
      }

      const data: UserStats = await response.json();
      setStats(data);
    } catch (err) {
      console.error('Failed to fetch stats:', err);
    } finally {
      setStatsLoading(false);
    }
  };

  const mfaStats = useMemo(() => {
    const mfaEnabled = users.filter(u => u.isMfaRegistered).length;
    const mfaDisabled = users.filter(u => !u.isMfaRegistered).length;
    return { mfaEnabled, mfaDisabled };
  }, [users]);

  const noUsageLocationCount = useMemo(() => {
    return users.filter(u => !u.usageLocation && u.userType !== 'Guest').length;
  }, [users]);

  const filteredAndSortedUsers = useMemo(() => {
    let result = [...users];

    if (searchTerm) {
      const term = searchTerm.toLowerCase();
      result = result.filter(
        (user) =>
          user.displayName.toLowerCase().includes(term) ||
          user.userPrincipalName.toLowerCase().includes(term) ||
          (user.mail?.toLowerCase().includes(term) ?? false) ||
          (user.department?.toLowerCase().includes(term) ?? false) ||
          (user.jobTitle?.toLowerCase().includes(term) ?? false)
      );
    }

    switch (filterType) {
      case 'members':
        result = result.filter((user) => user.userType === 'Member');
        break;
      case 'guests':
        result = result.filter((user) => user.userType === 'Guest');
        break;
      case 'enabled':
        result = result.filter((user) => user.accountEnabled);
        break;
      case 'disabled':
        result = result.filter((user) => !user.accountEnabled);
        break;
      case 'licensed':
        result = result.filter((user) => user.assignedLicenses && user.assignedLicenses.length > 0);
        break;
      case 'unlicensed':
        result = result.filter((user) => !user.assignedLicenses || user.assignedLicenses.length === 0);
        break;
      case 'active':
        const thirtyDaysAgo = new Date();
        thirtyDaysAgo.setDate(thirtyDaysAgo.getDate() - 30);
        result = result.filter((user) => {
          const lastSignIn = user.lastSignInDateTime ? new Date(user.lastSignInDateTime) : null;
          const lastNonInteractive = user.lastNonInteractiveSignInDateTime ? new Date(user.lastNonInteractiveSignInDateTime) : null;
          return (lastSignIn && lastSignIn > thirtyDaysAgo) || (lastNonInteractive && lastNonInteractive > thirtyDaysAgo);
        });
        break;
      case 'mfaEnabled':
        result = result.filter((user) => user.isMfaRegistered);
        break;
      case 'mfaDisabled':
        result = result.filter((user) => !user.isMfaRegistered);
        break;
      case 'noUsageLocation':
        result = result.filter((user) => !user.usageLocation && user.userType !== 'Guest');
        break;
    }

    result.sort((a, b) => {
      let aValue: string | boolean | null = null;
      let bValue: string | boolean | null = null;

      switch (sortField) {
        case 'displayName':
          aValue = a.displayName;
          bValue = b.displayName;
          break;
        case 'userPrincipalName':
          aValue = a.userPrincipalName;
          bValue = b.userPrincipalName;
          break;
        case 'userType':
          aValue = a.userType;
          bValue = b.userType;
          break;
        case 'accountEnabled':
          aValue = a.accountEnabled;
          bValue = b.accountEnabled;
          break;
        case 'lastSignInDateTime':
          aValue = a.lastSignInDateTime;
          bValue = b.lastSignInDateTime;
          break;
        case 'createdDateTime':
          aValue = a.createdDateTime;
          bValue = b.createdDateTime;
          break;
        case 'isMfaRegistered':
          aValue = a.isMfaRegistered;
          bValue = b.isMfaRegistered;
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

      const comparison = String(aValue).localeCompare(String(bValue));
      return sortDirection === 'asc' ? comparison : -comparison;
    });

    return result;
  }, [users, searchTerm, filterType, sortField, sortDirection]);

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
    const now = new Date();
    const diffMs = now.getTime() - date.getTime();
    const diffDays = Math.floor(diffMs / (1000 * 60 * 60 * 24));

    if (diffDays === 0) return 'Today';
    if (diffDays === 1) return 'Yesterday';
    if (diffDays < 7) return `${diffDays}d ago`;
    if (diffDays < 30) return `${Math.floor(diffDays / 7)}w ago`;
    if (diffDays < 365) return `${Math.floor(diffDays / 30)}mo ago`;
    return date.toLocaleDateString();
  };

  const openInEntraPortal = (userId: string) => {
    const url = filterType === 'noUsageLocation'
      ? `https://entra.microsoft.com/#view/Microsoft_AAD_UsersAndTenants/UserProfileMenuBlade/~/Properties/userId/${userId}/tenantId/${tenantId}`
      : `https://entra.microsoft.com/#view/Microsoft_AAD_UsersAndTenants/UserProfileMenuBlade/~/overview/userId/${userId}/tenantId/${tenantId}`;
    window.open(url, '_blank');
  };

  const getFilterLabel = (filter: FilterType): string => {
    switch (filter) {
      case 'all': return 'All Users';
      case 'members': return 'Members';
      case 'guests': return 'Guests';
      case 'enabled': return 'Enabled';
      case 'disabled': return 'Disabled';
      case 'licensed': return 'Licensed';
      case 'unlicensed': return 'Unlicensed';
      case 'active': return 'Active (30d)';
      case 'mfaEnabled': return 'MFA On';
      case 'mfaDisabled': return 'MFA Off';
      case 'noUsageLocation': return 'No Usage Location';
      default: return 'All Users';
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

  // Compact stat card
  const StatCard: React.FC<{
    title: string;
    value: number;
    icon: React.ReactNode;
    color: string;
    isActive: boolean;
    onClick: () => void;
  }> = ({ title, value, icon, color, isActive, onClick }) => (
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
          <p className={`text-sm font-semibold leading-tight ${color}`}>{value}</p>
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
          <h1 className="text-xl font-semibold text-slate-900 dark:text-white">Users</h1>
          <p className="text-sm text-slate-500 dark:text-slate-400 hidden sm:block">
            Manage and view all users in your Microsoft 365 tenant
          </p>
        </div>
        <a
          href="https://entra.microsoft.com/#view/Microsoft_AAD_UsersAndTenants/UserManagementMenuBlade/~/AllUsers/menuId/"
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
          value={stats?.totalUsers ?? 0}
          icon={<PersonAccountsRegular className="w-4 h-4 text-blue-600" />}
          color="text-blue-600"
          isActive={filterType === 'all'}
          onClick={() => setFilterType('all')}
        />
        <StatCard
          title="Members"
          value={stats?.memberUsers ?? 0}
          icon={<PersonFilled className="w-4 h-4 text-green-600" />}
          color="text-green-600"
          isActive={filterType === 'members'}
          onClick={() => setFilterType('members')}
        />
        <StatCard
          title="Guests"
          value={stats?.guestUsers ?? 0}
          icon={<GuestRegular className="w-4 h-4 text-purple-600" />}
          color="text-purple-600"
          isActive={filterType === 'guests'}
          onClick={() => setFilterType('guests')}
        />
        <StatCard
          title="Licensed"
          value={stats?.licensedUsers ?? 0}
          icon={<ShieldCheckmarkRegular className="w-4 h-4 text-cyan-600" />}
          color="text-cyan-600"
          isActive={filterType === 'licensed'}
          onClick={() => setFilterType('licensed')}
        />
      </div>

      {/* Stats Cards - Row 2 */}
      <div className="grid grid-cols-4 gap-2">
        <StatCard
          title="Active"
          value={stats?.usersSignedInLast30Days ?? 0}
          icon={<CalendarRegular className="w-4 h-4 text-amber-600" />}
          color="text-amber-600"
          isActive={filterType === 'active'}
          onClick={() => setFilterType('active')}
        />
        <StatCard
          title="MFA On"
          value={loading ? 0 : mfaStats.mfaEnabled}
          icon={<ShieldCheckmarkRegular className="w-4 h-4 text-emerald-600" />}
          color="text-emerald-600"
          isActive={filterType === 'mfaEnabled'}
          onClick={() => setFilterType('mfaEnabled')}
        />
        <StatCard
          title="MFA Off"
          value={loading ? 0 : mfaStats.mfaDisabled}
          icon={<ShieldErrorRegular className="w-4 h-4 text-red-600" />}
          color="text-red-600"
          isActive={filterType === 'mfaDisabled'}
          onClick={() => setFilterType('mfaDisabled')}
        />
        <StatCard
          title="No Location"
          value={loading ? 0 : noUsageLocationCount}
          icon={<LocationRegular className="w-4 h-4 text-orange-600" />}
          color={noUsageLocationCount > 0 ? 'text-orange-600' : 'text-slate-400'}
          isActive={filterType === 'noUsageLocation'}
          onClick={() => setFilterType('noUsageLocation')}
        />
      </div>

      {/* Filters and Search */}
      <div className="bg-white dark:bg-slate-800 rounded-lg border border-slate-200 dark:border-slate-700 p-3">
        <div className="flex flex-col sm:flex-row gap-2">
          {/* Search */}
          <div className="relative flex-1">
            <SearchRegular className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-400" />
            <input
              type="text"
              placeholder="Search users..."
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
            <option value="all">All Users</option>
            <option value="members">Members</option>
            <option value="guests">Guests</option>
            <option value="enabled">Enabled</option>
            <option value="disabled">Disabled</option>
            <option value="licensed">Licensed</option>
            <option value="unlicensed">Unlicensed</option>
            <option value="active">Active (30d)</option>
            <option value="mfaEnabled">MFA On</option>
            <option value="mfaDisabled">MFA Off</option>
            <option value="noUsageLocation">No Usage Location</option>
          </select>
        </div>

        <div className="mt-2 text-xs text-slate-500 dark:text-slate-400">
          {filteredAndSortedUsers.length} of {users.length} users
          {filterType !== 'all' && (
            <span className="text-blue-600 dark:text-blue-400 ml-1">
              • {getFilterLabel(filterType)}
            </span>
          )}
        </div>
      </div>

      {/* No Usage Location Warning Banner */}
      {filterType === 'noUsageLocation' && (
        <div className="bg-orange-50 dark:bg-orange-900/20 border border-orange-200 dark:border-orange-800 rounded-lg p-3 flex items-start gap-3">
          <WarningRegular className="w-5 h-5 text-orange-600 dark:text-orange-400 flex-shrink-0 mt-0.5" />
          <div>
            <p className="text-sm font-medium text-orange-800 dark:text-orange-200">
              Usage location required for Microsoft 365 licence assignment
            </p>
            <p className="text-xs text-orange-700 dark:text-orange-300 mt-0.5">
              Showing all non-guest accounts with no usage location set.
              A usage location must be set before Microsoft 365 licences can be assigned to a user.
              To fix, open the user in the{' '}
              <a
                href={`https://entra.microsoft.com/#view/Microsoft_AAD_UsersAndTenants/UserManagementMenuBlade/~/AllUsers`}
                target="_blank"
                rel="noopener noreferrer"
                className="underline font-medium hover:text-orange-900 dark:hover:text-orange-200"
              >
                Entra admin centre
              </a>
              {' '}and set Properties → Usage location.
            </p>
          </div>
        </div>
      )}

      {/* Users Table / Cards */}
      <div className="bg-white dark:bg-slate-800 rounded-lg border border-slate-200 dark:border-slate-700 overflow-hidden">
        {loading ? (
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
              {filteredAndSortedUsers.map((user) => (
                <div key={user.id} className="p-3 hover:bg-slate-50 dark:hover:bg-slate-700/50">
                  <div className="flex items-start gap-3">
                    <div className="w-10 h-10 rounded-full bg-blue-100 dark:bg-blue-900 flex items-center justify-center text-blue-600 dark:text-blue-300 font-medium flex-shrink-0">
                      {user.displayName.charAt(0).toUpperCase()}
                    </div>
                    <div className="flex-1 min-w-0">
                      <div className="flex items-start justify-between gap-2">
                        <div className="min-w-0">
                          <p className="font-medium text-slate-900 dark:text-white truncate">
                            {user.displayName}
                          </p>
                          <p className="text-xs text-slate-500 dark:text-slate-400 truncate">
                            {user.userPrincipalName}
                          </p>
                        </div>
                        <button
                          onClick={() => openInEntraPortal(user.id)}
                          className="p-1.5 text-slate-400 hover:text-blue-600 hover:bg-blue-50 dark:hover:bg-blue-900/30 rounded-lg transition-colors flex-shrink-0"
                          title="Open in Entra ID"
                        >
                          <OpenRegular className="w-4 h-4" />
                        </button>
                      </div>
                      
                      <div className="flex flex-wrap gap-1.5 mt-2">
                        <span className={`inline-flex items-center gap-1 px-1.5 py-0.5 rounded text-xs font-medium ${
                          user.userType === 'Member'
                            ? 'bg-green-100 text-green-700 dark:bg-green-900/30 dark:text-green-400'
                            : 'bg-purple-100 text-purple-700 dark:bg-purple-900/30 dark:text-purple-400'
                        }`}>
                          {user.userType}
                        </span>
                        <span className={`inline-flex items-center gap-1 px-1.5 py-0.5 rounded text-xs font-medium ${
                          user.accountEnabled
                            ? 'bg-green-100 text-green-700 dark:bg-green-900/30 dark:text-green-400'
                            : 'bg-red-100 text-red-700 dark:bg-red-900/30 dark:text-red-400'
                        }`}>
                          {user.accountEnabled ? 'Enabled' : 'Disabled'}
                        </span>
                        <span className={`inline-flex items-center gap-1 px-1.5 py-0.5 rounded text-xs font-medium ${
                          user.isMfaRegistered
                            ? 'bg-emerald-100 text-emerald-700 dark:bg-emerald-900/30 dark:text-emerald-400'
                            : 'bg-red-100 text-red-700 dark:bg-red-900/30 dark:text-red-400'
                        }`}>
                          MFA {user.isMfaRegistered ? 'On' : 'Off'}
                        </span>
                      </div>

                      <div className="flex flex-wrap gap-x-3 gap-y-1 mt-2 text-xs text-slate-500 dark:text-slate-400">
                        {user.department && <span>{user.department}</span>}
                        <span>Sign-in: {formatDate(user.lastSignInDateTime)}</span>
                        {user.defaultMfaMethod && <span>MFA: {user.defaultMfaMethod}</span>}
                      </div>

                      {user.assignedLicenses && user.assignedLicenses.length > 0 && (
                        <div className="flex flex-wrap gap-1 mt-2">
                          {user.assignedLicenses.slice(0, 2).map((license, idx) => (
                            <span
                              key={idx}
                              className="px-1.5 py-0.5 bg-blue-100 text-blue-700 dark:bg-blue-900/30 dark:text-blue-400 rounded text-xs"
                            >
                              {license.length > 15 ? license.substring(0, 15) + '...' : license}
                            </span>
                          ))}
                          {user.assignedLicenses.length > 2 && (
                            <span className="px-1.5 py-0.5 bg-slate-100 text-slate-600 dark:bg-slate-700 dark:text-slate-400 rounded text-xs">
                              +{user.assignedLicenses.length - 2}
                            </span>
                          )}
                        </div>
                      )}
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
                        User
                        <SortIcon field="displayName" />
                      </div>
                    </th>
                    <th
                      className="px-4 py-3 text-left text-sm font-medium text-slate-600 dark:text-slate-300 cursor-pointer hover:bg-slate-100 dark:hover:bg-slate-600"
                      onClick={() => handleSort('userType')}
                    >
                      <div className="flex items-center">
                        Type
                        <SortIcon field="userType" />
                      </div>
                    </th>
                    <th
                      className="px-4 py-3 text-left text-sm font-medium text-slate-600 dark:text-slate-300 cursor-pointer hover:bg-slate-100 dark:hover:bg-slate-600"
                      onClick={() => handleSort('accountEnabled')}
                    >
                      <div className="flex items-center">
                        Status
                        <SortIcon field="accountEnabled" />
                      </div>
                    </th>
                    <th
                      className="px-4 py-3 text-left text-sm font-medium text-slate-600 dark:text-slate-300 cursor-pointer hover:bg-slate-100 dark:hover:bg-slate-600"
                      onClick={() => handleSort('isMfaRegistered')}
                    >
                      <div className="flex items-center">
                        MFA
                        <SortIcon field="isMfaRegistered" />
                      </div>
                    </th>
                    <th className="px-4 py-3 text-left text-sm font-medium text-slate-600 dark:text-slate-300">
                      Department
                    </th>
                    <th
                      className="px-4 py-3 text-left text-sm font-medium text-slate-600 dark:text-slate-300 cursor-pointer hover:bg-slate-100 dark:hover:bg-slate-600"
                      onClick={() => handleSort('lastSignInDateTime')}
                    >
                      <div className="flex items-center">
                        Last Sign-in
                        <SortIcon field="lastSignInDateTime" />
                      </div>
                    </th>
                    <th className="px-4 py-3 text-left text-sm font-medium text-slate-600 dark:text-slate-300">
                      Licenses
                    </th>
                    {filterType === 'noUsageLocation' && (
                      <th className="px-4 py-3 text-left text-sm font-medium text-orange-600 dark:text-orange-400">
                        Usage Location
                      </th>
                    )}
                    <th className="px-4 py-3 text-right text-sm font-medium text-slate-600 dark:text-slate-300">
                      Actions
                    </th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-200 dark:divide-slate-700">
                  {filteredAndSortedUsers.map((user) => (
                    <tr
                      key={user.id}
                      className="hover:bg-slate-50 dark:hover:bg-slate-700/50 transition-colors"
                    >
                      <td className="px-4 py-3">
                        <div className="flex items-center gap-3">
                          <div className="w-10 h-10 rounded-full bg-blue-100 dark:bg-blue-900 flex items-center justify-center text-blue-600 dark:text-blue-300 font-medium">
                            {user.displayName.charAt(0).toUpperCase()}
                          </div>
                          <div>
                            <p className="font-medium text-slate-900 dark:text-white">
                              {user.displayName}
                            </p>
                            <p className="text-sm text-slate-500 dark:text-slate-400">
                              {user.userPrincipalName}
                            </p>
                          </div>
                        </div>
                      </td>
                      <td className="px-4 py-3">
                        <span
                          className={`inline-flex items-center gap-1 px-2 py-1 rounded-full text-xs font-medium ${
                            user.userType === 'Member'
                              ? 'bg-green-100 text-green-700 dark:bg-green-900/30 dark:text-green-400'
                              : 'bg-purple-100 text-purple-700 dark:bg-purple-900/30 dark:text-purple-400'
                          }`}
                        >
                          {user.userType === 'Member' ? (
                            <PersonRegular className="w-3 h-3" />
                          ) : (
                            <GuestRegular className="w-3 h-3" />
                          )}
                          {user.userType}
                        </span>
                      </td>
                      <td className="px-4 py-3">
                        <span
                          className={`inline-flex items-center gap-1 px-2 py-1 rounded-full text-xs font-medium ${
                            user.accountEnabled
                              ? 'bg-green-100 text-green-700 dark:bg-green-900/30 dark:text-green-400'
                              : 'bg-red-100 text-red-700 dark:bg-red-900/30 dark:text-red-400'
                          }`}
                        >
                          {user.accountEnabled ? (
                            <CheckmarkCircleFilled className="w-3 h-3" />
                          ) : (
                            <DismissCircleFilled className="w-3 h-3" />
                          )}
                          {user.accountEnabled ? 'Enabled' : 'Disabled'}
                        </span>
                      </td>
                      <td className="px-4 py-3">
                        <div>
                          <span
                            className={`inline-flex items-center gap-1 px-2 py-1 rounded-full text-xs font-medium ${
                              user.isMfaRegistered
                                ? 'bg-emerald-100 text-emerald-700 dark:bg-emerald-900/30 dark:text-emerald-400'
                                : 'bg-red-100 text-red-700 dark:bg-red-900/30 dark:text-red-400'
                            }`}
                          >
                            {user.isMfaRegistered ? (
                              <ShieldCheckmarkRegular className="w-3 h-3" />
                            ) : (
                              <ShieldErrorRegular className="w-3 h-3" />
                            )}
                            {user.isMfaRegistered ? 'Enabled' : 'Not Set'}
                          </span>
                          {user.defaultMfaMethod && (
                            <p className="text-xs text-slate-500 dark:text-slate-400 mt-1">
                              {user.defaultMfaMethod}
                            </p>
                          )}
                        </div>
                      </td>
                      <td className="px-4 py-3">
                        <div>
                          <p className="text-sm text-slate-900 dark:text-white">
                            {user.department || '-'}
                          </p>
                          <p className="text-xs text-slate-500 dark:text-slate-400">
                            {user.jobTitle || ''}
                          </p>
                        </div>
                      </td>
                      <td className="px-4 py-3">
                        <div>
                          <p className="text-sm text-slate-900 dark:text-white flex items-center gap-1">
                            <PersonRegular className="w-3 h-3 text-slate-400" />
                            {formatDate(user.lastSignInDateTime)}
                          </p>
                          <p className="text-xs text-slate-500 dark:text-slate-400 flex items-center gap-1">
                            <LockClosedRegular className="w-3 h-3" />
                            {formatDate(user.lastNonInteractiveSignInDateTime)}
                          </p>
                        </div>
                      </td>
                      <td className="px-4 py-3">
                        {user.assignedLicenses && user.assignedLicenses.length > 0 ? (
                          <div className="flex flex-wrap gap-1">
                            {user.assignedLicenses.slice(0, 2).map((license, idx) => (
                              <span
                                key={idx}
                                className="px-2 py-0.5 bg-blue-100 text-blue-700 dark:bg-blue-900/30 dark:text-blue-400 rounded text-xs"
                                title={license}
                              >
                                {license.length > 20 ? license.substring(0, 20) + '...' : license}
                              </span>
                            ))}
                            {user.assignedLicenses.length > 2 && (
                              <span className="px-2 py-0.5 bg-slate-100 text-slate-600 dark:bg-slate-700 dark:text-slate-400 rounded text-xs">
                                +{user.assignedLicenses.length - 2}
                              </span>
                            )}
                          </div>
                        ) : (
                          <span className="text-sm text-slate-400">No licenses</span>
                        )}
                      </td>
                      {filterType === 'noUsageLocation' && (
                        <td className="px-4 py-3">
                          <span className="inline-flex items-center gap-1 px-2 py-1 rounded-full text-xs font-medium bg-orange-100 text-orange-700 dark:bg-orange-900/30 dark:text-orange-400">
                            <LocationRegular className="w-3 h-3" />
                            Not set
                          </span>
                        </td>
                      )}
                      <td className="px-4 py-3 text-right">
                        <button
                          onClick={() => openInEntraPortal(user.id)}
                          className="p-2 text-slate-400 hover:text-blue-600 hover:bg-blue-50 dark:hover:bg-blue-900/30 rounded-lg transition-colors"
                          title="Open in Entra ID"
                        >
                          <OpenRegular className="w-5 h-5" />
                        </button>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </>
        )}
      </div>
    </div>
  );
};

export default UsersPage;
