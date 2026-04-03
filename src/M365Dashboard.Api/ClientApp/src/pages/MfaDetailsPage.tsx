import React, { useState, useEffect, useMemo } from 'react';
import { useNavigate } from 'react-router-dom';
import {
  ArrowLeftRegular,
  LockClosedRegular,
  LockOpenRegular,
  SearchRegular,
  CheckmarkCircleFilled,
  DismissCircleFilled,
  ShieldCheckmarkRegular,
  PersonRegular,
  OpenRegular,
  ChevronUpRegular,
  ChevronDownRegular,
} from '@fluentui/react-icons';
import { useAppContext } from '../contexts/AppContext';

interface MfaUserDetail {
  id: string;
  userPrincipalName: string;
  displayName: string | null;
  isMfaRegistered: boolean;
  isMfaCapable: boolean;
  defaultMfaMethod: string | null;
  methodsRegistered: string[] | null;
  isAdmin: boolean;
}

interface MfaRegistrationList {
  users: MfaUserDetail[];
  totalCount: number;
  mfaRegisteredCount: number;
  mfaNotRegisteredCount: number;
  mfaRegistrationPercentage: number;
  lastUpdated: string;
}

type FilterType = 'all' | 'registered' | 'not-registered' | 'admins';
type SortField = 'displayName' | 'userPrincipalName' | 'isMfaRegistered' | 'defaultMfaMethod';

const MfaDetailsPage: React.FC = () => {
  const navigate = useNavigate();
  const { getAccessToken } = useAppContext();
  const [data, setData] = useState<MfaRegistrationList | null>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [searchTerm, setSearchTerm] = useState('');
  const [filterType, setFilterType] = useState<FilterType>('all');
  const [sortField, setSortField] = useState<SortField>('isMfaRegistered');
  const [sortAscending, setSortAscending] = useState(true);

  useEffect(() => {
    fetchMfaDetails();
  }, []);

  const fetchMfaDetails = async () => {
    try {
      setLoading(true);
      const token = await getAccessToken();
      const response = await fetch('/api/security/mfa', {
        headers: {
          'Authorization': `Bearer ${token}`,
          'Content-Type': 'application/json',
        },
      });

      if (!response.ok) {
        throw new Error('Failed to fetch MFA details');
      }

      const result: MfaRegistrationList = await response.json();
      setData(result);
      setError(null);
    } catch (err) {
      setError(err instanceof Error ? err.message : 'An error occurred');
    } finally {
      setLoading(false);
    }
  };

  const handleSort = (field: SortField) => {
    if (sortField === field) {
      setSortAscending(!sortAscending);
    } else {
      setSortField(field);
      setSortAscending(true);
    }
  };

  const SortIcon = ({ field }: { field: SortField }) => {
    if (sortField !== field) return null;
    return sortAscending ? 
      <ChevronUpRegular className="w-4 h-4" /> : 
      <ChevronDownRegular className="w-4 h-4" />;
  };

  const filteredAndSortedUsers = useMemo(() => {
    if (!data) return [];

    let filtered = data.users;

    // Apply search filter
    if (searchTerm) {
      const lower = searchTerm.toLowerCase();
      filtered = filtered.filter(u =>
        (u.displayName?.toLowerCase().includes(lower)) ||
        (u.userPrincipalName.toLowerCase().includes(lower)) ||
        (u.defaultMfaMethod?.toLowerCase().includes(lower))
      );
    }

    // Apply type filter
    switch (filterType) {
      case 'registered':
        filtered = filtered.filter(u => u.isMfaRegistered);
        break;
      case 'not-registered':
        filtered = filtered.filter(u => !u.isMfaRegistered);
        break;
      case 'admins':
        filtered = filtered.filter(u => u.isAdmin);
        break;
    }

    // Apply sorting
    filtered = [...filtered].sort((a, b) => {
      let comparison = 0;
      switch (sortField) {
        case 'displayName':
          comparison = (a.displayName || '').localeCompare(b.displayName || '');
          break;
        case 'userPrincipalName':
          comparison = a.userPrincipalName.localeCompare(b.userPrincipalName);
          break;
        case 'isMfaRegistered':
          comparison = (a.isMfaRegistered === b.isMfaRegistered) ? 0 : a.isMfaRegistered ? 1 : -1;
          break;
        case 'defaultMfaMethod':
          comparison = (a.defaultMfaMethod || '').localeCompare(b.defaultMfaMethod || '');
          break;
      }
      return sortAscending ? comparison : -comparison;
    });

    return filtered;
  }, [data, searchTerm, filterType, sortField, sortAscending]);

  if (loading) {
    return (
      <div className="p-4 flex items-center justify-center h-64">
        <div className="animate-spin rounded-full h-8 w-8 border-b-2 border-blue-600"></div>
      </div>
    );
  }

  if (error) {
    return (
      <div className="p-4">
        <div className="bg-red-50 dark:bg-red-900/20 border border-red-200 dark:border-red-800 rounded-lg p-4">
          <p className="text-red-600 dark:text-red-400">{error}</p>
          <button 
            onClick={fetchMfaDetails}
            className="mt-2 text-sm text-red-600 dark:text-red-400 underline"
          >
            Retry
          </button>
        </div>
      </div>
    );
  }

  return (
    <div className="p-4 space-y-4 w-full max-w-full overflow-hidden">
      {/* Header */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
        <div className="flex items-center gap-3">
          <button
            onClick={() => navigate('/security')}
            className="p-2 hover:bg-slate-100 dark:hover:bg-slate-800 rounded-lg transition-colors"
          >
            <ArrowLeftRegular className="w-5 h-5 text-slate-600 dark:text-slate-400" />
          </button>
          <div>
            <h1 className="text-xl font-semibold text-slate-900 dark:text-white">MFA Registration</h1>
            <p className="text-sm text-slate-500 dark:text-slate-400">
              Multi-factor authentication status for all users
            </p>
          </div>
        </div>
        <a
          href="https://entra.microsoft.com/#view/Microsoft_AAD_IAM/AuthenticationMethodsMenuBlade/~/UserRegistrationDetails"
          target="_blank"
          rel="noopener noreferrer"
          className="inline-flex items-center justify-center gap-2 px-3 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 transition-colors text-sm whitespace-nowrap"
        >
          <OpenRegular className="w-4 h-4" />
          <span className="hidden sm:inline">Entra Admin Center</span>
          <span className="sm:hidden">Entra</span>
        </a>
      </div>

      {/* Stats Cards */}
      <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
        <div 
          className={`bg-white dark:bg-slate-800 rounded-lg border p-3 cursor-pointer transition-all ${
            filterType === 'all' 
              ? 'border-blue-500 ring-2 ring-blue-500/20' 
              : 'border-slate-200 dark:border-slate-700 hover:border-slate-300'
          }`}
          onClick={() => setFilterType('all')}
        >
          <div className="flex items-center gap-2">
            <PersonRegular className="w-5 h-5 text-blue-600" />
            <span className="text-sm text-slate-600 dark:text-slate-400">Total Users</span>
          </div>
          <p className="text-2xl font-bold text-slate-900 dark:text-white mt-1">{data?.totalCount ?? 0}</p>
        </div>

        <div 
          className={`bg-white dark:bg-slate-800 rounded-lg border p-3 cursor-pointer transition-all ${
            filterType === 'registered' 
              ? 'border-green-500 ring-2 ring-green-500/20' 
              : 'border-slate-200 dark:border-slate-700 hover:border-slate-300'
          }`}
          onClick={() => setFilterType('registered')}
        >
          <div className="flex items-center gap-2">
            <LockClosedRegular className="w-5 h-5 text-green-600" />
            <span className="text-sm text-slate-600 dark:text-slate-400">MFA Enabled</span>
          </div>
          <p className="text-2xl font-bold text-green-600 mt-1">{data?.mfaRegisteredCount ?? 0}</p>
        </div>

        <div 
          className={`bg-white dark:bg-slate-800 rounded-lg border p-3 cursor-pointer transition-all ${
            filterType === 'not-registered' 
              ? 'border-red-500 ring-2 ring-red-500/20' 
              : 'border-slate-200 dark:border-slate-700 hover:border-slate-300'
          }`}
          onClick={() => setFilterType('not-registered')}
        >
          <div className="flex items-center gap-2">
            <LockOpenRegular className="w-5 h-5 text-red-600" />
            <span className="text-sm text-slate-600 dark:text-slate-400">No MFA</span>
          </div>
          <p className="text-2xl font-bold text-red-600 mt-1">{data?.mfaNotRegisteredCount ?? 0}</p>
        </div>

        <div className="bg-white dark:bg-slate-800 rounded-lg border border-slate-200 dark:border-slate-700 p-3">
          <div className="flex items-center gap-2">
            <ShieldCheckmarkRegular className="w-5 h-5 text-blue-600" />
            <span className="text-sm text-slate-600 dark:text-slate-400">Coverage</span>
          </div>
          <p className={`text-2xl font-bold mt-1 ${
            (data?.mfaRegistrationPercentage ?? 0) >= 90 ? 'text-green-600' : 
            (data?.mfaRegistrationPercentage ?? 0) >= 70 ? 'text-amber-600' : 'text-red-600'
          }`}>
            {data?.mfaRegistrationPercentage ?? 0}%
          </p>
        </div>
      </div>

      {/* Search and Filter */}
      <div className="bg-white dark:bg-slate-800 rounded-lg border border-slate-200 dark:border-slate-700 p-4">
        <div className="flex flex-col sm:flex-row gap-3">
          <div className="relative flex-1">
            <SearchRegular className="absolute left-3 top-1/2 transform -translate-y-1/2 w-4 h-4 text-slate-400" />
            <input
              type="text"
              placeholder="Search by name, email, or method..."
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
              className="w-full pl-9 pr-3 py-2 border border-slate-200 dark:border-slate-600 rounded-lg bg-white dark:bg-slate-700 text-slate-900 dark:text-white focus:outline-none focus:ring-2 focus:ring-blue-500 text-sm"
            />
          </div>
          <select
            value={filterType}
            onChange={(e) => setFilterType(e.target.value as FilterType)}
            className="px-3 py-2 border border-slate-200 dark:border-slate-600 rounded-lg bg-white dark:bg-slate-700 text-slate-900 dark:text-white focus:outline-none focus:ring-2 focus:ring-blue-500 text-sm"
          >
            <option value="all">All Users</option>
            <option value="registered">MFA Enabled</option>
            <option value="not-registered">No MFA</option>
            <option value="admins">Admins Only</option>
          </select>
        </div>

        <div className="mt-2 text-xs text-slate-500 dark:text-slate-400">
          Showing {filteredAndSortedUsers.length} of {data?.totalCount ?? 0} users
        </div>
      </div>

      {/* Users Table */}
      <div className="bg-white dark:bg-slate-800 rounded-lg border border-slate-200 dark:border-slate-700 overflow-hidden">
        {/* Mobile Card View */}
        <div className="sm:hidden divide-y divide-slate-200 dark:divide-slate-700 max-h-[60vh] overflow-y-auto">
          {filteredAndSortedUsers.map((user) => (
            <div key={user.id} className="p-4">
              <div className="flex items-start justify-between">
                <div className="min-w-0 flex-1">
                  <div className="flex items-center gap-2">
                    <p className="font-medium text-slate-900 dark:text-white truncate">
                      {user.displayName || user.userPrincipalName}
                    </p>
                    {user.isAdmin && (
                      <span className="px-1.5 py-0.5 bg-purple-100 text-purple-700 dark:bg-purple-900/30 dark:text-purple-400 text-xs rounded">
                        Admin
                      </span>
                    )}
                  </div>
                  <p className="text-xs text-slate-500 dark:text-slate-400 truncate mt-0.5">
                    {user.userPrincipalName}
                  </p>
                </div>
                {user.isMfaRegistered ? (
                  <CheckmarkCircleFilled className="w-5 h-5 text-green-500 flex-shrink-0" />
                ) : (
                  <DismissCircleFilled className="w-5 h-5 text-red-500 flex-shrink-0" />
                )}
              </div>
              {user.isMfaRegistered && (
                <div className="mt-2">
                  <p className="text-xs text-slate-500 dark:text-slate-400">
                    <span className="font-medium">Default:</span> {user.defaultMfaMethod || 'Not set'}
                  </p>
                  {user.methodsRegistered && user.methodsRegistered.length > 0 && (
                    <div className="flex flex-wrap gap-1 mt-1">
                      {user.methodsRegistered.map((method, idx) => (
                        <span 
                          key={idx}
                          className="px-1.5 py-0.5 bg-slate-100 dark:bg-slate-700 text-slate-600 dark:text-slate-300 text-xs rounded"
                        >
                          {method}
                        </span>
                      ))}
                    </div>
                  )}
                </div>
              )}
            </div>
          ))}
        </div>

        {/* Desktop Table View */}
        <div className="hidden sm:block overflow-x-auto max-h-[60vh] overflow-y-auto">
          <table className="w-full">
            <thead className="bg-slate-50 dark:bg-slate-900 sticky top-0">
              <tr>
                <th 
                  className="px-4 py-3 text-left text-xs font-medium text-slate-500 dark:text-slate-400 uppercase tracking-wider cursor-pointer hover:bg-slate-100 dark:hover:bg-slate-800"
                  onClick={() => handleSort('displayName')}
                >
                  <div className="flex items-center gap-1">
                    User
                    <SortIcon field="displayName" />
                  </div>
                </th>
                <th 
                  className="px-4 py-3 text-left text-xs font-medium text-slate-500 dark:text-slate-400 uppercase tracking-wider cursor-pointer hover:bg-slate-100 dark:hover:bg-slate-800"
                  onClick={() => handleSort('isMfaRegistered')}
                >
                  <div className="flex items-center gap-1">
                    MFA Status
                    <SortIcon field="isMfaRegistered" />
                  </div>
                </th>
                <th 
                  className="px-4 py-3 text-left text-xs font-medium text-slate-500 dark:text-slate-400 uppercase tracking-wider cursor-pointer hover:bg-slate-100 dark:hover:bg-slate-800"
                  onClick={() => handleSort('defaultMfaMethod')}
                >
                  <div className="flex items-center gap-1">
                    Default Method
                    <SortIcon field="defaultMfaMethod" />
                  </div>
                </th>
                <th className="px-4 py-3 text-left text-xs font-medium text-slate-500 dark:text-slate-400 uppercase tracking-wider">
                  Registered Methods
                </th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-200 dark:divide-slate-700">
              {filteredAndSortedUsers.map((user) => (
                <tr key={user.id} className="hover:bg-slate-50 dark:hover:bg-slate-700/50">
                  <td className="px-4 py-3">
                    <div className="flex items-center gap-2">
                      <div className="min-w-0">
                        <div className="flex items-center gap-2">
                          <p className="font-medium text-slate-900 dark:text-white truncate">
                            {user.displayName || user.userPrincipalName}
                          </p>
                          {user.isAdmin && (
                            <span className="px-1.5 py-0.5 bg-purple-100 text-purple-700 dark:bg-purple-900/30 dark:text-purple-400 text-xs rounded flex-shrink-0">
                              Admin
                            </span>
                          )}
                        </div>
                        <p className="text-xs text-slate-500 dark:text-slate-400 truncate">
                          {user.userPrincipalName}
                        </p>
                      </div>
                    </div>
                  </td>
                  <td className="px-4 py-3">
                    {user.isMfaRegistered ? (
                      <span className="inline-flex items-center gap-1 px-2 py-1 bg-green-100 text-green-700 dark:bg-green-900/30 dark:text-green-400 text-xs font-medium rounded-full">
                        <CheckmarkCircleFilled className="w-3 h-3" />
                        Enabled
                      </span>
                    ) : (
                      <span className="inline-flex items-center gap-1 px-2 py-1 bg-red-100 text-red-700 dark:bg-red-900/30 dark:text-red-400 text-xs font-medium rounded-full">
                        <DismissCircleFilled className="w-3 h-3" />
                        Not Enabled
                      </span>
                    )}
                  </td>
                  <td className="px-4 py-3">
                    <span className="text-sm text-slate-600 dark:text-slate-300">
                      {user.defaultMfaMethod || '-'}
                    </span>
                  </td>
                  <td className="px-4 py-3">
                    {user.methodsRegistered && user.methodsRegistered.length > 0 ? (
                      <div className="flex flex-wrap gap-1">
                        {user.methodsRegistered.slice(0, 3).map((method, idx) => (
                          <span 
                            key={idx}
                            className="px-1.5 py-0.5 bg-slate-100 dark:bg-slate-700 text-slate-600 dark:text-slate-300 text-xs rounded"
                          >
                            {method}
                          </span>
                        ))}
                        {user.methodsRegistered.length > 3 && (
                          <span className="px-1.5 py-0.5 bg-slate-100 dark:bg-slate-700 text-slate-500 text-xs rounded">
                            +{user.methodsRegistered.length - 3} more
                          </span>
                        )}
                      </div>
                    ) : (
                      <span className="text-sm text-slate-400">-</span>
                    )}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>

        {filteredAndSortedUsers.length === 0 && (
          <div className="p-8 text-center text-slate-500 dark:text-slate-400">
            <p>No users found matching your criteria</p>
          </div>
        )}
      </div>
    </div>
  );
};

export default MfaDetailsPage;
