import React, { useState, useEffect, useRef, useCallback, useMemo } from 'react';
import {
  LocationRegular,
  PersonRegular,
  CheckmarkCircleRegular,
  DismissCircleRegular,
  ClockRegular,
  GlobeRegular,
  FilterRegular,
  ArrowSyncRegular,
  InfoRegular,
  WarningRegular,
  ChevronDownRegular,
  ChevronUpRegular,
  ListRegular,
  MapRegular,
} from '@fluentui/react-icons';
import { useAppContext } from '../contexts/AppContext';
import { useSearchParams } from 'react-router-dom';

interface SignInDetail {
  id: string;
  userPrincipalName: string;
  displayName?: string;
  createdDateTime?: string;
  ipAddress?: string;
  city?: string;
  state?: string;
  countryOrRegion?: string;
  latitude?: number;
  longitude?: number;
  isSuccess: boolean;
  errorCode?: number;
  failureReason?: string;
  clientAppUsed?: string;
  browser?: string;
  operatingSystem?: string;
  deviceDisplayName?: string;
  isCompliant?: boolean;
  isManaged?: boolean;
  riskLevel?: string;
  riskState?: string;
  mfaRequired?: boolean;
  conditionalAccessStatus?: string;
}

interface SignInLocation {
  latitude: number;
  longitude: number;
  city: string;
  state?: string;
  countryOrRegion: string;
  signInCount: number;
  successCount: number;
  failureCount: number;
  signIns: SignInDetail[];
}

interface SignInsMapData {
  locations: SignInLocation[];
  totalSignIns: number;
  successfulSignIns: number;
  failedSignIns: number;
  uniqueUsers: number;
  uniqueLocations: number;
  startDate: string;
  endDate: string;
  lastUpdated: string;
}

declare global {
  interface Window {
    atlas: any;
  }
}

const SignInsPage: React.FC = () => {
  const { getAccessToken } = useAppContext();
  const [searchParams, setSearchParams] = useSearchParams();
  const mapRef = useRef<HTMLDivElement>(null);
  const mapInstanceRef = useRef<any>(null);
  const dataSourceRef = useRef<any>(null);
  const popupRef = useRef<any>(null);

  const [mapData, setMapData] = useState<SignInsMapData | null>(null);
  const [allSignIns, setAllSignIns] = useState<SignInDetail[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [timeRange, setTimeRange] = useState<number>(168); // Default to 7 days for list view
  const [azureMapsKey, setAzureMapsKey] = useState<string | null>(null);
  const [mapReady, setMapReady] = useState(false);
  const [selectedLocation, setSelectedLocation] = useState<SignInLocation | null>(null);
  const [showSignInList, setShowSignInList] = useState(true);
  
  // View mode and status filter from URL params
  const statusFilter = searchParams.get('status');
  const [viewMode, setViewMode] = useState<'map' | 'list'>(statusFilter ? 'list' : 'map');
  const [statusFilterState, setStatusFilterState] = useState<'all' | 'success' | 'failure'>(
    statusFilter === 'failure' ? 'failure' : statusFilter === 'success' ? 'success' : 'all'
  );

  // Active map filter (success/failure pill on map view)
  const [mapFilter, setMapFilter] = useState<'all' | 'success' | 'failure'>('all');

  // Toggle map filter — clicking the active filter deselects it
  const handleMapFilterToggle = (filter: 'success' | 'failure') => {
    setMapFilter(prev => prev === filter ? 'all' : filter);
    setSelectedLocation(null);
  };

  // Filter sign-ins based on status
  const filteredSignIns = useMemo(() => {
    if (statusFilterState === 'all') return allSignIns;
    if (statusFilterState === 'failure') return allSignIns.filter(s => !s.isSuccess);
    if (statusFilterState === 'success') return allSignIns.filter(s => s.isSuccess);
    return allSignIns;
  }, [allSignIns, statusFilterState]);

  // Update URL when filter changes
  const handleStatusFilterChange = (newStatus: 'all' | 'success' | 'failure') => {
    setStatusFilterState(newStatus);
    if (newStatus === 'all') {
      searchParams.delete('status');
    } else {
      searchParams.set('status', newStatus);
    }
    setSearchParams(searchParams);
  };

  // Load Azure Maps SDK
  useEffect(() => {
    const loadAzureMapsSDK = async () => {
      if (window.atlas) {
        setMapReady(true);
        return;
      }

      const cssLink = document.createElement('link');
      cssLink.rel = 'stylesheet';
      cssLink.href = 'https://atlas.microsoft.com/sdk/javascript/mapcontrol/3/atlas.min.css';
      document.head.appendChild(cssLink);

      const script = document.createElement('script');
      script.src = 'https://atlas.microsoft.com/sdk/javascript/mapcontrol/3/atlas.min.js';
      script.async = true;
      script.onload = () => setMapReady(true);
      document.head.appendChild(script);
    };

    loadAzureMapsSDK();
  }, []);

  // Fetch Azure Maps key
  useEffect(() => {
    const fetchMapsKey = async () => {
      try {
        const token = await getAccessToken();
        const response = await fetch('/api/config/azure-maps-key', {
          headers: { 'Authorization': `Bearer ${token}` },
        });
        if (response.ok) {
          const data = await response.json();
          setAzureMapsKey(data.key);
        }
      } catch (err) {
        console.error('Failed to fetch Azure Maps key:', err);
      }
    };
    fetchMapsKey();
  }, [getAccessToken]);

  // Fetch sign-in data
  const fetchSignInData = useCallback(async () => {
    try {
      setLoading(true);
      setError(null);
      const token = await getAccessToken();
      
      // Always fetch map data for stats
      const mapResponse = await fetch(`/api/signins/map?hours=${timeRange}`, {
        headers: {
          'Authorization': `Bearer ${token}`,
          'Content-Type': 'application/json',
        },
      });

      if (!mapResponse.ok) throw new Error('Failed to fetch sign-in data');

      const mapData: SignInsMapData = await mapResponse.json();
      setMapData(mapData);
      
      // For list view, fetch all sign-ins (not just those with location data)
      const listResponse = await fetch(`/api/signins?hours=${timeRange}&take=500`, {
        headers: {
          'Authorization': `Bearer ${token}`,
          'Content-Type': 'application/json',
        },
      });
      
      if (listResponse.ok) {
        const listData = await listResponse.json();
        // Sort by date descending
        const sortedSignIns = (listData.signIns || []).sort((a: SignInDetail, b: SignInDetail) => {
          const dateA = a.createdDateTime ? new Date(a.createdDateTime).getTime() : 0;
          const dateB = b.createdDateTime ? new Date(b.createdDateTime).getTime() : 0;
          return dateB - dateA;
        });
        setAllSignIns(sortedSignIns);
      } else {
        // Fallback to map data sign-ins if list endpoint fails
        const allSignInsFromLocations: SignInDetail[] = [];
        mapData.locations.forEach(location => {
          if (location.signIns) {
            allSignInsFromLocations.push(...location.signIns);
          }
        });
        allSignInsFromLocations.sort((a, b) => {
          const dateA = a.createdDateTime ? new Date(a.createdDateTime).getTime() : 0;
          const dateB = b.createdDateTime ? new Date(b.createdDateTime).getTime() : 0;
          return dateB - dateA;
        });
        setAllSignIns(allSignInsFromLocations);
      }
    } catch (err) {
      setError(err instanceof Error ? err.message : 'An error occurred');
    } finally {
      setLoading(false);
    }
  }, [getAccessToken, timeRange]);

  useEffect(() => {
    fetchSignInData();
  }, [fetchSignInData]);

  // Initialize map
  useEffect(() => {
    if (!mapReady || !azureMapsKey || !mapRef.current || mapInstanceRef.current || viewMode !== 'map') {
      return;
    }

    try {
      const map = new window.atlas.Map(mapRef.current, {
        authOptions: {
          authType: 'subscriptionKey',
          subscriptionKey: azureMapsKey,
        },
        center: [0, 30],
        zoom: 1.5,
        style: 'grayscale_dark',
        language: 'en-US',
      });

      map.events.add('ready', () => {
        const dataSource = new window.atlas.source.DataSource();
        map.sources.add(dataSource);
        dataSourceRef.current = dataSource;

        const bubbleLayer = new window.atlas.layer.BubbleLayer(dataSource, null, {
          radius: ['interpolate', ['linear'], ['get', 'signInCount'], 1, 8, 10, 15, 50, 25, 100, 35, 500, 50],
          color: ['get', 'displayColor'],
          strokeColor: 'white',
          strokeWidth: 2,
          opacity: 0.8,
        });
        map.layers.add(bubbleLayer);

        const popup = new window.atlas.Popup({ closeButton: true, pixelOffset: [0, -20] });
        popupRef.current = popup;

        map.events.add('click', bubbleLayer, (e: any) => {
          if (e.shapes && e.shapes.length > 0) {
            const properties = e.shapes[0].getProperties();
            setSelectedLocation(properties as SignInLocation);
            popup.setOptions({
              position: e.shapes[0].getCoordinates(),
              content: `<div style="padding: 12px; max-width: 250px;"><h3 style="margin: 0 0 8px 0; font-weight: 600;">${properties.city}, ${properties.countryOrRegion}</h3><div style="display: grid; grid-template-columns: 1fr 1fr; gap: 8px; font-size: 14px;"><div><span style="color: #6b7280;">Total:</span><strong style="margin-left: 4px;">${properties.signInCount}</strong></div><div><span style="color: #22c55e;">✓ ${properties.successCount}</span><span style="margin-left: 8px; color: #ef4444;">✕ ${properties.failureCount}</span></div></div></div>`,
            });
            popup.open(map);
          }
        });

        map.events.add('mouseover', bubbleLayer, () => map.getCanvasContainer().style.cursor = 'pointer');
        map.events.add('mouseout', bubbleLayer, () => map.getCanvasContainer().style.cursor = 'grab');
      });

      mapInstanceRef.current = map;
    } catch (err) {
      console.error('Failed to initialize Azure Maps:', err);
      setError('Failed to initialize map.');
    }
  }, [mapReady, azureMapsKey, viewMode]);

  // Update map data (re-runs when data or filter changes)
  useEffect(() => {
    if (!dataSourceRef.current || !mapData) return;
    dataSourceRef.current.clear();

    const locations = mapData.locations.filter(loc => {
      if (mapFilter === 'success') return loc.successCount > 0;
      if (mapFilter === 'failure') return loc.failureCount > 0;
      return true;
    });

    const features = locations.map((location) => {
      // When filtered, show only the relevant count so bubble sizing is correct
      const displayCount =
        mapFilter === 'success' ? location.successCount
        : mapFilter === 'failure' ? location.failureCount
        : location.signInCount;

      // Colour: filter overrides the mixed-result amber logic
      const displayColor =
        mapFilter === 'success' ? '#22c55e'
        : mapFilter === 'failure' ? '#ef4444'
        : location.failureCount === 0 ? '#22c55e'
        : location.failureCount > location.successCount ? '#ef4444'
        : '#f59e0b';

      return new window.atlas.data.Feature(
        new window.atlas.data.Point([location.longitude, location.latitude]),
        { ...location, signInCount: displayCount, displayColor }
      );
    });

    dataSourceRef.current.add(features);
  }, [mapData, mapFilter]);

  const timeRangeOptions = [
    { value: 24, label: 'Last 24 hours' },
    { value: 168, label: 'Last 7 days' },
    { value: 720, label: 'Last 30 days' },
  ];

  // Warning for no Azure Maps key (only for map view)
  if (viewMode === 'map' && !azureMapsKey && !loading) {
    return (
      <div className="p-4">
        <div className="bg-amber-50 dark:bg-amber-900/20 border border-amber-200 dark:border-amber-800 rounded-lg p-6">
          <div className="flex items-start gap-3">
            <WarningRegular className="w-6 h-6 text-amber-600 flex-shrink-0" />
            <div>
              <h3 className="font-semibold text-amber-800 dark:text-amber-200">Azure Maps Not Configured</h3>
              <p className="text-sm text-amber-700 dark:text-amber-300 mt-1">
                To display the sign-in map, configure an Azure Maps subscription key. You can switch to List view instead.
              </p>
              <button
                onClick={() => setViewMode('list')}
                className="mt-3 px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700"
              >
                Switch to List View
              </button>
            </div>
          </div>
        </div>
      </div>
    );
  }

  return (
    <div className="h-full flex flex-col">
      {/* Header */}
      <div className="flex-shrink-0 p-4 border-b border-slate-200 dark:border-slate-700 bg-white dark:bg-slate-800">
        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4">
          <div>
            <h1 className="text-xl font-semibold text-slate-900 dark:text-white flex items-center gap-2">
              <GlobeRegular className="w-6 h-6" />
              Sign-in Activity
            </h1>
            <p className="text-sm text-slate-500 dark:text-slate-400">
              {viewMode === 'map' ? 'Visualize sign-in locations across your organization' : 'View detailed sign-in logs'}
            </p>
          </div>
          <div className="flex items-center gap-3">
            {/* View Mode Toggle */}
            <div className="flex items-center bg-slate-100 dark:bg-slate-700 rounded-lg p-1">
              <button
                onClick={() => setViewMode('map')}
                className={`flex items-center gap-1 px-3 py-1.5 rounded-md text-sm font-medium transition-colors ${
                  viewMode === 'map' 
                    ? 'bg-white dark:bg-slate-600 text-slate-900 dark:text-white shadow-sm' 
                    : 'text-slate-600 dark:text-slate-400 hover:text-slate-900 dark:hover:text-white'
                }`}
              >
                <MapRegular className="w-4 h-4" />
                Map
              </button>
              <button
                onClick={() => { setViewMode('list'); setMapFilter('all'); }}
                className={`flex items-center gap-1 px-3 py-1.5 rounded-md text-sm font-medium transition-colors ${
                  viewMode === 'list' 
                    ? 'bg-white dark:bg-slate-600 text-slate-900 dark:text-white shadow-sm' 
                    : 'text-slate-600 dark:text-slate-400 hover:text-slate-900 dark:hover:text-white'
                }`}
              >
                <ListRegular className="w-4 h-4" />
                List
              </button>
            </div>

            {/* Status Filter (List view only) */}
            {viewMode === 'list' && (
              <select
                value={statusFilterState}
                onChange={(e) => handleStatusFilterChange(e.target.value as 'all' | 'success' | 'failure')}
                className="px-3 py-1.5 border border-slate-300 dark:border-slate-600 rounded-lg bg-white dark:bg-slate-700 text-slate-900 dark:text-white text-sm"
              >
                <option value="all">All Sign-ins</option>
                <option value="success">Successful Only</option>
                <option value="failure">Failed Only</option>
              </select>
            )}

            {/* Time Range */}
            <div className="flex items-center gap-2">
              <FilterRegular className="w-4 h-4 text-slate-400" />
              <select
                value={timeRange}
                onChange={(e) => setTimeRange(Number(e.target.value))}
                className="px-3 py-1.5 border border-slate-300 dark:border-slate-600 rounded-lg bg-white dark:bg-slate-700 text-slate-900 dark:text-white text-sm"
              >
                {timeRangeOptions.map((opt) => (
                  <option key={opt.value} value={opt.value}>{opt.label}</option>
                ))}
              </select>
            </div>

            <button
              onClick={fetchSignInData}
              disabled={loading}
              className="p-2 text-slate-600 hover:text-blue-600 hover:bg-blue-50 dark:text-slate-400 dark:hover:text-blue-400 dark:hover:bg-blue-900/20 rounded-lg transition-colors disabled:opacity-50"
              title="Refresh"
            >
              <ArrowSyncRegular className={`w-5 h-5 ${loading ? 'animate-spin' : ''}`} />
            </button>
          </div>
        </div>

        {/* Stats Bar */}
        {mapData && (
          <div className="mt-4 grid grid-cols-2 sm:grid-cols-5 gap-3">
            <StatCard icon={<PersonRegular className="w-4 h-4" />} label="Total Sign-ins" value={mapData.totalSignIns.toLocaleString()} />
            <StatCard
              icon={<CheckmarkCircleRegular className="w-4 h-4 text-green-500" />}
              label="Successful"
              value={mapData.successfulSignIns.toLocaleString()}
              color="text-green-600"
              onClick={viewMode === 'map' ? () => handleMapFilterToggle('success') : undefined}
              active={mapFilter === 'success'}
              activeRing="ring-2 ring-green-500"
            />
            <StatCard
              icon={<DismissCircleRegular className="w-4 h-4 text-red-500" />}
              label="Failed"
              value={mapData.failedSignIns.toLocaleString()}
              color="text-red-600"
              onClick={viewMode === 'map' ? () => handleMapFilterToggle('failure') : undefined}
              active={mapFilter === 'failure'}
              activeRing="ring-2 ring-red-500"
            />
            <StatCard icon={<PersonRegular className="w-4 h-4" />} label="Unique Users" value={mapData.uniqueUsers.toLocaleString()} />
            <StatCard icon={<LocationRegular className="w-4 h-4" />} label="Locations" value={mapData.uniqueLocations.toLocaleString()} />
          </div>
        )}
      </div>

      {/* Main Content */}
      {viewMode === 'map' ? (
        <div className="flex-1 flex min-h-0">
          {/* Map */}
          <div className="flex-1 relative">
            {loading && (
              <div className="absolute inset-0 bg-white/80 dark:bg-slate-900/80 z-10 flex items-center justify-center">
                <div className="flex items-center gap-3">
                  <div className="animate-spin rounded-full h-8 w-8 border-b-2 border-blue-600"></div>
                  <span className="text-slate-600 dark:text-slate-300">Loading sign-in data...</span>
                </div>
              </div>
            )}
            {error && (
              <div className="absolute inset-0 bg-white dark:bg-slate-900 z-10 flex items-center justify-center">
                <div className="text-center">
                  <DismissCircleRegular className="w-12 h-12 text-red-500 mx-auto mb-3" />
                  <p className="text-red-600 dark:text-red-400">{error}</p>
                  <button onClick={fetchSignInData} className="mt-3 px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700">Retry</button>
                </div>
              </div>
            )}
            <div ref={mapRef} className="w-full h-full" />
          </div>

          {/* Side Panel */}
          <div className={`w-80 border-l flex flex-col overflow-hidden ${
            mapFilter === 'failure' ? 'border-red-200 dark:border-red-900 bg-red-50 dark:bg-red-950/30'
            : mapFilter === 'success' ? 'border-green-200 dark:border-green-900 bg-green-50 dark:bg-green-950/30'
            : 'border-slate-200 dark:border-slate-700 bg-white dark:bg-slate-800'
          }`}>
            <button
              onClick={() => setShowSignInList(!showSignInList)}
              className={`flex items-center justify-between px-4 py-3 border-b ${
                mapFilter === 'failure' ? 'border-red-200 dark:border-red-900 hover:bg-red-100 dark:hover:bg-red-900/30'
                : mapFilter === 'success' ? 'border-green-200 dark:border-green-900 hover:bg-green-100 dark:hover:bg-green-900/30'
                : 'border-slate-200 dark:border-slate-700 hover:bg-slate-50 dark:hover:bg-slate-700/50'
              }`}
            >
              <span className="font-medium text-slate-900 dark:text-white">
                {selectedLocation ? `${selectedLocation.city}, ${selectedLocation.countryOrRegion}` : (
                  mapFilter === 'failure' ? 'Failed Sign-ins'
                  : mapFilter === 'success' ? 'Successful Sign-ins'
                  : 'Recent Sign-ins'
                )}
              </span>
              {showSignInList ? <ChevronUpRegular className="w-5 h-5" /> : <ChevronDownRegular className="w-5 h-5" />}
            </button>

            {showSignInList && (
              <div className="flex-1 overflow-y-auto">
                {selectedLocation ? (
                  <div className="p-4">
                    <div className="flex items-center justify-between mb-4">
                      <p className="text-sm text-slate-500 dark:text-slate-400">{selectedLocation.signInCount} sign-ins</p>
                      <button onClick={() => setSelectedLocation(null)} className="text-xs text-blue-600 hover:underline">View all</button>
                    </div>
                    <div className="space-y-2">
                      {selectedLocation.signIns
                        .filter(s => mapFilter === 'failure' ? !s.isSuccess : mapFilter === 'success' ? s.isSuccess : true)
                        .map((signIn) => <SignInCard key={signIn.id} signIn={signIn} forceColor={mapFilter !== 'all' ? mapFilter : undefined} />)}
                    </div>
                  </div>
                ) : mapData?.locations ? (
                  <div className={`divide-y ${
                    mapFilter === 'failure' ? 'divide-red-100 dark:divide-red-900/50'
                    : mapFilter === 'success' ? 'divide-green-100 dark:divide-green-900/50'
                    : 'divide-slate-100 dark:divide-slate-700'
                  }`}>
                    {mapData.locations
                      .filter(loc => mapFilter === 'failure' ? loc.failureCount > 0 : mapFilter === 'success' ? loc.successCount > 0 : true)
                      .slice(0, 20)
                      .map((location, idx) => (
                        <button key={idx} onClick={() => setSelectedLocation(location)} className={`w-full px-4 py-3 flex items-center justify-between text-left ${
                          mapFilter === 'failure' ? 'hover:bg-red-100 dark:hover:bg-red-900/30'
                          : mapFilter === 'success' ? 'hover:bg-green-100 dark:hover:bg-green-900/30'
                          : 'hover:bg-slate-50 dark:hover:bg-slate-700/50'
                        }`}>
                          <div>
                            <p className="font-medium text-slate-900 dark:text-white text-sm">{location.city}</p>
                            <p className="text-xs text-slate-500 dark:text-slate-400">{location.countryOrRegion}</p>
                          </div>
                          <div className="text-right">
                            <p className={`text-sm font-medium ${
                              mapFilter === 'failure' ? 'text-red-600' : mapFilter === 'success' ? 'text-green-600' : 'text-slate-900 dark:text-white'
                            }`}>
                              {mapFilter === 'failure' ? location.failureCount : mapFilter === 'success' ? location.successCount : location.signInCount}
                            </p>
                            {mapFilter === 'all' && (
                              <div className="flex items-center gap-2 text-xs">
                                <span className="text-green-600">{location.successCount}</span>
                                <span className="text-red-600">{location.failureCount}</span>
                              </div>
                            )}
                          </div>
                        </button>
                      ))}
                  </div>
                ) : (
                  <div className="p-4 text-center text-slate-500 dark:text-slate-400">No sign-in data available</div>
                )}
              </div>
            )}
          </div>
        </div>
      ) : (
        /* List View */
        <div className="flex-1 overflow-auto p-4">
          {loading ? (
            <div className="flex items-center justify-center py-12">
              <div className="flex items-center gap-3">
                <div className="animate-spin rounded-full h-8 w-8 border-b-2 border-blue-600"></div>
                <span className="text-slate-600 dark:text-slate-300">Loading sign-in data...</span>
              </div>
            </div>
          ) : error ? (
            <div className="flex items-center justify-center py-12">
              <div className="text-center">
                <DismissCircleRegular className="w-12 h-12 text-red-500 mx-auto mb-3" />
                <p className="text-red-600 dark:text-red-400">{error}</p>
                <button onClick={fetchSignInData} className="mt-3 px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700">Retry</button>
              </div>
            </div>
          ) : filteredSignIns.length === 0 ? (
            <div className="flex items-center justify-center py-12">
              <div className="text-center">
                <CheckmarkCircleRegular className="w-12 h-12 text-green-500 mx-auto mb-3" />
                <p className="text-slate-600 dark:text-slate-400">
                  {statusFilterState === 'failure' ? 'No failed sign-ins found' : 'No sign-ins found'}
                </p>
              </div>
            </div>
          ) : (
            <div className="bg-white dark:bg-slate-800 rounded-lg border border-slate-200 dark:border-slate-700 overflow-hidden">
              <div className="px-4 py-3 bg-slate-50 dark:bg-slate-700/50 border-b border-slate-200 dark:border-slate-700">
                <p className="text-sm font-medium text-slate-700 dark:text-slate-300">
                  {filteredSignIns.length.toLocaleString()} {statusFilterState === 'failure' ? 'failed ' : statusFilterState === 'success' ? 'successful ' : ''}sign-ins
                </p>
              </div>
              <div className="divide-y divide-slate-100 dark:divide-slate-700">
                {filteredSignIns.map((signIn) => (
                  <SignInListRow key={signIn.id} signIn={signIn} />
                ))}
              </div>
            </div>
          )}
        </div>
      )}
    </div>
  );
};

// Helper components
interface StatCardProps {
  icon: React.ReactNode;
  label: string;
  value: string;
  color?: string;
  onClick?: () => void;
  active?: boolean;
  activeRing?: string;
}

const StatCard: React.FC<StatCardProps> = ({ icon, label, value, color = 'text-slate-900 dark:text-white', onClick, active, activeRing = 'ring-blue-400' }) => {
  if (onClick) {
    return (
      <button
        onClick={onClick}
        className={`bg-slate-50 dark:bg-slate-700/50 rounded-lg px-3 py-2 text-left w-full transition-all ${
          active ? `${activeRing} bg-white dark:bg-slate-700` : 'ring-2 ring-transparent hover:ring-slate-300 dark:hover:ring-slate-500'
        }`}
        title={active ? 'Click to clear filter' : 'Click to filter map'}
      >
        <div className="flex items-center gap-2 text-slate-500 dark:text-slate-400 text-xs">
          {icon}
          <span>{label}</span>
          {active && <span className="ml-auto text-[10px] font-medium uppercase tracking-wide opacity-60">active</span>}
        </div>
        <p className={`text-lg font-semibold mt-0.5 ${color}`}>{value}</p>
      </button>
    );
  }
  return (
    <div className="bg-slate-50 dark:bg-slate-700/50 rounded-lg px-3 py-2">
      <div className="flex items-center gap-2 text-slate-500 dark:text-slate-400 text-xs">
        {icon}
        <span>{label}</span>
      </div>
      <p className={`text-lg font-semibold mt-0.5 ${color}`}>{value}</p>
    </div>
  );
};

interface SignInCardProps {
  signIn: SignInDetail;
  forceColor?: 'success' | 'failure';
}

const SignInCard: React.FC<SignInCardProps> = ({ signIn, forceColor }) => {
  const isRed = forceColor === 'failure' || (!forceColor && !signIn.isSuccess);
  const isGreen = forceColor === 'success' || (!forceColor && signIn.isSuccess);
  return (
  <div className={`p-3 rounded-lg border ${isGreen ? 'bg-green-50 dark:bg-green-900/20 border-green-200 dark:border-green-800' : 'bg-red-50 dark:bg-red-900/20 border-red-200 dark:border-red-800'}`}>
    <div className="flex items-start justify-between">
      <div className="min-w-0 flex-1">
        <p className="font-medium text-slate-900 dark:text-white text-sm truncate">{signIn.displayName || signIn.userPrincipalName}</p>
        <p className="text-xs text-slate-500 dark:text-slate-400 truncate">{signIn.userPrincipalName}</p>
      </div>
      {isGreen ? <CheckmarkCircleRegular className="w-5 h-5 text-green-500 flex-shrink-0" /> : <DismissCircleRegular className="w-5 h-5 text-red-500 flex-shrink-0" />}
    </div>
    <div className="mt-2 text-xs text-slate-500 dark:text-slate-400 space-y-1">
      <div className="flex items-center gap-1">
        <ClockRegular className="w-3 h-3 flex-shrink-0" />
        <span>{signIn.createdDateTime ? new Date(signIn.createdDateTime).toLocaleString() : 'N/A'}</span>
      </div>
      {(signIn.city || signIn.countryOrRegion) && (
        <div className="flex items-center gap-1">
          <LocationRegular className="w-3 h-3 flex-shrink-0" />
          <span>{[signIn.city, signIn.countryOrRegion].filter(Boolean).join(', ')}</span>
        </div>
      )}
      {signIn.ipAddress && (
        <div className="flex items-center gap-1">
          <GlobeRegular className="w-3 h-3 flex-shrink-0" />
          <span>IP: {signIn.ipAddress}</span>
        </div>
      )}
      {signIn.clientAppUsed && <div className="truncate">{signIn.clientAppUsed}</div>}
      {(signIn.browser || signIn.operatingSystem) && (
        <div className="truncate">{[signIn.browser, signIn.operatingSystem].filter(Boolean).join(' · ')}</div>
      )}
      {isRed && signIn.failureReason && (
        <div className="text-red-600 dark:text-red-400 pt-1 border-t border-red-200 dark:border-red-800 mt-1">
          {signIn.failureReason}
        </div>
      )}
    </div>
  </div>
  );
};

const SignInListRow: React.FC<SignInCardProps> = ({ signIn }) => (
  <div className="px-4 py-3 flex items-center gap-4 hover:bg-slate-50 dark:hover:bg-slate-700/50">
    {/* Status Icon */}
    <div className="flex-shrink-0">
      {signIn.isSuccess ? (
        <CheckmarkCircleRegular className="w-6 h-6 text-green-500" />
      ) : (
        <DismissCircleRegular className="w-6 h-6 text-red-500" />
      )}
    </div>

    {/* User Info */}
    <div className="flex-1 min-w-0">
      <p className="font-medium text-slate-900 dark:text-white text-sm truncate">
        {signIn.displayName || signIn.userPrincipalName}
      </p>
      <p className="text-xs text-slate-500 dark:text-slate-400 truncate">{signIn.userPrincipalName}</p>
    </div>

    {/* Location */}
    <div className="hidden sm:block w-40">
      <p className="text-sm text-slate-700 dark:text-slate-300 truncate">
        {signIn.city && signIn.countryOrRegion ? `${signIn.city}, ${signIn.countryOrRegion}` : signIn.countryOrRegion || '-'}
      </p>
      <p className="text-xs text-slate-500 dark:text-slate-400">{signIn.ipAddress || '-'}</p>
    </div>

    {/* App / Client */}
    <div className="hidden md:block w-32">
      <p className="text-sm text-slate-700 dark:text-slate-300 truncate">{signIn.clientAppUsed || '-'}</p>
      <p className="text-xs text-slate-500 dark:text-slate-400 truncate">{signIn.browser || signIn.operatingSystem || '-'}</p>
    </div>

    {/* Time */}
    <div className="w-40 text-right">
      <p className="text-sm text-slate-700 dark:text-slate-300">
        {signIn.createdDateTime ? new Date(signIn.createdDateTime).toLocaleDateString() : '-'}
      </p>
      <p className="text-xs text-slate-500 dark:text-slate-400">
        {signIn.createdDateTime ? new Date(signIn.createdDateTime).toLocaleTimeString() : '-'}
      </p>
    </div>

    {/* Failure Reason (if failed) */}
    {!signIn.isSuccess && signIn.failureReason && (
      <div className="hidden lg:block w-48">
        <p className="text-xs text-red-600 dark:text-red-400 truncate" title={signIn.failureReason}>
          {signIn.failureReason}
        </p>
      </div>
    )}
  </div>
);

export default SignInsPage;
