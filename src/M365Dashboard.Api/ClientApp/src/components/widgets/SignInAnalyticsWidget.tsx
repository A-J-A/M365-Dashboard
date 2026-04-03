import React from 'react';
import { BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, Cell } from 'recharts';
import { WidgetCard, Metric, ProgressBar } from '../common/WidgetCard';
import { useSignInAnalytics } from '../../hooks/useWidgetData';
import { useTheme } from '../../contexts/AppContext';

export function SignInAnalyticsWidget() {
  const { data, isLoading, error, refresh } = useSignInAnalytics();
  const { resolvedTheme } = useTheme();

  const chartData = data?.trend.slice(-7).map(item => ({
    date: new Date(item.date).toLocaleDateString('en-GB', { weekday: 'short' }),
    successful: item.successful,
    failed: item.failed,
  })) ?? [];

  const gridColor = resolvedTheme === 'dark' ? '#374151' : '#e5e7eb';
  const textColor = resolvedTheme === 'dark' ? '#9ca3af' : '#6b7280';

  return (
    <WidgetCard
      title="Sign-in Analytics"
      subtitle="Authentication activity"
      isLoading={isLoading}
      error={error}
      onRefresh={refresh}
    >
      {data && (
        <div className="space-y-4">
          {/* Success rate */}
          <div className="flex items-center justify-between">
            <span className="text-sm text-gray-600 dark:text-gray-400">Success Rate</span>
            <span className="text-lg font-bold text-gray-900 dark:text-white">
              {data.successRate.toFixed(1)}%
            </span>
          </div>
          <ProgressBar 
            value={data.successRate} 
            color={data.successRate >= 95 ? 'green' : data.successRate >= 90 ? 'orange' : 'red'} 
            showLabel={false}
          />

          {/* Stats grid */}
          <div className="grid grid-cols-3 gap-3 py-2">
            <div className="text-center">
              <p className="text-lg font-bold text-green-600">{data.successfulSignIns.toLocaleString()}</p>
              <p className="text-xs text-gray-500 dark:text-gray-400">Successful</p>
            </div>
            <div className="text-center">
              <p className="text-lg font-bold text-red-600">{data.failedSignIns.toLocaleString()}</p>
              <p className="text-xs text-gray-500 dark:text-gray-400">Failed</p>
            </div>
            <div className="text-center">
              <p className="text-lg font-bold text-orange-500">{data.riskySignIns.toLocaleString()}</p>
              <p className="text-xs text-gray-500 dark:text-gray-400">Risky</p>
            </div>
          </div>

          {/* Chart */}
          <div className="h-32">
            <ResponsiveContainer width="100%" height="100%">
              <BarChart data={chartData} barGap={2}>
                <XAxis 
                  dataKey="date" 
                  tick={{ fontSize: 10, fill: textColor }}
                  axisLine={{ stroke: gridColor }}
                  tickLine={false}
                />
                <YAxis hide />
                <Tooltip
                  contentStyle={{
                    backgroundColor: resolvedTheme === 'dark' ? '#1f2937' : '#ffffff',
                    border: `1px solid ${gridColor}`,
                    borderRadius: '8px',
                  }}
                />
                <Bar dataKey="successful" stackId="a" fill="#22c55e" radius={[4, 4, 0, 0]} />
                <Bar dataKey="failed" stackId="a" fill="#ef4444" radius={[4, 4, 0, 0]} />
              </BarChart>
            </ResponsiveContainer>
          </div>

          {/* Top locations */}
          {data.topLocations.length > 0 && (
            <div className="pt-2 border-t border-gray-200 dark:border-gray-700">
              <p className="text-xs font-medium text-gray-500 dark:text-gray-400 mb-2">Top Locations</p>
              <div className="space-y-1">
                {data.topLocations.slice(0, 3).map((location, idx) => (
                  <div key={idx} className="flex items-center justify-between text-xs">
                    <span className="text-gray-600 dark:text-gray-400 truncate">{location.location}</span>
                    <span className="font-medium text-gray-900 dark:text-white">{location.count.toLocaleString()}</span>
                  </div>
                ))}
              </div>
            </div>
          )}
        </div>
      )}
    </WidgetCard>
  );
}
