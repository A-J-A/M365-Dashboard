import React from 'react';
import { LineChart, Line, XAxis, YAxis, Tooltip, ResponsiveContainer } from 'recharts';
import { WidgetCard, Metric } from '../common/WidgetCard';
import { useActiveUsers } from '../../hooks/useWidgetData';
import { useTheme } from '../../contexts/AppContext';

export function ActiveUsersWidget() {
  const { data, isLoading, error, refresh } = useActiveUsers();
  const { resolvedTheme } = useTheme();

  const chartData = data?.trend.slice(-14).map(item => ({
    date: new Date(item.date).toLocaleDateString('en-GB', { day: '2-digit', month: 'short' }),
    users: item.count,
  })) ?? [];

  const gridColor = resolvedTheme === 'dark' ? '#374151' : '#e5e7eb';
  const textColor = resolvedTheme === 'dark' ? '#9ca3af' : '#6b7280';

  return (
    <WidgetCard
      title="Active Users"
      subtitle="User activity trends"
      isLoading={isLoading}
      error={error}
      onRefresh={refresh}
    >
      {data && (
        <div className="space-y-4">
          {/* Metrics row */}
          <div className="grid grid-cols-3 gap-4">
            <Metric label="Daily" value={data.dailyActiveUsers} size="sm" />
            <Metric label="Weekly" value={data.weeklyActiveUsers} size="sm" />
            <Metric label="Monthly" value={data.monthlyActiveUsers} size="sm" />
          </div>

          {/* Chart */}
          <div className="h-40">
            <ResponsiveContainer width="100%" height="100%">
              <LineChart data={chartData}>
                <XAxis 
                  dataKey="date" 
                  tick={{ fontSize: 10, fill: textColor }}
                  axisLine={{ stroke: gridColor }}
                  tickLine={false}
                />
                <YAxis 
                  hide 
                  domain={['dataMin - 10', 'dataMax + 10']}
                />
                <Tooltip
                  contentStyle={{
                    backgroundColor: resolvedTheme === 'dark' ? '#1f2937' : '#ffffff',
                    border: `1px solid ${gridColor}`,
                    borderRadius: '8px',
                    boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)',
                  }}
                  labelStyle={{ color: resolvedTheme === 'dark' ? '#fff' : '#111' }}
                />
                <Line
                  type="monotone"
                  dataKey="users"
                  stroke="#0078d4"
                  strokeWidth={2}
                  dot={false}
                  activeDot={{ r: 4, fill: '#0078d4' }}
                />
              </LineChart>
            </ResponsiveContainer>
          </div>
        </div>
      )}
    </WidgetCard>
  );
}
