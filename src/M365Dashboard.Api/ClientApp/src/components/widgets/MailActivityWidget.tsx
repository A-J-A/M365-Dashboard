import React from 'react';
import { AreaChart, Area, XAxis, YAxis, Tooltip, ResponsiveContainer } from 'recharts';
import { WidgetCard, Metric } from '../common/WidgetCard';
import { useMailActivity } from '../../hooks/useWidgetData';
import { useTheme } from '../../contexts/AppContext';

export function MailActivityWidget() {
  const { data, isLoading, error, refresh } = useMailActivity();
  const { resolvedTheme } = useTheme();

  const chartData = data?.trend.slice(-14).map(item => ({
    date: new Date(item.date).toLocaleDateString('en-GB', { day: '2-digit', month: 'short' }),
    sent: item.sent,
    received: item.received,
  })) ?? [];

  const gridColor = resolvedTheme === 'dark' ? '#374151' : '#e5e7eb';
  const textColor = resolvedTheme === 'dark' ? '#9ca3af' : '#6b7280';

  const formatNumber = (num: number): string => {
    if (num >= 1000000) return `${(num / 1000000).toFixed(1)}M`;
    if (num >= 1000) return `${(num / 1000).toFixed(1)}K`;
    return num.toString();
  };

  return (
    <WidgetCard
      title="Mail Activity"
      subtitle="Email trends"
      isLoading={isLoading}
      error={error}
      onRefresh={refresh}
    >
      {data && (
        <div className="space-y-4">
          {/* Metrics */}
          <div className="grid grid-cols-2 gap-4">
            <div className="text-center p-3 bg-blue-50 dark:bg-blue-900/20 rounded-lg">
              <p className="text-xl font-bold text-blue-600 dark:text-blue-400">
                {formatNumber(data.totalEmailsSent)}
              </p>
              <p className="text-xs text-gray-500 dark:text-gray-400">Sent</p>
            </div>
            <div className="text-center p-3 bg-green-50 dark:bg-green-900/20 rounded-lg">
              <p className="text-xl font-bold text-green-600 dark:text-green-400">
                {formatNumber(data.totalEmailsReceived)}
              </p>
              <p className="text-xs text-gray-500 dark:text-gray-400">Received</p>
            </div>
          </div>

          {/* Chart */}
          <div className="h-36">
            <ResponsiveContainer width="100%" height="100%">
              <AreaChart data={chartData}>
                <defs>
                  <linearGradient id="colorSent" x1="0" y1="0" x2="0" y2="1">
                    <stop offset="5%" stopColor="#0078d4" stopOpacity={0.3}/>
                    <stop offset="95%" stopColor="#0078d4" stopOpacity={0}/>
                  </linearGradient>
                  <linearGradient id="colorReceived" x1="0" y1="0" x2="0" y2="1">
                    <stop offset="5%" stopColor="#22c55e" stopOpacity={0.3}/>
                    <stop offset="95%" stopColor="#22c55e" stopOpacity={0}/>
                  </linearGradient>
                </defs>
                <XAxis 
                  dataKey="date" 
                  tick={{ fontSize: 10, fill: textColor }}
                  axisLine={{ stroke: gridColor }}
                  tickLine={false}
                  interval="preserveStartEnd"
                />
                <YAxis hide />
                <Tooltip
                  contentStyle={{
                    backgroundColor: resolvedTheme === 'dark' ? '#1f2937' : '#ffffff',
                    border: `1px solid ${gridColor}`,
                    borderRadius: '8px',
                  }}
                  formatter={(value: number) => formatNumber(value)}
                />
                <Area
                  type="monotone"
                  dataKey="sent"
                  stroke="#0078d4"
                  strokeWidth={2}
                  fillOpacity={1}
                  fill="url(#colorSent)"
                />
                <Area
                  type="monotone"
                  dataKey="received"
                  stroke="#22c55e"
                  strokeWidth={2}
                  fillOpacity={1}
                  fill="url(#colorReceived)"
                />
              </AreaChart>
            </ResponsiveContainer>
          </div>

          {/* Legend */}
          <div className="flex justify-center gap-6 text-xs">
            <div className="flex items-center gap-1.5">
              <div className="w-2.5 h-2.5 rounded-full bg-blue-600" />
              <span className="text-gray-600 dark:text-gray-400">Sent</span>
            </div>
            <div className="flex items-center gap-1.5">
              <div className="w-2.5 h-2.5 rounded-full bg-green-600" />
              <span className="text-gray-600 dark:text-gray-400">Received</span>
            </div>
          </div>
        </div>
      )}
    </WidgetCard>
  );
}
