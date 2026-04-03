import React from 'react';
import { BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer } from 'recharts';
import { WidgetCard } from '../common/WidgetCard';
import { useTeamsActivity } from '../../hooks/useWidgetData';
import { useTheme } from '../../contexts/AppContext';
import {
  Chat24Regular,
  Call24Regular,
  Video24Regular,
  People24Regular,
} from '@fluentui/react-icons';

export function TeamsActivityWidget() {
  const { data, isLoading, error, refresh } = useTeamsActivity();
  const { resolvedTheme } = useTheme();

  const chartData = data?.trend.slice(-7).map(item => ({
    date: new Date(item.date).toLocaleDateString('en-GB', { weekday: 'short' }),
    messages: item.messages,
    calls: item.calls,
    meetings: item.meetings,
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
      title="Teams Activity"
      subtitle="Collaboration metrics"
      isLoading={isLoading}
      error={error}
      onRefresh={refresh}
    >
      {data && (
        <div className="space-y-4">
          {/* Metric cards */}
          <div className="grid grid-cols-2 gap-2">
            <MetricCard 
              icon={Chat24Regular} 
              label="Messages" 
              value={formatNumber(data.totalMessages)} 
              color="purple"
            />
            <MetricCard 
              icon={Call24Regular} 
              label="Calls" 
              value={formatNumber(data.totalCalls)} 
              color="blue"
            />
            <MetricCard 
              icon={Video24Regular} 
              label="Meetings" 
              value={formatNumber(data.totalMeetings)} 
              color="green"
            />
            <MetricCard 
              icon={People24Regular} 
              label="Active Users" 
              value={data.activeUsers.toLocaleString()} 
              color="orange"
            />
          </div>

          {/* Chart */}
          <div className="h-32">
            <ResponsiveContainer width="100%" height="100%">
              <BarChart data={chartData} barSize={8}>
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
                  formatter={(value: number) => formatNumber(value)}
                />
                <Bar dataKey="messages" fill="#8b5cf6" radius={[4, 4, 0, 0]} />
                <Bar dataKey="calls" fill="#0078d4" radius={[4, 4, 0, 0]} />
                <Bar dataKey="meetings" fill="#22c55e" radius={[4, 4, 0, 0]} />
              </BarChart>
            </ResponsiveContainer>
          </div>

          {/* Legend */}
          <div className="flex justify-center gap-4 text-xs">
            <div className="flex items-center gap-1.5">
              <div className="w-2 h-2 rounded-sm bg-purple-500" />
              <span className="text-gray-500 dark:text-gray-400">Messages</span>
            </div>
            <div className="flex items-center gap-1.5">
              <div className="w-2 h-2 rounded-sm bg-blue-600" />
              <span className="text-gray-500 dark:text-gray-400">Calls</span>
            </div>
            <div className="flex items-center gap-1.5">
              <div className="w-2 h-2 rounded-sm bg-green-500" />
              <span className="text-gray-500 dark:text-gray-400">Meetings</span>
            </div>
          </div>
        </div>
      )}
    </WidgetCard>
  );
}

interface MetricCardProps {
  icon: React.ComponentType<{ className?: string }>;
  label: string;
  value: string;
  color: 'purple' | 'blue' | 'green' | 'orange';
}

const colorClasses = {
  purple: 'bg-purple-50 dark:bg-purple-900/20 text-purple-600 dark:text-purple-400',
  blue: 'bg-blue-50 dark:bg-blue-900/20 text-blue-600 dark:text-blue-400',
  green: 'bg-green-50 dark:bg-green-900/20 text-green-600 dark:text-green-400',
  orange: 'bg-orange-50 dark:bg-orange-900/20 text-orange-600 dark:text-orange-400',
};

function MetricCard({ icon: Icon, label, value, color }: MetricCardProps) {
  return (
    <div className={`p-2.5 rounded-lg ${colorClasses[color]}`}>
      <div className="flex items-center gap-2">
        <Icon className="w-4 h-4" />
        <div>
          <p className="text-sm font-bold">{value}</p>
          <p className="text-xs opacity-75">{label}</p>
        </div>
      </div>
    </div>
  );
}
