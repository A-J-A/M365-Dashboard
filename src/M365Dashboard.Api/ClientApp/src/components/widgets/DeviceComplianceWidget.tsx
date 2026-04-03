import React from 'react';
import { PieChart, Pie, Cell, ResponsiveContainer, Legend } from 'recharts';
import { WidgetCard } from '../common/WidgetCard';
import { useDeviceCompliance } from '../../hooks/useWidgetData';
import { useTheme } from '../../contexts/AppContext';

export function DeviceComplianceWidget() {
  const { data, isLoading, error, refresh } = useDeviceCompliance();
  const { resolvedTheme } = useTheme();

  const pieData = data ? [
    { name: 'Compliant', value: data.compliantDevices, color: '#22c55e' },
    { name: 'Non-compliant', value: data.nonCompliantDevices, color: '#ef4444' },
    { name: 'Unknown', value: data.unknownDevices, color: '#9ca3af' },
  ].filter(item => item.value > 0) : [];

  return (
    <WidgetCard
      title="Device Compliance"
      subtitle="Intune managed devices"
      isLoading={isLoading}
      error={error}
      onRefresh={refresh}
    >
      {data && (
        <div className="space-y-4">
          {/* Compliance rate */}
          <div className="text-center">
            <p className="text-3xl font-bold text-gray-900 dark:text-white">
              {data.complianceRate.toFixed(1)}%
            </p>
            <p className="text-sm text-gray-500 dark:text-gray-400">Compliance Rate</p>
          </div>

          {/* Pie chart */}
          <div className="h-32">
            <ResponsiveContainer width="100%" height="100%">
              <PieChart>
                <Pie
                  data={pieData}
                  cx="50%"
                  cy="50%"
                  innerRadius={25}
                  outerRadius={45}
                  paddingAngle={2}
                  dataKey="value"
                >
                  {pieData.map((entry, index) => (
                    <Cell key={`cell-${index}`} fill={entry.color} />
                  ))}
                </Pie>
              </PieChart>
            </ResponsiveContainer>
          </div>

          {/* Legend */}
          <div className="flex justify-center gap-4 text-xs">
            {pieData.map((item, idx) => (
              <div key={idx} className="flex items-center gap-1.5">
                <div className="w-2.5 h-2.5 rounded-full" style={{ backgroundColor: item.color }} />
                <span className="text-gray-600 dark:text-gray-400">
                  {item.name}: {item.value}
                </span>
              </div>
            ))}
          </div>

          {/* By platform */}
          {data.byPlatform.length > 0 && (
            <div className="pt-3 border-t border-gray-200 dark:border-gray-700">
              <p className="text-xs font-medium text-gray-500 dark:text-gray-400 mb-2">By Platform</p>
              <div className="space-y-1">
                {data.byPlatform.map((platform, idx) => (
                  <div key={idx} className="flex items-center justify-between text-xs">
                    <span className="text-gray-600 dark:text-gray-400">{platform.platform}</span>
                    <span className="font-medium">
                      <span className="text-green-600">{platform.compliant}</span>
                      <span className="text-gray-400 mx-1">/</span>
                      <span className="text-gray-900 dark:text-white">{platform.total}</span>
                    </span>
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
