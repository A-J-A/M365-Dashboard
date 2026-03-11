import React from 'react';
import { PieChart, Pie, Cell, ResponsiveContainer, Tooltip } from 'recharts';
import { WidgetCard, ProgressBar } from '../common/WidgetCard';
import { useLicenseUsage } from '../../hooks/useWidgetData';
import { useTheme } from '../../contexts/AppContext';

const COLORS = ['#0078d4', '#00bcf2', '#5c2d91', '#00b294', '#ff8c00', '#e81123'];

export function LicenseUsageWidget() {
  const { data, isLoading, error, refresh } = useLicenseUsage();
  const { resolvedTheme } = useTheme();

  const pieData = data?.licenses.slice(0, 5).map((license, idx) => ({
    name: license.skuName,
    value: license.consumedUnits,
    color: COLORS[idx % COLORS.length],
  })) ?? [];

  return (
    <WidgetCard
      title="License Usage"
      subtitle="Subscription utilization"
      isLoading={isLoading}
      error={error}
      onRefresh={refresh}
    >
      {data && (
        <div className="space-y-4">
          {/* Overall utilization */}
          <div className="flex items-center justify-between">
            <span className="text-sm text-gray-600 dark:text-gray-400">Overall Utilization</span>
            <span className="text-lg font-bold text-gray-900 dark:text-white">
              {data.overallUtilization.toFixed(1)}%
            </span>
          </div>
          <ProgressBar 
            value={data.overallUtilization} 
            color={data.overallUtilization >= 90 ? 'orange' : 'blue'} 
            showLabel={false}
          />

          {/* Summary */}
          <div className="flex items-center justify-between text-sm py-2">
            <span className="text-gray-500 dark:text-gray-400">
              {data.totalConsumed.toLocaleString()} / {data.totalAvailable.toLocaleString()} licenses
            </span>
          </div>

          {/* License breakdown */}
          <div className="space-y-3">
            {data.licenses.slice(0, 4).map((license, idx) => (
              <div key={license.skuId} className="space-y-1">
                <div className="flex items-center justify-between text-xs">
                  <span className="text-gray-600 dark:text-gray-400 truncate flex-1 mr-2">
                    {formatSkuName(license.skuName)}
                  </span>
                  <span className="text-gray-900 dark:text-white font-medium whitespace-nowrap">
                    {license.consumedUnits} / {license.prepaidUnits}
                  </span>
                </div>
                <div className="w-full bg-gray-200 dark:bg-gray-700 rounded-full h-1.5">
                  <div
                    className="h-1.5 rounded-full transition-all duration-300"
                    style={{ 
                      width: `${Math.min(license.utilizationPercent, 100)}%`,
                      backgroundColor: COLORS[idx % COLORS.length]
                    }}
                  />
                </div>
              </div>
            ))}
          </div>

          {data.licenses.length > 4 && (
            <p className="text-xs text-gray-500 dark:text-gray-400 text-center">
              +{data.licenses.length - 4} more licenses
            </p>
          )}
        </div>
      )}
    </WidgetCard>
  );
}

// Helper to format SKU names into readable format
function formatSkuName(skuName: string): string {
  const nameMap: Record<string, string> = {
    'ENTERPRISEPACK': 'Office 365 E3',
    'ENTERPRISEPREMIUM': 'Office 365 E5',
    'SPE_E3': 'Microsoft 365 E3',
    'SPE_E5': 'Microsoft 365 E5',
    'EMS': 'Enterprise Mobility + Security',
    'EMSPREMIUM': 'EMS E5',
    'AAD_PREMIUM': 'Azure AD Premium P1',
    'AAD_PREMIUM_P2': 'Azure AD Premium P2',
    'POWER_BI_PRO': 'Power BI Pro',
    'PROJECTPREMIUM': 'Project Plan 5',
    'VISIOCLIENT': 'Visio Plan 2',
  };

  return nameMap[skuName] || skuName.replace(/_/g, ' ');
}
