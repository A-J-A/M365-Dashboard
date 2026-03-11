import React from 'react';
import {
  People24Regular,
  ShieldCheckmark24Regular,
  Key24Regular,
  Laptop24Regular,
} from '@fluentui/react-icons';
import type { DashboardSummary } from '../../types';

interface SummaryCardsProps {
  summary: DashboardSummary | null;
  isLoading: boolean;
}

export function SummaryCards({ summary, isLoading }: SummaryCardsProps) {
  const cards: Array<{
    title: string;
    value: string;
    subtitle: string;
    icon: React.ComponentType<{ className?: string }>;
    color: 'blue' | 'green' | 'purple' | 'orange';
  }> = [
    {
      title: 'Active Users',
      value: summary?.activeUsers.toLocaleString() ?? '-',
      subtitle: 'Monthly active users',
      icon: People24Regular,
      color: 'blue',
    },
    {
      title: 'Sign-in Success',
      value: summary ? `${summary.signInSuccessRate.toFixed(1)}%` : '-',
      subtitle: 'Last 7 days',
      icon: ShieldCheckmark24Regular,
      color: 'green',
    },
    {
      title: 'License Usage',
      value: summary ? `${summary.licenseUtilization.toFixed(1)}%` : '-',
      subtitle: 'Overall utilization',
      icon: Key24Regular,
      color: 'purple',
    },
    {
      title: 'Device Compliance',
      value: summary ? `${summary.deviceComplianceRate.toFixed(1)}%` : '-',
      subtitle: 'Compliant devices',
      icon: Laptop24Regular,
      color: 'orange',
    },
  ];

  return (
    <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
      {cards.map((card, index) => (
        <SummaryCard key={index} {...card} isLoading={isLoading} />
      ))}
    </div>
  );
}

interface SummaryCardProps {
  title: string;
  value: string;
  subtitle: string;
  icon: React.ComponentType<{ className?: string }>;
  color: 'blue' | 'green' | 'purple' | 'orange';
  isLoading: boolean;
}

const colorClasses = {
  blue: {
    bg: 'bg-blue-50 dark:bg-blue-900/30',
    icon: 'text-blue-600 dark:text-blue-400',
    ring: 'ring-blue-600/20',
  },
  green: {
    bg: 'bg-green-50 dark:bg-green-900/30',
    icon: 'text-green-600 dark:text-green-400',
    ring: 'ring-green-600/20',
  },
  purple: {
    bg: 'bg-purple-50 dark:bg-purple-900/30',
    icon: 'text-purple-600 dark:text-purple-400',
    ring: 'ring-purple-600/20',
  },
  orange: {
    bg: 'bg-orange-50 dark:bg-orange-900/30',
    icon: 'text-orange-600 dark:text-orange-400',
    ring: 'ring-orange-600/20',
  },
};

function SummaryCard({ title, value, subtitle, icon: Icon, color, isLoading }: SummaryCardProps) {
  const colors = colorClasses[color];

  return (
    <div className="bg-white dark:bg-gray-800 rounded-xl border border-gray-200 dark:border-gray-700 p-5 transition-shadow hover:shadow-md">
      <div className="flex items-start justify-between">
        <div className="flex-1 min-w-0">
          <p className="text-sm font-medium text-gray-500 dark:text-gray-400 truncate">
            {title}
          </p>
          {isLoading ? (
            <div className="mt-2 h-8 w-20 skeleton rounded" />
          ) : (
            <p className="mt-1 text-2xl font-bold text-gray-900 dark:text-white">
              {value}
            </p>
          )}
          <p className="mt-1 text-xs text-gray-400 dark:text-gray-500">
            {subtitle}
          </p>
        </div>
        <div className={`p-2.5 rounded-lg ${colors.bg}`}>
          <Icon className={`w-5 h-5 ${colors.icon}`} />
        </div>
      </div>
    </div>
  );
}
