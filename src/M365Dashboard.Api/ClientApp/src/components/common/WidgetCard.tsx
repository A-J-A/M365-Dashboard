import React from 'react';
import { Spinner, Text } from '@fluentui/react-components';
import { ArrowSync20Regular, Warning20Regular } from '@fluentui/react-icons';

interface WidgetCardProps {
  title: string;
  subtitle?: string;
  isLoading?: boolean;
  error?: string | null;
  onRefresh?: () => void;
  children: React.ReactNode;
  className?: string;
}

export function WidgetCard({
  title,
  subtitle,
  isLoading,
  error,
  onRefresh,
  children,
  className = '',
}: WidgetCardProps) {
  return (
    <div className={`widget-card p-5 h-full flex flex-col ${className}`}>
      {/* Header */}
      <div className="flex items-start justify-between mb-4">
        <div className="flex-1 min-w-0">
          <h3 className="text-base font-semibold text-gray-900 dark:text-white truncate">
            {title}
          </h3>
          {subtitle && (
            <p className="text-xs text-gray-500 dark:text-gray-400 mt-0.5">
              {subtitle}
            </p>
          )}
        </div>
        {onRefresh && (
          <button
            onClick={onRefresh}
            disabled={isLoading}
            className="p-1.5 text-gray-400 hover:text-gray-600 dark:hover:text-gray-300 rounded-lg hover:bg-gray-100 dark:hover:bg-gray-700 transition-colors disabled:opacity-50"
            aria-label="Refresh"
          >
            <ArrowSync20Regular className={isLoading ? 'animate-spin' : ''} />
          </button>
        )}
      </div>

      {/* Content */}
      <div className="flex-1 min-h-0">
        {isLoading && !error ? (
          <div className="flex items-center justify-center h-full min-h-[120px]">
            <Spinner size="small" />
          </div>
        ) : error ? (
          <div className="flex flex-col items-center justify-center h-full min-h-[120px] text-center">
            <Warning20Regular className="w-8 h-8 text-red-500 mb-2" />
            <Text className="text-sm text-red-600 dark:text-red-400">{error}</Text>
            {onRefresh && (
              <button
                onClick={onRefresh}
                className="mt-2 text-sm text-blue-600 dark:text-blue-400 hover:underline"
              >
                Try again
              </button>
            )}
          </div>
        ) : (
          children
        )}
      </div>
    </div>
  );
}

// Metric display component
interface MetricProps {
  label: string;
  value: string | number;
  trend?: number;
  size?: 'sm' | 'md' | 'lg';
}

export function Metric({ label, value, trend, size = 'md' }: MetricProps) {
  const sizeClasses = {
    sm: { value: 'text-lg', label: 'text-xs' },
    md: { value: 'text-2xl', label: 'text-sm' },
    lg: { value: 'text-3xl', label: 'text-sm' },
  };

  return (
    <div>
      <p className={`font-bold text-gray-900 dark:text-white ${sizeClasses[size].value}`}>
        {typeof value === 'number' ? value.toLocaleString() : value}
        {trend !== undefined && (
          <span className={`ml-2 text-sm font-medium ${trend >= 0 ? 'text-green-600' : 'text-red-600'}`}>
            {trend >= 0 ? '↑' : '↓'} {Math.abs(trend)}%
          </span>
        )}
      </p>
      <p className={`text-gray-500 dark:text-gray-400 ${sizeClasses[size].label}`}>
        {label}
      </p>
    </div>
  );
}

// Progress bar component
interface ProgressBarProps {
  value: number;
  max?: number;
  color?: 'blue' | 'green' | 'red' | 'orange' | 'purple';
  showLabel?: boolean;
  size?: 'sm' | 'md';
}

const progressColors = {
  blue: 'bg-blue-600',
  green: 'bg-green-600',
  red: 'bg-red-600',
  orange: 'bg-orange-500',
  purple: 'bg-purple-600',
};

export function ProgressBar({ value, max = 100, color = 'blue', showLabel = true, size = 'md' }: ProgressBarProps) {
  const percentage = Math.min(Math.max((value / max) * 100, 0), 100);
  const heightClass = size === 'sm' ? 'h-1.5' : 'h-2.5';

  return (
    <div className="w-full">
      <div className={`w-full bg-gray-200 dark:bg-gray-700 rounded-full ${heightClass} overflow-hidden`}>
        <div
          className={`${progressColors[color]} ${heightClass} rounded-full transition-all duration-300`}
          style={{ width: `${percentage}%` }}
        />
      </div>
      {showLabel && (
        <p className="mt-1 text-xs text-gray-500 dark:text-gray-400 text-right">
          {percentage.toFixed(1)}%
        </p>
      )}
    </div>
  );
}
