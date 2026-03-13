import React, { createContext, useContext, useState, useEffect, useCallback } from 'react';
import type { UserSettings, WidgetConfiguration, UserProfile } from '../types';
import { settingsApi, userApi } from '../services/api';
import { useMsal, useIsAuthenticated } from '@azure/msal-react';

// Theme Context
type Theme = 'light' | 'dark' | 'system';

interface ThemeContextType {
  theme: Theme;
  resolvedTheme: 'light' | 'dark';
  setTheme: (theme: Theme) => void;
}

const ThemeContext = createContext<ThemeContextType | undefined>(undefined);

export function ThemeProvider({ children }: { children: React.ReactNode }) {
  const [theme, setThemeState] = useState<Theme>(() => {
    // Load saved theme from localStorage
    const saved = localStorage.getItem('theme') as Theme;
    return saved || 'system';
  });
  const [resolvedTheme, setResolvedTheme] = useState<'light' | 'dark'>('light');

  const setTheme = useCallback((newTheme: Theme) => {
    setThemeState(newTheme);
    localStorage.setItem('theme', newTheme);
  }, []);

  useEffect(() => {
    const updateResolvedTheme = () => {
      if (theme === 'system') {
        const isDark = window.matchMedia('(prefers-color-scheme: dark)').matches;
        setResolvedTheme(isDark ? 'dark' : 'light');
      } else {
        setResolvedTheme(theme);
      }
    };

    updateResolvedTheme();

    const mediaQuery = window.matchMedia('(prefers-color-scheme: dark)');
    mediaQuery.addEventListener('change', updateResolvedTheme);

    return () => mediaQuery.removeEventListener('change', updateResolvedTheme);
  }, [theme]);

  useEffect(() => {
    document.documentElement.classList.toggle('dark', resolvedTheme === 'dark');
  }, [resolvedTheme]);

  return (
    <ThemeContext.Provider value={{ theme, resolvedTheme, setTheme }}>
      {children}
    </ThemeContext.Provider>
  );
}

export function useTheme() {
  const context = useContext(ThemeContext);
  if (!context) {
    throw new Error('useTheme must be used within a ThemeProvider');
  }
  return context;
}

// App Context (for accessing tokens)
import { getAccessToken as getTokenFromApi } from '../services/api';

interface AppContextType {
  getAccessToken: () => Promise<string>;
}

const AppContext = createContext<AppContextType | undefined>(undefined);

export function useAppContext() {
  const context = useContext(AppContext);
  if (!context) {
    throw new Error('useAppContext must be used within an AppProvider');
  }
  return context;
}

// User Context
interface UserContextType {
  profile: UserProfile | null;
  isLoading: boolean;
  error: string | null;
  isAdmin: boolean;
  refreshProfile: () => Promise<void>;
}

const UserContext = createContext<UserContextType | undefined>(undefined);

export function UserProvider({ children }: { children: React.ReactNode }) {
  const { instance, accounts } = useMsal();
  const isAuthenticated = useIsAuthenticated();
  const [profile, setProfile] = useState<UserProfile | null>(null);
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const getAccessToken = useCallback(async (): Promise<string> => {
    return getTokenFromApi();
  }, []);

  const refreshProfile = useCallback(async () => {
    if (!isAuthenticated || accounts.length === 0) {
      setProfile(null);
      return;
    }

    setIsLoading(true);
    setError(null);

    try {
      const userProfile = await userApi.getProfile();
      setProfile(userProfile);
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Failed to load profile');
    } finally {
      setIsLoading(false);
    }
  }, [isAuthenticated, accounts]);

  useEffect(() => {
    refreshProfile();
  }, [refreshProfile]);

  const isAdmin = profile?.roles.includes('Dashboard.Admin') ?? false;

  return (
    <AppContext.Provider value={{ getAccessToken }}>
      <UserContext.Provider value={{ profile, isLoading, error, isAdmin, refreshProfile }}>
        {children}
      </UserContext.Provider>
    </AppContext.Provider>
  );
}

export function useUser() {
  const context = useContext(UserContext);
  if (!context) {
    throw new Error('useUser must be used within a UserProvider');
  }
  return context;
}

// Settings Context
interface SettingsContextType {
  settings: UserSettings | null;
  widgets: WidgetConfiguration[];
  isLoading: boolean;
  error: string | null;
  updateSettings: (settings: Partial<UserSettings>) => Promise<void>;
  updateWidget: (widgetType: string, config: Partial<WidgetConfiguration>) => Promise<void>;
  resetWidgets: () => Promise<void>;
  refreshSettings: () => Promise<void>;
}

const SettingsContext = createContext<SettingsContextType | undefined>(undefined);

export function SettingsProvider({ children }: { children: React.ReactNode }) {
  const isAuthenticated = useIsAuthenticated();
  const { setTheme } = useTheme();
  const [settings, setSettings] = useState<UserSettings | null>(null);
  const [widgets, setWidgets] = useState<WidgetConfiguration[]>([]);
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const refreshSettings = useCallback(async () => {
    if (!isAuthenticated) {
      setSettings(null);
      setWidgets([]);
      return;
    }

    setIsLoading(true);
    setError(null);

    try {
      const [userSettings, widgetConfigs] = await Promise.all([
        settingsApi.getSettings(),
        settingsApi.getWidgets(),
      ]);
      
      setSettings(userSettings);
      setWidgets(widgetConfigs);
      setTheme(userSettings.theme);
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Failed to load settings');
    } finally {
      setIsLoading(false);
    }
  }, [isAuthenticated, setTheme]);

  useEffect(() => {
    refreshSettings();
  }, [refreshSettings]);

  const updateSettings = async (newSettings: Partial<UserSettings>) => {
    try {
      const updated = await settingsApi.updateSettings(newSettings);
      setSettings(updated);
      if (newSettings.theme) {
        setTheme(newSettings.theme);
      }
    } catch (err) {
      throw err;
    }
  };

  const updateWidget = async (widgetType: string, config: Partial<WidgetConfiguration>) => {
    try {
      const updated = await settingsApi.updateWidget(widgetType, config);
      setWidgets(prev => 
        prev.map(w => w.widgetType === widgetType ? updated : w)
      );
    } catch (err) {
      throw err;
    }
  };

  const resetWidgets = async () => {
    try {
      const defaultWidgets = await settingsApi.resetWidgets();
      setWidgets(defaultWidgets);
    } catch (err) {
      throw err;
    }
  };

  return (
    <SettingsContext.Provider value={{
      settings,
      widgets,
      isLoading,
      error,
      updateSettings,
      updateWidget,
      resetWidgets,
      refreshSettings,
    }}>
      {children}
    </SettingsContext.Provider>
  );
}

export function useSettings() {
  const context = useContext(SettingsContext);
  if (!context) {
    throw new Error('useSettings must be used within a SettingsProvider');
  }
  return context;
}
