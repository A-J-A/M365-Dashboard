import React, { useState } from 'react';
import { Link, useLocation } from 'react-router-dom';
import { useMsal } from '@azure/msal-react';
import {
  Button,
  Avatar,
  Menu,
  MenuTrigger,
  MenuList,
  MenuItem,
  MenuPopover,
  Tooltip,
} from '@fluentui/react-components';
import {
  Home24Regular,
  Home24Filled,
  Settings24Regular,
  Settings24Filled,
  SignOut24Regular,
  WeatherMoon24Regular,
  WeatherSunny24Regular,
  Navigation24Regular,
  PersonAccounts24Regular,
  PersonAccounts24Filled,
  ShieldCheckmark24Regular,
  ShieldCheckmark24Filled,
  Mail24Regular,
  Mail24Filled,
  People24Regular,
  People24Filled,
  Laptop24Regular,
  Laptop24Filled,
  DocumentBulletList24Regular,
  DocumentBulletList24Filled,
  Grid24Regular,
  ChevronLeft24Regular,
  ChevronRight24Regular,
  Location24Regular,
  Location24Filled,
  ShareScreenStart24Regular,
  ShareScreenStart24Filled,
  Key24Regular,
  Key24Filled,
  Call24Regular,
  Call24Filled,
  Shield24Regular,
  Shield24Filled,
  ShieldError24Regular,
  ShieldError24Filled,

  MailInbox24Regular,
  MailInbox24Filled,
  Apps24Regular,
  Apps24Filled,
  PersonKey24Regular,
  PersonKey24Filled,
  DocumentText24Regular,
  DocumentText24Filled,
} from '@fluentui/react-icons';
import { useUser, useTheme } from '../../contexts/AppContext';

interface LayoutProps {
  children: React.ReactNode;
}

export function Layout({ children }: LayoutProps) {
  const location = useLocation();
  const { instance } = useMsal();
  const { profile } = useUser();
  const { resolvedTheme, setTheme } = useTheme();
  const [isCollapsed, setIsCollapsed] = useState(false);

  const handleLogout = () => {
    instance.logoutRedirect();
  };

  const toggleTheme = () => {
    setTheme(resolvedTheme === 'dark' ? 'light' : 'dark');
  };

  const navItems = [
    { 
      to: '/', 
      icon: Home24Regular, 
      iconFilled: Home24Filled, 
      label: 'Home',
      section: 'main'
    },
    { 
      to: '/users', 
      icon: PersonAccounts24Regular, 
      iconFilled: PersonAccounts24Filled, 
      label: 'Users',
      section: 'main'
    },
    { 
      to: '/teams', 
      icon: People24Regular, 
      iconFilled: People24Filled, 
      label: 'Teams & Groups',
      section: 'main'
    },
    { 
      to: '/devices', 
      icon: Laptop24Regular, 
      iconFilled: Laptop24Filled, 
      label: 'Devices',
      section: 'main'
    },
    { 
      to: '/mailflow', 
      icon: Mail24Regular, 
      iconFilled: Mail24Filled, 
      label: 'Exchange Online',
      section: 'main'
    },
    { 
      to: '/security', 
      icon: ShieldCheckmark24Regular, 
      iconFilled: ShieldCheckmark24Filled, 
      label: 'Security',
      section: 'main'
    },
    { 
      to: '/signins', 
      icon: Location24Regular, 
      iconFilled: Location24Filled, 
      label: 'Sign-ins Map',
      section: 'main'
    },
    { 
      to: '/teamsphone', 
      icon: Call24Regular, 
      iconFilled: Call24Filled, 
      label: 'Teams Phone',
      section: 'main'
    },
    { 
      to: '/conditional-access', 
      icon: Shield24Regular, 
      iconFilled: Shield24Filled, 
      label: 'Conditional Access',
      section: 'security'
    },
    { 
      to: '/privileged-access', 
      icon: PersonKey24Regular, 
      iconFilled: PersonKey24Filled, 
      label: 'Privileged Access',
      section: 'security'
    },
    { 
      to: '/threat-intelligence', 
      icon: ShieldError24Regular, 
      iconFilled: ShieldError24Filled, 
      label: 'Threat Intel',
      section: 'security'
    },
    { 
      to: '/app-consent', 
      icon: Apps24Regular, 
      iconFilled: Apps24Filled, 
      label: 'Applications',
      section: 'security'
    },
    { 
      to: '/defender-office', 
      icon: ShieldCheckmark24Regular, 
      iconFilled: ShieldCheckmark24Filled, 
      label: 'Defender for Office',
      section: 'security'
    },
    { 
      to: '/sharepoint', 
      icon: ShareScreenStart24Regular, 
      iconFilled: ShareScreenStart24Filled, 
      label: 'SharePoint',
      section: 'main'
    },
    { 
      to: '/licenses', 
      icon: Key24Regular, 
      iconFilled: Key24Filled, 
      label: 'Licenses',
      section: 'main'
    },
    { 
      to: '/reports', 
      icon: DocumentBulletList24Regular, 
      iconFilled: DocumentBulletList24Filled, 
      label: 'Reports',
      section: 'main'
    },
    { 
      to: '/executive-report', 
      icon: DocumentText24Regular, 
      iconFilled: DocumentText24Filled, 
      label: 'Executive Summary',
      section: 'main'
    },
    { 
      to: '/cis-benchmark', 
      icon: ShieldCheckmark24Regular, 
      iconFilled: ShieldCheckmark24Filled, 
      label: 'CIS Benchmark',
      section: 'security'
    },
    { 
      to: '/security-assessment', 
      icon: DocumentText24Regular, 
      iconFilled: DocumentText24Filled, 
      label: 'Security Assessment',
      section: 'security'
    },
    { 
      to: '/settings', 
      icon: Settings24Regular, 
      iconFilled: Settings24Filled, 
      label: 'Settings',
      section: 'bottom'
    },
  ];

  const mainNavItems = navItems.filter(item => item.section === 'main');
  const securityNavItems = navItems.filter(item => item.section === 'security');
  const bottomNavItems = navItems.filter(item => item.section === 'bottom');

  return (
    <div className="min-h-screen bg-gray-100 dark:bg-gray-900 flex">
      {/* Left Sidebar */}
      <aside 
        className={`
          fixed top-0 left-0 h-full bg-[#0f172a] dark:bg-gray-950 text-white z-50
          flex flex-col transition-all duration-300 ease-in-out
          ${isCollapsed ? 'w-16' : 'w-64'}
        `}
      >
        {/* Logo */}
        <div className={`h-12 flex items-center border-b border-gray-700 ${isCollapsed ? 'justify-center px-2' : 'px-4'}`}>
          <Link to="/" className="flex items-center gap-3 text-white hover:opacity-90">
            <div className="w-8 h-8 bg-blue-600 rounded flex items-center justify-center flex-shrink-0">
              <Grid24Regular className="w-5 h-5" />
            </div>
            {!isCollapsed && (
              <span className="font-semibold text-sm">M365 Dashboard</span>
            )}
          </Link>
        </div>

        {/* Main Navigation */}
        <nav className="flex-1 overflow-y-auto py-2">
          <ul className="space-y-1 px-2">
            {mainNavItems.map((item) => (
              <NavItem 
                key={item.to}
                {...item}
                active={location.pathname === item.to || (item.to === '/' && location.pathname === '/')}
                collapsed={isCollapsed}
              />
            ))}
          </ul>
          
          {/* Security Section */}
          {securityNavItems.length > 0 && (
            <>
              <div className={`mt-4 mb-2 ${isCollapsed ? 'px-2' : 'px-4'}`}>
                {!isCollapsed && (
                  <span className="text-xs font-semibold text-gray-500 uppercase tracking-wider">Security & SOC</span>
                )}
                {isCollapsed && <div className="border-t border-gray-700" />}
              </div>
              <ul className="space-y-1 px-2">
                {securityNavItems.map((item) => (
                  <NavItem 
                    key={item.to}
                    {...item}
                    active={location.pathname === item.to}
                    collapsed={isCollapsed}
                  />
                ))}
              </ul>
            </>
          )}
        </nav>

        {/* Bottom Navigation */}
        <div className="border-t border-gray-700 py-2 px-2">
          <ul className="space-y-1">
            {bottomNavItems.map((item) => (
              <NavItem 
                key={item.to}
                {...item}
                active={location.pathname === item.to}
                collapsed={isCollapsed}
              />
            ))}
          </ul>
          
          {/* Collapse Toggle */}
          <button
            onClick={() => setIsCollapsed(!isCollapsed)}
            className={`
              w-full mt-2 flex items-center gap-3 px-3 py-2 rounded-md text-sm
              text-gray-400 hover:text-white hover:bg-gray-700/50 transition-colors
              ${isCollapsed ? 'justify-center' : ''}
            `}
          >
            {isCollapsed ? (
              <ChevronRight24Regular className="w-5 h-5 flex-shrink-0" />
            ) : (
              <>
                <ChevronLeft24Regular className="w-5 h-5 flex-shrink-0" />
                <span>Collapse</span>
              </>
            )}
          </button>
        </div>
      </aside>

      {/* Main Content Area */}
      <div className={`flex-1 flex flex-col min-w-0 transition-all duration-300 ${isCollapsed ? 'ml-16' : 'ml-64'}`}>
        {/* Top Header */}
        <header className="sticky top-0 z-40 h-12 bg-white dark:bg-gray-800 border-b border-gray-200 dark:border-gray-700 flex items-center justify-between px-4">
          {/* Breadcrumb / Title */}
          <div className="flex items-center gap-2 text-sm">
            <span className="text-gray-500 dark:text-gray-400">Home</span>
            {location.pathname !== '/' && (
              <>
                <span className="text-gray-400">/</span>
                <span className="text-gray-900 dark:text-white font-medium">
                  {navItems.find(item => item.to === location.pathname)?.label || location.pathname.slice(1)}
                </span>
              </>
            )}
          </div>

          {/* Right side actions */}
          <div className="flex items-center gap-1">
            <Button
              appearance="subtle"
              icon={resolvedTheme === 'dark' ? <WeatherSunny24Regular /> : <WeatherMoon24Regular />}
              onClick={toggleTheme}
              size="small"
              aria-label="Toggle theme"
              title={resolvedTheme === 'dark' ? 'Switch to light mode' : 'Switch to dark mode'}
            />

            <div className="w-px h-6 bg-gray-200 dark:bg-gray-700 mx-2" />

            <Menu>
              <MenuTrigger disableButtonEnhancement>
                <Button appearance="subtle" size="small" className="!p-1">
                  <Avatar
                    name={profile?.displayName || 'User'}
                    image={profile?.profilePhoto ? { src: profile.profilePhoto } : undefined}
                    size={28}
                  />
                </Button>
              </MenuTrigger>
              <MenuPopover>
                <MenuList>
                  <div className="px-3 py-2 border-b border-gray-200 dark:border-gray-700">
                    <div className="font-medium text-gray-900 dark:text-white text-sm">
                      {profile?.displayName}
                    </div>
                    <div className="text-xs text-gray-500 dark:text-gray-400">
                      {profile?.email}
                    </div>
                    {profile?.roles.includes('Dashboard.Admin') && (
                      <div className="mt-1">
                        <span className="inline-flex items-center px-1.5 py-0.5 rounded text-xs font-medium bg-blue-100 text-blue-800 dark:bg-blue-900 dark:text-blue-200">
                          Admin
                        </span>
                      </div>
                    )}
                  </div>
                  <MenuItem icon={<Settings24Regular />} onClick={() => window.location.href = '/settings'}>
                    Settings
                  </MenuItem>
                  <MenuItem icon={<SignOut24Regular />} onClick={handleLogout}>
                    Sign out
                  </MenuItem>
                </MenuList>
              </MenuPopover>
            </Menu>
          </div>
        </header>

        {/* Main content */}
        <main className="flex-1 overflow-auto min-w-0">
          {children}
        </main>
      </div>
    </div>
  );
}

interface NavItemProps {
  to: string;
  icon: React.ComponentType<{ className?: string }>;
  iconFilled: React.ComponentType<{ className?: string }>;
  label: string;
  active: boolean;
  collapsed: boolean;
}

function NavItem({ to, icon: Icon, iconFilled: IconFilled, label, active, collapsed }: NavItemProps) {
  const ActiveIcon = active ? IconFilled : Icon;
  
  const content = (
    <Link
      to={to}
      className={`
        flex items-center gap-3 px-3 py-2 rounded-md text-sm transition-colors
        ${active 
          ? 'bg-blue-600 text-white' 
          : 'text-gray-300 hover:bg-gray-700/50 hover:text-white'
        }
        ${collapsed ? 'justify-center' : ''}
      `}
    >
      <ActiveIcon className="w-5 h-5 flex-shrink-0" />
      {!collapsed && <span>{label}</span>}
    </Link>
  );

  if (collapsed) {
    return (
      <li>
        <Tooltip content={label} relationship="label" positioning="after">
          {content}
        </Tooltip>
      </li>
    );
  }

  return <li>{content}</li>;
}
