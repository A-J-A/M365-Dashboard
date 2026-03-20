import { BrowserRouter, Routes, Route, Navigate } from 'react-router-dom';
import { MsalProvider, AuthenticatedTemplate, UnauthenticatedTemplate, useMsal, useIsAuthenticated } from '@azure/msal-react';
import { InteractionStatus } from '@azure/msal-browser';
import { FluentProvider, webLightTheme, webDarkTheme, Spinner } from '@fluentui/react-components';
import { msalInstance, loginRequest } from './services/authConfig';
import { ThemeProvider, UserProvider, SettingsProvider, useTheme } from './contexts/AppContext';
import { Layout } from './components/layout/Layout';
import { Dashboard } from './components/dashboard/Dashboard';
import { SettingsPage } from './components/settings/SettingsPage';
import { LoginPage } from './components/layout/LoginPage';
import UsersPage from './pages/UsersPage';
import TeamsGroupsPage from './pages/TeamsGroupsPage';
import DevicesPage from './pages/DevicesPage';
import MailflowPage from './pages/MailflowPage';
import SecurityPage from './pages/SecurityPage';
import MfaDetailsPage from './pages/MfaDetailsPage';
import ReportsPage from './pages/ReportsPage';
import SignInsPage from './pages/SignInsPage';
import SharePointPage from './pages/SharePointPage';
import LicensesPage from './pages/LicensesPage';
import TeamsPhonePage from './pages/TeamsPhonePage';
import ConditionalAccessPage from './pages/ConditionalAccessPage';
import PrivilegedAccessPage from './pages/PrivilegedAccessPage';
import ThreatIntelligencePage from './pages/ThreatIntelligencePage';
import MailFlowMonitorPage from './pages/MailFlowMonitorPage';
import ApplicationConsentPage from './pages/ApplicationConsentPage';
import ExecutiveReportPage from './pages/ExecutiveReportPage';
import CisBenchmarkPage from './pages/CisBenchmarkPage';
import SecurityAssessmentPage from './pages/SecurityAssessmentPage';

import DefenderForOfficePage from './pages/DefenderForOfficePage';
import MailboxAccessPage from './pages/MailboxAccessPage';

function AuthenticatedApp() {
  return (
    <UserProvider>
      <SettingsProvider>
        <Layout>
          <Routes>
            <Route path="/" element={<Dashboard />} />
            <Route path="/users" element={<UsersPage />} />
            <Route path="/teams" element={<TeamsGroupsPage />} />
            <Route path="/devices" element={<DevicesPage />} />
            <Route path="/mailflow" element={<MailflowPage />} />
            <Route path="/security" element={<SecurityPage />} />
            <Route path="/security/mfa" element={<MfaDetailsPage />} />
            <Route path="/signins" element={<SignInsPage />} />
            <Route path="/teamsphone" element={<TeamsPhonePage />} />
            <Route path="/conditional-access" element={<ConditionalAccessPage />} />
            <Route path="/privileged-access" element={<PrivilegedAccessPage />} />
            <Route path="/threat-intelligence" element={<ThreatIntelligencePage />} />
            <Route path="/mail-flow" element={<MailFlowMonitorPage />} />
            <Route path="/app-consent" element={<ApplicationConsentPage />} />
            <Route path="/sharepoint" element={<SharePointPage />} />
            <Route path="/licenses" element={<LicensesPage />} />
            <Route path="/reports" element={<ReportsPage />} />
            <Route path="/executive-report" element={<ExecutiveReportPage />} />
            <Route path="/cis-benchmark" element={<CisBenchmarkPage />} />
            <Route path="/security-assessment" element={<SecurityAssessmentPage />} />
            <Route path="/defender-office" element={<DefenderForOfficePage />} />
            <Route path="/mailbox-access" element={<MailboxAccessPage />} />
            <Route path="/settings/reports" element={<Navigate to="/settings" replace />} />
            <Route path="/settings" element={<SettingsPage />} />
            <Route path="*" element={<Navigate to="/" replace />} />
          </Routes>
        </Layout>
      </SettingsProvider>
    </UserProvider>
  );
}

function AppRoutes() {
  const { inProgress } = useMsal();
  const isAuthenticated = useIsAuthenticated();

  // Show loading spinner while MSAL is processing
  if (inProgress !== InteractionStatus.None) {
    return (
      <div style={{ 
        display: 'flex', 
        justifyContent: 'center', 
        alignItems: 'center', 
        height: '100vh' 
      }}>
        <Spinner size="large" label="Signing in..." />
      </div>
    );
  }

  return (
    <BrowserRouter>
      <AuthenticatedTemplate>
        <AuthenticatedApp />
      </AuthenticatedTemplate>
      
      <UnauthenticatedTemplate>
        <LoginPage />
      </UnauthenticatedTemplate>
    </BrowserRouter>
  );
}

function AppContent() {
  const { resolvedTheme } = useTheme();
  const theme = resolvedTheme === 'dark' ? webDarkTheme : webLightTheme;

  return (
    <FluentProvider theme={theme}>
      <AppRoutes />
    </FluentProvider>
  );
}

function App() {
  return (
    <MsalProvider instance={msalInstance}>
      <ThemeProvider>
        <AppContent />
      </ThemeProvider>
    </MsalProvider>
  );
}

export default App;
