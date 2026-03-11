import { Configuration, LogLevel, PublicClientApplication } from '@azure/msal-browser';

// MSAL configuration
const msalConfig: Configuration = {
  auth: {
    clientId: import.meta.env.VITE_AZURE_CLIENT_ID || 'YOUR_CLIENT_ID',
    authority: `https://login.microsoftonline.com/${import.meta.env.VITE_AZURE_TENANT_ID || 'YOUR_TENANT_ID'}`,
    redirectUri: window.location.origin,
    postLogoutRedirectUri: window.location.origin,
    navigateToLoginRequestUrl: true,
  },
  cache: {
    cacheLocation: 'sessionStorage',
    storeAuthStateInCookie: false,
  },
  system: {
    loggerOptions: {
      loggerCallback: (level, message, containsPii) => {
        if (containsPii) return;
        switch (level) {
          case LogLevel.Error:
            console.error(message);
            break;
          case LogLevel.Warning:
            console.warn(message);
            break;
          case LogLevel.Info:
            console.info(message);
            break;
          case LogLevel.Verbose:
            console.debug(message);
            break;
        }
      },
      logLevel: import.meta.env.DEV ? LogLevel.Verbose : LogLevel.Warning,
    },
  },
};

// Scopes for API access
export const loginRequest = {
  scopes: ['User.Read', 'openid', 'profile', 'email'],
};

export const apiRequest = {
  scopes: [`api://${import.meta.env.VITE_AZURE_CLIENT_ID || 'YOUR_CLIENT_ID'}/access_as_user`],
};

export const graphScopes = {
  scopes: [
    'User.Read',
    'Reports.Read.All',
    'AuditLog.Read.All',
    'Directory.Read.All',
    'DeviceManagementManagedDevices.Read.All',
  ],
};

// Create MSAL instance
export const msalInstance = new PublicClientApplication(msalConfig);

// Initialize MSAL
export const initializeMsal = async () => {
  await msalInstance.initialize();
  
  // Handle redirect promise
  const response = await msalInstance.handleRedirectPromise();
  
  if (response) {
    msalInstance.setActiveAccount(response.account);
  } else {
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length > 0) {
      msalInstance.setActiveAccount(accounts[0]);
    }
  }
  
  return msalInstance;
};
