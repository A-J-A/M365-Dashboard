import { useMsal } from '@azure/msal-react';
import { InteractionStatus } from '@azure/msal-browser';
import { Button, Spinner } from '@fluentui/react-components';
import { Grid24Regular, LockClosed24Regular } from '@fluentui/react-icons';
import { loginRequest } from '../../services/authConfig';

export function LoginPage() {
  const { instance, inProgress } = useMsal();
  const isLoading = inProgress !== InteractionStatus.None;

  const handleLogin = async () => {
    // Prevent multiple login attempts
    if (inProgress !== InteractionStatus.None) {
      console.log('Login already in progress, ignoring click');
      return;
    }
    
    try {
      await instance.loginRedirect(loginRequest);
    } catch (error) {
      console.error('Login failed:', error);
    }
  };

  return (
    <div className="min-h-screen bg-gradient-to-br from-blue-600 via-blue-700 to-indigo-800 flex items-center justify-center p-4">
      <div className="bg-white dark:bg-gray-800 rounded-2xl shadow-2xl max-w-md w-full p-8">
        {/* Logo */}
        <div className="text-center mb-8">
          <div className="inline-flex items-center justify-center w-16 h-16 rounded-2xl bg-blue-600 text-white mb-4">
            <Grid24Regular className="w-8 h-8" />
          </div>
          <h1 className="text-2xl font-bold text-gray-900 dark:text-white">
            M365 Dashboard
          </h1>
          <p className="mt-2 text-gray-600 dark:text-gray-400">
            Microsoft 365 Tenant Insights
          </p>
        </div>

        {/* Features */}
        <div className="space-y-4 mb-8">
          <Feature title="Real-time Analytics" description="Monitor user activity, sign-ins, and security events" />
          <Feature title="License Management" description="Track license consumption across your tenant" />
          <Feature title="Device Compliance" description="View Intune device compliance status" />
        </div>

        {/* Login button */}
        <Button
          appearance="primary"
          size="large"
          style={{ width: '100%' }}
          icon={isLoading ? <Spinner size="tiny" /> : <LockClosed24Regular />}
          onClick={handleLogin}
          disabled={isLoading}
        >
          {isLoading ? 'Signing in...' : 'Sign in with Microsoft'}
        </Button>

        {/* Footer */}
        <p className="mt-6 text-center text-sm text-gray-500 dark:text-gray-400">
          Secure authentication via Microsoft Entra ID
        </p>
      </div>
    </div>
  );
}

function Feature({ title, description }: { title: string; description: string }) {
  return (
    <div className="flex items-start gap-3">
      <div className="flex-shrink-0 w-5 h-5 rounded-full bg-blue-100 dark:bg-blue-900 flex items-center justify-center mt-0.5">
        <svg className="w-3 h-3 text-blue-600 dark:text-blue-400" fill="currentColor" viewBox="0 0 20 20">
          <path fillRule="evenodd" d="M16.707 5.293a1 1 0 010 1.414l-8 8a1 1 0 01-1.414 0l-4-4a1 1 0 011.414-1.414L8 12.586l7.293-7.293a1 1 0 011.414 0z" clipRule="evenodd" />
        </svg>
      </div>
      <div>
        <h3 className="text-sm font-medium text-gray-900 dark:text-white">{title}</h3>
        <p className="text-sm text-gray-500 dark:text-gray-400">{description}</p>
      </div>
    </div>
  );
}
