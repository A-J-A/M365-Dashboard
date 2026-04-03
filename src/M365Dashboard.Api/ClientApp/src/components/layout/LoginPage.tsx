import { useMsal } from '@azure/msal-react';
import { InteractionStatus } from '@azure/msal-browser';
import { loginRequest } from '../../services/authConfig';

const features = [
  {
    title: 'Security & Compliance',
    description: 'MFA status, risky users, Secure Score, and Conditional Access insights',
  },
  {
    title: 'Users & Groups',
    description: 'Licence assignment, sign-in activity, group membership, and guest access',
  },
  {
    title: 'Devices & Intune',
    description: 'Compliance posture, OS versions, encryption status, and enrolment health',
  },
  {
    title: 'Licences & Reports',
    description: 'Subscription utilisation, mailflow analytics, and executive reporting',
  },
];

export function LoginPage() {
  const { instance, inProgress } = useMsal();
  const isLoading = inProgress !== InteractionStatus.None;

  const handleLogin = async () => {
    if (inProgress !== InteractionStatus.None) return;
    try {
      await instance.loginRedirect(loginRequest);
    } catch (error) {
      console.error('Login failed:', error);
    }
  };

  return (
    <div className="min-h-screen flex bg-slate-900">

      {/* Left panel — hidden on mobile, visible on lg+ */}
      <div className="hidden lg:flex flex-1 flex-col justify-between p-12 border-r border-white/5">

        {/* Logo */}
        <div className="flex items-center gap-3">
          <div className="w-9 h-9 rounded-xl bg-blue-600 flex items-center justify-center flex-shrink-0">
            <GridIcon />
          </div>
          <span className="text-white font-semibold text-sm tracking-tight">M365 Dashboard</span>
        </div>

        {/* Headline */}
        <div>
          <p className="text-xs font-semibold text-blue-400 uppercase tracking-widest mb-4">Open Source</p>
          <h1 className="text-4xl font-bold text-white leading-tight tracking-tight mb-4">
            Your Microsoft 365<br />tenant, at a glance.
          </h1>
          <p className="text-slate-400 text-base leading-relaxed">
            A unified dashboard for security posture, user management,
            device compliance, and operational reporting — all in one place.
          </p>
        </div>

        {/* Feature list */}
        <div className="space-y-5">
          {features.map((f) => (
            <div key={f.title} className="flex gap-4 items-start">
              <div className="flex-shrink-0 w-2 h-2 rounded-full bg-blue-500 mt-2" />
              <div>
                <p className="text-sm font-semibold text-slate-200 mb-0.5">{f.title}</p>
                <p className="text-sm text-slate-500 leading-relaxed">{f.description}</p>
              </div>
            </div>
          ))}
        </div>

        {/* Footer */}
        <p className="text-xs text-slate-700">
          Powered by Microsoft Graph API · Read-only tenant access
        </p>
      </div>

      {/* Right panel — sign in */}
      <div className="w-full lg:max-w-md flex items-center justify-center p-8 mx-auto">
        <div className="w-full">

          {/* Card */}
          <div className="bg-slate-800 border border-white/8 rounded-2xl p-10">

            {/* App icon */}
            <div className="w-14 h-14 rounded-2xl bg-blue-600 flex items-center justify-center mb-6">
              <GridIcon className="w-7 h-7" />
            </div>

            <h2 className="text-2xl font-bold text-slate-100 tracking-tight mb-2">Sign in</h2>
            <p className="text-sm text-slate-400 leading-relaxed mb-8">
              Use your Microsoft 365 work account to access the dashboard.
            </p>

            {/* Divider */}
            <div className="border-t border-white/5 mb-8" />

            {/* Button */}
            <button
              onClick={handleLogin}
              disabled={isLoading}
              className={`
                w-full flex items-center justify-center gap-3 px-5 py-3.5
                rounded-xl text-white text-sm font-semibold tracking-tight
                transition-all duration-150
                ${isLoading
                  ? 'bg-blue-600/50 cursor-not-allowed'
                  : 'bg-blue-600 hover:bg-blue-500 active:scale-[0.99] cursor-pointer'
                }
              `}
            >
              {isLoading ? (
                <>
                  <svg className="w-4 h-4 animate-spin" viewBox="0 0 24 24" fill="none">
                    <circle cx="12" cy="12" r="10" stroke="rgba(255,255,255,0.3)" strokeWidth="3" />
                    <path d="M12 2a10 10 0 0110 10" stroke="white" strokeWidth="3" strokeLinecap="round" />
                  </svg>
                  Signing in...
                </>
              ) : (
                <>
                  <MicrosoftLogo />
                  Sign in with Microsoft
                </>
              )}
            </button>


          </div>

          <p className="text-center mt-6 text-xs text-slate-700">
            By signing in you agree to your organisation's acceptable use policy.
          </p>
        </div>
      </div>
    </div>
  );
}

function GridIcon({ className }: { className?: string }) {
  return (
    <svg viewBox="0 0 24 24" fill="none" stroke="white" strokeWidth="2" strokeLinecap="round" className={className ?? 'w-5 h-5'}>
      <rect x="3" y="3" width="7" height="7" rx="1" />
      <rect x="14" y="3" width="7" height="7" rx="1" />
      <rect x="3" y="14" width="7" height="7" rx="1" />
      <rect x="14" y="14" width="7" height="7" rx="1" />
    </svg>
  );
}

function MicrosoftLogo() {
  return (
    <svg viewBox="0 0 21 21" className="w-4 h-4 flex-shrink-0">
      <rect x="1" y="1" width="9" height="9" fill="#f25022" />
      <rect x="11" y="1" width="9" height="9" fill="#7fba00" />
      <rect x="1" y="11" width="9" height="9" fill="#00a4ef" />
      <rect x="11" y="11" width="9" height="9" fill="#ffb900" />
    </svg>
  );
}
