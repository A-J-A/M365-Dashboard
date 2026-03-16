import { useMsal } from '@azure/msal-react';
import { InteractionStatus } from '@azure/msal-browser';
import { loginRequest } from '../../services/authConfig';

const features = [
  {
    icon: (
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" style={{ width: 20, height: 20 }}>
        <path d="M3 13.5L7.5 9l3 3L15 7.5l4.5 4.5" />
        <rect x="3" y="3" width="18" height="18" rx="2" />
      </svg>
    ),
    title: 'Security & Compliance',
    description: 'MFA status, risky users, Secure Score, and Conditional Access insights',
  },
  {
    icon: (
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" style={{ width: 20, height: 20 }}>
        <circle cx="9" cy="7" r="4" />
        <path d="M3 21v-2a4 4 0 014-4h4a4 4 0 014 4v2" />
        <path d="M16 3.13a4 4 0 010 7.75" />
        <path d="M21 21v-2a4 4 0 00-3-3.87" />
      </svg>
    ),
    title: 'Users & Groups',
    description: 'Licence assignment, sign-in activity, group membership, and guest access',
  },
  {
    icon: (
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" style={{ width: 20, height: 20 }}>
        <rect x="5" y="2" width="14" height="20" rx="2" />
        <path d="M12 18h.01" />
        <path d="M9 6h6M9 10h6M9 14h4" />
      </svg>
    ),
    title: 'Devices & Intune',
    description: 'Compliance posture, OS versions, encryption status, and enrolment health',
  },
  {
    icon: (
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" style={{ width: 20, height: 20 }}>
        <path d="M20 7H4a2 2 0 00-2 2v6a2 2 0 002 2h16a2 2 0 002-2V9a2 2 0 00-2-2z" />
        <path d="M16 21V5a2 2 0 00-2-2h-4a2 2 0 00-2 2v16" />
      </svg>
    ),
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
    <div style={{
      minHeight: '100vh',
      display: 'flex',
      backgroundColor: '#0f172a',
      fontFamily: "'Segoe UI', system-ui, -apple-system, sans-serif",
    }}>

      {/* Left panel — branding & features */}
      <div style={{
        display: 'none',
        flex: 1,
        padding: '3rem',
        flexDirection: 'column',
        justifyContent: 'space-between',
        borderRight: '1px solid rgba(255,255,255,0.06)',
      }} className="login-left-panel">

        {/* Logo mark */}
        <div style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
          <div style={{
            width: 36, height: 36, borderRadius: 10,
            background: 'linear-gradient(135deg, #3b82f6, #6366f1)',
            display: 'flex', alignItems: 'center', justifyContent: 'center',
          }}>
            <svg viewBox="0 0 24 24" fill="none" stroke="white" strokeWidth="2" strokeLinecap="round" style={{ width: 18, height: 18 }}>
              <rect x="3" y="3" width="7" height="7" rx="1" />
              <rect x="14" y="3" width="7" height="7" rx="1" />
              <rect x="3" y="14" width="7" height="7" rx="1" />
              <rect x="14" y="14" width="7" height="7" rx="1" />
            </svg>
          </div>
          <span style={{ color: 'white', fontWeight: 600, fontSize: 15, letterSpacing: '-0.01em' }}>M365 Dashboard</span>
        </div>

        {/* Headline */}
        <div>
          <p style={{ fontSize: 12, color: '#60a5fa', fontWeight: 500, letterSpacing: '0.08em', textTransform: 'uppercase', marginBottom: 16 }}>Open Source</p>
          <h1 style={{ fontSize: 36, fontWeight: 700, color: 'white', lineHeight: 1.15, letterSpacing: '-0.02em', margin: '0 0 16px' }}>
            Your Microsoft 365<br />tenant, at a glance.
          </h1>
          <p style={{ fontSize: 15, color: '#94a3b8', lineHeight: 1.7, margin: 0 }}>
            A unified dashboard for security posture, user management, device compliance, and operational reporting — all in one place.
          </p>
        </div>

        {/* Feature list */}
        <div style={{ display: 'flex', flexDirection: 'column', gap: 20 }}>
          {features.map((f) => (
            <div key={f.title} style={{ display: 'flex', gap: 14, alignItems: 'flex-start' }}>
              <div style={{
                flexShrink: 0, width: 38, height: 38, borderRadius: 10,
                background: 'rgba(59,130,246,0.12)',
                border: '1px solid rgba(59,130,246,0.2)',
                display: 'flex', alignItems: 'center', justifyContent: 'center',
                color: '#60a5fa',
              }}>
                {f.icon}
              </div>
              <div>
                <p style={{ margin: '0 0 3px', fontSize: 14, fontWeight: 600, color: '#e2e8f0' }}>{f.title}</p>
                <p style={{ margin: 0, fontSize: 13, color: '#64748b', lineHeight: 1.5 }}>{f.description}</p>
              </div>
            </div>
          ))}
        </div>

        {/* Bottom tagline */}
        <p style={{ fontSize: 12, color: '#334155', margin: 0 }}>
          Powered by Microsoft Graph API &middot; Read-only tenant access
        </p>
      </div>

      {/* Right panel — sign in card */}
      <div style={{
        width: '100%',
        maxWidth: 480,
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'center',
        padding: '2rem',
        margin: '0 auto',
      }}>
        <div style={{ width: '100%' }}>

          {/* Card */}
          <div style={{
            background: '#1e293b',
            border: '1px solid rgba(255,255,255,0.08)',
            borderRadius: 20,
            padding: '2.5rem',
          }}>

            {/* Icon */}
            <div style={{ marginBottom: 28 }}>
              <div style={{
                width: 52, height: 52, borderRadius: 14,
                background: 'linear-gradient(135deg, #3b82f6 0%, #6366f1 100%)',
                display: 'flex', alignItems: 'center', justifyContent: 'center',
                marginBottom: 24,
              }}>
                <svg viewBox="0 0 24 24" fill="none" stroke="white" strokeWidth="2" strokeLinecap="round" style={{ width: 24, height: 24 }}>
                  <rect x="3" y="3" width="7" height="7" rx="1" />
                  <rect x="14" y="3" width="7" height="7" rx="1" />
                  <rect x="3" y="14" width="7" height="7" rx="1" />
                  <rect x="14" y="14" width="7" height="7" rx="1" />
                </svg>
              </div>
              <h2 style={{ margin: '0 0 8px', fontSize: 22, fontWeight: 700, color: '#f1f5f9', letterSpacing: '-0.02em' }}>
                Sign in
              </h2>
              <p style={{ margin: 0, fontSize: 14, color: '#64748b', lineHeight: 1.5 }}>
                Use your Microsoft 365 work account to access the dashboard.
              </p>
            </div>

            {/* Divider */}
            <div style={{ height: 1, background: 'rgba(255,255,255,0.06)', marginBottom: 28 }} />

            {/* Sign in button */}
            <button
              onClick={handleLogin}
              disabled={isLoading}
              style={{
                width: '100%',
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'center',
                gap: 10,
                padding: '13px 20px',
                borderRadius: 12,
                border: 'none',
                background: isLoading
                  ? 'rgba(59,130,246,0.5)'
                  : 'linear-gradient(135deg, #3b82f6 0%, #6366f1 100%)',
                color: 'white',
                fontSize: 15,
                fontWeight: 600,
                letterSpacing: '-0.01em',
                cursor: isLoading ? 'not-allowed' : 'pointer',
                transition: 'opacity 0.15s, transform 0.1s',
                fontFamily: 'inherit',
              }}
              onMouseEnter={e => { if (!isLoading) (e.currentTarget as HTMLButtonElement).style.opacity = '0.9'; }}
              onMouseLeave={e => { (e.currentTarget as HTMLButtonElement).style.opacity = '1'; }}
              onMouseDown={e => { if (!isLoading) (e.currentTarget as HTMLButtonElement).style.transform = 'scale(0.99)'; }}
              onMouseUp={e => { (e.currentTarget as HTMLButtonElement).style.transform = 'scale(1)'; }}
            >
              {isLoading ? (
                <>
                  <svg style={{ width: 18, height: 18, animation: 'spin 0.8s linear infinite' }} viewBox="0 0 24 24" fill="none">
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

            {/* Info chips */}
            <div style={{ display: 'flex', gap: 8, marginTop: 20, flexWrap: 'wrap' }}>
              <Chip icon="🔒" label="Read-only access" />
              <Chip icon="🛡️" label="Entra ID auth" />
              <Chip icon="✓" label="No data stored" />
            </div>
          </div>

          {/* Below card note */}
          <p style={{ textAlign: 'center', marginTop: 24, fontSize: 12, color: '#334155' }}>
            By signing in you agree to your organisation&apos;s acceptable use policy.
          </p>
        </div>
      </div>

      <style>{`
        @keyframes spin { to { transform: rotate(360deg); } }
        @media (min-width: 900px) {
          .login-left-panel { display: flex !important; }
        }
      `}</style>
    </div>
  );
}

function MicrosoftLogo() {
  return (
    <svg viewBox="0 0 21 21" style={{ width: 18, height: 18, flexShrink: 0 }}>
      <rect x="1" y="1" width="9" height="9" fill="#f25022" />
      <rect x="11" y="1" width="9" height="9" fill="#7fba00" />
      <rect x="1" y="11" width="9" height="9" fill="#00a4ef" />
      <rect x="11" y="11" width="9" height="9" fill="#ffb900" />
    </svg>
  );
}

function Chip({ icon, label }: { icon: string; label: string }) {
  return (
    <div style={{
      display: 'inline-flex',
      alignItems: 'center',
      gap: 5,
      padding: '5px 10px',
      borderRadius: 8,
      background: 'rgba(255,255,255,0.04)',
      border: '1px solid rgba(255,255,255,0.07)',
      fontSize: 12,
      color: '#64748b',
    }}>
      <span style={{ fontSize: 11 }}>{icon}</span>
      {label}
    </div>
  );
}
