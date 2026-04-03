# M365 Dashboard

A modern, open-source Microsoft 365 tenant dashboard built with .NET 8 and React.

![License](https://img.shields.io/badge/license-MIT-blue.svg)
![.NET](https://img.shields.io/badge/.NET-8.0-purple.svg)
![React](https://img.shields.io/badge/React-18-blue.svg)

## Features

- 📊 **Real-time Analytics** - Monitor active users, sign-ins, and usage patterns
- 🔐 **Security Assessment** - Comprehensive tenant security health checks
- 🛡️ **CIS Benchmark** - Microsoft 365 CIS Benchmark compliance scoring
- 👥 **User Management** - View and analyze user accounts, guests, and sign-in activity
- 📱 **Device Compliance** - Track Intune managed device compliance
- 📧 **Email Security** - Domain security auditing (SPF, DKIM, DMARC, MTA-STS)
- 💬 **Teams & Groups** - Monitor Teams, Microsoft 365 Groups, and Teams Phone
- 📜 **License Management** - View license consumption and utilization
- 📄 **PDF Reports** - Generate branded security assessment reports
- 🎨 **Dark Mode** - Full dark mode support
- ⚙️ **Customizable** - Per-user settings and branding options

## Quick Start

### Option 1: Deploy to Azure (Recommended)

The easiest way to deploy is using the automated PowerShell scripts:

```powershell
# 1. Clone the repository
git clone https://github.com/YourOrg/m365-dashboard.git
cd m365-dashboard

# 2. Register the Entra ID App (requires Azure CLI)
.\scripts\Register-EntraApp.ps1

# 3. Deploy to Azure
.\scripts\Deploy-M365Dashboard.ps1
```

The deployment script will:
- Create all Azure resources (Container App, SQL Database, Container Registry, Key Vault)
- Build and deploy the application
- Configure authentication automatically
- Provide you with the application URL

### Option 2: GitHub Actions (CI/CD)

1. Fork this repository
2. Add the following secrets to your GitHub repository:

| Secret | Description |
|--------|-------------|
| `AZURE_CREDENTIALS` | Azure service principal JSON |
| `ACR_LOGIN_SERVER` | Container Registry URL (e.g., `myacr.azurecr.io`) |
| `ACR_USERNAME` | Container Registry username |
| `ACR_PASSWORD` | Container Registry password |
| `CONTAINER_APP_NAME` | Name of your Container App |
| `RESOURCE_GROUP` | Azure resource group name |

3. Push to `main` branch - GitHub Actions will build and deploy automatically

### Option 3: Local Development

```bash
# Prerequisites: .NET 8 SDK, Node.js 20+

# Backend
cd src/M365Dashboard.Api
dotnet run

# Frontend (new terminal)
cd src/M365Dashboard.Api/ClientApp
npm install
npm run dev
```

## Architecture

```
┌─────────────────┐     ┌─────────────────┐     ┌─────────────────┐
│   React SPA     │────▶│   .NET 8 API    │────▶│  Microsoft      │
│   (Frontend)    │     │   (Backend)     │     │  Graph API      │
│                 │     │                 │     │                 │
│ • MSAL.js       │     │ • JWT Validation│     │ • Application   │
│ • Fluent UI     │     │ • App Roles     │     │   Permissions   │
│ • Recharts      │     │ • Caching       │     │                 │
└─────────────────┘     └─────────────────┘     └─────────────────┘
        │                       │
        │                       ▼
        │               ┌─────────────────┐
        │               │   Azure SQL     │
        │               │   Database      │
        │               │                 │
        │               │ • User Settings │
        │               │ • Report Config │
        │               │ • Cache         │
        │               └─────────────────┘
        │
        ▼
┌─────────────────┐
│   Entra ID      │
│   (Auth)        │
│                 │
│ • User Sign-in  │
│ • App Roles     │
│ • Token Issue   │
└─────────────────┘
```

## Permission Model

This dashboard uses **Application Permissions** for Microsoft Graph API access:

| Aspect | How It Works |
|--------|--------------|
| **User Sign-in** | Users authenticate via Entra ID |
| **Data Access** | Backend uses app identity to read tenant data |
| **Benefit** | Helpdesk users see all data without needing M365 admin rights |
| **Control** | Access controlled via App Roles (Dashboard.Admin / Dashboard.Reader) |

### Required Graph API Permissions

| Permission | Purpose |
|------------|---------|
| `User.Read.All` | Read user profiles |
| `Group.Read.All` | Read groups and Teams |
| `Directory.Read.All` | Read directory data |
| `DeviceManagementManagedDevices.Read.All` | Read Intune devices |
| `DeviceManagementConfiguration.Read.All` | Read device policies |
| `DeviceManagementApps.Read.All` | Read app management |
| `SecurityEvents.Read.All` | Read security alerts |
| `IdentityRiskyUser.Read.All` | Read risky users |
| `IdentityRiskEvent.Read.All` | Read risk events |
| `Reports.Read.All` | Read usage reports |
| `AuditLog.Read.All` | Read audit logs |
| `Mail.Read` | Read mail flow rules |
| `Domain.Read.All` | Read domain configuration |
| `Organization.Read.All` | Read organization info |
| `Policy.Read.All` | Read security policies |

## Project Structure

```
m365-dashboard/
├── .github/workflows/       # GitHub Actions CI/CD
├── docs/                    # Documentation
├── infra/                   # Azure Bicep templates
├── scripts/                 # PowerShell deployment scripts
└── src/M365Dashboard.Api/   # .NET 8 Backend
    ├── Controllers/         # API endpoints
    ├── Services/            # Business logic & Graph API
    ├── Models/              # Data models
    └── ClientApp/           # React Frontend
        ├── src/components/  # React components
        ├── src/pages/       # Page components
        └── src/contexts/    # React contexts
```

## Screenshots

*Coming soon*

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Acknowledgments

- [Microsoft Graph API](https://docs.microsoft.com/en-us/graph/)
- [Fluent UI React](https://react.fluentui.dev/)
- [Recharts](https://recharts.org/)
- [QuestPDF](https://www.questpdf.com/) for PDF generation
