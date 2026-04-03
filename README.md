# M365 Dashboard

A modern, open-source Microsoft 365 tenant dashboard built with .NET 8 and React. Designed for IT administrators and MSPs to monitor security posture, device compliance, user activity, and licensing across Microsoft 365 tenants.

![License](https://img.shields.io/badge/license-MIT-blue.svg)
![.NET](https://img.shields.io/badge/.NET-8.0-purple.svg)
![React](https://img.shields.io/badge/React-18-blue.svg)

---

## Features

- 📊 **Executive Summary Reports** — Generate branded PDF security reports for clients
- 🔐 **Security Assessment** — Comprehensive tenant security health checks
- 🛡️ **CIS Benchmark** — Microsoft 365 CIS Benchmark compliance scoring
- 👥 **User Management** — Users, guests, sign-in activity, and MFA status
- 📱 **Device Compliance** — Intune managed device compliance and OS version tracking
- 📧 **Email Security** — Domain security auditing (SPF, DKIM, DMARC, MTA-STS)
- 🔑 **Privileged Access** — Entra ID admin role assignments with MFA status
- 💬 **Teams & Groups** — Monitor Teams, Microsoft 365 Groups, and Teams Phone
- 📜 **License Management** — View license consumption and utilisation
- 🗓️ **Scheduled Reports** — Automated email delivery of PDF reports
- 🎨 **Dark Mode** — Full dark mode support
- ⚙️ **Multi-tenant / MSP** — Deploy once, connect to multiple client tenants

---

## Deployment

### Prerequisites

- [Azure CLI](https://docs.microsoft.com/en-us/cli/azure/install-azure-cli)
- An Azure subscription
- Global Administrator access to the target Microsoft 365 tenant

### Deploy to Azure

The deployment script handles everything end-to-end — no manual Azure portal steps required.

```powershell
git clone https://github.com/A-J-A/M365-Dashboard.git
cd M365-Dashboard
.\scripts\Deploy-M365Dashboard.ps1
```

The script will:

1. **Create an Entra ID app registration** with all required API permissions
2. **Deploy Azure infrastructure** — Container App, SQL Database, Container Registry, Key Vault
3. **Build and push the Docker image** using Azure Container Registry Build (no local Docker required)
4. **Configure authentication** — redirect URIs, admin consent, app roles
5. **Output the dashboard URL** when complete

#### Deployment Modes

| Mode | When to use |
|------|-------------|
| **Standard** | App registration and Azure resources in the same tenant |
| **MSP / Multi-tenant** | App registration in the client's M365 tenant, Azure resources in your own subscription |

MSP mode allows a single Azure deployment to serve multiple client tenants. The script guides you through both logins.

#### What gets created in Azure

| Resource | Purpose |
|----------|---------|
| Resource Group | Container for all resources |
| Container App | Hosts the dashboard (auto-scaling) |
| Container Registry | Stores the Docker image |
| SQL Database | Stores settings, schedules, and report history |
| Key Vault | Stores app credentials and configuration securely |

---

## CI/CD with GitHub Actions

If you want automatic deployments on every code push, add the following secrets to your GitHub repository under **Settings → Secrets and variables → Actions**.

> **These secrets are completely secure.** GitHub encrypts secrets at rest and only injects them into workflow runners during Actions jobs. They are never exposed in repository code, git history, or to anyone who clones the repository — even collaborators cannot read secret values.

| Secret | Description | How to get it |
|--------|-------------|---------------|
| `AZURE_CREDENTIALS` | Service principal JSON for Azure login | Output by deploy script, or run `az ad sp create-for-rbac --sdk-auth` |
| `ACR_LOGIN_SERVER` | Container Registry login server | Output by deploy script (e.g. `myacr.azurecr.io`) |
| `ACR_USERNAME` | Container Registry username | Run `az acr credential show --name <acr-name> --query username -o tsv` |
| `ACR_PASSWORD` | Container Registry password | Run `az acr credential show --name <acr-name> --query passwords[0].value -o tsv` |
| `CONTAINER_APP_NAME` | Name of the Container App | Output by deploy script |
| `RESOURCE_GROUP` | Azure resource group name | Output by deploy script |
| `VITE_AZURE_CLIENT_ID` | Entra app registration client ID | Output by deploy script |
| `VITE_AZURE_TENANT_ID` | Entra tenant ID | Output by deploy script |

The deploy script can set these automatically if you have the [GitHub CLI](https://cli.github.com) installed and authenticated. Otherwise all values are printed at the end of deployment for manual entry.

---

## Local Development

```powershell
# Prerequisites: .NET 8 SDK, Node.js 20+

# Backend
cd src/M365Dashboard.Api
dotnet run

# Frontend (new terminal)
cd src/M365Dashboard.Api/ClientApp
npm install
npm run dev
```

The frontend dev server proxies API calls to `https://localhost:5001`. You will need a valid `appsettings.Development.json` with Entra ID credentials and a SQL connection string — the deploy script generates this automatically.

---

## Architecture

```
┌─────────────────┐     ┌─────────────────┐     ┌─────────────────┐
│   React SPA     │────▶│   .NET 8 API    │────▶│  Microsoft      │
│   (Frontend)    │     │   (Backend)     │     │  Graph API      │
│                 │     │                 │     │                 │
│ • MSAL.js       │     │ • JWT Validation│     │ • Application   │
│ • Fluent UI     │     │ • App Roles     │     │   Permissions   │
│ • Recharts      │     │ • EF Core       │     │                 │
└─────────────────┘     └─────────────────┘     └─────────────────┘
        │                       │
        │                       ▼
        │               ┌─────────────────┐
        │               │   Azure SQL     │
        │               │   Database      │
        │               │                 │
        │               │ • User Settings │
        │               │ • Report Config │
        │               │ • Schedules     │
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

---

## Permission Model

This dashboard uses **Application Permissions** — the backend authenticates as itself, not on behalf of individual users. This means helpdesk staff can use the dashboard without needing M365 admin rights.

| Aspect | Detail |
|--------|--------|
| **User Sign-in** | Users authenticate via Entra ID |
| **Data Access** | Backend uses its own app identity to read tenant data |
| **Access Control** | Managed via App Roles (`Dashboard.Admin` / `Dashboard.Reader`) |

### Required Microsoft Graph Permissions

| Permission | Purpose |
|------------|---------|
| `User.Read.All` | User profiles and sign-in activity |
| `Group.Read.All` | Groups and Teams |
| `Directory.Read.All` | Directory data and admin roles |
| `Organization.Read.All` | Organisation information |
| `Policy.Read.All` | Conditional access and security policies |
| `Domain.Read.All` | Domain configuration for email security |
| `DeviceManagementManagedDevices.Read.All` | Intune managed devices |
| `DeviceManagementConfiguration.Read.All` | Device compliance policies |
| `DeviceManagementApps.Read.All` | App management |
| `DeviceManagementServiceConfig.Read.All` | Intune service configuration |
| `SecurityEvents.Read.All` | Security alerts |
| `IdentityRiskyUser.Read.All` | Risky users |
| `IdentityRiskEvent.Read.All` | Risk events |
| `AuditLog.Read.All` | Audit logs |
| `Reports.Read.All` | Usage reports and mailbox statistics |
| `Mail.Read` | Mail flow rules |
| `Mail.Send` | Scheduled report email delivery |
| `UserAuthenticationMethod.Read.All` | MFA registration status |
| `AttackSimulation.Read.All` | Attack simulation training data |
| `Sites.Read.All` | SharePoint usage |

### Additional Permissions

| API | Permission | Purpose |
|-----|------------|---------|
| Microsoft Defender for Endpoint | `Machine.Read.All`, `Vulnerability.Read.All`, `Score.Read.All` | Defender exposure score and vulnerabilities |
| Exchange Online | `Exchange.ManageAsApp` | Mailbox statistics and mail flow |

---

## Project Structure

```
M365-Dashboard/
├── .github/workflows/       # GitHub Actions CI/CD pipeline
├── infra/                   # Azure Bicep infrastructure templates
├── scripts/                 # PowerShell deployment scripts
│   ├── Deploy-M365Dashboard.ps1   # Main deployment script
│   └── Update-M365Dashboard.ps1  # Update to new release
└── src/M365Dashboard.Api/   # .NET 8 backend + React frontend
    ├── Controllers/         # API endpoints
    ├── Services/            # Business logic and Graph API integration
    ├── Models/              # Data models
    └── ClientApp/           # React frontend (Vite + Tailwind)
        ├── src/components/  # Shared components
        ├── src/pages/       # Page components
        └── src/contexts/    # React context providers
```

---

## Updating

The dashboard includes a built-in update mechanism. When a new release is published to this repository, the Settings page will show an update notification. Clicking **Update** will pull the new image and restart the Container App automatically — no manual steps required.

To update manually:

```powershell
.\scripts\Update-M365Dashboard.ps1
```

---

## Post-Deployment Steps

After running the deploy script, two steps require manual action in the Microsoft 365 admin centre:

1. **Grant admin consent** for the app registration API permissions if not auto-granted
   - Entra admin centre → App registrations → your app → API permissions → Grant admin consent

2. **Exchange Security Reader role** (for Defender for Office 365 data)
   - Exchange admin centre → Roles → Admin roles → View-Only Organization Management → Add your app registration as a member

---

## Contributing

Contributions are welcome. Please open an issue first to discuss what you would like to change, then submit a pull request.

---

## License

This project is licensed under the MIT License — see the [LICENSE](LICENSE) file for details.

---

## Built With

- [Microsoft Graph API](https://docs.microsoft.com/en-us/graph/)
- [Fluent UI React](https://react.fluentui.dev/)
- [Recharts](https://recharts.org/)
- [QuestPDF](https://www.questpdf.com/)
- [Azure Container Apps](https://azure.microsoft.com/en-us/products/container-apps/)
