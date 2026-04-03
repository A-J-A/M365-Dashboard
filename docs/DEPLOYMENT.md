# M365 Dashboard - Deployment Guide

Deploy your own M365 Dashboard in **under 15 minutes** with two automated scripts. No manual Azure Portal configuration required.

## Prerequisites

- **Azure subscription** with Owner or Contributor access
- **Microsoft 365 tenant** with Global Administrator access
- **Azure CLI** installed ([Install Guide](https://docs.microsoft.com/en-us/cli/azure/install-azure-cli))

> **No Docker required.** The build runs in Azure Container Registry — nothing needs to be installed locally beyond the Azure CLI.

---

## Quick Deploy (2 Steps)

### Step 1: Register the Entra ID App (2 minutes)

```powershell
# Clone the repository
git clone https://github.com/cloud1st/m365-dashboard.git
cd m365-dashboard

# Login to Azure
az login

# Run the registration script
.\scripts\Register-EntraApp.ps1
```

This script automatically:
- ✅ Creates the App Registration in Entra ID
- ✅ Adds all required Microsoft Graph API permissions
- ✅ Creates App Roles (Admin & Reader)
- ✅ Generates a client secret
- ✅ Grants admin consent
- ✅ Saves config to `entra-app-config.json` for the next step

### Step 2: Deploy Everything (10 minutes)

```powershell
.\scripts\Deploy-M365Dashboard.ps1
```

The script reads your saved config automatically and walks you through a few prompts (resource prefix, region, SQL password). It then creates everything in Azure:

| Resource | Purpose |
|---|---|
| Azure Key Vault | Stores all secrets — app never holds secrets in config |
| Container App (Managed Identity) | Pulls secrets from Key Vault at startup |
| Azure Container Registry | Hosts the Docker image |
| Azure SQL Database | Application data |
| Azure Maps Account | Map widgets |
| Log Analytics | Container logs |

At the end the script outputs your app URL. Open it and sign in.

### 🎉 Done!

No manual portal steps required. All secrets are stored in Key Vault and accessed via Managed Identity — no credentials exist in environment variables or config files.

---

## How Secrets Work

All sensitive values are stored in Azure Key Vault at deploy time:

```
Key Vault Secret Name              → App Config Path
─────────────────────────────────────────────────────
AzureAd--TenantId                  → AzureAd:TenantId
AzureAd--ClientId                  → AzureAd:ClientId
AzureAd--ClientSecret              → AzureAd:ClientSecret
AzureAd--Audience                  → AzureAd:Audience
ConnectionStrings--DefaultConnection → ConnectionStrings:DefaultConnection
AzureMaps--SubscriptionKey         → AzureMaps:SubscriptionKey
```

The Container App has a **System-Assigned Managed Identity** with **Key Vault Secrets User** role. On startup, `Program.cs` calls `AddAzureKeyVault()` using `DefaultAzureCredential` — no passwords or connection strings ever appear in environment variables or config files.

---

## What Gets Created

| Resource | SKU | Monthly Cost |
|----------|-----|--------------|
| Container App | Consumption | ~$0-20 (scales to zero) |
| SQL Database | Basic (5 DTU) | ~$5 |
| Container Registry | Basic | ~$5 |
| Azure Maps | Gen2 | Free tier (250k transactions) |
| Key Vault | Standard | ~$0.03/secret |
| Log Analytics | Pay-per-GB | ~$2-5 |
| **Total** | | **~$15-35/month** |

---

## Manual Deployment (Alternative)

If you prefer to set things up manually or the scripts don't work in your environment:

<details>
<summary>Click to expand manual steps</summary>

### 1. Create App Registration Manually

1. Go to [Azure Portal](https://portal.azure.com) → **Microsoft Entra ID** → **App registrations**
2. Click **New registration**
   - Name: `M365 Dashboard`
   - Supported account types: **Single tenant**
   - Redirect URI: Leave blank
3. Note the **Application (client) ID** and **Directory (tenant) ID**

4. Go to **Certificates & secrets** → **New client secret**
   - Description: `M365 Dashboard`
   - Expires: 24 months
   - **Copy the secret value immediately**

5. Go to **API permissions** → **Add a permission** → **Microsoft Graph** → **Application permissions**

   Add these permissions:
   | Permission | Purpose |
   |------------|---------|
   | User.Read.All | Read user profiles |
   | Group.Read.All | Read groups |
   | Directory.Read.All | Read directory data |
   | Device.Read.All | Read device info |
   | DeviceManagementManagedDevices.Read.All | Intune devices |
   | DeviceManagementConfiguration.Read.All | Intune config |
   | DeviceManagementApps.Read.All | Intune apps |
   | SecurityEvents.Read.All | Security events |
   | IdentityRiskyUser.Read.All | Risky users |
   | IdentityRiskEvent.Read.All | Risk events |
   | Reports.Read.All | Usage reports |
   | AuditLog.Read.All | Audit logs |
   | Mail.Read | Email security |
   | Domain.Read.All | Domain info |
   | Organization.Read.All | Org info |
   | Policy.Read.All | Policies |

6. Click **Grant admin consent**

7. Go to **App roles** → **Create app role**
   - Display name: `Dashboard Admin`
   - Value: `Dashboard.Admin`
   - Description: `Full admin access`
   
   Create another:
   - Display name: `Dashboard Reader`
   - Value: `Dashboard.Reader`
   - Description: `Read-only access`

### 2. Deploy Infrastructure via Azure Portal

Use the Bicep templates in `/infra` or create resources manually:
- Resource Group
- Azure Container Registry (Basic)
- Azure Container App Environment
- Azure Container App
- Azure SQL Server + Database (Basic)
- Azure Maps Account (Gen2)
- Azure Key Vault

### 3. Build and Push Docker Image

```bash
az acr login --name <your-acr-name>
docker build -t <your-acr>.azurecr.io/m365dashboard:latest .
docker push <your-acr>.azurecr.io/m365dashboard:latest
```

### 4. Configure Container App

Set these environment variables:
- `ASPNETCORE_ENVIRONMENT`: `Production`
- `ConnectionStrings__DefaultConnection`: SQL connection string
- `AzureAd__TenantId`: Your tenant ID
- `AzureAd__ClientId`: Your client ID
- `AzureAd__ClientSecret`: Your client secret
- `AzureMaps__SubscriptionKey`: Maps subscription key

</details>

---

## Updating the Dashboard

To update to a new version:

```powershell
cd m365-dashboard
git pull

# Rebuild and push
az acr login --name <your-acr-name>
docker build -t <your-acr>.azurecr.io/m365dashboard:latest .
docker push <your-acr>.azurecr.io/m365dashboard:latest

# Update container app
az containerapp update `
  --name m365dash-prod-app `
  --resource-group m365dash-prod-rg `
  --image <your-acr>.azurecr.io/m365dashboard:latest
```

Or enable GitHub Actions for automatic deployments (see `.github/workflows/deploy.yml`).

---

## Troubleshooting

### "Failed to fetch report data"
- Ensure all Graph API permissions are granted with admin consent
- Check the App Registration has the correct permissions

### "Authentication failed"
- Verify the redirect URI matches exactly (including https://)
- Check client secret hasn't expired
- Ensure Access and ID tokens are enabled in Authentication settings

### PDF reports not generating
- PDFs only work in the Docker container (Linux)
- Local development automatically falls back to Word documents

### Container App not starting
- Check Log Analytics for error logs
- Verify SQL connection string is correct
- Ensure Container Registry credentials are valid

---

## Support

- 📖 [Full Documentation](./docs/)
- 🐛 [Report Issues](https://github.com/cloud1st/m365-dashboard/issues)
- 💬 [Discussions](https://github.com/cloud1st/m365-dashboard/discussions)

---

## License

MIT License - Free to use, modify, and distribute.
