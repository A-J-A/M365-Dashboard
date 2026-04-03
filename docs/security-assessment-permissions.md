# Security Assessment - Required Permissions (Read-Only)

This document outlines the Microsoft Graph API permissions required for the Security Assessment feature.
We use read-only permissions wherever possible to follow the principle of least privilege.

## Microsoft Graph API Permissions

| Permission | Type | Purpose | Cloud1st Equivalent |
|------------|------|---------|---------------------|
| `User.Read.All` | Application | Read user profiles, MFA status, license assignments | User.ReadWrite.All |
| `Directory.Read.All` | Application | Read directory data including roles | Directory.ReadWrite.All |
| `Organization.Read.All` | Application | Read organization settings and domains | Organization.ReadWrite.All |
| `Policy.Read.All` | Application | Read authorization, authentication policies | Policy.ReadWrite.* |
| `Group.Read.All` | Application | Read groups for Teams/M365 Groups | Group.ReadWrite.All |
| `RoleManagement.Read.All` | Application | Read admin role assignments | RoleManagement.ReadWrite.Directory |
| `IdentityRiskyUser.Read.All` | Application | Read risky users (optional) | IdentityRiskyUser.ReadWrite.All |
| `AuditLog.Read.All` | Application | Read audit logs for sign-in analysis | AuditLog.Read.All ✓ |
| `Reports.Read.All` | Application | Read usage reports | Reports.Read.All ✓ |
| `DeviceManagementConfiguration.Read.All` | Application | Read Intune compliance policies | DeviceManagementConfiguration.ReadWrite.All |
| `DeviceManagementManagedDevices.Read.All` | Application | Read managed devices | DeviceManagementManagedDevices.ReadWrite.All |
| `IdentityProvider.Read.All` | Application | Read identity providers config | - |
| `SecurityEvents.Read.All` | Application | Read security events | SecurityEvents.ReadWrite.All |

## Exchange Online Permissions

| Permission | Type | Purpose |
|------------|------|---------|
| `Exchange.ManageAsApp` | Application | Required for Exchange Online PowerShell cmdlets |

Note: Exchange Online checks require connecting via ExchangeOnlineManagement module.

## Permissions NOT Required (Read-Only Alternative)

These permissions from the Cloud1st solution are NOT needed for read-only assessment:

- `*.ReadWrite.*` - All write permissions can be replaced with `.Read.` variants
- `User.Invite.All` - We don't invite users
- `User.Export.All` - We don't export user data
- `MailboxSettings.ReadWrite` - Only reading settings
- `SharePointTenantSettings.ReadWrite.All` - Only reading settings

## API Endpoints Used

### Entra ID Checks
- `GET /organization` - Tenant info
- `GET /users` - User statistics
- `GET /subscribedSkus` - License info
- `GET /directoryRoles` - Admin roles
- `GET /policies/authorizationPolicy` - Auth settings
- `GET /policies/identitySecurityDefaultsEnforcementPolicy` - Security defaults
- `GET /identity/conditionalAccess/policies` - CA policies
- `GET /domains` - Domain settings

### Intune Checks
- `GET /deviceManagement/deviceCompliancePolicies` - Compliance policies
- `GET /deviceAppManagement/managedAppPolicies` - App protection policies

### Exchange Online Checks (via PowerShell)
- `Get-TransportConfig` - Transport settings
- `Get-OrganizationConfig` - Org settings
- `Get-AntiPhishPolicy` - Anti-phishing settings

## How to Add Permissions

1. Go to Azure Portal > App Registrations
2. Select your app
3. Go to API Permissions
4. Click "Add a permission"
5. Select "Microsoft Graph"
6. Choose "Application permissions"
7. Add each permission listed above
8. Click "Grant admin consent"

## Minimal Permission Set (Essential Only)

For a minimal deployment, these are the essential permissions:

```
User.Read.All
Directory.Read.All
Organization.Read.All
Policy.Read.All
DeviceManagementConfiguration.Read.All
```

Additional checks become available with more permissions.
