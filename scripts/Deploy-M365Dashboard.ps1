<# 
.SYNOPSIS
    Deploy M365 Dashboard to Azure
.DESCRIPTION
    This script deploys the M365 Dashboard infrastructure and application to Azure.
.PARAMETER NamePrefix
    Prefix for all Azure resources (default: m365dash)
.PARAMETER Location
    Azure region (default: uksouth)
.PARAMETER Environment
    Environment name (default: prod)
.PARAMETER TenantId
    Entra ID Tenant ID
.PARAMETER ClientId
    Entra ID App Registration Client ID
.PARAMETER ClientSecret
    Entra ID App Registration Client Secret
.PARAMETER SqlPassword
    SQL Server admin password
.EXAMPLE
    .\Deploy-M365Dashboard.ps1 -TenantId "xxx" -ClientId "xxx" -ClientSecret "xxx" -SqlPassword "xxx"
#>

param(
    [string]$NamePrefix,
    [string]$Location,
    [string]$Environment = "prod",
    [string]$TenantId,
    [string]$ClientId,
    [string]$ClientSecret,
    [string]$SqlPassword
)

$ErrorActionPreference = "Stop"

# ============================================================================
# Deployment Mode & Login
# ============================================================================
Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "M365 Dashboard - Deployment Script" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Deployment Mode" -ForegroundColor Cyan
Write-Host "---------------" -ForegroundColor Cyan
Write-Host "  [1] Standard          - App registration and Azure resources in the same tenant" -ForegroundColor White
Write-Host "  [2] MSP / Multi-tenant - App registration in client's tenant, Azure resources in your subscription" -ForegroundColor White
Write-Host ""
$deployMode = Read-Host "Select mode (1-2)"
$isMspMode = ($deployMode -eq "2")

if ($isMspMode) {
    Write-Host ""
    Write-Host "  MSP mode selected." -ForegroundColor Cyan
    Write-Host "  Step 1 of 2: Login as a Global Admin in the CLIENT'S Microsoft 365 tenant" -ForegroundColor Yellow
    Write-Host "  (This is used to create the app registration in their tenant)" -ForegroundColor Gray
    Write-Host ""
    Read-Host "  Press Enter to open browser login for the CLIENT tenant"
    cmd /c "az logout 2>nul" | Out-Null
    cmd /c "az login" | Out-Null
    if ($LASTEXITCODE -ne 0) {
        Write-Host "  Login failed." -ForegroundColor Red; exit 1
    }
    $clientTenantAccountJson = cmd /c "az account show 2>nul"
    $clientTenantAccount = $clientTenantAccountJson | ConvertFrom-Json
    Write-Host "  Logged in as: $($clientTenantAccount.user.name) (tenant: $($clientTenantAccount.tenantId))" -ForegroundColor Green
    $clientTenantId = $clientTenantAccount.tenantId
} else {
    Write-Host ""
    Write-Host "Checking Azure CLI login..." -ForegroundColor Yellow
    $ErrorActionPreference = "Continue"
    $currentAccountJson = cmd /c "az account show 2>nul"
    $ErrorActionPreference = "Stop"

    if ($currentAccountJson) {
        $currentAccount = $currentAccountJson | ConvertFrom-Json
        $currentUser = $currentAccount.user.name
        $currentTenant = $currentAccount.tenantId
        Write-Host ""
        Write-Host "  Currently logged in as: $currentUser" -ForegroundColor White
        Write-Host "  Tenant ID:              $currentTenant" -ForegroundColor White
        Write-Host ""
        Write-Host "  [1] Continue as $currentUser" -ForegroundColor White
        Write-Host "  [2] Login as a different user" -ForegroundColor White
        Write-Host ""
        $loginChoice = Read-Host "Select option (1-2)"
        if ($loginChoice -eq "2") {
            cmd /c "az logout 2>nul" | Out-Null
            cmd /c "az login" | Out-Null
            if ($LASTEXITCODE -ne 0) {
                Write-Host "  Login failed." -ForegroundColor Red; exit 1
            }
            $currentAccountJson = cmd /c "az account show 2>nul"
            $currentAccount = $currentAccountJson | ConvertFrom-Json
            Write-Host "  Logged in as: $($currentAccount.user.name)" -ForegroundColor Green
        } else {
            Write-Host "  Continuing as $currentUser" -ForegroundColor Green
        }
    } else {
        Write-Host "  Not logged in. Launching browser login..." -ForegroundColor Yellow
        cmd /c "az login" | Out-Null
        if ($LASTEXITCODE -ne 0) {
            Write-Host "  Login failed." -ForegroundColor Red; exit 1
        }
        $currentAccountJson = cmd /c "az account show 2>nul"
        $currentAccount = $currentAccountJson | ConvertFrom-Json
        Write-Host "  Logged in as: $($currentAccount.user.name)" -ForegroundColor Green
    }
}

# Azure region options
$regionOptions = @{
    "1"  = @{ Code = "uksouth";       Name = "UK South" }
    "2"  = @{ Code = "ukwest";        Name = "UK West" }
    "3"  = @{ Code = "northeurope";   Name = "North Europe (Ireland)" }
    "4"  = @{ Code = "westeurope";    Name = "West Europe (Netherlands)" }
    "5"  = @{ Code = "eastus";        Name = "East US" }
    "6"  = @{ Code = "eastus2";       Name = "East US 2" }
    "7"  = @{ Code = "westus";        Name = "West US" }
    "8"  = @{ Code = "westus2";       Name = "West US 2" }
    "9"  = @{ Code = "centralus";     Name = "Central US" }
    "10" = @{ Code = "australiaeast"; Name = "Australia East" }
}

# ============================================================================
# Entra ID App Registration
# ============================================================================
Write-Host ""
Write-Host "Entra ID App Registration" -ForegroundColor Cyan
Write-Host "-------------------------" -ForegroundColor Cyan

$configPath = Join-Path (Join-Path $PSScriptRoot "..") "entra-app-config.json"
$configExists = Test-Path $configPath

if ($TenantId -and $ClientId -and $ClientSecret) {
    # All values passed as parameters - skip prompt
    Write-Host "  Using credentials passed as parameters" -ForegroundColor Gray
} else {
    Write-Host ""
    if ($configExists) {
        $savedConfig = Get-Content $configPath | ConvertFrom-Json
        Write-Host "  [1] Create a new app registration in this tenant" -ForegroundColor White
        Write-Host "  [2] Use existing config ($($savedConfig.AppName), created $($savedConfig.CreatedAt))" -ForegroundColor White
        Write-Host "  [3] Enter app details manually" -ForegroundColor White
        Write-Host ""
        $appChoice = Read-Host "Select option (1-3)"
    } else {
        Write-Host "  No existing app config found." -ForegroundColor Yellow
        Write-Host "  [1] Create a new app registration in this tenant" -ForegroundColor White
        Write-Host "  [2] Enter app details manually" -ForegroundColor White
        Write-Host ""
        $appChoice = Read-Host "Select option (1-2)"
        # Remap so '2' = manual in both branches
        if ($appChoice -eq "2") { $appChoice = "3" }
    }

    switch ($appChoice) {
        "1" {
            # ----------------------------------------------------------------
            # Create a new app registration
            # ----------------------------------------------------------------
            Write-Host ""
            $appNameInput = Read-Host "App registration name (default: M365 Dashboard)"
            if ([string]::IsNullOrWhiteSpace($appNameInput)) { $appNameInput = "M365 Dashboard" }

            $graphAppId = "00000003-0000-0000-c000-000000000000"
            $ErrorActionPreference = "Continue"

            Write-Host "  Creating app registration '$appNameInput'..." -ForegroundColor Gray
            $newAppRaw = cmd /c "az ad app create --display-name `"$appNameInput`" --sign-in-audience AzureADMyOrg --enable-access-token-issuance true --enable-id-token-issuance true 2>&1"
            $newAppJson = ($newAppRaw | Where-Object { $_ -notmatch '^WARNING:' }) -join "`n"
            if ($LASTEXITCODE -ne 0 -or -not $newAppJson -or $newAppJson -notmatch '"appId"') {
                Write-Host "  Failed to create app registration:" -ForegroundColor Red
                Write-Host $newAppRaw -ForegroundColor Red
                exit 1
            }
            $newApp = $newAppJson | ConvertFrom-Json
            $ClientId = $newApp.appId
            $appObjectIdNew = $newApp.id
            Write-Host "  App created. Client ID: $ClientId" -ForegroundColor Green

            # Create service principal
            Write-Host "  Creating service principal..." -ForegroundColor Gray
            cmd /c "az ad sp create --id $ClientId 2>nul" | Out-Null

            # Add Microsoft Graph permissions
            Write-Host "  Adding Microsoft Graph permissions..." -ForegroundColor Gray
            $graphPermissions = @(
                # Core
                @{ id = "df021288-bdef-4463-88db-98f22de89214"; name = "User.Read.All" }
                @{ id = "5b567255-7703-4780-807c-7be8301ae99b"; name = "Group.Read.All" }
                @{ id = "7ab1d382-f21e-4acd-a863-ba3e13f7da61"; name = "Directory.Read.All" }
                @{ id = "498476ce-e0fe-48b0-b801-37ba7e2685c6"; name = "Organization.Read.All" }
                @{ id = "246dd0d5-5bd0-4def-940b-0421030a5b68"; name = "Policy.Read.All" }
                @{ id = "dbb9058a-0e50-45d7-ae91-66909b5d4664"; name = "Domain.Read.All" }
                # Devices & Intune
                @{ id = "7438b122-aefc-4978-80ed-43db9fcc7715"; name = "Device.Read.All" }
                @{ id = "2f51be20-0bb4-4fed-bf7b-db946066c75e"; name = "DeviceManagementManagedDevices.Read.All" }
                @{ id = "dc377aa6-52d8-4e23-b271-2a7ae04cedf3"; name = "DeviceManagementConfiguration.Read.All" }
                @{ id = "7a6ee1e7-141e-4cec-ae74-d9db155731ff"; name = "DeviceManagementApps.Read.All" }
                @{ id = "06a5fe6d-c49d-46a7-b082-56b1b14103c7"; name = "DeviceManagementServiceConfig.Read.All" }
                # Mail & Reports
                @{ id = "810c84a8-4a9e-49e6-bf7d-12d183f40d01"; name = "Mail.Read" }
                @{ id = "230c1aed-a721-4c5d-9cb4-a90514e508ef"; name = "Reports.Read.All" }
                # Security
                @{ id = "bf394140-e372-4bf9-a898-299cfc7564e5"; name = "SecurityEvents.Read.All" }
                @{ id = "dc5007c0-2d7d-4c42-879c-2dab87571379"; name = "IdentityRiskyUser.Read.All" }
                @{ id = "6e472fd1-ad78-48da-a0f0-97ab2c6b769e"; name = "IdentityRiskEvent.Read.All" }
                @{ id = "b0afded3-3588-46d8-8b3d-9842eff778da"; name = "AuditLog.Read.All" }
                @{ id = "e0b77adb-e790-44a3-b0a0-257d06303687"; name = "UserAuthenticationMethod.Read.All" }
                @{ id = "93283d0a-6322-4fa8-966b-8c121624760d"; name = "AttackSimulation.Read.All" }
                # SharePoint
                @{ id = "332a536c-c7ef-4017-ab91-336970924f0d"; name = "Sites.Read.All" }
                # Teams
                @{ id = "45bbb07e-7321-4fd7-a8f6-3ff27e6a81c8"; name = "CallRecords.Read.All" }
            )
            foreach ($perm in $graphPermissions) {
                cmd /c "az ad app permission add --id $ClientId --api $graphAppId --api-permissions $($perm.id)=Role 2>nul" | Out-Null
                Write-Host "    + $($perm.name)" -ForegroundColor Gray
            }
            Write-Host "  Graph permissions added (21 permissions)" -ForegroundColor Green

            # Add Microsoft Defender for Endpoint permissions (separate API)
            Write-Host "  Adding Microsoft Defender for Endpoint permissions..." -ForegroundColor Gray
            $defenderAppId = "fc780465-2017-40d4-a0c5-307022471b92" # WindowsDefenderATP
            $defenderPermissions = @(
                @{ id = "ea8291d3-4b9a-44b5-bc3a-6cea3026dc79"; name = "Machine.Read.All" }
                @{ id = "41269fc5-d04d-4bfd-bce7-43a51cea049a"; name = "Vulnerability.Read.All" }
                @{ id = "02b005dd-f804-43b4-8fc7-078460413f74"; name = "Score.Read.All" }
            )
            foreach ($perm in $defenderPermissions) {
                cmd /c "az ad app permission add --id $ClientId --api $defenderAppId --api-permissions $($perm.id)=Role 2>nul" | Out-Null
                Write-Host "    + $($perm.name)" -ForegroundColor Gray
            }
            Write-Host "  Defender for Endpoint permissions added" -ForegroundColor Green

            # Add Exchange Online permissions (separate API)
            Write-Host "  Adding Exchange Online permissions..." -ForegroundColor Gray
            $exchangeAppId = "00000002-0000-0ff1-ce00-000000000000" # Office 365 Exchange Online
            $exchangePermId = "dc50a0fb-09a3-484d-be87-e023b12c6440" # Exchange.ManageAsApp
            cmd /c "az ad app permission add --id $ClientId --api $exchangeAppId --api-permissions $exchangePermId=Role 2>nul" | Out-Null
            Write-Host "    + Exchange.ManageAsApp" -ForegroundColor Gray
            Write-Host "  Exchange Online permissions added" -ForegroundColor Green

            # Add app roles
            Write-Host "  Adding app roles..." -ForegroundColor Gray
            $adminRoleId = [guid]::NewGuid().ToString()
            $readerRoleId = [guid]::NewGuid().ToString()
            $appRolesJson = "[{`"id`":`"$adminRoleId`",`"allowedMemberTypes`":[`"User`"],`"description`":`"Full administrative access to M365 Dashboard`",`"displayName`":`"Dashboard Admin`",`"isEnabled`":true,`"value`":`"Dashboard.Admin`"},{`"id`":`"$readerRoleId`",`"allowedMemberTypes`":[`"User`"],`"description`":`"Read-only access to M365 Dashboard`",`"displayName`":`"Dashboard Reader`",`"isEnabled`":true,`"value`":`"Dashboard.Reader`"}]"
            $rolesFile = [System.IO.Path]::GetTempFileName()
            [System.IO.File]::WriteAllText($rolesFile, $appRolesJson, [System.Text.Encoding]::UTF8)
            cmd /c "az ad app update --id $ClientId --app-roles @`"$rolesFile`" 2>nul" | Out-Null
            Remove-Item $rolesFile -ErrorAction SilentlyContinue
            Write-Host "  App roles added" -ForegroundColor Green

            # Create client secret
            Write-Host "  Creating client secret..." -ForegroundColor Gray
            $newSecretRaw = cmd /c "az ad app credential reset --id $ClientId --append --display-name M365Dashboard-Secret --years 2 2>&1"
            $newSecretJson = ($newSecretRaw | Where-Object { $_ -notmatch '^WARNING:' }) -join "`n"
            if ($LASTEXITCODE -ne 0 -or -not $newSecretJson -or $newSecretJson -notmatch '"password"') {
                Write-Host "  Failed to create client secret" -ForegroundColor Red
                exit 1
            }
            $newSecret = $newSecretJson | ConvertFrom-Json
            $ClientSecret = $newSecret.password
            if ([string]::IsNullOrWhiteSpace($ClientSecret)) {
                Write-Host "  Failed to extract client secret from response" -ForegroundColor Red
                Write-Host "  Raw response: $newSecretRaw" -ForegroundColor Red
                exit 1
            }
            Write-Host "  Client secret created (valid 2 years)" -ForegroundColor Green

            # Set Application ID URI and expose access_as_user scope
            # The identifier URI must be set first so Entra can resolve api://<clientId> as a resource
            Write-Host "  Setting Application ID URI..." -ForegroundColor Gray
            cmd /c "az ad app update --id $ClientId --identifier-uris `"api://$ClientId`" 2>nul" | Out-Null
            Write-Host "  Application ID URI set: api://$ClientId" -ForegroundColor Green

            Write-Host "  Exposing access_as_user scope..." -ForegroundColor Gray
            $appObjectIdForScope = ($newApp.id)
            $scopeId = [guid]::NewGuid().ToString()
            $scopeBody = "{`"api`":{`"oauth2PermissionScopes`":[{`"adminConsentDescription`":`"Allow the application to access M365 Dashboard on behalf of the signed-in user`",`"adminConsentDisplayName`":`"Access M365 Dashboard`",`"id`":`"$scopeId`",`"isEnabled`":true,`"type`":`"User`",`"userConsentDescription`":`"Allow the application to access M365 Dashboard on your behalf`",`"userConsentDisplayName`":`"Access M365 Dashboard`",`"value`":`"access_as_user`"} ] }}"
            $scopeFile = [System.IO.Path]::GetTempFileName() + ".json"
            [System.IO.File]::WriteAllText($scopeFile, $scopeBody, [System.Text.Encoding]::UTF8)
            cmd /c "az rest --method PATCH --uri `"https://graph.microsoft.com/v1.0/applications/$appObjectIdForScope`" --body @`"$scopeFile`" --headers Content-Type=application/json 2>nul" | Out-Null
            Remove-Item $scopeFile -ErrorAction SilentlyContinue
            Write-Host "  access_as_user scope exposed" -ForegroundColor Green

            # Grant admin consent - try az CLI first, fall back to Graph API
            Write-Host "  Granting admin consent..." -ForegroundColor Gray
            cmd /c "az ad app permission admin-consent --id $ClientId 2>nul" | Out-Null
            if ($LASTEXITCODE -eq 0) {
                Write-Host "  Admin consent granted" -ForegroundColor Green
            } else {
                # Fallback: use Graph API oAuth2PermissionGrants
                $spRaw = cmd /c "az ad sp show --id $ClientId --query id -o tsv 2>nul"
                $spObjId = ($spRaw | Where-Object { $_ -notmatch '^WARNING:' }) -join '' | ForEach-Object { $_.Trim() }
                $graphSpId = (cmd /c "az ad sp show --id $graphAppId --query id -o tsv 2>nul").Trim()
                if ($spObjId) {
                    $consentBody = "{`"clientId`":`"$spObjId`",`"consentType`":`"AllPrincipals`",`"resourceId`":`"$graphSpId`",`"scope`":`"openid profile`"}"
                    $consentFile = [System.IO.Path]::GetTempFileName()
                    [System.IO.File]::WriteAllText($consentFile, $consentBody, [System.Text.Encoding]::UTF8)
                    cmd /c "az rest --method POST --uri `"https://graph.microsoft.com/v1.0/oauth2PermissionGrants`" --body @`"$consentFile`" --headers Content-Type=application/json 2>nul" | Out-Null
                    Remove-Item $consentFile -ErrorAction SilentlyContinue
                }
                if ($LASTEXITCODE -eq 0) {
                    Write-Host "  Admin consent granted" -ForegroundColor Green
                } else {
                    Write-Host "  Could not auto-grant consent - grant manually after deployment:" -ForegroundColor Yellow
                    Write-Host "  Azure Portal > App registrations > $appNameInput > API permissions > Grant admin consent" -ForegroundColor Yellow
                }
            }

            $ErrorActionPreference = "Stop"

            # Get tenant ID from current login
            $TenantId = (cmd /c "az account show --query tenantId -o tsv 2>nul").Trim()

            # Assign Exchange Recipient Administrator role to the service principal
            # This is an Entra directory role and can be automated via Graph API
            # Role template ID: 31392ffb-586c-42d1-9346-e59415a2cc4e
            Write-Host "  Assigning Exchange Recipient Administrator role..." -ForegroundColor Gray
            $ErrorActionPreference = "Continue"
            $exchangeRoleTemplateId = "31392ffb-586c-42d1-9346-e59415a2cc4e"
            # Activate the role in the tenant if not already active
            $activeRoleRaw = cmd /c "az rest --method GET --uri `"https://graph.microsoft.com/v1.0/directoryRoles?`$filter=roleTemplateId eq '$exchangeRoleTemplateId'`" 2>nul"
            $activeRoleJson = ($activeRoleRaw | Where-Object { $_ -notmatch '^WARNING:' }) -join "`n"
            $activeRole = $activeRoleJson | ConvertFrom-Json
            $roleId = $activeRole.value[0].id
            if (-not $roleId) {
                # Role not yet activated in tenant - activate it
                $activateBody = "{`"roleTemplateId`":`"$exchangeRoleTemplateId`"}"
                $activateFile = [System.IO.Path]::GetTempFileName() + ".json"
                [System.IO.File]::WriteAllText($activateFile, $activateBody, [System.Text.Encoding]::UTF8)
                $activateRaw = cmd /c "az rest --method POST --uri `"https://graph.microsoft.com/v1.0/directoryRoles`" --body @`"$activateFile`" --headers Content-Type=application/json 2>nul"
                Remove-Item $activateFile -ErrorAction SilentlyContinue
                $activateJson = ($activateRaw | Where-Object { $_ -notmatch '^WARNING:' }) -join "`n"
                $roleId = ($activateJson | ConvertFrom-Json).id
            }
            # Get the service principal object ID
            $spObjIdForRole = (cmd /c "az ad sp show --id $ClientId --query id -o tsv 2>nul").Trim()
            if ($roleId -and $spObjIdForRole) {
                $memberBody = "{`"@odata.id`":`"https://graph.microsoft.com/v1.0/directoryObjects/$spObjIdForRole`"}"
                $memberFile = [System.IO.Path]::GetTempFileName() + ".json"
                [System.IO.File]::WriteAllText($memberFile, $memberBody, [System.Text.Encoding]::UTF8)
                $assignResult = cmd /c "az rest --method POST --uri `"https://graph.microsoft.com/v1.0/directoryRoles/$roleId/members/`$ref`" --body @`"$memberFile`" --headers Content-Type=application/json 2>&1"
                Remove-Item $memberFile -ErrorAction SilentlyContinue
                if ($LASTEXITCODE -eq 0 -or ($assignResult -join "") -match "already exists") {
                    Write-Host "  Exchange Recipient Administrator role assigned" -ForegroundColor Green
                } else {
                    Write-Host "  Could not assign Exchange Recipient Administrator role automatically" -ForegroundColor Yellow
                    Write-Host "  Assign manually: Entra admin centre > Roles & admins > Exchange Recipient Administrator > Add assignments" -ForegroundColor Yellow
                }
            } else {
                Write-Host "  Could not resolve role or service principal - assign Exchange Recipient Administrator manually" -ForegroundColor Yellow
            }
            $ErrorActionPreference = "Stop"

            # Save new config
            $newConfig = @{ TenantId = $TenantId; ClientId = $ClientId; ClientSecret = $ClientSecret; AppName = $appNameInput; CreatedAt = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss") }
            $newConfig | ConvertTo-Json | Out-File $configPath -Encoding UTF8
            Write-Host "  Config saved to entra-app-config.json" -ForegroundColor Green
        }
        "2" {
            # ----------------------------------------------------------------
            # Use existing config
            # ----------------------------------------------------------------
            $TenantId     = $savedConfig.TenantId
            $ClientId     = $savedConfig.ClientId
            $ClientSecret = $savedConfig.ClientSecret
            Write-Host "  Using Tenant ID:  $TenantId" -ForegroundColor Gray
            Write-Host "  Using Client ID:  $ClientId" -ForegroundColor Gray
            Write-Host "  Using Client Secret: ********" -ForegroundColor Gray
        }
        default {
            # ----------------------------------------------------------------
            # Enter manually
            # ----------------------------------------------------------------
            $TenantId = Read-Host "Enter Entra ID Tenant ID"
            $ClientId = Read-Host "Enter Entra ID Client ID"
            $secureSecret = Read-Host "Enter Entra ID Client Secret" -AsSecureString
            $ClientSecret = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($secureSecret))
        }
    }
}

# Prompt for resource name prefix
if (-not $NamePrefix) {
    Write-Host ""
    Write-Host "Resource Naming" -ForegroundColor Cyan
    Write-Host "---------------" -ForegroundColor Cyan
    Write-Host "Enter a short prefix for your Azure resources (e.g., 'm365dash', 'contoso')."
    Write-Host "This will be used to name: resource group, container app, SQL server, etc."
    Write-Host ""
    $NamePrefix = Read-Host "Resource name prefix (3-10 chars, letters/numbers only)"
    
    while ($NamePrefix -notmatch "^[a-zA-Z][a-zA-Z0-9]{2,9}$") {
        Write-Host "  Invalid prefix - must be 3-10 chars, start with a letter, letters/numbers only." -ForegroundColor Yellow
        $NamePrefix = Read-Host "Resource name prefix"
    }
    # Force lowercase - Azure Container Apps require lowercase names
    $NamePrefix = $NamePrefix.ToLower()
}

# Prompt for Azure region
if (-not $Location) {
    Write-Host ""
    Write-Host "Azure Region" -ForegroundColor Cyan
    Write-Host "------------" -ForegroundColor Cyan
    Write-Host "Select the Azure region for deployment:"
    Write-Host ""
    foreach ($key in ($regionOptions.Keys | Sort-Object { [int]$_ })) {
        Write-Host "  [$key] $($regionOptions[$key].Name)"
    }
    Write-Host ""
    $regionChoice = Read-Host "Enter number (1-10)"
    
    if ($regionOptions.ContainsKey($regionChoice)) {
        $Location = $regionOptions[$regionChoice].Code
        Write-Host "  Selected: $($regionOptions[$regionChoice].Name)" -ForegroundColor Green
    } else {
        Write-Host "  Invalid choice. Using default: UK South" -ForegroundColor Yellow
        $Location = "uksouth"
    }
}

# Select subscription
Write-Host ""
Write-Host "Azure Subscription" -ForegroundColor Cyan
Write-Host "------------------" -ForegroundColor Cyan

$accountJson = $currentAccountJson

# Get all subscriptions
$subscriptionsJson = cmd /c "az account list --query [?state=='Enabled'] -o json 2>nul"
$subscriptions = $subscriptionsJson | ConvertFrom-Json
$selectedSubscriptionName = ""

if ($subscriptions.Count -gt 1) {
    Write-Host "Multiple subscriptions found. Select one for deployment:"
    Write-Host ""
    
    $i = 1
    foreach ($sub in $subscriptions) {
        $isDefault = if ($sub.isDefault) { " (current)" } else { "" }
        Write-Host "  [$i] $($sub.name)$isDefault" -ForegroundColor White
        Write-Host "      $($sub.id)" -ForegroundColor Gray
        $i++
    }
    Write-Host ""
    $subChoice = Read-Host "Enter number (1-$($subscriptions.Count))"
    
    $selectedIndex = [int]$subChoice - 1
    if ($selectedIndex -ge 0 -and $selectedIndex -lt $subscriptions.Count) {
        $selectedSub = $subscriptions[$selectedIndex]
        $selectedSubscriptionName = $selectedSub.name
        Write-Host "  Selected: $selectedSubscriptionName" -ForegroundColor Green
        cmd /c "az account set --subscription $($selectedSub.id) 2>nul"
    } else {
        $currentAccount = $accountJson | ConvertFrom-Json
        $selectedSubscriptionName = $currentAccount.name
        Write-Host "  Invalid choice. Using current: $selectedSubscriptionName" -ForegroundColor Yellow
    }
} else {
    $currentAccount = $accountJson | ConvertFrom-Json
    $selectedSubscriptionName = $currentAccount.name
    Write-Host "Using subscription: $selectedSubscriptionName" -ForegroundColor Green
}

# Prompt for SQL password
if (-not $SqlPassword) {
    Write-Host ""
    Write-Host "SQL Server Password" -ForegroundColor Cyan
    Write-Host "-------------------" -ForegroundColor Cyan
    Write-Host "Create a strong password for the SQL Server admin account."
    Write-Host "Requirements: Min 8 chars, uppercase, lowercase, number, special char"
    Write-Host ""
    $securePassword = Read-Host "Enter SQL Admin Password" -AsSecureString
    $SqlPassword = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($securePassword))
}

# Get region display name
$locationName = $Location
foreach ($key in $regionOptions.Keys) {
    if ($regionOptions[$key].Code -eq $Location) {
        $locationName = $regionOptions[$key].Name
        break
    }
}

# Check if resource group already exists
$resourceGroupName = "$NamePrefix-$Environment-rg"
$existingRg = cmd /c "az group show --name $resourceGroupName --query location -o tsv 2>nul"

if ($existingRg) {
    $existingRg = $existingRg.Trim()
    Write-Host ""
    Write-Host "WARNING: Resource group '$resourceGroupName' already exists!" -ForegroundColor Yellow
    Write-Host "  Existing location: $existingRg" -ForegroundColor Yellow
    Write-Host "  Selected location: $Location" -ForegroundColor Yellow
    Write-Host ""
    
    if ($existingRg -ne $Location) {
        Write-Host "Location mismatch! You cannot deploy to a different region." -ForegroundColor Red
        Write-Host ""
        Write-Host "Options:" -ForegroundColor Cyan
        Write-Host "  [1] Delete existing resource group and start fresh"
        Write-Host "  [2] Use existing resource group (deploy to $existingRg)"
        Write-Host "  [3] Cancel and choose a different resource prefix"
        Write-Host ""
        $rgChoice = Read-Host "Enter choice (1-3)"
        
        switch ($rgChoice) {
            "1" {
                Write-Host "Deleting resource group '$resourceGroupName'..." -ForegroundColor Yellow
                Write-Host "  This may take a few minutes..." -ForegroundColor Gray
                cmd /c "az group delete --name $resourceGroupName --yes 2>nul"
                Write-Host "  Resource group deleted" -ForegroundColor Green
            }
            "2" {
                Write-Host "  Using existing location: $existingRg" -ForegroundColor Green
                $Location = $existingRg
                foreach ($key in $regionOptions.Keys) {
                    if ($regionOptions[$key].Code -eq $Location) {
                        $locationName = $regionOptions[$key].Name
                        break
                    }
                }
            }
            default {
                Write-Host "Deployment cancelled." -ForegroundColor Yellow
                exit 0
            }
        }
    } else {
        Write-Host "Existing resource group is in the same location. Will update existing resources." -ForegroundColor Green
    }
}

# Check for soft-deleted Key Vaults that might conflict
Write-Host ""
Write-Host "Checking for soft-deleted Key Vaults..." -ForegroundColor Yellow
$deletedVaults = cmd /c "az keyvault list-deleted --query `"[?starts_with(name, '$NamePrefix')].name`" -o tsv 2>nul"
if ($deletedVaults) {
    $vaultList = $deletedVaults -split "`n" | Where-Object { $_ -ne "" }
    foreach ($vault in $vaultList) {
        Write-Host "  Purging deleted vault: $vault" -ForegroundColor Gray
        cmd /c "az keyvault purge --name $vault 2>nul"
    }
    Write-Host "  Purged $($vaultList.Count) deleted vault(s)" -ForegroundColor Green
} else {
    Write-Host "  No conflicting deleted vaults found" -ForegroundColor Green
}

Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "M365 Dashboard Deployment" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "Subscription:    $selectedSubscriptionName"
Write-Host "Resource Prefix: $NamePrefix"
Write-Host "Resource Group:  $NamePrefix-$Environment-rg"
Write-Host "Location:        $locationName ($Location)"
Write-Host "Tenant ID:       $TenantId"
Write-Host "Client ID:       $ClientId"
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""

$confirm = Read-Host "Proceed with deployment? (Y/n)"
if ($confirm -eq "n" -or $confirm -eq "N") {
    Write-Host "Deployment cancelled." -ForegroundColor Yellow
    exit 0
}
Write-Host ""

# In MSP mode, switch from client tenant to your own Azure subscription for infrastructure deployment
if ($isMspMode) {
    Write-Host ""
    Write-Host "Step 2 of 2: Login to YOUR Azure subscription for infrastructure deployment" -ForegroundColor Yellow
    Write-Host "  (The app registration in the client tenant is complete)" -ForegroundColor Gray
    Write-Host "  Now logging in to your Cloud1st Azure subscription..." -ForegroundColor Gray
    Write-Host ""
    Read-Host "  Press Enter to open browser login for YOUR Azure subscription"
    cmd /c "az logout 2>nul" | Out-Null
    cmd /c "az login" | Out-Null
    if ($LASTEXITCODE -ne 0) {
        Write-Host "  Login failed." -ForegroundColor Red; exit 1
    }
    $yourAccountJson = cmd /c "az account show 2>nul"
    $yourAccount = $yourAccountJson | ConvertFrom-Json
    Write-Host "  Logged in as: $($yourAccount.user.name) (tenant: $($yourAccount.tenantId))" -ForegroundColor Green
    Write-Host ""
}

# Final validation before deploying
if ([string]::IsNullOrWhiteSpace($TenantId) -or [string]::IsNullOrWhiteSpace($ClientId) -or [string]::IsNullOrWhiteSpace($ClientSecret)) {
    Write-Host "ERROR: Missing required Entra ID credentials before deployment:" -ForegroundColor Red
    Write-Host "  TenantId:     $(if ($TenantId) { $TenantId } else { '(empty)' })" -ForegroundColor Red
    Write-Host "  ClientId:     $(if ($ClientId) { $ClientId } else { '(empty)' })" -ForegroundColor Red
    Write-Host "  ClientSecret: $(if ($ClientSecret) { '(set)' } else { '(empty)' })" -ForegroundColor Red
    exit 1
}

# Deploy infrastructure
Write-Host "Deploying Azure infrastructure..." -ForegroundColor Yellow
Write-Host "  This may take 5-10 minutes..." -ForegroundColor Gray

$infraPath = Join-Path (Join-Path (Join-Path $PSScriptRoot "..") "infra") "main.bicep"
$infraPath = (Resolve-Path $infraPath).Path

$deploymentName = "$NamePrefix-$Environment-$(Get-Date -Format 'yyyyMMddHHmmss')"

$deployingUserObjectId = cmd /c "az ad signed-in-user show --query id -o tsv 2>nul"
$deployingUserObjectId = $deployingUserObjectId.Trim()

$ErrorActionPreference = "Continue"
$deploymentResult = az deployment sub create `
    --name $deploymentName `
    --location $Location `
    --template-file $infraPath `
    --parameters namePrefix=$NamePrefix `
    --parameters location=$Location `
    --parameters environment=$Environment `
    --parameters entraIdTenantId=$TenantId `
    --parameters entraIdClientId=$ClientId `
    --parameters entraIdClientSecret="$ClientSecret" `
    --parameters sqlAdminPassword="$SqlPassword" `
    --parameters deployingUserObjectId=$deployingUserObjectId `
    --query properties.outputs -o json 2>&1

$ErrorActionPreference = "Stop"

if ($deploymentResult -match "ERROR" -or -not $deploymentResult -or $LASTEXITCODE -ne 0) {
    Write-Host "Deployment failed:" -ForegroundColor Red
    Write-Host $deploymentResult -ForegroundColor Red
    exit 1
}

$deploymentJson = $deploymentResult | Where-Object { $_ -notmatch "^WARNING:" } | Out-String
$deploymentOutput = $deploymentJson | ConvertFrom-Json

$resourceGroup = $deploymentOutput.resourceGroupName.value
$acrName = $deploymentOutput.containerRegistryName.value
$acrServer = $deploymentOutput.containerRegistryLoginServer.value
$appUrl = $deploymentOutput.containerAppUrl.value

Write-Host ""
Write-Host "Infrastructure deployed successfully!" -ForegroundColor Green
Write-Host ""

# Build and push Docker image using ACR Build (no local Docker required)
Write-Host "Building Docker image in Azure..." -ForegroundColor Yellow
Write-Host "  This may take 5-10 minutes..." -ForegroundColor Gray

# Must run from repo root so az acr build can find the Dockerfile
$repoRootPath = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$ErrorActionPreference = "Continue"
Push-Location $repoRootPath
try {
    cmd /c "az acr build --registry $acrName --image m365dashboard:latest --build-arg VITE_AZURE_CLIENT_ID=$ClientId --build-arg VITE_AZURE_TENANT_ID=$TenantId . 2>&1"
    if ($LASTEXITCODE -ne 0) {
        Write-Host "  Warning: Docker image build reported errors - check output above" -ForegroundColor Yellow
    } else {
        Write-Host "  Docker image built and pushed successfully" -ForegroundColor Green
    }
} finally {
    Pop-Location
}
$ErrorActionPreference = "Stop"

Write-Host ""
Write-Host "Updating Container App with new image..." -ForegroundColor Yellow
cmd /c "az containerapp update --name $NamePrefix-$Environment-app --resource-group $resourceGroup --image $acrServer/m365dashboard:latest 2>nul"

# Configure redirect URI and enable tokens
Write-Host ""
Write-Host "Configuring App Registration..." -ForegroundColor Yellow

# The MSAL redirectUri is window.location.origin (bare URL, no path)
# Register that as the SPA redirect URI in Entra
$appUrlClean = $appUrl.TrimEnd('/')
Write-Host "  Setting redirect URI: $appUrlClean" -ForegroundColor Gray

# Get the app's object ID (different from the client/app ID)
$appObjectId = (cmd /c "az ad app show --id $ClientId --query id -o tsv 2>nul").Trim()

# Write body to temp file to avoid JSON escaping issues in cmd
$redirectBodyFile = [System.IO.Path]::GetTempFileName()
$redirectBody = "{`"spa`":{`"redirectUris`":[`"$appUrlClean`"]}}"
[System.IO.File]::WriteAllText($redirectBodyFile, $redirectBody, [System.Text.Encoding]::UTF8)

$ErrorActionPreference = "Continue"
$uriUpdateResult = cmd /c "az rest --method PATCH --uri `"https://graph.microsoft.com/v1.0/applications/$appObjectId`" --body @`"$redirectBodyFile`" --headers Content-Type=application/json 2>&1"
Remove-Item $redirectBodyFile -ErrorAction SilentlyContinue
if ($LASTEXITCODE -eq 0) {
    Write-Host "  Redirect URI configured" -ForegroundColor Green
} else {
    Write-Host "  Warning: Could not set redirect URI: $uriUpdateResult" -ForegroundColor Yellow
}
$ErrorActionPreference = "Stop"

Write-Host "  Enabling access tokens and ID tokens..." -ForegroundColor Gray
cmd /c "az ad app update --id $ClientId --enable-access-token-issuance true --enable-id-token-issuance true 2>nul"
Write-Host "  Tokens enabled" -ForegroundColor Green

Write-Host "  Granting admin consent for API permissions..." -ForegroundColor Gray
$ErrorActionPreference = "Continue"
$consentResult = cmd /c "az ad app permission admin-consent --id $ClientId 2>&1"
if ($LASTEXITCODE -eq 0) {
    Write-Host "  Admin consent granted" -ForegroundColor Green
} else {
    # Fallback: grant via Graph API directly
    $graphConsentResult = cmd /c "az rest --method POST --uri `"https://graph.microsoft.com/v1.0/oauth2PermissionGrants`" --body `"{`\`"clientId`\`":`\`"$spId`\`",`\`"consentType`\`":`\`"AllPrincipals`\`",`\`"resourceId`\`":`\`"$spId`\`",`\`"scope`\`":`\`"openid profile`\`"}`" 2>&1"
    if ($LASTEXITCODE -eq 0) {
        Write-Host "  Admin consent granted via Graph API" -ForegroundColor Green
    } else {
        Write-Host "  Could not grant admin consent automatically - grant manually:" -ForegroundColor Yellow
        Write-Host "  Azure Portal > Entra ID > App registrations > $ClientId > API permissions > Grant admin consent" -ForegroundColor Yellow
    }
}
$ErrorActionPreference = "Stop"

Write-Host "  Assigning Dashboard Admin role to current user..." -ForegroundColor Gray
$currentUser = cmd /c "az ad signed-in-user show --query id -o tsv 2>nul"
$spId = cmd /c "az ad sp show --id $ClientId --query id -o tsv 2>nul"

if ($currentUser -and $spId) {
    $appRoles = cmd /c "az ad app show --id $ClientId --query appRoles -o json 2>nul" | ConvertFrom-Json
    $adminRole = $appRoles | Where-Object { $_.value -eq "Dashboard.Admin" }
    
    if ($adminRole) {
        $roleId = $adminRole.id
        # Write body to temp file to avoid shell escaping issues
        $roleBodyFile = [System.IO.Path]::GetTempFileName()
        $roleBody = "{`"principalId`":`"$currentUser`",`"resourceId`":`"$spId`",`"appRoleId`":`"$roleId`"}"
        [System.IO.File]::WriteAllText($roleBodyFile, $roleBody, [System.Text.Encoding]::UTF8)
        
        $ErrorActionPreference = "Continue"
        $assignResult = cmd /c "az rest --method POST --uri https://graph.microsoft.com/v1.0/users/$currentUser/appRoleAssignments --body @`"$roleBodyFile`" --headers Content-Type=application/json 2>&1"
        $ErrorActionPreference = "Stop"
        Remove-Item $roleBodyFile -ErrorAction SilentlyContinue
        
        if ($LASTEXITCODE -eq 0) {
            Write-Host "  Dashboard Admin role assigned to current user" -ForegroundColor Green
        } elseif (($assignResult -join "") -match "already exists") {
            Write-Host "  Dashboard Admin role already assigned" -ForegroundColor Green
        } else {
            Write-Host "  Could not assign role automatically (may need to assign manually in Entra Enterprise Apps)" -ForegroundColor Yellow
        }
    }
}

# Update local dev config
Write-Host ""
Write-Host "Updating local development config..." -ForegroundColor Yellow
$devSettingsPath = Join-Path $PSScriptRoot "..\src\M365Dashboard.Api\appsettings.Development.json"
if (Test-Path $devSettingsPath) {
    try {
        $sqlConnStr = "Server=tcp:$($deploymentOutput.sqlServerFqdn.value),1433;Initial Catalog=$($deploymentOutput.sqlDatabaseName.value);Persist Security Info=False;User ID=sqladmin;Password=$SqlPassword;MultipleActiveResultSets=True;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;"
        $devSettingsRaw = Get-Content $devSettingsPath -Raw
        $devSettingsObj = $devSettingsRaw | ConvertFrom-Json
        # Add/update properties (compatible with PS5 and PS7)
        if (-not $devSettingsObj.PSObject.Properties['ConnectionStrings']) {
            $devSettingsObj | Add-Member -NotePropertyName 'ConnectionStrings' -NotePropertyValue ([PSCustomObject]@{ DefaultConnection = $sqlConnStr })
        } else {
            $devSettingsObj.ConnectionStrings = [PSCustomObject]@{ DefaultConnection = $sqlConnStr }
        }
        if (-not $devSettingsObj.PSObject.Properties['KeyVault']) {
            $devSettingsObj | Add-Member -NotePropertyName 'KeyVault' -NotePropertyValue ([PSCustomObject]@{ Uri = $deploymentOutput.keyVaultUri.value })
        } else {
            $devSettingsObj.KeyVault = [PSCustomObject]@{ Uri = $deploymentOutput.keyVaultUri.value }
        }
        $devSettingsObj | ConvertTo-Json -Depth 10 | Out-File $devSettingsPath -Encoding UTF8
        Write-Host "  Local dev config updated with SQL connection string and Key Vault URI" -ForegroundColor Green
    } catch {
        Write-Host "  Could not update local dev config (non-critical): $_" -ForegroundColor Yellow
    }
} else {
    Write-Host "  appsettings.Development.json not found - run Register-EntraApp.ps1 first to generate it" -ForegroundColor Yellow
}

# ============================================================================
# Configure GitHub Actions secrets
# ============================================================================
Write-Host ""
Write-Host "Configuring GitHub Actions CI/CD..." -ForegroundColor Yellow

$acrUsername     = (cmd /c "az acr credential show --name $acrName --query username -o tsv 2>nul").Trim()
$acrPassword     = (cmd /c "az acr credential show --name $acrName --query passwords[0].value -o tsv 2>nul").Trim()
$subscriptionId  = (cmd /c "az account show --query id -o tsv 2>nul").Trim()
$containerAppName = "$NamePrefix-$Environment-app"
$spName          = "$NamePrefix-$Environment-github-actions"

# Create service principal for GitHub Actions (Contributor on the resource group)
Write-Host "  Creating service principal '$spName'..." -ForegroundColor Gray
$ErrorActionPreference = "Continue"
$spJson = cmd /c "az ad sp create-for-rbac --name `"$spName`" --role contributor --scopes /subscriptions/$subscriptionId/resourceGroups/$resourceGroup --sdk-auth 2>&1"
if ($LASTEXITCODE -ne 0 -or ($spJson -join "") -match '"error"') {
    Write-Host "  SP may already exist - resetting credentials..." -ForegroundColor Gray
    $spJson = cmd /c "az ad sp credential reset --name `"$spName`" --sdk-auth 2>&1"
}
$ErrorActionPreference = "Stop"
$spJson = ($spJson | Where-Object { $_ -notmatch "^WARNING:" }) -join "`n"

if (-not $spJson -or ($spJson -notmatch '"clientId"')) {
    Write-Host "  Could not create service principal automatically." -ForegroundColor Yellow
    $spJson = "<run manually: az ad sp create-for-rbac --name $spName --role contributor --scopes /subscriptions/$subscriptionId/resourceGroups/$resourceGroup --sdk-auth>"
}

# Detect GitHub repo slug from git remote
$repoRoot  = Split-Path $PSScriptRoot -Parent
$gitRemote = (cmd /c "git -C `"$repoRoot`" remote get-url origin 2>nul").Trim()
$repoSlug  = ""
if ($gitRemote -match "github\.com[:/](.+?)(\.git)?$") {
    $repoSlug = $Matches[1].Trim()
}

# Helper: print manual instructions
function Write-GitHubSecretsInstructions {
    $url = if ($repoSlug) { "https://github.com/$repoSlug/settings/secrets/actions" } else { "https://github.com/<owner>/<repo>/settings/secrets/actions" }
    Write-Host ""
    Write-Host "  Set these 6 secrets at:" -ForegroundColor Cyan
    Write-Host "  $url" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  Secret Name            Value" -ForegroundColor White
    Write-Host "  ---------------------  -----" -ForegroundColor DarkGray
    Write-Host "  AZURE_CREDENTIALS      (JSON block printed below)" -ForegroundColor White
    Write-Host "  ACR_LOGIN_SERVER       $acrServer" -ForegroundColor White
    Write-Host "  ACR_USERNAME           $acrUsername" -ForegroundColor White
    Write-Host "  ACR_PASSWORD           $acrPassword" -ForegroundColor White
    Write-Host "  CONTAINER_APP_NAME     $containerAppName" -ForegroundColor White
    Write-Host "  RESOURCE_GROUP         $resourceGroup" -ForegroundColor White
    Write-Host "  VITE_AZURE_CLIENT_ID   $ClientId" -ForegroundColor White
    Write-Host "  VITE_AZURE_TENANT_ID   $TenantId" -ForegroundColor White
    Write-Host ""
    Write-Host "  AZURE_CREDENTIALS value:" -ForegroundColor Cyan
    Write-Host $spJson -ForegroundColor DarkGray
    Write-Host ""
}

# Try gh CLI (fully automatic path)
$ghAvailable = $null
$ErrorActionPreference = "Continue"
$ghAvailable = cmd /c "gh --version 2>nul"
$ErrorActionPreference = "Stop"

$secretsSet = $false

if ($ghAvailable -and $repoSlug) {
    $ErrorActionPreference = "Continue"
    cmd /c "gh auth status 2>nul" | Out-Null
    $ghAuthed = ($LASTEXITCODE -eq 0)
    $ErrorActionPreference = "Stop"

    if ($ghAuthed) {
        Write-Host "  Setting GitHub Actions secrets via gh CLI..." -ForegroundColor Gray

        # Write AZURE_CREDENTIALS to a temp file to avoid JSON quoting issues
        $tempFile = [System.IO.Path]::GetTempFileName()
        [System.IO.File]::WriteAllText($tempFile, $spJson, [System.Text.Encoding]::UTF8)

        try {
            $ErrorActionPreference = "Continue"
            # Use PowerShell pipeline instead of cmd stdin redirection (more reliable cross-platform)
            Get-Content $tempFile -Raw | & gh secret set AZURE_CREDENTIALS --repo $repoSlug
            & gh secret set ACR_LOGIN_SERVER --body $acrServer --repo $repoSlug
            & gh secret set ACR_USERNAME --body $acrUsername --repo $repoSlug
            & gh secret set ACR_PASSWORD --body $acrPassword --repo $repoSlug
            & gh secret set CONTAINER_APP_NAME --body $containerAppName --repo $repoSlug
            & gh secret set RESOURCE_GROUP --body $resourceGroup --repo $repoSlug
            & gh secret set VITE_AZURE_CLIENT_ID --body $ClientId --repo $repoSlug
            & gh secret set VITE_AZURE_TENANT_ID --body $TenantId --repo $repoSlug
            $ErrorActionPreference = "Stop"

            # Verify
            $secretList = cmd /c "gh secret list --repo $repoSlug 2>nul"
            $expected = @("AZURE_CREDENTIALS", "ACR_LOGIN_SERVER", "ACR_USERNAME", "ACR_PASSWORD", "CONTAINER_APP_NAME", "RESOURCE_GROUP", "VITE_AZURE_CLIENT_ID", "VITE_AZURE_TENANT_ID")
            $missing  = $expected | Where-Object { $secretList -notmatch $_ }

            if ($missing.Count -eq 0) {
                Write-Host "  All 6 GitHub Actions secrets configured for: $repoSlug" -ForegroundColor Green
                Write-Host "  CI/CD is ready - every push to 'main' will auto-deploy" -ForegroundColor Green
                $secretsSet = $true
            } else {
                Write-Host "  Warning: could not verify secrets: $($missing -join ', ')" -ForegroundColor Yellow
            }
        } finally {
            Remove-Item $tempFile -ErrorAction SilentlyContinue
        }
    } else {
        Write-Host "  GitHub CLI not authenticated - run 'gh auth login' to enable automatic secret setup." -ForegroundColor Yellow
    }
} elseif (-not $ghAvailable) {
    Write-Host "  GitHub CLI not installed (https://cli.github.com) - set secrets manually." -ForegroundColor Yellow
} elseif (-not $repoSlug) {
    Write-Host "  Could not detect GitHub remote URL from git config - set secrets manually." -ForegroundColor Yellow
}

if (-not $secretsSet) {
    Write-GitHubSecretsInstructions
}

Write-Host ""
Write-Host "============================================" -ForegroundColor Green
Write-Host "Deployment Complete!" -ForegroundColor Green
Write-Host "============================================" -ForegroundColor Green
Write-Host ""
Write-Host "Your M365 Dashboard is available at:"
Write-Host "  $appUrl" -ForegroundColor Cyan
Write-Host ""
Write-Host "============================================" -ForegroundColor Yellow
Write-Host "Manual Steps Required" -ForegroundColor Yellow
Write-Host "============================================" -ForegroundColor Yellow
Write-Host ""
Write-Host "The following steps require manual configuration:" -ForegroundColor White
Write-Host ""
Write-Host "1. Grant Admin Consent (required)" -ForegroundColor Cyan
Write-Host "   Azure Portal > Entra ID > App registrations > $ClientId" -ForegroundColor Gray
Write-Host "   > API permissions > Grant admin consent for [your tenant]" -ForegroundColor Gray
Write-Host ""
Write-Host "2. Security Reader role in Exchange (required for Defender for Office data)" -ForegroundColor Cyan
Write-Host "   Exchange Admin Centre > Roles > Admin roles" -ForegroundColor Gray
Write-Host "   > View-Only Organization Management > Members tab > Add" -ForegroundColor Gray
Write-Host "   > Search for app registration by name and add it" -ForegroundColor Gray
Write-Host "   Exchange Admin Centre: https://admin.exchange.microsoft.com/#/adminRoles" -ForegroundColor Gray
Write-Host ""
Write-Host "3. Defender for Endpoint permissions (only if Defender P1/P2 licensed)" -ForegroundColor Cyan
Write-Host "   Azure Portal > App registrations > $ClientId > API permissions" -ForegroundColor Gray
Write-Host "   > Add permission > APIs my org uses > WindowsDefenderATP" -ForegroundColor Gray
Write-Host "   > Add: Machine.Read.All, Vulnerability.Read.All, Score.Read.All" -ForegroundColor Gray
Write-Host "   > Grant admin consent" -ForegroundColor Gray
Write-Host ""
Write-Host "Once complete, open the dashboard and sign in:" -ForegroundColor White
Write-Host "  $appUrl" -ForegroundColor Cyan
Write-Host ""
