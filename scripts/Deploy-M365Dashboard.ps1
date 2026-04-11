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

# Update script is separate — see scripts/Update-M365Dashboard.ps1

param(
    [string]$NamePrefix,
    [string]$Location,
    [string]$Environment = "prod",
    [string]$TenantId,
    [string]$ClientId,
    [string]$ClientSecret,
    [string]$SqlPassword,
    [string]$DeployMode = "Standard",        # Standard or MSP
    [string]$CredentialType = "Secret",      # Secret or Certificate
    [string]$SubscriptionId = "",            # If set, az account set to this before deploying
    [switch]$NonInteractive                  # Skip all prompts (used by wizard)
)

$ErrorActionPreference = "Stop"

# ============================================================================
# Helper: Interactive Azure Login (browser or device code)
# ============================================================================
function Invoke-AzLogin {
    param(
        [string]$Prompt = "Login",
        [switch]$AllowNoSubscriptions
    )
    Write-Host ""
    Write-Host "  Login method" -ForegroundColor Cyan
    Write-Host "  [1] Browser  - opens a browser window (default)" -ForegroundColor White
    Write-Host "  [2] Device code - visit https://microsoft.com/devicelogin and enter a code" -ForegroundColor White
    Write-Host ""
    $loginMethod = Read-Host "  Select login method (1-2, default 1)"
    $useDeviceCode = ($loginMethod -eq "2")

    cmd /c "az logout 2>nul"

    $ErrorActionPreference = "Continue"
    if ($useDeviceCode) {
        Write-Host "  Starting device code login..." -ForegroundColor Yellow
        Write-Host "  Watch for the code and URL to appear below, then open https://microsoft.com/devicelogin" -ForegroundColor Gray
        if ($AllowNoSubscriptions) {
            az login --use-device-code --allow-no-subscriptions | Out-Null
        } else {
            az login --use-device-code | Out-Null
        }
    } else {
        Write-Host "  Opening browser..." -ForegroundColor Yellow
        if ($AllowNoSubscriptions) {
            az login --allow-no-subscriptions | Out-Null
        } else {
            az login | Out-Null
        }
    }
    $loginExit = $LASTEXITCODE
    $ErrorActionPreference = "Stop"

    # Fetch account details separately with explicit JSON output
    $ErrorActionPreference = "Continue"
    $accountRaw = az account show -o json 2>$null
    $ErrorActionPreference = "Stop"
    $accountJson = ($accountRaw | Where-Object { $_ -notmatch '^WARNING:' }) -join "`n"
    if (-not $accountJson -or $accountJson -notmatch '"tenantId"') {
        Write-Host "  Login failed." -ForegroundColor Red
        if (-not $useDeviceCode) {
            Write-Host "  Tip: try again and select option [2] (device code) if the browser is not working." -ForegroundColor Yellow
        }
        exit 1
    }
    return $accountJson
}

# ============================================================================
# Deployment Mode & Login
# ============================================================================
# ============================================================================
# Banner
# ============================================================================
Write-Host "" 
Write-Host "  ███╗   ███╗██████╗  ██████╗ ███████╗" -ForegroundColor Cyan
Write-Host "  ████╗ ████║╚════██╗██╔════╝ ██╔════╝" -ForegroundColor Cyan
Write-Host "  ██╔████╔██║ █████╔╝███████╗ ███████╗" -ForegroundColor Cyan
Write-Host "  ██║╚██╔╝██║ ╚═══██╗██╔═══██╗╚════██║" -ForegroundColor Cyan
Write-Host "  ██║ ╚═╝ ██║██████╔╝╚██████╔╝███████║" -ForegroundColor Cyan
Write-Host "  ╚═╝     ╚═╝╚═════╝  ╚═════╝ ╚══════╝" -ForegroundColor Cyan
Write-Host "" 
Write-Host "         Dashboard  Deployment" -ForegroundColor White
Write-Host "  ─────────────────────────────────────" -ForegroundColor DarkGray
Write-Host "  Microsoft 365 Tenant Management Portal" -ForegroundColor DarkGray
Write-Host "  Open Source | github.com/A-J-A/M365-Dashboard" -ForegroundColor DarkGray
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

if ($NonInteractive) {
    $deployMode = if ($DeployMode -eq "MSP") { "2" } else { "1" }
    Write-Host "  Non-interactive mode: using $DeployMode deployment" -ForegroundColor Gray
} else {
    $deployMode = Read-Host "Select mode (1-2)"
}
$isMspMode = ($deployMode -eq "2")

if ($isMspMode) {
    Write-Host ""
    Write-Host "  MSP mode selected." -ForegroundColor Cyan
    Write-Host "  Step 1 of 2: Login as a Global Admin in the CLIENT'S Microsoft 365 tenant" -ForegroundColor Yellow
    if ($NonInteractive) {
        # In NonInteractive mode the wizard already handled client tenant login.
        # If TenantId was passed directly, use it — otherwise read from current session.
        if ($TenantId) {
            $clientTenantId = $TenantId
            $clientTenantAccountJson = "{`"tenantId`":`"$TenantId`",`"user`":{`"name`":`"(passed via parameter)`"}}"
            Write-Host "  Using client tenant ID from parameter: $clientTenantId" -ForegroundColor Gray
        } else {
            $ErrorActionPreference = "Continue"
            $clientTenantAccountJson = (cmd /c "az account show -o json 2>nul" | Where-Object { $_ -notmatch '^WARNING:' }) -join "`n"
            $ErrorActionPreference = "Stop"
            Write-Host "  Using existing login session" -ForegroundColor Gray
        }
    } else {
        Read-Host "  Press Enter when ready to log in to the CLIENT tenant"
        $clientTenantAccountJson = Invoke-AzLogin -AllowNoSubscriptions
    }
    $clientTenantAccount = $clientTenantAccountJson | ConvertFrom-Json
    Write-Host "  Logged in as: $($clientTenantAccount.user.name) (tenant: $($clientTenantAccount.tenantId))" -ForegroundColor Green
    $clientTenantId = $clientTenantAccount.tenantId
    # Also set currentAccountJson so later subscription code works
    $currentAccountJson = $clientTenantAccountJson
} else {
    Write-Host ""
    Write-Host "Checking Azure CLI login..." -ForegroundColor Yellow
    $ErrorActionPreference = "Continue"
    $currentAccountJson = (cmd /c "az account show -o json 2>nul" | Where-Object { $_ -notmatch '^WARNING:' }) -join "`n"
    $ErrorActionPreference = "Stop"

    # In NonInteractive mode, if account show fails but SubscriptionId was provided,
    # try setting the subscription first (handles --allow-no-subscriptions login state)
    if (-not $currentAccountJson -and $NonInteractive -and $SubscriptionId) {
        cmd /c "az account set --subscription $SubscriptionId 2>nul" | Out-Null
        $ErrorActionPreference = "Continue"
        $currentAccountJson = (cmd /c "az account show -o json 2>nul" | Where-Object { $_ -notmatch '^WARNING:' }) -join "`n"
        $ErrorActionPreference = "Stop"
    }

    if ($currentAccountJson) {
        $currentAccount = $currentAccountJson | ConvertFrom-Json
        $currentUser = $currentAccount.user.name
        $currentTenant = $currentAccount.tenantId
        Write-Host ""
        Write-Host "  Currently logged in as: $currentUser" -ForegroundColor White
        Write-Host "  Tenant ID:              $currentTenant" -ForegroundColor White
        Write-Host ""
        $loginChoice = if ($NonInteractive) { "1" } else {
            Write-Host "  [1] Continue as $currentUser" -ForegroundColor White
            Write-Host "  [2] Login as a different user" -ForegroundColor White
            Write-Host ""
            Read-Host "Select option (1-2)"
        }
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
        if ($NonInteractive) {
            Write-Host "  ERROR: Not logged in to Azure. Run 'az login' before launching the wizard." -ForegroundColor Red
            exit 1
        }
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

# In MSP mode, use a per-client-tenant config file so it doesn't collide with the
# MSP's own standard deployment config.
if ($isMspMode -and $clientTenantId) {
    $configPath = Join-Path (Join-Path $PSScriptRoot "..") "entra-app-config-$clientTenantId.json"
} else {
    $configPath = Join-Path (Join-Path $PSScriptRoot "..") "entra-app-config.json"
}
$configExists = Test-Path $configPath

# In MSP mode, also check if saved config belongs to the current client tenant
if ($configExists -and $isMspMode) {
    $savedConfigCheck = Get-Content $configPath | ConvertFrom-Json
    if ($savedConfigCheck.TenantId -ne $clientTenantId) {
        $configExists = $false  # Wrong tenant - ignore it
    }
}

if ($TenantId -and $ClientId -and $ClientSecret) {
    # All values passed as parameters - skip prompt
    Write-Host "  Using credentials passed as parameters" -ForegroundColor Gray
} else {
    Write-Host ""
    if ($NonInteractive) {
        # Non-interactive: always create new app registration
        $appChoice = "1"
        Write-Host "  Non-interactive mode: creating new app registration" -ForegroundColor Gray
    } elseif ($configExists) {
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
            $appNameInput = if ($NonInteractive) { "M365 Dashboard" } else {
                Write-Host ""
                $n = Read-Host "App registration name (default: M365 Dashboard)"
                if ([string]::IsNullOrWhiteSpace($n)) { "M365 Dashboard" } else { $n }
            }

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
                @{ id = "b633e1c5-b582-4048-a93e-9f11b44c7e96"; name = "Mail.Send" }
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
            Write-Host "  Graph permissions added (22 permissions)" -ForegroundColor Green

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

            # ----------------------------------------------------------------
            # Credential type selection
            # ----------------------------------------------------------------
            Write-Host ""
            if ($NonInteractive) {
                $useCertAuth = ($CredentialType -eq "Certificate")
                Write-Host "  Using credential type: $CredentialType" -ForegroundColor Gray
            } else {
                Write-Host "  Credential Type" -ForegroundColor Cyan
                Write-Host "  [1] Client Secret  - simpler, but may be blocked by tenant credential policies" -ForegroundColor White
                Write-Host "  [2] Certificate    - more secure, works even when client secrets are blocked" -ForegroundColor White
                Write-Host ""
                $credChoice = Read-Host "  Select credential type (1-2, default 1)"
                $useCertAuth = ($credChoice -eq "2")
            }
            $certThumbprint = $null
            $certPfxBase64 = $null

            if (-not $useCertAuth) {
                # ---- Client secret path ----
                Write-Host "  Creating client secret..." -ForegroundColor Gray
                Start-Sleep -Seconds 5
                $newSecretRaw = cmd /c "az ad app credential reset --id $ClientId --append --display-name M365Dashboard-Secret --years 2 2>&1"
                $newSecretJson = ($newSecretRaw | Where-Object { $_ -notmatch '^WARNING:' }) -join "`n"

                if ($LASTEXITCODE -ne 0 -or -not $newSecretJson -or $newSecretJson -notmatch '"password"') {
                    # Retry once for propagation delay
                    Write-Host "  Retrying..." -ForegroundColor Yellow
                    Start-Sleep -Seconds 10
                    $newSecretRaw = cmd /c "az ad app credential reset --id $ClientId --append --display-name M365Dashboard-Secret --years 2 2>&1"
                    $newSecretJson = ($newSecretRaw | Where-Object { $_ -notmatch '^WARNING:' }) -join "`n"
                }

                if ($LASTEXITCODE -ne 0 -or -not $newSecretJson -or $newSecretJson -notmatch '"password"') {
                    $errMsg = ($newSecretRaw | Where-Object { $_ -match 'ERROR:' }) -join ''
                    Write-Host ""
                    Write-Host "  Client secret creation failed:" -ForegroundColor Red
                    Write-Host "  $errMsg" -ForegroundColor Red
                    if ($errMsg -match 'policy|Credential type not allowed') {
                        Write-Host "  This tenant has a credential policy blocking client secrets." -ForegroundColor Yellow
                        Write-Host "  Re-run the script and select option [2] Certificate instead." -ForegroundColor Yellow
                    }
                    exit 1
                }
            }

            if ($useCertAuth) {
                # Generate a self-signed certificate using PowerShell (no external tools needed).
                # The public key (.cer) is uploaded to the app registration in the CLIENT tenant.
                # The private key (.pfx, no password) is stored in Key Vault in the MSP Azure subscription.
                # The app reads the PFX from Key Vault at runtime via the config provider.
                Write-Host "  Generating self-signed certificate..." -ForegroundColor Gray
                $certSubject = "CN=M365Dashboard-$ClientId"
                $cert = New-SelfSignedCertificate `
                    -Subject $certSubject `
                    -CertStoreLocation "Cert:\CurrentUser\My" `
                    -KeyExportPolicy Exportable `
                    -KeySpec Signature `
                    -KeyLength 2048 `
                    -HashAlgorithm SHA256 `
                    -NotAfter (Get-Date).AddYears(2)
                $certThumbprint = $cert.Thumbprint
                Write-Host "  Certificate generated. Thumbprint: $certThumbprint" -ForegroundColor Green

                # Export public key (.cer) for upload to Entra
                $cerPath = [System.IO.Path]::GetTempFileName() + ".cer"
                Export-Certificate -Cert $cert -FilePath $cerPath -Type CERT | Out-Null

                # Export private key (.pfx) without password for Key Vault storage
                $pfxPath = [System.IO.Path]::GetTempFileName() + ".pfx"
                $emptyPwd = [System.Security.SecureString]::new()
                Export-PfxCertificate -Cert $cert -FilePath $pfxPath -Password $emptyPwd | Out-Null
                $certPfxBase64 = [Convert]::ToBase64String([System.IO.File]::ReadAllBytes($pfxPath))

                # Upload public key to the app registration
                Write-Host "  Uploading certificate to app registration..." -ForegroundColor Gray
                $uploadRaw = cmd /c "az ad app credential reset --id $ClientId --append --display-name M365Dashboard-Cert --cert @`"$cerPath`" --create-cert 2>&1"
                # Note: --create-cert is ignored when --cert is provided; this just uploads the .cer
                # Try the correct form: upload existing cert
                $uploadRaw = cmd /c "az rest --method POST --uri `"https://graph.microsoft.com/v1.0/applications/$appObjectIdNew/addKey`" 2>&1"
                # Simpler: use az ad app credential reset with the cert file
                $cerB64 = [Convert]::ToBase64String([System.IO.File]::ReadAllBytes($cerPath))
                $keyBody = "{`"keyCredentials`":[{`"type`":`"AsymmetricX509Cert`",`"usage`":`"Verify`",`"key`":`"$cerB64`",`"displayName`":`"M365Dashboard-Cert`"}]}"
                $keyFile = [System.IO.Path]::GetTempFileName() + ".json"
                [System.IO.File]::WriteAllText($keyFile, $keyBody, [System.Text.Encoding]::UTF8)
                $uploadResult = cmd /c "az rest --method PATCH --uri `"https://graph.microsoft.com/v1.0/applications/$appObjectIdNew`" --body @`"$keyFile`" --headers Content-Type=application/json 2>&1"
                Remove-Item $keyFile -ErrorAction SilentlyContinue
                Remove-Item $cerPath -ErrorAction SilentlyContinue
                Remove-Item $pfxPath -ErrorAction SilentlyContinue
                # Remove cert from local store (no longer needed here - Key Vault is the store)
                Remove-Item "Cert:\CurrentUser\My\$certThumbprint" -ErrorAction SilentlyContinue

                $uploadResultClean = ($uploadResult | Where-Object { $_ -notmatch '^WARNING:' }) -join ''
                if ($LASTEXITCODE -eq 0 -or [string]::IsNullOrWhiteSpace($uploadResultClean)) {
                    Write-Host "  Certificate uploaded to app registration." -ForegroundColor Green
                } else {
                    Write-Host "  Warning uploading certificate: $uploadResultClean" -ForegroundColor Yellow
                    Write-Host "  You may need to upload the certificate manually in Entra Portal." -ForegroundColor Yellow
                }

                # ClientSecret is empty - cert auth doesn't use a secret
                $ClientSecret = ""
            } else {
                $newSecret = $newSecretJson | ConvertFrom-Json
                $ClientSecret = $newSecret.password
                if ([string]::IsNullOrWhiteSpace($ClientSecret)) {
                    Write-Host "  Failed to extract client secret" -ForegroundColor Red
                    exit 1
                }
                Write-Host "  Client secret created (valid 2 years)" -ForegroundColor Green
            }

            # Upload logo to app registration
            Write-Host "  Uploading app logo..." -ForegroundColor Gray
            try {
                # Logo embedded as base64 PNG (215x215, generated from m365-icon.svg)
                $logoB64 = "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"
                $logoBytes = [Convert]::FromBase64String($logoB64)
                $tempLogo  = [System.IO.Path]::GetTempFileName() + ".png"
                [System.IO.File]::WriteAllBytes($tempLogo, $logoBytes)
                $logoResult = cmd /c "az rest --method PUT --uri `"https://graph.microsoft.com/v1.0/applications/$appObjectIdNew/logo`" --body `"@$tempLogo`" --headers Content-Type=image/png 2>&1"
                Remove-Item $tempLogo -ErrorAction SilentlyContinue
                if ($LASTEXITCODE -eq 0) {
                    Write-Host "  App logo uploaded" -ForegroundColor Green
                } else {
                    Write-Host "  Could not upload logo (non-critical): $logoResult" -ForegroundColor Yellow
                }
            } catch {
                Write-Host "  Could not upload logo (non-critical): $_" -ForegroundColor Yellow
            }

            # Set Application ID URI and expose access_as_user scope
            # The identifier URI must be set first
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

            # Wait for SP to propagate before attempting consent — without this,
            # admin-consent silently exits 0 but grants nothing because the SP
            # isn't visible yet across all Entra directory nodes.
            Write-Host "  Waiting for service principal to propagate..." -ForegroundColor Gray
            $spObjIdForConsent = $null
            for ($attempt = 1; $attempt -le 12; $attempt++) {
                Start-Sleep -Seconds 5
                $spCheckRaw = cmd /c "az ad sp show --id $ClientId --query id -o tsv 2>nul"
                $spObjIdForConsent = ($spCheckRaw | Where-Object { $_ -notmatch '^WARNING:' }) -join '' | ForEach-Object { $_.Trim() }
                if ($spObjIdForConsent) {
                    Write-Host "  Service principal confirmed (${attempt}x5s)" -ForegroundColor Gray
                    break
                }
                Write-Host "  Still waiting... ($attempt/12)" -ForegroundColor Gray
            }

            if (-not $spObjIdForConsent) {
                Write-Host "  WARNING: Service principal did not appear within 60s — consent may fail" -ForegroundColor Yellow
            }

            Write-Host "  Granting admin consent..." -ForegroundColor Gray
            # Run without stderr suppression so failures are visible in the log
            $consentRaw = cmd /c "az ad app permission admin-consent --id $ClientId 2>&1"
            $consentErr  = ($consentRaw | Where-Object { $_ -match 'ERROR|error|Insufficient|forbidden' }) -join ''
            if ($LASTEXITCODE -eq 0 -and -not $consentErr) {
                Write-Host "  Admin consent granted" -ForegroundColor Green
            } else {
                Write-Host "  az admin-consent failed: $consentErr" -ForegroundColor Yellow
                Write-Host "  Falling back to Graph API appRoleAssignments..." -ForegroundColor Gray

                # Correct fallback: grant each application permission via appRoleAssignments.
                # oauth2PermissionGrants only covers delegated permissions — it does nothing
                # for Role-type (application) permissions like User.Read.All.
                $ErrorActionPreference = "Continue"
                $graphToken = (cmd /c "az account get-access-token --resource https://graph.microsoft.com --query accessToken -o tsv 2>nul").Trim()

                if ($graphToken -and $spObjIdForConsent) {
                    # Resolve the Graph SP object ID
                    $graphSpRaw = cmd /c "az ad sp show --id 00000003-0000-0000-c000-000000000000 --query id -o tsv 2>nul"
                    $graphSpObjId = ($graphSpRaw | Where-Object { $_ -notmatch '^WARNING:' }) -join '' | ForEach-Object { $_.Trim() }

                    if ($graphSpObjId) {
                        $consentOk = $true
                        foreach ($perm in $graphPermissions) {
                            $roleBody = "{`"principalId`":`"$spObjIdForConsent`",`"resourceId`":`"$graphSpObjId`",`"appRoleId`":`"$($perm.id)`"}"
                            $roleFile = [System.IO.Path]::GetTempFileName() + ".json"
                            [System.IO.File]::WriteAllText($roleFile, $roleBody, [System.Text.Encoding]::UTF8)
                            $roleResult = cmd /c "az rest --method POST --uri `"https://graph.microsoft.com/v1.0/servicePrincipals/$spObjIdForConsent/appRoleAssignments`" --body @`"$roleFile`" --headers Content-Type=application/json 2>&1"
                            Remove-Item $roleFile -ErrorAction SilentlyContinue
                            $roleResultClean = ($roleResult | Where-Object { $_ -notmatch '^WARNING:' }) -join ''
                            if ($LASTEXITCODE -ne 0 -and $roleResultClean -notmatch 'already exists|Permission being assigned already exists') {
                                Write-Host "    ! $($perm.name): $roleResultClean" -ForegroundColor Yellow
                                $consentOk = $false
                            } else {
                                Write-Host "    + $($perm.name)" -ForegroundColor Gray
                            }
                        }

                        # Also grant Exchange.ManageAsApp via appRoleAssignments
                        $exSpRaw = cmd /c "az ad sp show --id 00000002-0000-0ff1-ce00-000000000000 --query id -o tsv 2>nul"
                        $exSpObjId = ($exSpRaw | Where-Object { $_ -notmatch '^WARNING:' }) -join '' | ForEach-Object { $_.Trim() }
                        if ($exSpObjId) {
                            $exRoleBody = "{`"principalId`":`"$spObjIdForConsent`",`"resourceId`":`"$exSpObjId`",`"appRoleId`":`"dc50a0fb-09a3-484d-be87-e023b12c6440`"}"
                            $exRoleFile = [System.IO.Path]::GetTempFileName() + ".json"
                            [System.IO.File]::WriteAllText($exRoleFile, $exRoleBody, [System.Text.Encoding]::UTF8)
                            $exResult = cmd /c "az rest --method POST --uri `"https://graph.microsoft.com/v1.0/servicePrincipals/$spObjIdForConsent/appRoleAssignments`" --body @`"$exRoleFile`" --headers Content-Type=application/json 2>&1"
                            Remove-Item $exRoleFile -ErrorAction SilentlyContinue
                        }

                        if ($consentOk) {
                            Write-Host "  Admin consent granted via Graph API" -ForegroundColor Green
                        } else {
                            Write-Host "  Some permissions could not be granted automatically." -ForegroundColor Yellow
                            Write-Host "  Grant manually: https://entra.microsoft.com/#view/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/~/CallAnAPI/appId/$ClientId" -ForegroundColor Cyan
                        }
                    } else {
                        Write-Host "  Could not resolve Graph service principal — grant consent manually:" -ForegroundColor Yellow
                        Write-Host "  https://entra.microsoft.com/#view/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/~/CallAnAPI/appId/$ClientId" -ForegroundColor Cyan
                    }
                } else {
                    Write-Host "  No Graph token or SP ID available — grant consent manually:" -ForegroundColor Yellow
                    Write-Host "  https://entra.microsoft.com/#view/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/~/CallAnAPI/appId/$ClientId" -ForegroundColor Cyan
                }
                $ErrorActionPreference = "Stop"
            }

            $ErrorActionPreference = "Stop"

            # Get tenant ID — in MSP NonInteractive mode use the passed-in client tenant ID
            # (az account show at this point returns the MSP tenant, not the client tenant)
            if ($NonInteractive -and $isMspMode -and $clientTenantId) {
                $TenantId = $clientTenantId
            } else {
                $TenantId = (cmd /c "az account show --query tenantId -o tsv 2>nul").Trim()
            }

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

            # Save new config (include cert details if using certificate auth)
            $newConfig = @{
                TenantId         = $TenantId
                ClientId         = $ClientId
                ClientSecret     = $ClientSecret
                UseCertAuth      = $useCertAuth
                CertThumbprint   = $certThumbprint
                AppObjectId      = $appObjectIdNew
                AppName          = $appNameInput
                CreatedAt        = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
            }
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

# Prompt for resource name prefix — skip in non-interactive mode (passed as parameter)
if ($NonInteractive -and -not $NamePrefix) {
    Write-Host "ERROR: -NamePrefix is required in non-interactive mode" -ForegroundColor Red; exit 1
}
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

# Prompt for Azure region — skip in non-interactive mode (passed as parameter)
if ($NonInteractive -and -not $Location) {
    Write-Host "ERROR: -Location is required in non-interactive mode" -ForegroundColor Red; exit 1
}
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

# In MSP mode, subscription selection happens after the Step 2 login (below).
# In standard mode, select subscription now.
$selectedSubscriptionName = ""

# If SubscriptionId was passed (from wizard), set it immediately
if ($SubscriptionId -and -not $isMspMode) {
    Write-Host "Setting subscription: $SubscriptionId" -ForegroundColor Gray
    cmd /c "az account set --subscription $SubscriptionId 2>nul"
    $ErrorActionPreference = "Continue"
    $currentAccountJson = (cmd /c "az account show -o json 2>nul" | Where-Object { $_ -notmatch '^WARNING:' }) -join "`n"
    $ErrorActionPreference = "Stop"
}

if (-not $isMspMode) {
    Write-Host ""
    Write-Host "Azure Subscription" -ForegroundColor Cyan
    Write-Host "------------------" -ForegroundColor Cyan

    $subscriptionsJson = cmd /c "az account list --query [?state=='Enabled'] -o json 2>nul"
    $subscriptions = $subscriptionsJson | ConvertFrom-Json

    if ($subscriptions.Count -gt 1 -and -not $NonInteractive) {
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
            $currentAccount = $currentAccountJson | ConvertFrom-Json
            $selectedSubscriptionName = $currentAccount.name
            Write-Host "  Invalid choice. Using current: $selectedSubscriptionName" -ForegroundColor Yellow
        }
    } else {
        # Non-interactive or single subscription: use default/current
        $defaultSub = $subscriptions | Where-Object { $_.isDefault } | Select-Object -First 1
        if (-not $defaultSub) { $defaultSub = $subscriptions[0] }
        $selectedSubscriptionName = $defaultSub.name
        Write-Host "Using subscription: $selectedSubscriptionName" -ForegroundColor Green
    }
}

# Prompt for SQL password — skip in non-interactive mode
if ($NonInteractive -and -not $SqlPassword) {
    $SqlPassword = $env:WIZARD_SQL_PASSWORD
    if (-not $SqlPassword) { Write-Host "ERROR: SQL password required in non-interactive mode" -ForegroundColor Red; exit 1 }
}
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

if (-not $NonInteractive) {
    $confirm = Read-Host "Proceed with deployment? (Y/n)"
    if ($confirm -eq "n" -or $confirm -eq "N") {
        Write-Host "Deployment cancelled." -ForegroundColor Yellow
        exit 0
    }
}
Write-Host ""

# In MSP mode, capture a Graph token for the client tenant NOW before we log out of it.
# az logout in Invoke-AzLogin will clear all cached credentials, so we must get this token first.
$mspGraphToken = $null
if ($isMspMode) {
    Write-Host "Capturing Graph token for client tenant (needed for post-deployment config)..." -ForegroundColor Yellow
    $ErrorActionPreference = "Continue"
    $tokenRaw = cmd /c "az account get-access-token --resource https://graph.microsoft.com -o json 2>nul"
    $tokenJson = ($tokenRaw | Where-Object { $_ -notmatch '^WARNING:' }) -join ''
    if ($tokenJson -match '"accessToken"') {
        $mspGraphToken = ($tokenJson | ConvertFrom-Json).accessToken
        Write-Host "  Graph token obtained for client tenant" -ForegroundColor Green
    } else {
        Write-Host "  Warning: Could not obtain Graph token - redirect URI will need to be set manually" -ForegroundColor Yellow
    }
    $ErrorActionPreference = "Stop"
}

# In MSP mode, switch from client tenant to your own Azure subscription for infrastructure deployment
if ($isMspMode) {
    Write-Host ""
    Write-Host "Step 2 of 2: Login to YOUR Azure subscription for infrastructure deployment" -ForegroundColor Yellow
    Write-Host "  (The app registration in the client tenant is complete)" -ForegroundColor Gray
    Write-Host "  Now logging in to your MSP Azure subscription for resource deployment..." -ForegroundColor Gray
    Write-Host ""
    if (-not $NonInteractive) { Read-Host "  Press Enter when ready to log in to YOUR Azure subscription" }
    if ($NonInteractive) {
        # Wizard already logged in to MSP subscription — read the current session
        $ErrorActionPreference = "Continue"
        $yourAccountJson = (cmd /c "az account show -o json 2>nul" | Where-Object { $_ -notmatch '^WARNING:' }) -join "`n"
        $ErrorActionPreference = "Stop"
        Write-Host "  Using existing login session" -ForegroundColor Gray
    } else {
        $yourAccountJson = Invoke-AzLogin
    }
    $yourAccount = $yourAccountJson | ConvertFrom-Json
    Write-Host "  Logged in as: $($yourAccount.user.name) (tenant: $($yourAccount.tenantId))" -ForegroundColor Green
    Write-Host ""

    # Select subscription from the MSP's Azure account
    Write-Host "Azure Subscription" -ForegroundColor Cyan
    Write-Host "------------------" -ForegroundColor Cyan
    $subscriptionsJson = cmd /c "az account list --query [?state=='Enabled'] -o json 2>nul"
    $subscriptions = $subscriptionsJson | ConvertFrom-Json

    # In NonInteractive mode, honour the SubscriptionId passed from the wizard
    if ($NonInteractive -and $SubscriptionId) {
        cmd /c "az account set --subscription $SubscriptionId 2>nul" | Out-Null
        $selectedSub = $subscriptions | Where-Object { $_.id -eq $SubscriptionId } | Select-Object -First 1
        $selectedSubscriptionName = if ($selectedSub) { $selectedSub.name } else { $SubscriptionId }
        Write-Host "Using subscription: $selectedSubscriptionName" -ForegroundColor Green
    } elseif ($subscriptions.Count -gt 1 -and -not $NonInteractive) {
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
            $selectedSubscriptionName = $yourAccount.name
            Write-Host "  Invalid choice. Using current: $selectedSubscriptionName" -ForegroundColor Yellow
        }
    } else {
        $defaultSub = $subscriptions | Where-Object { $_.isDefault } | Select-Object -First 1
        if (-not $defaultSub) { $defaultSub = $subscriptions[0] }
        $selectedSubscriptionName = $defaultSub.name
        Write-Host "Using subscription: $selectedSubscriptionName" -ForegroundColor Green
    }
    Write-Host ""
}

# Final validation before deploying
if ([string]::IsNullOrWhiteSpace($TenantId) -or [string]::IsNullOrWhiteSpace($ClientId)) {
    Write-Host "ERROR: Missing required Entra ID credentials before deployment:" -ForegroundColor Red
    Write-Host "  TenantId: $(if ($TenantId) { $TenantId } else { '(empty)' })" -ForegroundColor Red
    Write-Host "  ClientId: $(if ($ClientId) { $ClientId } else { '(empty)' })" -ForegroundColor Red
    exit 1
}
if (-not $useCertAuth -and [string]::IsNullOrWhiteSpace($ClientSecret)) {
    Write-Host "ERROR: No client secret and no certificate configured." -ForegroundColor Red
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

# Write sensitive parameters to a JSON file to avoid shell quoting/escaping issues
$paramsFile = [System.IO.Path]::GetTempFileName() + ".json"
$paramsObj = @{
    namePrefix             = @{ value = $NamePrefix }
    location               = @{ value = $Location }
    environment            = @{ value = $Environment }
    entraIdTenantId        = @{ value = $TenantId }
    entraIdClientId        = @{ value = $ClientId }
    entraIdClientSecret    = @{ value = $ClientSecret }
    sqlAdminPassword       = @{ value = $SqlPassword }
    deployingUserObjectId  = @{ value = $deployingUserObjectId }
}
[System.IO.File]::WriteAllText($paramsFile, ($paramsObj | ConvertTo-Json -Depth 5), [System.Text.Encoding]::UTF8)

$ErrorActionPreference = "Continue"
$deploymentResult = az deployment sub create `
    --name $deploymentName `
    --location $Location `
    --template-file $infraPath `
    --parameters "@$paramsFile" `
    --query properties.outputs -o json 2>&1

Remove-Item $paramsFile -ErrorAction SilentlyContinue
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

# If using certificate auth, store the PFX in Key Vault now that it exists
if ($useCertAuth -and $certPfxBase64) {
    Write-Host "Storing certificate in Key Vault..." -ForegroundColor Yellow
    $kvName = $deploymentOutput.keyVaultName.value

    if ([string]::IsNullOrWhiteSpace($kvName)) {
        Write-Host "  ERROR: Could not determine Key Vault name from deployment output." -ForegroundColor Red
        Write-Host "  Certificate secrets were NOT stored. The app will fail to authenticate." -ForegroundColor Red
        Write-Host "  Deployment output keys: $($deploymentOutput.PSObject.Properties.Name -join ', ')" -ForegroundColor Yellow
        exit 1
    }

    Write-Host "  Key Vault: $kvName" -ForegroundColor Gray
    $ErrorActionPreference = "Continue"

    # Store PFX as base64 secret - the app reads it as AzureAd:ClientCertificatePfx
    $pfxResult = cmd /c "az keyvault secret set --vault-name $kvName --name AzureAd--ClientCertificatePfx --value `"$certPfxBase64`" 2>&1"
    if ($LASTEXITCODE -ne 0) {
        Write-Host "  ERROR storing PFX: $pfxResult" -ForegroundColor Red
        exit 1
    }
    Write-Host "  Certificate PFX stored in Key Vault" -ForegroundColor Green

    $thumbResult = cmd /c "az keyvault secret set --vault-name $kvName --name AzureAd--ClientCertificateThumbprint --value `"$certThumbprint`" 2>&1"
    if ($LASTEXITCODE -ne 0) {
        Write-Host "  ERROR storing thumbprint: $thumbResult" -ForegroundColor Red
        exit 1
    }
    Write-Host "  Thumbprint stored in Key Vault" -ForegroundColor Green

    # Set ClientSecret to a single space so IsNullOrWhiteSpace correctly skips it
    cmd /c "az keyvault secret set --vault-name $kvName --name AzureAd--ClientSecret --value `" `" 2>nul" | Out-Null

    $ErrorActionPreference = "Stop"
    Write-Host "  Auth mode: certificate (no client secret required)" -ForegroundColor Green
    Write-Host ""
}

# Build and push Docker image using ACR Build (no local Docker required)
Write-Host "Building Docker image in Azure..." -ForegroundColor Yellow
Write-Host "  This may take 5-10 minutes..." -ForegroundColor Gray

# Must run from repo root so az acr build can find the Dockerfile
$repoRootPath = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$ErrorActionPreference = "Continue"
Push-Location $repoRootPath
try {
    $buildVersion = "deploy-$(Get-Date -Format 'yyyy.MM.dd')-$(git -C $repoRootPath rev-parse --short HEAD 2>$null)"
    cmd /c "az acr build --registry $acrName --image m365dashboard:latest --build-arg VITE_AZURE_CLIENT_ID=$ClientId --build-arg VITE_AZURE_TENANT_ID=$TenantId --build-arg BUILD_VERSION=$buildVersion . 2>&1"
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

$appUrlClean = $appUrl.TrimEnd('/')
Write-Host "  Setting redirect URI: $appUrlClean" -ForegroundColor Gray

# In MSP mode use the Graph token captured before Step 2 login (before az logout cleared credentials).
$ErrorActionPreference = "Continue"
if ($isMspMode) {
    $graphToken = $mspGraphToken

    if ($graphToken) {
        # Use the app object ID we saved during registration
        $appObjectId = if ($appObjectIdNew) { $appObjectIdNew } else { $null }

        if ($appObjectId) {
            # Set redirect URI
            $redirectBody = "{`"spa`":{`"redirectUris`":[`"$appUrlClean`"]}}"
            $redirectFile = [System.IO.Path]::GetTempFileName() + ".json"
            [System.IO.File]::WriteAllText($redirectFile, $redirectBody, [System.Text.Encoding]::UTF8)
            $redirectResult = cmd /c "az rest --method PATCH --uri `"https://graph.microsoft.com/v1.0/applications/$appObjectId`" --body @`"$redirectFile`" --headers Content-Type=application/json Authorization=`"Bearer $graphToken`" 2>&1"
            Remove-Item $redirectFile -ErrorAction SilentlyContinue
            if ($LASTEXITCODE -eq 0) {
                Write-Host "  Redirect URI configured" -ForegroundColor Green
            } else {
                Write-Host "  Warning: Could not set redirect URI: $redirectResult" -ForegroundColor Yellow
            }

            # Enable access tokens and ID tokens
            $tokenBody = '{"web":{"implicitGrantSettings":{"enableAccessTokenIssuance":true,"enableIdTokenIssuance":true}},"spa":{"redirectUris":["' + $appUrlClean + '"]}}'
            $tokenFile = [System.IO.Path]::GetTempFileName() + ".json"
            [System.IO.File]::WriteAllText($tokenFile, $tokenBody, [System.Text.Encoding]::UTF8)
            cmd /c "az rest --method PATCH --uri `"https://graph.microsoft.com/v1.0/applications/$appObjectId`" --body @`"$tokenFile`" --headers Content-Type=application/json Authorization=`"Bearer $graphToken`" 2>nul" | Out-Null
            Remove-Item $tokenFile -ErrorAction SilentlyContinue
            Write-Host "  Tokens enabled" -ForegroundColor Green
        } else {
            Write-Host "  Warning: app object ID not available - redirect URI must be set manually." -ForegroundColor Yellow
            Write-Host "  https://entra.microsoft.com/#view/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/~/Authentication/appId/$ClientId" -ForegroundColor Cyan
        }

        # Grant admin consent via Graph token
        Write-Host "  Granting admin consent..." -ForegroundColor Gray
        $spIdRaw = cmd /c "az rest --method GET --uri `"https://graph.microsoft.com/v1.0/servicePrincipals?`$filter=appId eq '$ClientId'`" --headers Authorization=`"Bearer $graphToken`" --query value[0].id -o tsv 2>nul"
        $spIdForConsent = ($spIdRaw | Where-Object { $_ -notmatch '^WARNING:' }) -join '' | ForEach-Object { $_.Trim() }
        if ($spIdForConsent) {
            $consentBody = "{`"clientId`":`"$spIdForConsent`",`"consentType`":`"AllPrincipals`",`"resourceId`":`"$spIdForConsent`",`"scope`":`"openid profile`"}"
            $consentFile = [System.IO.Path]::GetTempFileName() + ".json"
            [System.IO.File]::WriteAllText($consentFile, $consentBody, [System.Text.Encoding]::UTF8)
            $consentResult = cmd /c "az rest --method POST --uri `"https://graph.microsoft.com/v1.0/oauth2PermissionGrants`" --body @`"$consentFile`" --headers Content-Type=application/json Authorization=`"Bearer $graphToken`" 2>&1"
            Remove-Item $consentFile -ErrorAction SilentlyContinue
            if ($LASTEXITCODE -eq 0) {
                Write-Host "  Admin consent granted" -ForegroundColor Green
            } else {
                Write-Host "  Could not grant admin consent automatically - grant manually:" -ForegroundColor Yellow
                Write-Host "  https://entra.microsoft.com/#view/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/~/CallAnAPI/appId/$ClientId" -ForegroundColor Cyan
            }
        }
    } else {
        Write-Host "  Could not obtain Graph token for client tenant." -ForegroundColor Yellow
        Write-Host "  Set redirect URI manually: https://entra.microsoft.com/#view/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/~/Authentication/appId/$ClientId" -ForegroundColor Cyan
    }
} else {
    # Standard mode - CLI is in the correct tenant
    $appObjectIdRaw = cmd /c "az ad app show --id $ClientId --query id -o tsv 2>nul"
    $appObjectId = ($appObjectIdRaw | Where-Object { $_ -notmatch '^WARNING:' }) -join '' | ForEach-Object { $_.Trim() }

    if ($appObjectId) {
        $redirectBodyFile = [System.IO.Path]::GetTempFileName()
        $redirectBody = "{`"spa`":{`"redirectUris`":[`"$appUrlClean`"]}}"
        [System.IO.File]::WriteAllText($redirectBodyFile, $redirectBody, [System.Text.Encoding]::UTF8)
        $uriUpdateResult = cmd /c "az rest --method PATCH --uri `"https://graph.microsoft.com/v1.0/applications/$appObjectId`" --body @`"$redirectBodyFile`" --headers Content-Type=application/json 2>&1"
        Remove-Item $redirectBodyFile -ErrorAction SilentlyContinue
        if ($LASTEXITCODE -eq 0) {
            Write-Host "  Redirect URI configured" -ForegroundColor Green
        } else {
            Write-Host "  Warning: Could not set redirect URI: $uriUpdateResult" -ForegroundColor Yellow
        }
    }

    Write-Host "  Enabling access tokens and ID tokens..." -ForegroundColor Gray
    cmd /c "az ad app update --id $ClientId --enable-access-token-issuance true --enable-id-token-issuance true 2>nul"
    Write-Host "  Tokens enabled" -ForegroundColor Green

    Write-Host "  Granting admin consent for API permissions..." -ForegroundColor Gray
    $consentResult = cmd /c "az ad app permission admin-consent --id $ClientId 2>&1"
    if ($LASTEXITCODE -eq 0) {
        Write-Host "  Admin consent granted" -ForegroundColor Green
    } else {
        Write-Host "  Could not grant admin consent automatically - grant manually:" -ForegroundColor Yellow
        Write-Host "  Azure Portal > Entra ID > App registrations > $ClientId > API permissions > Grant admin consent" -ForegroundColor Yellow
    }
}
$ErrorActionPreference = "Stop"

# Dashboard Admin role assignment - in MSP mode this assigns to the client tenant admin
# who performed the app registration (we can't assign to the MSP user as they're in a different tenant)
Write-Host "  Assigning Dashboard Admin role..." -ForegroundColor Gray
$ErrorActionPreference = "Continue"
$spId = (cmd /c "az ad sp show --id $ClientId --query id -o tsv $entraTarget 2>nul" | Where-Object { $_ -notmatch '^WARNING:' }) -join '' | ForEach-Object { $_.Trim() }

if ($spId) {
    $appRoles = cmd /c "az ad app show --id $ClientId --query appRoles -o json $entraTarget 2>nul" | ConvertFrom-Json
    $adminRole = $appRoles | Where-Object { $_.value -eq "Dashboard.Admin" }

    if ($adminRole) {
        $roleId = $adminRole.id
        if ($isMspMode) {
            # In MSP mode the CLI is in the MSP tenant - we can't get the client tenant's signed-in user.
            # The client admin will need to assign themselves the Dashboard.Admin role via Entra Enterprise Apps.
            Write-Host "  MSP mode: Dashboard Admin role must be assigned manually in the client tenant." -ForegroundColor Yellow
            Write-Host "  https://entra.microsoft.com/#view/Microsoft_AAD_IAM/ManagedAppMenuBlade/~/Users/objectId/$spId" -ForegroundColor Cyan
        } else {
            $currentUser = (cmd /c "az ad signed-in-user show --query id -o tsv 2>nul").Trim()
            if ($currentUser) {
                $roleBodyFile = [System.IO.Path]::GetTempFileName()
                $roleBody = "{`"principalId`":`"$currentUser`",`"resourceId`":`"$spId`",`"appRoleId`":`"$roleId`"}"
                [System.IO.File]::WriteAllText($roleBodyFile, $roleBody, [System.Text.Encoding]::UTF8)
                $assignResult = cmd /c "az rest --method POST --uri https://graph.microsoft.com/v1.0/users/$currentUser/appRoleAssignments --body @`"$roleBodyFile`" --headers Content-Type=application/json 2>&1"
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
    }
}
$ErrorActionPreference = "Stop"

# $containerAppName is defined further down in the GitHub Actions section.
# Save deploy config after it is set - see below.

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

# Save deploy config now that $containerAppName is defined
Write-Host ""
Write-Host "Saving deploy config..." -ForegroundColor Yellow
$deployConfig = @{
    ResourceGroup    = $resourceGroup
    ContainerAppName = $containerAppName
    AcrName          = $acrName
    AppUrl           = $appUrl
    DeployedAt       = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
}
$deployConfigPath = Join-Path (Split-Path $PSScriptRoot -Parent) "deploy-config.json"
$deployConfig | ConvertTo-Json | Out-File $deployConfigPath -Encoding UTF8
Write-Host "  Deploy config saved to deploy-config.json" -ForegroundColor Green

# Store Container App details in Key Vault so the app can self-update
Write-Host ""
Write-Host "Storing Container App config in Key Vault for self-update..." -ForegroundColor Yellow
$kvName = $deploymentOutput.keyVaultName.value
Write-Host "  Key Vault: $kvName" -ForegroundColor Gray
$ErrorActionPreference = "Continue"
cmd /c "az keyvault secret set --vault-name $kvName --name ContainerApp--SubscriptionId --value `"$subscriptionId`" 2>nul" | Out-Null
cmd /c "az keyvault secret set --vault-name $kvName --name ContainerApp--ResourceGroup --value `"$resourceGroup`" 2>nul" | Out-Null
cmd /c "az keyvault secret set --vault-name $kvName --name ContainerApp--Name --value `"$containerAppName`" 2>nul" | Out-Null
$ErrorActionPreference = "Stop"
Write-Host "  Container App config stored in Key Vault" -ForegroundColor Green

# Grant the Container App's managed identity Contributor on itself so it can self-update
Write-Host "  Granting Container App managed identity Contributor role for self-update..." -ForegroundColor Gray
$ErrorActionPreference = "Continue"
$containerAppResourceId = cmd /c "az containerapp show --name $containerAppName --resource-group $resourceGroup --query id -o tsv 2>nul"
$containerAppResourceId = ($containerAppResourceId | Where-Object { $_ -notmatch '^WARNING:' }) -join '' | ForEach-Object { $_.Trim() }
$managedIdentityPrincipalId = cmd /c "az containerapp show --name $containerAppName --resource-group $resourceGroup --query identity.principalId -o tsv 2>nul"
$managedIdentityPrincipalId = ($managedIdentityPrincipalId | Where-Object { $_ -notmatch '^WARNING:' }) -join '' | ForEach-Object { $_.Trim() }
Write-Host "  Container App resource ID: $containerAppResourceId" -ForegroundColor Gray
Write-Host "  Managed Identity principal ID: $managedIdentityPrincipalId" -ForegroundColor Gray
if ($containerAppResourceId -and $managedIdentityPrincipalId) {
    cmd /c "az role assignment create --assignee $managedIdentityPrincipalId --role Contributor --scope `"$containerAppResourceId`" 2>nul" | Out-Null
    if ($LASTEXITCODE -eq 0) {
        Write-Host "  Self-update permission granted" -ForegroundColor Green
    } else {
        Write-Host "  Could not grant self-update permission automatically" -ForegroundColor Yellow
    }
} else {
    Write-Host "  Could not resolve Container App ID or Managed Identity - self-update not configured" -ForegroundColor Yellow
}
$ErrorActionPreference = "Stop"

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
Write-Host "  DASHBOARD_URL: $appUrl" -ForegroundColor Cyan
Write-Host ""
Write-Host "============================================" -ForegroundColor Yellow
Write-Host "Post-Deployment Checks" -ForegroundColor Yellow
Write-Host "============================================" -ForegroundColor Yellow
Write-Host ""

# ── Check 1: Admin consent ──────────────────────────────────────────────────
Write-Host "Checking admin consent status..." -ForegroundColor Gray
$consentGranted = $false
try {
    $ErrorActionPreference = "Continue"
    # Get the service principal object ID for our app (use client tenant in MSP mode)
    $spObjRaw = cmd /c "az ad sp show --id $ClientId --query id -o tsv $entraTarget 2>nul"
    $spObjId  = ($spObjRaw | Where-Object { $_ -notmatch '^WARNING:' }) -join '' | ForEach-Object { $_.Trim() }

    if ($spObjId) {
        # Query oauth2PermissionGrants - if any AllPrincipals grant exists, consent was given
        $grantsRaw = cmd /c "az rest --method GET --uri `"https://graph.microsoft.com/v1.0/oauth2PermissionGrants?`$filter=clientId eq '$spObjId'`" --query value -o json $entraTarget 2>nul"
        $grantsJson = ($grantsRaw | Where-Object { $_ -notmatch '^WARNING:' }) -join ''
        if ($grantsJson -and $grantsJson -ne '[]' -and $grantsJson -ne 'null') {
            $grants = $grantsJson | ConvertFrom-Json
            if ($grants -and $grants.Count -gt 0) {
                $consentGranted = $true
            }
        }

        # Also check appRoleAssignments - application permissions granted show up here
        if (-not $consentGranted) {
            $assignmentsRaw = cmd /c "az rest --method GET --uri `"https://graph.microsoft.com/v1.0/servicePrincipals/$spObjId/appRoleAssignments`" --query value -o json $entraTarget 2>nul"
            $assignmentsJson = ($assignmentsRaw | Where-Object { $_ -notmatch '^WARNING:' }) -join ''
            if ($assignmentsJson -and $assignmentsJson -ne '[]' -and $assignmentsJson -ne 'null') {
                $assignments = $assignmentsJson | ConvertFrom-Json
                if ($assignments -and $assignments.Count -gt 0) {
                    $consentGranted = $true
                }
            }
        }
    }
    $ErrorActionPreference = "Stop"
} catch {
    # Consent check failed non-fatally - default to showing the manual step
    $consentGranted = $false
}

Write-Host ""
Write-Host "============================================" -ForegroundColor Yellow
Write-Host "Manual Steps Required" -ForegroundColor Yellow
Write-Host "============================================" -ForegroundColor Yellow
Write-Host ""
Write-Host "The following steps require manual configuration:" -ForegroundColor White
Write-Host ""

if ($consentGranted) {
    Write-Host "1. Grant Admin Consent" -NoNewline -ForegroundColor Cyan
    Write-Host "  [ALREADY GRANTED]" -ForegroundColor Green
    Write-Host "   All permissions have been consented for this app registration." -ForegroundColor Gray
} else {
    Write-Host "1. Grant Admin Consent (required)" -ForegroundColor Cyan
    Write-Host "   https://entra.microsoft.com/#view/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/~/CallAnAPI/appId/$ClientId" -ForegroundColor Gray
    Write-Host "   > Grant admin consent for [your tenant]" -ForegroundColor Gray
}
Write-Host ""
Write-Host "2. Security Reader role in Exchange (required for Defender for Office data)" -ForegroundColor Cyan
Write-Host "   https://admin.cloud.microsoft/exchange#/adminRoles" -ForegroundColor Gray
Write-Host "   > View-Only Organization Management > Members tab > Add" -ForegroundColor Gray
Write-Host "   > Search for app registration by name and add it" -ForegroundColor Gray
Write-Host ""
# ── Check 3: Defender for Endpoint licensing and consent ─────────────────────────
$defenderLicensed    = $false
$defenderConsentDone = $false
$defenderAppId       = "fc780465-2017-40d4-a0c5-307022471b92" # WindowsDefenderATP

try {
    $ErrorActionPreference = "Continue"

    # Defender P1/P2 SKU prefixes to look for in subscribedSkus
    $defenderSkus = @(
        "WIN_DEF_ATP",           # Microsoft Defender for Endpoint P1/P2
        "MDATP_Server",          # Defender for Servers
        "DEFENDER_ENDPOINT_P1",  # Standalone P1
        "MDATP_XPLAT"            # Cross-platform
    )

    Write-Host "Checking Defender for Endpoint licensing..." -ForegroundColor Gray
    $skusRaw  = cmd /c "az rest --method GET --uri `"https://graph.microsoft.com/v1.0/subscribedSkus`" --query value -o json 2>nul"
    $skusJson = ($skusRaw | Where-Object { $_ -notmatch '^WARNING:' }) -join ''

    if ($skusJson -and $skusJson -ne '[]' -and $skusJson -ne 'null') {
        $skus = $skusJson | ConvertFrom-Json
        foreach ($sku in $skus) {
            $partNumber = $sku.skuPartNumber
            if ($defenderSkus | Where-Object { $partNumber -like "*$_*" }) {
                if ($sku.prepaidUnits.enabled -gt 0) {
                    $defenderLicensed = $true
                    Write-Host "  Defender licence detected: $partNumber" -ForegroundColor Green
                    break
                }
            }
        }
    }

    if ($defenderLicensed) {
        Write-Host "  Granting Defender for Endpoint API consent..." -ForegroundColor Gray

        # Ensure the WindowsDefenderATP service principal exists in this tenant
        $defSpRaw = cmd /c "az ad sp show --id $defenderAppId --query id -o tsv 2>nul"
        $defSpId  = ($defSpRaw | Where-Object { $_ -notmatch '^WARNING:' }) -join '' | ForEach-Object { $_.Trim() }

        if (-not $defSpId) {
            # SP not yet in tenant - create it
            cmd /c "az ad sp create --id $defenderAppId 2>nul" | Out-Null
            Start-Sleep -Seconds 5
            $defSpRaw = cmd /c "az ad sp show --id $defenderAppId --query id -o tsv 2>nul"
            $defSpId  = ($defSpRaw | Where-Object { $_ -notmatch '^WARNING:' }) -join '' | ForEach-Object { $_.Trim() }
        }

        if ($spObjId -and $defSpId) {
            # Grant each Defender app role assignment
            $defenderRoles = @(
                @{ id = "ea8291d3-4b9a-44b5-bc3a-6cea3026dc79"; name = "Machine.Read.All" }
                @{ id = "41269fc5-d04d-4bfd-bce7-43a51cea049a"; name = "Vulnerability.Read.All" }
                @{ id = "02b005dd-f804-43b4-8fc7-078460413f74"; name = "Score.Read.All" }
            )

            $allRolesGranted = $true
            foreach ($role in $defenderRoles) {
                # Check if already assigned
                $existingRaw  = cmd /c "az rest --method GET --uri `"https://graph.microsoft.com/v1.0/servicePrincipals/$spObjId/appRoleAssignments`" --query `"value[?appRoleId=='$($role.id)']`" -o json 2>nul"
                $existingJson = ($existingRaw | Where-Object { $_ -notmatch '^WARNING:' }) -join ''
                if ($existingJson -and $existingJson -ne '[]' -and $existingJson -ne 'null') {
                    Write-Host "    + $($role.name) [already granted]" -ForegroundColor Gray
                    continue
                }

                $body    = "{`"principalId`":`"$spObjId`",`"resourceId`":`"$defSpId`",`"appRoleId`":`"$($role.id)`"}"
                $grantRaw = cmd /c "az rest --method POST --uri `"https://graph.microsoft.com/v1.0/servicePrincipals/$spObjId/appRoleAssignments`" --body `"$body`" --headers Content-Type=application/json 2>&1"
                if ($LASTEXITCODE -eq 0) {
                    Write-Host "    + $($role.name) [granted]" -ForegroundColor Green
                } else {
                    Write-Host "    ! $($role.name) [failed]" -ForegroundColor Yellow
                    $allRolesGranted = $false
                }
            }
            $defenderConsentDone = $allRolesGranted
        }
    }
    $ErrorActionPreference = "Stop"
} catch {
    Write-Host "  Defender check encountered an error: $($_.Exception.Message)" -ForegroundColor Yellow
}

Write-Host "3. Defender for Endpoint permissions" -NoNewline -ForegroundColor Cyan
if (-not $defenderLicensed) {
    Write-Host "  [NOT LICENSED - SKIPPED]" -ForegroundColor Gray
    Write-Host "   No Defender for Endpoint P1/P2 licence detected in this tenant." -ForegroundColor Gray
    Write-Host "   If you add a Defender licence later, grant consent manually:" -ForegroundColor Gray
    Write-Host "   https://entra.microsoft.com/#view/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/~/CallAnAPI/appId/$ClientId" -ForegroundColor Gray
} elseif ($defenderConsentDone) {
    Write-Host "  [CONSENT GRANTED]" -ForegroundColor Green
    Write-Host "   Machine.Read.All, Vulnerability.Read.All, Score.Read.All - all consented." -ForegroundColor Gray
} else {
    Write-Host "  [REQUIRES MANUAL CONSENT]" -ForegroundColor Yellow
    Write-Host "   Defender licence found but consent could not be fully automated." -ForegroundColor Yellow
    Write-Host "   https://entra.microsoft.com/#view/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/~/CallAnAPI/appId/$ClientId" -ForegroundColor Gray
    Write-Host "   > Grant admin consent for [your tenant]" -ForegroundColor Gray
}
Write-Host ""
Write-Host "Once complete, open the dashboard and sign in:" -ForegroundColor White
Write-Host "  DASHBOARD_URL: $appUrl" -ForegroundColor Cyan
Write-Host ""

# ============================================================================
# Configure GitHub Actions Secrets
# ============================================================================
Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "GitHub Actions CI/CD Setup" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Gathering values for GitHub Actions secrets..." -ForegroundColor Yellow

$acrUsername    = (cmd /c "az acr credential show --name $acrName --query username -o tsv 2>nul").Trim()
$acrPassword    = (cmd /c "az acr credential show --name $acrName --query passwords[0].value -o tsv 2>nul").Trim()
$subscriptionId = (cmd /c "az account show --query id -o tsv 2>nul").Trim()
$containerAppName = "$NamePrefix-$Environment-app"
$spNameGh       = "$NamePrefix-$Environment-github-actions"

# Create service principal for GitHub Actions
Write-Host "  Creating GitHub Actions service principal '$spNameGh'..." -ForegroundColor Gray
$ErrorActionPreference = "Continue"
$spJsonGh = cmd /c "az ad sp create-for-rbac --name `"$spNameGh`" --role contributor --scopes /subscriptions/$subscriptionId/resourceGroups/$resourceGroup --sdk-auth 2>&1"
if ($LASTEXITCODE -ne 0 -or ($spJsonGh -join "") -match '"error"') {
    Write-Host "  SP may already exist - resetting credentials..." -ForegroundColor Gray
    $spJsonGh = cmd /c "az ad sp credential reset --name `"$spNameGh`" --sdk-auth 2>&1"
}
$ErrorActionPreference = "Stop"
$spJsonGh = ($spJsonGh | Where-Object { $_ -notmatch '^WARNING:' }) -join "`n"

# Detect GitHub repo slug from git remote
$repoRoot  = Split-Path $PSScriptRoot -Parent
$gitRemote = (cmd /c "git -C `"$repoRoot`" remote get-url origin 2>nul").Trim()
$repoSlug  = ""
if ($gitRemote -match "github\.com[:/](.+?)(\.git)?$") {
    $repoSlug = $Matches[1].Trim()
}

# Check for GitHub CLI - install if missing
$ErrorActionPreference = "Continue"
$ghAvailable = cmd /c "gh --version 2>nul"
$ErrorActionPreference = "Stop"

if (-not $ghAvailable) {
    Write-Host "  GitHub CLI not found - attempting to install..." -ForegroundColor Yellow
    $ErrorActionPreference = "Continue"

    # Try winget first (built into Windows 10/11)
    $winget = cmd /c "winget --version 2>nul"
    if ($winget) {
        Write-Host "  Installing via winget..." -ForegroundColor Gray
        cmd /c "winget install --id GitHub.cli --silent --accept-package-agreements --accept-source-agreements 2>&1"
        if ($LASTEXITCODE -eq 0) {
            # Refresh PATH so gh is available in this session
            $env:PATH = [System.Environment]::GetEnvironmentVariable('PATH', 'Machine') + ';' + [System.Environment]::GetEnvironmentVariable('PATH', 'User')
            $ghAvailable = cmd /c "gh --version 2>nul"
            if ($ghAvailable) {
                Write-Host "  GitHub CLI installed successfully" -ForegroundColor Green
            }
        }
    }

    # Try Chocolatey if winget failed
    if (-not $ghAvailable) {
        $choco = cmd /c "choco --version 2>nul"
        if ($choco) {
            Write-Host "  Installing via Chocolatey..." -ForegroundColor Gray
            cmd /c "choco install gh --yes 2>&1"
            $env:PATH = [System.Environment]::GetEnvironmentVariable('PATH', 'Machine') + ';' + [System.Environment]::GetEnvironmentVariable('PATH', 'User')
            $ghAvailable = cmd /c "gh --version 2>nul"
            if ($ghAvailable) {
                Write-Host "  GitHub CLI installed successfully" -ForegroundColor Green
            }
        }
    }

    # Try Scoop if still not available
    if (-not $ghAvailable) {
        $scoop = cmd /c "scoop --version 2>nul"
        if ($scoop) {
            Write-Host "  Installing via Scoop..." -ForegroundColor Gray
            cmd /c "scoop install gh 2>&1"
            $env:PATH = [System.Environment]::GetEnvironmentVariable('PATH', 'Machine') + ';' + [System.Environment]::GetEnvironmentVariable('PATH', 'User')
            $ghAvailable = cmd /c "gh --version 2>nul"
            if ($ghAvailable) {
                Write-Host "  GitHub CLI installed successfully" -ForegroundColor Green
            }
        }
    }

    if (-not $ghAvailable) {
        Write-Host "  Could not install GitHub CLI automatically." -ForegroundColor Yellow
        Write-Host "  Install manually from https://cli.github.com then run 'gh auth login'" -ForegroundColor Yellow
        Write-Host "  Secrets will be printed below for manual entry." -ForegroundColor Yellow
    }
    $ErrorActionPreference = "Stop"
}

# Authenticate gh CLI if installed but not authenticated
if ($ghAvailable) {
    $ErrorActionPreference = "Continue"
    cmd /c "gh auth status 2>nul" | Out-Null
    $ghAuthed = ($LASTEXITCODE -eq 0)
    $ErrorActionPreference = "Stop"

    if (-not $ghAuthed) {
        Write-Host ""
        Write-Host "  ┌─────────────────────────────────────────────┐" -ForegroundColor Cyan
        Write-Host "  │        GitHub Authentication Required         │" -ForegroundColor Cyan
        Write-Host "  └─────────────────────────────────────────────┘" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "  The deployment script needs to configure GitHub" -ForegroundColor White
        Write-Host "  Actions secrets to enable automatic CI/CD." -ForegroundColor White
        Write-Host ""
        $ghAccountHint = if ($repoSlug -and $repoSlug -match '^([a-zA-Z0-9_.-]+)/') { $Matches[1] } else { "the repository owner" }
        Write-Host "  Please sign in with the GitHub account that" -ForegroundColor White
        Write-Host "  owns the repository ($ghAccountHint)." -ForegroundColor White
        Write-Host ""
        if ($repoSlug) {
            Write-Host "  Repository: github.com/$repoSlug" -ForegroundColor DarkGray
            Write-Host ""
        }
        Write-Host "  A browser window will open to complete login." -ForegroundColor Yellow
        Write-Host ""
        Read-Host "  Press Enter to continue to GitHub login"
        & gh auth login
        $ErrorActionPreference = "Continue"
        cmd /c "gh auth status 2>nul" | Out-Null
        $ghAuthed = ($LASTEXITCODE -eq 0)
        $ErrorActionPreference = "Stop"
        if ($ghAuthed) {
            Write-Host "  GitHub CLI authenticated successfully" -ForegroundColor Green
        } else {
            Write-Host "  Authentication failed - secrets will be printed for manual entry" -ForegroundColor Yellow
        }
    }
}

# Try to set secrets automatically via gh CLI
$secretsSet = $false

if ($ghAvailable -and $repoSlug -and $spJsonGh -match '"clientId"') {
    $ErrorActionPreference = "Continue"
    cmd /c "gh auth status 2>nul" | Out-Null
    $ghAuthed = ($LASTEXITCODE -eq 0)
    $ErrorActionPreference = "Stop"

    if ($ghAuthed) {
        Write-Host "  Setting GitHub Actions secrets via gh CLI..." -ForegroundColor Gray
        $tempFile = [System.IO.Path]::GetTempFileName()
        [System.IO.File]::WriteAllText($tempFile, $spJsonGh, [System.Text.Encoding]::UTF8)
        try {
            $ErrorActionPreference = "Continue"
            Get-Content $tempFile -Raw | & gh secret set AZURE_CREDENTIALS --repo $repoSlug
            & gh secret set ACR_LOGIN_SERVER   --body $acrServer        --repo $repoSlug
            & gh secret set ACR_USERNAME        --body $acrUsername      --repo $repoSlug
            & gh secret set ACR_PASSWORD        --body $acrPassword      --repo $repoSlug
            & gh secret set CONTAINER_APP_NAME  --body $containerAppName --repo $repoSlug
            & gh secret set RESOURCE_GROUP      --body $resourceGroup    --repo $repoSlug
            & gh secret set VITE_AZURE_CLIENT_ID --body $ClientId        --repo $repoSlug
            & gh secret set VITE_AZURE_TENANT_ID --body $TenantId        --repo $repoSlug
            $ErrorActionPreference = "Stop"
            Write-Host "  All 8 GitHub Actions secrets configured for: $repoSlug" -ForegroundColor Green
            Write-Host "  CI/CD is ready - every push to 'main' will auto-deploy" -ForegroundColor Green
            $secretsSet = $true
        } catch {
            Write-Host "  Failed to set secrets via gh CLI: $_" -ForegroundColor Yellow
        } finally {
            Remove-Item $tempFile -ErrorAction SilentlyContinue
        }
    } else {
        Write-Host "  GitHub CLI not authenticated." -ForegroundColor Yellow
        Write-Host "  Run 'gh auth login' then re-run this script, or set secrets manually below." -ForegroundColor Yellow
    }
} elseif (-not $ghAvailable) {
    Write-Host "  GitHub CLI not installed (https://cli.github.com) - set secrets manually below." -ForegroundColor Yellow
} elseif (-not $repoSlug) {
    Write-Host "  Could not detect GitHub remote URL - set secrets manually below." -ForegroundColor Yellow
}

if (-not $secretsSet) {
    $secretsUrl = if ($repoSlug) { "https://github.com/$repoSlug/settings/secrets/actions" } else { "https://github.com/<owner>/<repo>/settings/secrets/actions" }
    Write-Host ""
    Write-Host "  Add these secrets at:" -ForegroundColor Cyan
    Write-Host "  $secretsUrl" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  Secret Name              Value" -ForegroundColor White
    Write-Host "  -----------------------  -----" -ForegroundColor DarkGray
    Write-Host "  AZURE_CREDENTIALS        (JSON printed below)" -ForegroundColor White
    Write-Host "  ACR_LOGIN_SERVER         $acrServer" -ForegroundColor White
    Write-Host "  ACR_USERNAME             $acrUsername" -ForegroundColor White
    Write-Host "  ACR_PASSWORD             $acrPassword" -ForegroundColor White
    Write-Host "  CONTAINER_APP_NAME       $containerAppName" -ForegroundColor White
    Write-Host "  RESOURCE_GROUP           $resourceGroup" -ForegroundColor White
    Write-Host "  VITE_AZURE_CLIENT_ID     $ClientId" -ForegroundColor White
    Write-Host "  VITE_AZURE_TENANT_ID     $TenantId" -ForegroundColor White
    Write-Host ""
    Write-Host "  AZURE_CREDENTIALS value:" -ForegroundColor Cyan
    Write-Host $spJsonGh -ForegroundColor DarkGray
    Write-Host ""
}
