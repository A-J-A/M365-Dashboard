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

# Try to load config from Register-EntraApp.ps1 output
$configPath = Join-Path (Join-Path $PSScriptRoot "..") "entra-app-config.json"
if (Test-Path $configPath) {
    Write-Host "Found Entra app config from previous registration..." -ForegroundColor Green
    $savedConfig = Get-Content $configPath | ConvertFrom-Json
    
    if (-not $TenantId -and $savedConfig.TenantId) {
        $TenantId = $savedConfig.TenantId
        Write-Host "  Using Tenant ID: $TenantId" -ForegroundColor Gray
    }
    if (-not $ClientId -and $savedConfig.ClientId) {
        $ClientId = $savedConfig.ClientId
        Write-Host "  Using Client ID: $ClientId" -ForegroundColor Gray
    }
    if (-not $ClientSecret -and $savedConfig.ClientSecret) {
        $ClientSecret = $savedConfig.ClientSecret
        Write-Host "  Using Client Secret: ********" -ForegroundColor Gray
    }
}

# Prompt for any still-missing values
if (-not $TenantId) {
    $TenantId = Read-Host "Enter Entra ID Tenant ID"
}
if (-not $ClientId) {
    $ClientId = Read-Host "Enter Entra ID Client ID"
}
if (-not $ClientSecret) {
    $secureSecret = Read-Host "Enter Entra ID Client Secret" -AsSecureString
    $ClientSecret = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($secureSecret))
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

# Check Azure CLI login and select subscription early
Write-Host ""
Write-Host "Azure Subscription" -ForegroundColor Cyan
Write-Host "------------------" -ForegroundColor Cyan

$ErrorActionPreference = "Continue"
$accountJson = cmd /c "az account show 2>nul"
$ErrorActionPreference = "Stop"

if (-not $accountJson) {
    Write-Host "Please run 'az login' first" -ForegroundColor Red
    exit 1
}

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
    cmd /c "az acr build --registry $acrName --image m365dashboard:latest . 2>&1"
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
            $ErrorActionPreference = "Stop"

            # Verify
            $secretList = cmd /c "gh secret list --repo $repoSlug 2>nul"
            $expected = @("AZURE_CREDENTIALS", "ACR_LOGIN_SERVER", "ACR_USERNAME", "ACR_PASSWORD", "CONTAINER_APP_NAME", "RESOURCE_GROUP")
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
Write-Host "Open the URL above and sign in with your Microsoft 365 account!" -ForegroundColor White
Write-Host ""
