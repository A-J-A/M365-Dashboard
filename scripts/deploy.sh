#!/bin/bash
# ============================================================================
# M365 Dashboard - Quick Deploy Script
# ============================================================================
# Usage: ./deploy.sh -n <name-prefix> -l <location> -t <tenant-id> -c <client-id> -s <client-secret>
# ============================================================================

set -e

# Default values
NAME_PREFIX="m365dash"
LOCATION="uksouth"
ENVIRONMENT="prod"

# Parse arguments
while getopts "n:l:e:t:c:s:p:" opt; do
  case $opt in
    n) NAME_PREFIX="$OPTARG" ;;
    l) LOCATION="$OPTARG" ;;
    e) ENVIRONMENT="$OPTARG" ;;
    t) TENANT_ID="$OPTARG" ;;
    c) CLIENT_ID="$OPTARG" ;;
    s) CLIENT_SECRET="$OPTARG" ;;
    p) SQL_PASSWORD="$OPTARG" ;;
    \?) echo "Invalid option -$OPTARG" >&2; exit 1 ;;
  esac
done

# Prompt for missing values
if [ -z "$TENANT_ID" ]; then
  read -p "Enter Entra ID Tenant ID: " TENANT_ID
fi

if [ -z "$CLIENT_ID" ]; then
  read -p "Enter Entra ID Client ID: " CLIENT_ID
fi

if [ -z "$CLIENT_SECRET" ]; then
  read -sp "Enter Entra ID Client Secret: " CLIENT_SECRET
  echo
fi

if [ -z "$SQL_PASSWORD" ]; then
  read -sp "Enter SQL Admin Password (min 8 chars, uppercase, lowercase, number): " SQL_PASSWORD
  echo
fi

echo ""
echo "============================================"
echo "M365 Dashboard Deployment"
echo "============================================"
echo "Name Prefix:  $NAME_PREFIX"
echo "Location:     $LOCATION"
echo "Environment:  $ENVIRONMENT"
echo "Tenant ID:    $TENANT_ID"
echo "Client ID:    $CLIENT_ID"
echo "============================================"
echo ""

# Check Azure CLI is logged in
echo "Checking Azure CLI login..."
az account show > /dev/null 2>&1 || { echo "Please run 'az login' first"; exit 1; }

# Deploy infrastructure
echo "Deploying Azure infrastructure..."
DEPLOYMENT_OUTPUT=$(az deployment sub create \
  --location "$LOCATION" \
  --template-file infra/main.bicep \
  --parameters namePrefix="$NAME_PREFIX" \
  --parameters location="$LOCATION" \
  --parameters environment="$ENVIRONMENT" \
  --parameters entraIdTenantId="$TENANT_ID" \
  --parameters entraIdClientId="$CLIENT_ID" \
  --parameters entraIdClientSecret="$CLIENT_SECRET" \
  --parameters sqlAdminPassword="$SQL_PASSWORD" \
  --query properties.outputs -o json)

# Extract outputs
RESOURCE_GROUP=$(echo $DEPLOYMENT_OUTPUT | jq -r '.resourceGroupName.value')
ACR_NAME=$(echo $DEPLOYMENT_OUTPUT | jq -r '.containerRegistryName.value')
ACR_SERVER=$(echo $DEPLOYMENT_OUTPUT | jq -r '.containerRegistryLoginServer.value')
APP_URL=$(echo $DEPLOYMENT_OUTPUT | jq -r '.containerAppUrl.value')

echo ""
echo "Infrastructure deployed successfully!"
echo ""

# Build and push Docker image
echo "Building Docker image..."
az acr login --name "$ACR_NAME"

docker build -t "$ACR_SERVER/m365dashboard:latest" .
docker push "$ACR_SERVER/m365dashboard:latest"

echo ""
echo "Updating Container App with new image..."
az containerapp update \
  --name "${NAME_PREFIX}-${ENVIRONMENT}-app" \
  --resource-group "$RESOURCE_GROUP" \
  --image "$ACR_SERVER/m365dashboard:latest"

echo ""
echo "============================================"
echo "Deployment Complete!"
echo "============================================"
echo ""
echo "Your M365 Dashboard is available at:"
echo "  $APP_URL"
echo ""
echo "Next steps:"
echo "1. Add redirect URI to your App Registration:"
echo "   ${APP_URL}/authentication/login-callback"
echo ""
echo "2. Enable tokens in App Registration > Authentication:"
echo "   - Access tokens"
echo "   - ID tokens"
echo ""
echo "3. Open $APP_URL and sign in!"
echo ""
