# Build stage for frontend
FROM node:20-alpine AS frontend-build
WORKDIR /app/frontend
COPY src/M365Dashboard.Api/ClientApp/package*.json ./
RUN npm ci
COPY src/M365Dashboard.Api/ClientApp/ ./
# Accept Entra app credentials as build args so Vite can bake them into the JS bundle
ARG VITE_AZURE_CLIENT_ID
ARG VITE_AZURE_TENANT_ID
ENV VITE_AZURE_CLIENT_ID=$VITE_AZURE_CLIENT_ID
ENV VITE_AZURE_TENANT_ID=$VITE_AZURE_TENANT_ID
# Fix permissions for node_modules binaries and run build
RUN chmod -R +x node_modules/.bin && npm run build

# Build stage for backend
FROM mcr.microsoft.com/dotnet/sdk:8.0 AS backend-build
WORKDIR /app
COPY src/M365Dashboard.Api/*.csproj ./
RUN dotnet restore
COPY src/M365Dashboard.Api/ ./
# Skip SPA build in dotnet publish - frontend is built separately
RUN dotnet publish -c Release -o /app/publish -p:SpaRoot= -p:SkipBuildWebpack=true

# Write version file - populated by the release workflow via BUILD_VERSION arg
ARG BUILD_VERSION=unknown
RUN echo "${BUILD_VERSION}" > /app/publish/version.txt

# Runtime stage
FROM mcr.microsoft.com/dotnet/aspnet:8.0
WORKDIR /app

# Install dependencies for QuestPDF / SkiaSharp (native PDF rendering)
RUN apt-get update && apt-get install -y \
    libfontconfig1 \
    fontconfig \
    fonts-dejavu-core \
    fonts-liberation \
    libssl3 \
    libicu72 \
    libgdiplus \
    libc6-dev \
    && rm -rf /var/lib/apt/lists/*

# Copy published backend
COPY --from=backend-build /app/publish ./

# Copy built frontend to wwwroot
COPY --from=frontend-build /app/frontend/build ./wwwroot

# Expose ports
EXPOSE 8080
EXPOSE 8081

# Set environment variables
ENV ASPNETCORE_URLS=http://+:8080
ENV ASPNETCORE_ENVIRONMENT=Production

ENTRYPOINT ["dotnet", "M365Dashboard.Api.dll"]
