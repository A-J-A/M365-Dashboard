@echo off
echo.
echo ======================================================
echo   M365 Dashboard - Azure SQL Database Setup
echo ======================================================
echo.
echo This script will:
echo   1. Create an Azure SQL Server
echo   2. Create the M365Dashboard database
echo   3. Configure firewall rules
echo   4. Update your configuration files
echo   5. Apply database migrations
echo.
echo Prerequisites:
echo   - Azure CLI installed (https://aka.ms/installazurecli)
echo   - Logged in to Azure (script will prompt if not)
echo.
echo Estimated cost: ~£3.77/month (Basic tier)
echo.
pause

PowerShell -ExecutionPolicy Bypass -File "%~dp0Setup-AzureSql.ps1"

echo.
pause
