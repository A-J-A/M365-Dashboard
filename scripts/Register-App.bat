@echo off
echo.
echo ======================================================
echo   M365 Dashboard - Entra ID App Registration
echo ======================================================
echo.
echo This script will:
echo   1. Install Microsoft.Graph PowerShell module (if needed)
echo   2. Create an Entra ID App Registration
echo   3. Configure all required permissions
echo   4. Create app roles (Admin/Reader)
echo   5. Generate a client secret
echo   6. Update your configuration files
echo.
echo You will be prompted to sign in with an account that has
echo Application Administrator or Global Administrator role.
echo.
pause

PowerShell -ExecutionPolicy Bypass -File "%~dp0Register-M365DashboardApp.ps1"

echo.
pause
