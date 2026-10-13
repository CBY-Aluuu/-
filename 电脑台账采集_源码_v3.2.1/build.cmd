@echo off
setlocal
cd /d "%~dp0"
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0tools\pc_inventory_portable\build-ios.ps1" -OutputDirectory "%~dp0build"
if errorlevel 1 (
    echo Build failed. See the message above.
    pause
    exit /b 1
)
echo Build succeeded. The application is in the build folder.
pause
