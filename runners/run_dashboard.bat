@echo off
cd /d "%~dp0.."
echo [%date% %time%] Generating monitoring dashboard...
python scripts/generate_dashboard.py
if %ERRORLEVEL% NEQ 0 (
    echo [%date% %time%] Dashboard generation FAILED with exit code %ERRORLEVEL%
    exit /b %ERRORLEVEL%
)
echo [%date% %time%] Dashboard updated at docs\dashboard.html
