@echo off
cd /d "%~dp0.."
echo [%date% %time%] Starting evening workflow (GPO)...
python scripts/evening/run_evening.py
if %ERRORLEVEL% NEQ 0 (
    echo [%date% %time%] Evening script FAILED with exit code %ERRORLEVEL%
    exit /b %ERRORLEVEL%
)
echo [%date% %time%] Evening workflow complete.
