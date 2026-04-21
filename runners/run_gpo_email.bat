@echo off
cd /d "%~dp0.."

:: Rotate logs older than 30 days
if exist "logs\" (
    forfiles /p "logs" /m *.log /d -30 /c "cmd /c del @path" 2>nul
)

echo [%date% %time%] Starting GPO email...
python scripts/evening/send_gpo_email.py
if %ERRORLEVEL% NEQ 0 (
    echo [%date% %time%] GPO email FAILED with exit code %ERRORLEVEL%
    exit /b %ERRORLEVEL%
)
echo [%date% %time%] GPO email complete.
