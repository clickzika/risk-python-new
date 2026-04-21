@echo off
cd /d "%~dp0.."

:: Rotate logs older than 30 days
if exist "logs\" (
    forfiles /p "logs" /m *.log /d -30 /c "cmd /c del @path" 2>nul
)

echo [%date% %time%] Starting evening workflow (GPO)...
python scripts/evening/run_evening.py
if %ERRORLEVEL% NEQ 0 (
    echo [%date% %time%] Evening script FAILED with exit code %ERRORLEVEL%
    exit /b %ERRORLEVEL%
)
echo [%date% %time%] Evening workflow complete.
