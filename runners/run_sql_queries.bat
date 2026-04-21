@echo off
cd /d "%~dp0.."

:: Rotate logs older than 30 days
if exist "logs\" (
    forfiles /p "logs" /m *.log /d -30 /c "cmd /c del @path" 2>nul
)

echo [%date% %time%] Starting SQL query runner...
python scripts/run_sql_queries.py
if %ERRORLEVEL% NEQ 0 (
    echo [%date% %time%] SQL query runner FAILED with exit code %ERRORLEVEL%
    exit /b %ERRORLEVEL%
)
echo [%date% %time%] SQL queries complete.
