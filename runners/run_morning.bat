@echo off
cd /d "%~dp0.."

:: Rotate logs older than 30 days
if exist "logs\" (
    forfiles /p "logs" /m *.log /d -30 /c "cmd /c del @path" 2>nul
)

echo [%date% %time%] Starting morning workflow...
python scripts/morning/run_morning_part1.py
if %ERRORLEVEL% NEQ 0 (
    echo [%date% %time%] Part 1 FAILED with exit code %ERRORLEVEL%
    exit /b %ERRORLEVEL%
)
echo [%date% %time%] Part 1 complete. Starting Part 2...
python scripts/morning/run_morning_part2.py
if %ERRORLEVEL% NEQ 0 (
    echo [%date% %time%] Part 2 FAILED with exit code %ERRORLEVEL%
    exit /b %ERRORLEVEL%
)
echo [%date% %time%] Morning workflow complete.
