@echo off
:: ============================================================
:: setup_scheduler.bat — Register LHFund Risk automation tasks
:: Run ONCE as Administrator on the workstation.
:: To remove tasks: schtasks /delete /tn "LHFund\RiskMorning" /f
::                  schtasks /delete /tn "LHFund\RiskEvening" /f
:: ============================================================

set PYTHON=C:\ProgramData\anaconda3\python.exe
set PROJECT=%~dp0..
set RUNNER_MORNING="%PROJECT%\runners\run_morning.bat"
set RUNNER_EVENING="%PROJECT%\runners\run_evening.bat"
set RUNNER_GPO_EMAIL="%PROJECT%\runners\run_gpo_email.bat"

echo Registering LHFund Risk scheduled tasks...

:: Morning — Mon-Fri at 07:30
schtasks /create ^
  /tn "LHFund\RiskMorning" ^
  /tr "%RUNNER_MORNING%" ^
  /sc WEEKLY /d MON,TUE,WED,THU,FRI ^
  /st 07:30 ^
  /rl HIGHEST ^
  /f
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Failed to register RiskMorning task.
    exit /b %ERRORLEVEL%
)
echo [OK] RiskMorning registered — runs Mon-Fri at 07:30

:: Evening — Mon-Fri at 17:00
schtasks /create ^
  /tn "LHFund\RiskEvening" ^
  /tr "%RUNNER_EVENING%" ^
  /sc WEEKLY /d MON,TUE,WED,THU,FRI ^
  /st 17:00 ^
  /rl HIGHEST ^
  /f
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Failed to register RiskEvening task.
    exit /b %ERRORLEVEL%
)
echo [OK] RiskEvening registered  — runs Mon-Fri at 17:00

:: GPO standalone email — Mon-Fri at 17:15
schtasks /create ^
  /tn "LHFund\RiskGPOEmail" ^
  /tr "%RUNNER_GPO_EMAIL%" ^
  /sc WEEKLY /d MON,TUE,WED,THU,FRI ^
  /st 17:15 ^
  /rl HIGHEST ^
  /f
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Failed to register RiskGPOEmail task.
    exit /b %ERRORLEVEL%
)
echo [OK] RiskGPOEmail registered  — runs Mon-Fri at 17:15

:: SQL queries — Mon-Fri at 18:00 (after market data is settled)
schtasks /create ^
  /tn "LHFund\RiskSQLQueries" ^
  /tr "%PROJECT%\runners\run_sql_queries.bat" ^
  /sc WEEKLY /d MON,TUE,WED,THU,FRI ^
  /st 18:00 ^
  /rl HIGHEST ^
  /f
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Failed to register RiskSQLQueries task.
    exit /b %ERRORLEVEL%
)
echo [OK] RiskSQLQueries registered — runs Mon-Fri at 18:00

:: Dashboard — Mon-Fri at 17:30 (after evening + GPO email)
schtasks /create ^
  /tn "LHFund\RiskDashboard" ^
  /tr "%PROJECT%\runners\run_dashboard.bat" ^
  /sc WEEKLY /d MON,TUE,WED,THU,FRI ^
  /st 17:30 ^
  /rl HIGHEST ^
  /f
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Failed to register RiskDashboard task.
    exit /b %ERRORLEVEL%
)
echo [OK] RiskDashboard registered — runs Mon-Fri at 17:30

echo.
echo Done. Verify in Task Scheduler under LHFund\.
echo To run manually: schtasks /run /tn "LHFund\RiskMorning"
echo                  schtasks /run /tn "LHFund\RiskEvening"
echo                  schtasks /run /tn "LHFund\RiskGPOEmail"
echo                  schtasks /run /tn "LHFund\RiskDashboard"
echo                  schtasks /run /tn "LHFund\RiskSQLQueries"
