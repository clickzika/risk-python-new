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

echo.
echo Done. Verify in Task Scheduler under LHFund\.
echo To run manually: schtasks /run /tn "LHFund\RiskMorning"
echo                  schtasks /run /tn "LHFund\RiskEvening"
