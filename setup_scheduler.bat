@echo off
:: ============================================================
:: Weekly Update Tracker - Windows Task Scheduler Setup
:: ============================================================
:: Double-click this file to register the daily reminder task.
:: It will run every weekday at 9:00 AM automatically.
:: ============================================================

set PYTHON=C:\Users\kandhanr\AppData\Local\Programs\Python\Python310\python.exe
set SCRIPT=C:\Users\kandhanr\Weekly_update+tracking\weekly_update_tracker.py
set TASKNAME=Weekly_Update_Daily_Reminder

echo.
echo Setting up Task Scheduler for Weekly Update reminders...
echo.

:: Delete existing task if it exists (so we can re-register cleanly)
schtasks /delete /tn "%TASKNAME%" /f >nul 2>&1

:: Create the scheduled task — runs Mon-Fri at 09:00 AM
schtasks /create ^
  /tn "%TASKNAME%" ^
  /tr "\"%PYTHON%\" \"%SCRIPT%\" --send" ^
  /sc WEEKLY ^
  /d MON,TUE,WED,THU,FRI ^
  /st 09:00 ^
  /rl HIGHEST ^
  /f

if %ERRORLEVEL% == 0 (
    echo.
    echo [SUCCESS] Task registered successfully!
    echo   Name    : %TASKNAME%
    echo   Runs    : Monday to Friday at 09:00 AM
    echo   Action  : Sends reminder and overdue emails automatically
    echo.
    echo To test it RIGHT NOW without waiting, run:
    echo   schtasks /run /tn "%TASKNAME%"
    echo.
    echo To remove the task later, run:
    echo   schtasks /delete /tn "%TASKNAME%" /f
) else (
    echo.
    echo [ERROR] Failed to create the task.
    echo   Try right-clicking this file and selecting "Run as administrator".
)

echo.
pause
