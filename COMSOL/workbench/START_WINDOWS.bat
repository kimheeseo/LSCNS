@echo off
cd /d "%~dp0"
where py >nul 2>nul
if errorlevel 1 (
  echo Python 3.12 is required. See README.md.
  pause
  exit /b 1
)
if not exist ".venv\Scripts\python.exe" (
  py -3.12 -m venv .venv
  if errorlevel 1 goto failed
)
if not exist ".venv\READY_0_1" (
  .venv\Scripts\python.exe -m pip install -r requirements.txt
  if errorlevel 1 goto failed
  echo ready > ".venv\READY_0_1"
)
.venv\Scripts\python.exe launch.py
if errorlevel 1 goto failed
exit /b 0
:failed
echo Setup or execution failed. See README.md for manual installation.
pause
exit /b 1
