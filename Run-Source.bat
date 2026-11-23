@echo off
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
  py -3.12 -m venv .venv
  if errorlevel 1 goto failed
)
if not exist ".venv\mcq-ready" (
  ".venv\Scripts\python.exe" -m pip install -r requirements.txt
  if errorlevel 1 goto failed
  type nul > ".venv\mcq-ready"
)
".venv\Scripts\python.exe" main.py
if errorlevel 1 goto failed
exit /b 0
:failed
echo Setup or startup failed. See the error above. Source mode needs Python 3.12 x64.
pause
exit /b 1
