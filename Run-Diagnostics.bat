@echo off
cd /d "%~dp0"
"Runtime\python.exe" main.py
if errorlevel 1 pause
