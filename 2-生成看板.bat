@echo off
cd /d "%~dp0"

set PYTHON=C:\Program Files\Python312\python.exe
if not exist "%PYTHON%" set PYTHON=python

"%PYTHON%" "%~dp0generate_data.py"
if errorlevel 1 pause

copy /Y "%~dp0dashboard_v2.html" "%~dp0index.html" >nul
echo Done.
pause
