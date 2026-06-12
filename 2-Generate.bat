@echo off
cd /d "%~dp0"

echo ============================================
echo   Step 1: Generate Dashboard Data
echo ============================================

set PYTHON=C:\Program Files\Python312\python.exe
if not exist "%PYTHON%" set PYTHON=python

echo Generating data...
"%PYTHON%" "%~dp0generate_data.py"
if errorlevel 1 (
    echo [ERROR] Data generation failed!
    pause
    exit /b 1
)

echo.
echo Step 2: Sync index.html...
copy /Y "%~dp0dashboard_v2.html" "%~dp0index.html" >/dev/null
echo [OK] Done!

echo.
echo Next: Run 3-upload.bat to push to GitHub
pause
