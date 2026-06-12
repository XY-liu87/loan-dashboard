@echo off
cd /d "%~dp0"

echo ============================================
echo   Upload to GitHub Pages
echo ============================================

set GIT=D:\Git\cmd\git.exe
if not exist "%GIT%" set GIT=git

echo Step 1: Sync index.html...
copy /Y "%~dp0dashboard_v2.html" "%~dp0index.html" >/dev/null

echo Step 2: Git commit...
"%GIT%" add index.html dashboard_v2.html data_v2_enc.js name_map.json .gitignore 3-upload.bat 2-Generate.bat
"%GIT%" commit -m "update dashboard"

echo Step 3: Git push...
"%GIT%" push origin master

echo.
echo [OK] Done! Wait 1 min then visit:
echo https://xy-liu87.github.io/loan-dashboard/
pause
