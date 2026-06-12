@echo off
cd /d "%~dp0"

set GIT=D:\Git\cmd\git.exe
if not exist "%GIT%" set GIT=git

copy /Y "%~dp0dashboard_v2.html" "%~dp0index.html" >nul
"%GIT%" add -A
"%GIT%" commit -m "update"
"%GIT%" push origin master

echo [OK] https://xy-liu87.github.io/loan-dashboard/
pause
