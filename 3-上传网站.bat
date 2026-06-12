@echo off
cd /d "%~dp0"

set GIT=D:\Git\cmd\git.exe
if not exist "%GIT%" set GIT=git

copy /Y "%~dp0dashboard_v2.html" "%~dp0index.html" >nul
"%GIT%" add index.html dashboard_v2.html data_v2_enc.js name_map.json .gitignore 3-upload.bat 2-Generate.bat
"%GIT%" commit -m "update"
"%GIT%" push origin master

echo Done. https://xy-liu87.github.io/loan-dashboard/
pause
