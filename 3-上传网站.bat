@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo ============================================
echo   上传看板到 GitHub Pages
echo ============================================

:: 检查 git 仓库
if not exist "%~dp0.git" (
    echo [错误] 未找到 .git 目录
    echo 请先运行: git init ^&^& git remote add origin 你的仓库地址
    pause
    exit /b 1
)

:: 检查 index.html
if not exist "%~dp0index.html" (
    echo [错误] 未找到 index.html
    pause
    exit /b 1
)

:: 同步 dashboard_v2.html 到 index.html
echo [1/3] 同步 index.html...
copy /Y "%~dp0dashboard_v2.html" "%~dp0index.html" >nul
echo [OK] index.html 已同步

:: 提交并推送
echo [2/3] 提交到 git...
git add index.html dashboard_v2.html data_v2_enc.js name_map.json .gitignore 3-上传网站.bat
git commit -m "更新看板 %date% %time%" 2>nul

echo [3/3] 推送到 GitHub...
git push origin main

echo.
echo [完成] 等待 1-2 分钟后访问:
echo https://xy-liu87.github.io/loan-dashboard/
pause
