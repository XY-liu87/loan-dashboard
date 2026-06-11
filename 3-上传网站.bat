@echo off
chcp 65001 >nul
title 看板 - 上传网站
echo.
echo ============================================
echo   贷后数据看板 - 发布到网站
echo ============================================
echo.

cd /d "%~dp0"

if not exist "%~dp0index.html" (
    echo [WARN] 未找到 index.html, 请先运行 2-生成看板.bat
    pause
    exit /b 1
)

git --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [FAIL] 未找到 Git!
    echo 请安装 Git: https://git-scm.com/download/win
    start https://git-scm.com/download/win
    pause
    exit /b 1
)

git remote -v 2>nul | findstr "loan-dashboard" >nul
if %errorlevel% neq 0 (
    echo [INFO] 正在关联 GitHub 仓库...
    if not exist "%~dp0.git" (
        git init
        git checkout -b master
    )
    git remote add origin https://github.com/XY-liu87/loan-dashboard.git 2>nul
    if %errorlevel% neq 0 (
        git remote set-url origin https://github.com/XY-liu87/loan-dashboard.git
    )
    echo [OK] 已关联远程仓库
)

echo [1/3] 检查变更...
git status --short

echo [2/3] 提交更新...
git add index.html dashboard_v2.html data_v2.js data_v2_enc.js 看板结果
git commit -m "更新看板数据"
if %errorlevel% equ 0 (
    echo [OK] 新变更已提交
) else (
    echo [INFO] 没有新变更需要提交, 直接推送已有提交...
)

echo [3/3] 上传到 GitHub...
git push -u origin master
if %errorlevel% equ 0 (
    echo ============================================
    echo   [OK] 发布成功!
    echo   https://xy-liu87.github.io/loan-dashboard/
    echo ============================================
) else (
    echo [FAIL] 上传失败,请检查网络或 GitHub 权限
)
echo.
pause
