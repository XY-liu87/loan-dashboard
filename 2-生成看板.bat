@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo ============================================
echo   生成看板数据
echo ============================================

:: 查找 Python
set PYTHON=
for %%p in (python python3 py) do (
    where %%p >nul 2>&1
    if not errorlevel 1 (
        set PYTHON=%%p
        goto :found
    )
)
echo [错误] 未找到 Python，请先运行 1-安装环境.bat
pause
exit /b 1

:found
echo [1/2] 生成数据文件...
%PYTHON% 生成看板数据.py
if errorlevel 1 (
    echo [错误] 数据生成失败
    pause
    exit /b 1
)

echo [2/2] 同步 index.html...
copy /Y "%~dp0dashboard_v2.html" "%~dp0index.html" >nul
echo [OK] index.html 已更新

echo.
echo [完成] 数据已更新
echo   本地查看: 双击 dashboard_v2.html
echo   上传网站: 运行 3-上传网站.bat
pause
