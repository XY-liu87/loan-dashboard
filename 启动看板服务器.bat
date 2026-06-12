@echo off
chcp 65001 >nul
title 贷后数据看板 v3 - 权限服务器

echo.
echo ╔══════════════════════════════════════════════╗
echo ║     贷后数据看板 v3 权限服务器             ║
echo ╚══════════════════════════════════════════════╝
echo.

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
echo [信息] 使用 Python: %PYTHON%

:: 检查 Flask
%PYTHON% -c "import flask" >nul 2>&1
if errorlevel 1 (
    echo [信息] 正在安装 Flask...
    %PYTHON% -m pip install flask -i https://pypi.tuna.tsinghua.edu.cn/simple
    if errorlevel 1 (
        echo [错误] Flask 安装失败
        pause
        exit /b 1
    )
)

:: 启动服务器
echo [信息] 启动看板服务器...
echo.
%PYTHON% server.py --port 5000
pause
