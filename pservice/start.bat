@echo off
chcp 65001 >nul
title OfficeHelper Python Service

echo ======================================
echo   OfficeHelper Python Service
echo   端口: http://127.0.0.1:8765
echo ======================================
echo.
echo 启动中...
cd /d "%~dp0pservice"

REM 检查 Python
python --version 2>nul || (echo [错误] 未找到 Python，请先安装 Python 3.9+ && pause && exit /b 1)

REM 安装依赖（如需要）
pip install -q -r requirements.txt 2>nul

REM 启动服务
python main.py

pause
