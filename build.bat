@echo off
chcp 65001 >nul 2>&1
setlocal enabledelayedexpansion

set "LOGFILE=%~dp0build.log"
set "ERROR_LOG=%~dp0build_error.log"
set "PYTHON=D:\soft\python3.14\python.exe"
set "NPM=D:\soft\nodeJS\npm.cmd"
set "CARGO=C:\Users\gaotianyu.35\.cargo\bin\cargo.exe"
set "RUSTUP=C:\Users\gaotianyu.35\.cargo\bin\rustup.exe"
set "MSVC_BIN=D:\soft\MSVC\VC\Tools\MSVC\14.50.35717\bin\Hostx64\x64"
set "CARGO_BIN=C:\Users\gaotianyu.35\.cargo\bin"

echo ======================================== > "%LOGFILE%"
echo OfficeHelper 完整构建脚本 >> "%LOGFILE%"
date /t >> "%LOGFILE%"
time /t >> "%LOGFILE%"
echo ======================================== >> "%LOGFILE%"
echo.

echo [INFO] 日志文件: %LOGFILE%
echo [INFO] 错误日志: %ERROR_LOG%
echo.

REM ── 步骤 1：PyInstaller 打包 Python 后端 ───────────────────────────────
echo [1/3] 打包 Python 后端 (PyInstaller)...
if exist "dist\OfficeHelperBackend\OfficeHelperBackend.exe" (
    echo     后端已存在，跳过
) else (
    echo     运行: pyinstaller OfficeHelperBackend.spec --clean
    "%PYTHON%" -m PyInstaller OfficeHelperBackend.spec --clean >> "%LOGFILE%" 2>&1
    if errorlevel 1 (
        echo     [错误] PyInstaller 打包失败，请查看 %LOGFILE%
        echo [错误] PyInstaller 打包失败 >> "%ERROR_LOG%"
        echo ======================================== >> "%ERROR_LOG%"
        type "%LOGFILE%" >> "%ERROR_LOG%"
        pause
        exit /b 1
    )
    echo     PyInstaller 完成
)
echo.

REM ── 步骤 2：前端构建 ───────────────────────────────────────────────────
echo [2/3] 构建前端 (Vite)...
if exist "front\dist\index.html" (
    echo     前端已存在，跳过
) else (
    echo     运行: cd front ^&^& npm run build
    cd front
    call "%NPM%" run build >> "%LOGFILE%" 2>&1
    cd ..
    if errorlevel 1 (
        echo     [错误] 前端构建失败，请查看 %LOGFILE%
        echo [错误] 前端构建失败 >> "%ERROR_LOG%"
        echo ======================================== >> "%ERROR_LOG%"
        type "%LOGFILE%" >> "%ERROR_LOG%"
        pause
        exit /b 1
    )
    echo     前端构建完成
)
echo.

REM ── 步骤 3：Tauri 打包 ─────────────────────────────────────────────────
echo [3/3] 打包 Tauri 应用...
echo     设置 Cargo 环境变量...
set "CARGO_HOME=C:\Users\gaotianyu.35\.cargo"
set "RUSTUP_HOME=C:\Users\gaotianyu.35\.cargo"
set "PATH=%CARGO_BIN%;%MSVC_BIN%;%PATH%"
echo     CARGO_HOME=%CARGO_HOME%
echo     运行: npm run tauri build
cd front
call "%NPM%" run tauri build >> "%LOGFILE%" 2>&1
cd ..
if errorlevel 1 (
    echo     [错误] Tauri 打包失败，请查看 %LOGFILE%
    echo [错误] Tauri 打包失败 >> "%ERROR_LOG%"
    echo ======================================== >> "%ERROR_LOG%"
    type "%LOGFILE%" >> "%ERROR_LOG%"
    pause
    exit /b 1
)
echo     Tauri 打包完成
echo.

REM ── 步骤 4：查找输出文件 ───────────────────────────────────────────────
echo [4/4] 查找安装包...
for /r "front\src-tauri\target\release\bundle" %%f in (*-setup.exe) do (
    echo     找到安装包: %%f
)
echo.

echo ========================================
echo 构建完成！
echo ========================================
echo 后端目录: dist\OfficeHelperBackend\
echo 前端目录: front\dist\
echo Tauri输出: front\src-tauri\target\release\bundle\
echo 日志文件: %LOGFILE%
echo.
pause
