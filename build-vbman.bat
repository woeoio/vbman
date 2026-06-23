@echo off
chcp 65001 >nul

:: 检查管理员权限
>nul 2>&1 net session
if %errorlevel% neq 0 (
    echo 正在请求管理员权限...
    powershell -Command "Start-Process cmd -ArgumentList '/c \"%~f0\"' -Verb RunAs"
    exit /b
)

echo ============================================
echo  VBMAN DLL 编译与复制脚本
echo ============================================
echo.

:: VB6.exe 路径
set VB6="D:\pro\Microsoft Visual Studio\VB98\VB6.EXE"

if not exist %VB6% (
    echo [错误] 未找到 vb6.exe: %VB6%
    pause
    exit /b 1
)

echo [1/2] 编译 VBMAN.dll ...
cd /d "%~dp0"
%VB6% /make "src\VBMAN.vbp"
if %errorlevel% neq 0 (
    echo [错误] 编译失败！错误码: %errorlevel%
    pause
    exit /b 1
)

echo [1/2] 编译成功！
echo.

:: 目标DLL路径
set SRC_DLL=%~dp0dist\DLL\VBMAN.dll
set DST_DIR=D:\code\vi\vbmanlib\vbman-demo\_bin\DLL

if not exist "%SRC_DLL%" (
    echo [错误] 编译产物不存在: %SRC_DLL%
    pause
    exit /b 1
)

echo [2/2] 复制 VBMAN.dll 到 %DST_DIR% ...

:: 确保目标目录存在
if not exist "%DST_DIR%" (
    mkdir "%DST_DIR%"
)

copy /y "%SRC_DLL%" "%DST_DIR%\VBMAN.dll"
if %errorlevel% neq 0 (
    echo [错误] 复制失败！
    pause
    exit /b 1
)

echo.
echo ============================================
echo  完成！
echo  From - %SRC_DLL%
echo  To   - %DST_DIR%\VBMAN.dll
echo ============================================
pause
