@echo off
chcp 65001 >nul

:: ============================================
::  Elevate to admin
:: ============================================
fltmc >nul 2>&1 || (
    echo Requesting admin privileges...
    powershell -NoProfile -Command "Start-Process cmd.exe -ArgumentList '/c \"\"%~f0\"\"' -Verb RunAs"
    exit /b
)

echo ============================================
echo  VBMAN.DLL - Unregister, Clean and Register
echo ============================================
echo.

set "DLL_PATH=%~dp0VBMAN.dll"

if not exist "%DLL_PATH%" (
    echo [ERROR] VBMAN.dll not found!
    echo Expected -^> %DLL_PATH%
    pause
    exit /b 1
)

:: ---- Step 1 - Unregister DLL ----
echo [1/3] Unregistering DLL...
%windir%\SysWOW64\regsvr32.exe /u /s "%DLL_PATH%" 2>nul
%windir%\System32\regsvr32.exe /u /s "%DLL_PATH%" 2>nul
echo    Done.
echo.

:: ---- Step 2 - Clean leftover registry ----
echo [2/3] Cleaning VBMANLIB registry leftovers...

:: Delete VBMANLIB ProgID subtree (all ProgIDs like VBMANLIB.cJson etc.)
reg delete "HKCU\SOFTWARE\Classes\VBMANLIB"     /f >nul 2>&1
reg delete "HKLM\SOFTWARE\Classes\VBMANLIB"     /f >nul 2>&1
reg delete "HKLM\SOFTWARE\Classes\Wow6432Node\VBMANLIB" /f >nul 2>&1

:: Delete TypeLib entries named VBMANLIB
:: We query each root for TypeLib GUIDs whose default value is VBMANLIB
for /f "tokens=*" %%i in ('reg query "HKCU\SOFTWARE\Classes\TypeLib" /s /f "VBMANLIB" /e 2^>nul ^| findstr /i "HKEY_CURRENT_USER"') do reg delete "%%i" /f >nul 2>&1
for /f "tokens=*" %%i in ('reg query "HKLM\SOFTWARE\Classes\TypeLib" /s /f "VBMANLIB" /e 2^>nul ^| findstr /i "HKEY_LOCAL_MACHINE"') do reg delete "%%i" /f >nul 2>&1
for /f "tokens=*" %%i in ('reg query "HKLM\SOFTWARE\Classes\Wow6432Node\TypeLib" /s /f "VBMANLIB" /e 2^>nul ^| findstr /i "HKEY_LOCAL_MACHINE"') do reg delete "%%i" /f >nul 2>&1

echo    Done.
echo.

:: ---- Step 3 - Re-register DLL ----
echo [3/3] Registering DLL...

:: VB6 DLL is 32-bit, use SysWOW64 regsvr32
%windir%\SysWOW64\regsvr32.exe /s "%DLL_PATH%"
if errorlevel 1 (
    echo   [FAIL] VBMAN.dll registration failed!
) else (
    echo   [OK] VBMAN.dll registered successfully!
)

echo.
echo ============================================
echo  All done.
echo ============================================
pause
