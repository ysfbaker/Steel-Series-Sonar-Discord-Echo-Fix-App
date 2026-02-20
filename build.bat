@echo off
title Sonar Discord Echo Fix - CXA Studio's
cd /d "%~dp0"
REM normalize script dir without trailing backslash
set "ROOTDIR=%~dp0"
set "ROOT=%ROOTDIR:~0,-1%"

echo ============================================
echo  SteelSeries Sonar Discord Echo Fix
echo  Development by CodeXart Studio's
echo ============================================
echo.

echo "%ROOT%"

echo [1/3] Installing dependencies...
py -m pip install pycaw comtypes psutil pillow pystray pyinstaller --quiet
if %errorlevel% neq 0 (
    echo ERROR: pip failed. Make sure Python is installed.
    pause
    exit /b 1
)

echo [2/3] Building EXE...
py -3 -m PyInstaller --onefile --windowed --name "SonarDiscordFix" --distpath "%ROOT%\dist" --workpath "%ROOT%\build" --specpath "%ROOT%" --hidden-import pystray._win32 "%ROOT%\sonar_fix.py"

if %errorlevel% neq 0 (
    echo ERROR: Build failed!
    pause
    exit /b 1
)

echo [3/3] Done!
echo.
echo ============================================
echo  EXE: "%ROOT%\dist\SonarDiscordFix.exe"
echo ============================================
echo.
pause
