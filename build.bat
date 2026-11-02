@echo off
rem ============================================================
rem  Tool calibration photo tool (SharePoint) - build script
rem  Double-click -> dist\(Korean name).exe is created
rem ============================================================
cd /d "%~dp0"

echo.
echo [1/2] Installing required packages...
python -m pip install -r requirements-desktop.txt
if errorlevel 1 goto :fail

echo.
echo [2/2] Building exe (takes 1-2 minutes)...
python build.py
if errorlevel 1 goto :fail

echo.
echo ============================================================
echo  Done. Opening the dist folder.
echo ============================================================
explorer dist
pause
exit /b 0

:fail
echo.
echo  *** Build failed. Please send a screenshot of this window. ***
pause
exit /b 1
