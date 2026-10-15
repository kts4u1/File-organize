@echo off
cd /d "%~dp0"
echo Creating Python environment (.venv) ...
python -m venv .venv || goto :error
call .venv\Scripts\activate.bat
python -m pip install --upgrade pip
python -m pip install -r requirements.txt || goto :error
echo.
echo Done. Start the program with run.bat
pause
exit /b 0
:error
echo.
echo Installation failed. Check that Python 3.10+ is installed and the network/proxy allows pip.
pause
exit /b 1
