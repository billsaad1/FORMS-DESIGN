@echo off
echo ====================================================================
echo      Jeel Al Bena Bilingual Protection Forms Manager
echo ====================================================================
echo.
echo Installing required Python libraries...
pip install -r requirements.txt
if %ERRORLEVEL% neq 0 (
    echo.
    echo [ERROR] Python installation or pip command failed.
    echo Please make sure Python 3 is installed on this PC and added to PATH.
    pause
    exit /b
)
echo.
echo Launching local Flask Server...
echo.
python app/server.py
pause
