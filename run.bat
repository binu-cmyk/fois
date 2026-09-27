@echo off
title FOIS Tracker & MIS Dashboard
echo =======================================================
echo          FOIS Tracker & MIS Dashboard Setup
echo =======================================================
echo.
echo Installing/verifying Python dependencies...
python -m pip install -r requirements.txt
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo WARNING: Dependency installation failed. Trying user-level install...
    python -m pip install -r requirements.txt --user
)
echo.
echo Launching FOIS Dashboard...
echo Starting server on http://127.0.0.1:5000/
start http://127.0.0.1:5000/
echo.
python app.py
pause
