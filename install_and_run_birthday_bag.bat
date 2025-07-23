@echo off
echo Birthday Bag Exporter - Installation and Launcher
echo ================================================
echo.

REM Check if Python is installed
python --version > nul 2>&1
if %errorlevel% neq 0 (
    echo Python is not installed or not in PATH.
    echo Please install Python 3.6 or higher from https://www.python.org/downloads/
    echo Make sure to check "Add Python to PATH" during installation.
    echo.
    pause
    exit /b 1
)

echo Installing required packages...
python -m pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo Failed to install required packages.
    echo Please try running: pip install -r requirements.txt
    echo.
    pause
    exit /b 1
)

echo.
echo Starting Birthday Bag Exporter...
echo.
python birthday_bag_exporter.py
if %errorlevel% neq 0 (
    echo Application exited with an error.
    echo.
    pause
)

exit /b 0