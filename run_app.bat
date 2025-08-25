@echo off
title Safety Stock Analyzer - Run Application
color 0B

echo.
echo ================================================
echo   SAFETY STOCK ANALYZER - RUN APPLICATION
echo ================================================
echo.
echo Starting the Safety Stock Analyzer...
echo.

:: Check if Python is available
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: Python not found!
    echo Please install Python 3.8+ first
    echo.
    pause
    exit /b 1
)

:: Check if required packages are installed
echo Checking dependencies...
python -c "import PyQt6, pandas, numpy, matplotlib" >nul 2>&1
if %errorlevel% neq 0 (
    echo Installing required packages...
    pip install -r requirements.txt
    if %errorlevel% neq 0 (
        echo ERROR: Failed to install packages!
        echo.
        pause
        exit /b 1
    )
)

echo ✓ All dependencies are ready!
echo.
echo Launching Safety Stock Analyzer...
echo.

:: Run the application
python safety_stock_analyzer.py

echo.
echo Application closed.
pause
