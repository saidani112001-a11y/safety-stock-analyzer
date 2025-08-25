@echo off
title Safety Stock Analyzer - Build Executable
color 0A

echo.
echo ================================================
echo   SAFETY STOCK ANALYZER - BUILD EXECUTABLE
echo ================================================
echo.
echo This script will convert the Python app to a standalone .exe file
echo.

:: Check if Python is available
echo [1/4] Checking Python environment...
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: Python not found!
    echo Please install Python 3.8+ and try again
    echo.
    pause
    exit /b 1
)

echo ✓ Python environment is ready!
echo.

:: Install required packages
echo [2/4] Installing required packages...
echo.
pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo ERROR: Failed to install packages!
    echo Please check your internet connection and try again
    echo.
    pause
    exit /b 1
)

echo ✓ Packages installed successfully!
echo.

:: Install PyInstaller if not present
echo [3/4] Installing PyInstaller...
echo.
pip install pyinstaller
if %errorlevel% neq 0 (
    echo ERROR: Failed to install PyInstaller!
    echo.
    pause
    exit /b 1
)

echo ✓ PyInstaller installed successfully!
echo.

:: Build executable
echo [4/4] Building standalone executable...
echo.
echo This may take several minutes...
echo.

pyinstaller --onefile --windowed --name "Safety_Stock_Analyzer" --icon=NONE safety_stock_analyzer.py

if %errorlevel% neq 0 (
    echo ERROR: Build failed!
    echo Please check the error messages above
    echo.
    pause
    exit /b 1
)

echo.
echo ================================================
echo   BUILD COMPLETED SUCCESSFULLY!
echo ================================================
echo.
echo Your executable is located at:
echo dist\Safety_Stock_Analyzer.exe
echo.
echo You can now run the app without Python installed!
echo.
pause

