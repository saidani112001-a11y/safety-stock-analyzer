@echo off
echo ========================================
echo    Safety Stock Analyzer - Deploy
echo ========================================
echo.

echo [1/5] Cleaning previous builds...
if exist "build" rmdir /s /q "build"
if exist "dist" rmdir /s /q "dist"
if exist "*.spec" del "*.spec"
echo ✓ Cleaned previous builds

echo.
echo [2/5] Installing dependencies...
pip install -r requirements.txt
pip install pyinstaller
echo ✓ Dependencies installed

echo.
echo [3/5] Building executable...
pyinstaller --onefile --windowed --name "SafetyStockAnalyzer" --icon=NONE safety_stock_analyzer.py
echo ✓ Executable built

echo.
echo [4/5] Creating distribution package...
if not exist "dist\package" mkdir "dist\package"
copy "dist\SafetyStockAnalyzer.exe" "dist\package\"
copy "README.md" "dist\package\"
copy "LICENSE" "dist\package\"
copy "requirements.txt" "dist\package\"
copy "run_app.bat" "dist\package\"
copy "build_exe.bat" "dist\package\"

echo.
echo [5/5] Creating ZIP archive...
cd dist
powershell Compress-Archive -Path "package\*" -DestinationPath "SafetyStockAnalyzer-v1.0.0.zip" -Force
cd ..
echo ✓ ZIP archive created

echo.
echo ========================================
echo    DEPLOYMENT COMPLETE!
echo ========================================
echo.
echo 📁 Executable: dist\SafetyStockAnalyzer.exe
echo 📦 Package: dist\SafetyStockAnalyzer-v1.0.0.zip
echo 📋 Source: dist\package\
echo.
echo 🚀 Ready for distribution!
echo.
pause
