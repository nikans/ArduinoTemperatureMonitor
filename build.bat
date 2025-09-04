@echo off
echo ========================================
echo Temperature Monitor Build Script
echo ========================================
echo.

REM Check if Python is available
python --version >nul 2>&1
if errorlevel 1 (
    echo Error: Python is not installed or not in PATH
    echo Please install Python 3.7 or higher https://www.python.org/downloads/windows/
    pause
    exit /b 1
)

echo Python found. Starting build process...
echo.

REM Run the build script
python build_exe.py

if errorlevel 1 (
    echo.
    echo Build failed! Check the error messages above.
    pause
    exit /b 1
)

echo.
echo ========================================
echo Build completed successfully!
echo ========================================
echo.
echo Files created:
echo   - dist\TemperatureMonitor.exe
echo   - dist\install.bat (admin installer)
echo   - dist\install_simple.bat (simple installer)
echo   - dist\uninstall.bat
echo   - TemperatureMonitor.zip
echo.
echo To test the application:
echo   1. Run: dist\TemperatureMonitor.exe
echo.
echo Distribution package ready:
echo   - TemperatureMonitor.zip contains everything needed
echo   - Share this ZIP file with users
echo   - Users extract and choose installer:
echo     - install_simple.bat (recommended - no admin needed)
echo     - install.bat (system-wide installation)
echo.
pause
