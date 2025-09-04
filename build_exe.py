#!/usr/bin/env python3
"""
Build script for Temperature Monitor application
Creates a standalone executable using PyInstaller
"""

import os
import sys
import subprocess
import shutil
import zipfile
from pathlib import Path

def install_requirements():
    """Install required packages for building"""
    print("Installing build requirements...")
    requirements = [
        'pyinstaller>=5.0',
        'pyserial>=3.5',
        'matplotlib>=3.7.0',
        'numpy>=1.24.0',
        'PyYAML>=6.0',
        'Pillow>=9.0.0'  # For icon creation
    ]

    for req in requirements:
        try:
            # Use --break-system-packages flag to install in system
            subprocess.check_call([
                sys.executable, '-m', 'pip', 'install',
                '--break-system-packages',
                '--user',  # Install to user directory to be safer
                req
            ])
            print(f"✓ Installed {req}")
        except subprocess.CalledProcessError as e:
            print(f"✗ Failed to install {req}: {e}")
            print("Trying alternative installation method...")
            try:
                # Fallback: try with --break-system-packages only
                subprocess.check_call([
                    sys.executable, '-m', 'pip', 'install',
                    '--break-system-packages',
                    req
                ])
                print(f"✓ Installed {req} (system-wide)")
            except subprocess.CalledProcessError as e2:
                print(f"✗ Failed to install {req} with both methods: {e2}")
                return False
    return True

def create_icon():
    """Create a simple icon file if it doesn't exist"""
    icon_path = Path('icon.ico')
    if not icon_path.exists():
        print("Creating application icon...")
        # Create a simple 32x32 icon using PIL if available
        try:
            from PIL import Image, ImageDraw

            # Create a simple thermometer icon
            img = Image.new('RGBA', (32, 32), (0, 0, 0, 0))
            draw = ImageDraw.Draw(img)

            # Draw thermometer body
            draw.rectangle([12, 4, 16, 24], fill=(0, 100, 200, 255))
            draw.rectangle([10, 24, 18, 28], fill=(0, 100, 200, 255))

            # Draw temperature line
            draw.rectangle([13, 6, 15, 20], fill=(255, 255, 255, 255))

            # Save as ICO
            img.save('icon.ico', format='ICO', sizes=[(32, 32)])
            print("✓ Created icon.ico")
        except ImportError:
            print("⚠ PIL not available, skipping icon creation")
            # Remove icon reference from spec file
            with open('temperature_monitor.spec', 'r') as f:
                content = f.read()
            content = content.replace("icon='icon.ico'", "# icon='icon.ico'")
            with open('temperature_monitor.spec', 'w') as f:
                f.write(content)

def build_executable():
    """Build the executable using PyInstaller"""
    print("Building executable...")

    try:
        # Run PyInstaller
        cmd = [sys.executable, '-m', 'PyInstaller', '--clean', 'temperature_monitor.spec']
        subprocess.check_call(cmd)
        print("✓ Executable built successfully!")
        return True
    except subprocess.CalledProcessError as e:
        print(f"✗ Build failed: {e}")
        return False

def create_installer():
    """Create a smart installer that adapts based on admin privileges"""
    print("Creating smart installer script...")

    installer_script = '''@echo off
echo Installing Temperature Monitor...
echo.

REM Check if we're running as administrator
net session >nul 2>&1
if %errorLevel% == 0 (
    echo Running with administrator privileges...
    echo Attempting system-wide installation to Program Files...
    echo.
    goto :ADMIN_INSTALL
) else (
    echo Running without administrator privileges.
    echo This installer will install to your user directory instead.
    echo.
    echo If you want system-wide installation, please:
    echo 1. Right-click this installer
    echo 2. Select "Run as administrator"
    echo.
    echo Continuing with user installation...
    echo.
    call install_for_user.bat
    exit /b %errorlevel%
)

:ADMIN_INSTALL
REM Try to create installation directory in Program Files
set "INSTALL_DIR=%PROGRAMFILES%\\TemperatureMonitor"
echo Attempting to install to: %INSTALL_DIR%

REM Try to create the directory
mkdir "%INSTALL_DIR%" 2>nul
if errorlevel 1 (
    echo Warning: Cannot create directory in Program Files.
    echo This might be due to system restrictions.
    echo.
    echo Falling back to user installation...
    echo.
    call install_for_user.bat
    exit /b %errorlevel%
) else (
    echo Directory created successfully in Program Files.
)

REM Copy all files from current directory
echo Copying files...
echo Current directory: %CD%
echo Target directory: %INSTALL_DIR%

REM First, try to copy all files except the executable
echo Copying configuration and language files...
xcopy "config.ini" "%INSTALL_DIR%\\" /Y
xcopy "uninstall.bat" "%INSTALL_DIR%\\" /Y
xcopy "lang\\*" "%INSTALL_DIR%\\lang\\" /E /I /Y

REM Try to copy the executable with retry mechanism
echo Copying executable...
set "RETRY_COUNT=0"
:RETRY_COPY
echo Attempt %RETRY_COUNT%: Copying TemperatureMonitor.exe...
xcopy "TemperatureMonitor.exe" "%INSTALL_DIR%\\" /Y
if errorlevel 1 (
    set /a RETRY_COUNT+=1
    if %RETRY_COUNT% LSS 3 (
        echo Sharing violation detected. Retrying in 2 seconds...
        timeout /t 2 /nobreak >nul
        goto RETRY_COPY
    ) else (
        echo.
        echo Error: Cannot copy TemperatureMonitor.exe after 3 attempts
        echo.
        echo Possible causes:
        echo 1. Temperature Monitor is currently running
        echo 2. Antivirus software is blocking the file
        echo 3. File permissions issue
        echo 4. File is locked by another process
        echo.
        echo Falling back to user installation...
        echo.
        call install_for_user.bat
        exit /b %errorlevel%
    )
)

echo All files copied successfully!

REM Create desktop shortcut
set "DESKTOP=%USERPROFILE%\\Desktop"
echo Set oWS = WScript.CreateObject("WScript.Shell") > CreateShortcut.vbs
echo sLinkFile = "%DESKTOP%\\Temperature Monitor.lnk" >> CreateShortcut.vbs
echo Set oLink = oWS.CreateShortcut(sLinkFile) >> CreateShortcut.vbs
echo oLink.TargetPath = "%INSTALL_DIR%\\TemperatureMonitor.exe" >> CreateShortcut.vbs
echo oLink.WorkingDirectory = "%INSTALL_DIR%" >> CreateShortcut.vbs
echo oLink.Description = "Arduino Temperature Monitor" >> CreateShortcut.vbs
echo oLink.Save >> CreateShortcut.vbs
cscript CreateShortcut.vbs
del CreateShortcut.vbs

REM Create Start Menu shortcut
set "START_MENU=%APPDATA%\\Microsoft\\Windows\\Start Menu\\Programs"
if not exist "%START_MENU%\\Temperature Monitor" mkdir "%START_MENU%\\Temperature Monitor"
echo Set oWS = WScript.CreateObject("WScript.Shell") > CreateStartMenuShortcut.vbs
echo sLinkFile = "%START_MENU%\\Temperature Monitor\\Temperature Monitor.lnk" >> CreateStartMenuShortcut.vbs
echo Set oLink = oWS.CreateShortcut(sLinkFile) >> CreateStartMenuShortcut.vbs
echo oLink.TargetPath = "%INSTALL_DIR%\\TemperatureMonitor.exe" >> CreateStartMenuShortcut.vbs
echo oLink.WorkingDirectory = "%INSTALL_DIR%" >> CreateStartMenuShortcut.vbs
echo oLink.Description = "Arduino Temperature Monitor" >> CreateStartMenuShortcut.vbs
echo oLink.Save >> CreateStartMenuShortcut.vbs
cscript CreateStartMenuShortcut.vbs
del CreateStartMenuShortcut.vbs

echo.
echo Installation completed successfully!
echo Temperature Monitor has been installed to: %INSTALL_DIR%
echo Desktop shortcut created.
echo Start Menu shortcut created.
echo.
echo You can now run Temperature Monitor from the desktop shortcut or Start Menu.
echo.
pause
'''

    with open('dist/install.bat', 'w') as f:
        f.write(installer_script)

    print("✓ Created dist/install.bat")

def create_uninstaller():
    """Create an improved uninstaller script that works with and without admin privileges"""
    print("Creating uninstaller script...")

    uninstaller_script = '''@echo off
echo Uninstalling Temperature Monitor...
echo.

REM Check if we're running as administrator
net session >nul 2>&1
if %errorLevel% == 0 (
    echo Running with administrator privileges...
    set "ADMIN_MODE=1"
) else (
    echo Running without administrator privileges...
    set "ADMIN_MODE=0"
)

REM Set installation directories
set "PROGRAM_FILES_DIR=%PROGRAMFILES%\\TemperatureMonitor"
set "USER_DIR=%USERPROFILE%\\TemperatureMonitor"

REM Check where the application is installed
set "INSTALL_DIR="
if exist "%PROGRAM_FILES_DIR%" (
    set "INSTALL_DIR=%PROGRAM_FILES_DIR%"
    echo Found installation in Program Files: %INSTALL_DIR%
) else if exist "%USER_DIR%" (
    set "INSTALL_DIR=%USER_DIR%"
    echo Found installation in user directory: %INSTALL_DIR%
) else (
    echo Temperature Monitor is not installed.
    pause
    exit /b 0
)

REM Check if we can access the installation directory
if "%ADMIN_MODE%"=="0" (
    if "%INSTALL_DIR%"=="%PROGRAM_FILES_DIR%" (
        echo.
        echo Error: Application is installed in Program Files but you don't have administrator privileges.
        echo.
        echo Please either:
        echo 1. Right-click this uninstaller and select "Run as administrator"
        echo 2. Or manually delete the folder: %INSTALL_DIR%
        echo.
        pause
        exit /b 1
    )
)

REM Remove desktop shortcut
set "DESKTOP=%USERPROFILE%\\Desktop"
if exist "%DESKTOP%\\Temperature Monitor.lnk" (
    del "%DESKTOP%\\Temperature Monitor.lnk"
    echo Desktop shortcut removed.
)

REM Remove Start Menu shortcut
set "START_MENU=%APPDATA%\\Microsoft\\Windows\\Start Menu\\Programs\\Temperature Monitor"
if exist "%START_MENU%" (
    rmdir /s /q "%START_MENU%"
    echo Start Menu shortcut removed.
)

REM Remove installation directory
if exist "%INSTALL_DIR%" (
    echo Removing installation directory: %INSTALL_DIR%
    rmdir /s /q "%INSTALL_DIR%"
    if exist "%INSTALL_DIR%" (
        echo Warning: Could not completely remove installation directory.
        echo You may need to manually delete: %INSTALL_DIR%
    ) else (
        echo Installation directory removed successfully.
    )
)

echo.
echo Temperature Monitor has been uninstalled successfully!
echo.
pause
'''

    with open('dist/uninstall.bat', 'w') as f:
        f.write(uninstaller_script)

    print("✓ Created dist/uninstall.bat")

def create_user_installer():
    """Create a user installer that doesn't require admin privileges"""
    print("Creating user installer script...")

    user_installer_script = '''@echo off
echo Installing Temperature Monitor (User Installation)...
echo.

REM Install to user's home directory (no admin privileges needed)
set "INSTALL_DIR=%USERPROFILE%\\TemperatureMonitor"
echo Installing to: %INSTALL_DIR%

REM Create installation directory
if not exist "%INSTALL_DIR%" mkdir "%INSTALL_DIR%"

REM Copy all files from current directory
echo Copying files...

REM First, try to copy all files except the executable
echo Copying configuration and language files...
xcopy "config.ini" "%INSTALL_DIR%\\" /Y >nul 2>&1
xcopy "uninstall.bat" "%INSTALL_DIR%\\" /Y >nul 2>&1
xcopy "lang\\*" "%INSTALL_DIR%\\lang\\" /E /I /Y >nul 2>&1

REM Try to copy the executable with retry mechanism
echo Copying executable...
set "RETRY_COUNT=0"
:RETRY_COPY
xcopy "TemperatureMonitor.exe" "%INSTALL_DIR%\\" /Y >nul 2>&1
if errorlevel 1 (
    set /a RETRY_COUNT+=1
    if %RETRY_COUNT% LSS 3 (
        echo Sharing violation detected. Retrying in 2 seconds...
        timeout /t 2 /nobreak >nul
        goto RETRY_COPY
    ) else (
        echo.
        echo Error: Cannot copy TemperatureMonitor.exe
        echo This usually happens when the application is currently running.
        echo.
        echo Please:
        echo 1. Close Temperature Monitor if it's running
        echo 2. Wait a few seconds
        echo 3. Run this installer again
        echo.
        pause
        exit /b 1
    )
)

echo All files copied successfully!

REM Create desktop shortcut
set "DESKTOP=%USERPROFILE%\\Desktop"
echo Set oWS = WScript.CreateObject("WScript.Shell") > CreateShortcut.vbs
echo sLinkFile = "%DESKTOP%\\Temperature Monitor.lnk" >> CreateShortcut.vbs
echo Set oLink = oWS.CreateShortcut(sLinkFile) >> CreateShortcut.vbs
echo oLink.TargetPath = "%INSTALL_DIR%\\TemperatureMonitor.exe" >> CreateShortcut.vbs
echo oLink.WorkingDirectory = "%INSTALL_DIR%" >> CreateShortcut.vbs
echo oLink.Description = "Arduino Temperature Monitor" >> CreateShortcut.vbs
echo oLink.Save >> CreateShortcut.vbs
cscript CreateShortcut.vbs
del CreateShortcut.vbs

REM Create Start Menu shortcut
set "START_MENU=%APPDATA%\\Microsoft\\Windows\\Start Menu\\Programs"
if not exist "%START_MENU%\\Temperature Monitor" mkdir "%START_MENU%\\Temperature Monitor"
echo Set oWS = WScript.CreateObject("WScript.Shell") > CreateStartMenuShortcut.vbs
echo sLinkFile = "%START_MENU%\\Temperature Monitor\\Temperature Monitor.lnk" >> CreateStartMenuShortcut.vbs
echo Set oLink = oWS.CreateShortcut(sLinkFile) >> CreateStartMenuShortcut.vbs
echo oLink.TargetPath = "%INSTALL_DIR%\\TemperatureMonitor.exe" >> CreateStartMenuShortcut.vbs
echo oLink.WorkingDirectory = "%INSTALL_DIR%" >> CreateStartMenuShortcut.vbs
echo oLink.Description = "Arduino Temperature Monitor" >> CreateStartMenuShortcut.vbs
echo oLink.Save >> CreateStartMenuShortcut.vbs
cscript CreateStartMenuShortcut.vbs
del CreateStartMenuShortcut.vbs

echo.
echo Installation completed successfully!
echo Temperature Monitor has been installed to: %INSTALL_DIR%
echo Desktop shortcut created.
echo Start Menu shortcut created.
echo.
echo You can now run Temperature Monitor from the desktop shortcut or Start Menu.
echo.
pause
'''

    with open('dist/install_for_user.bat', 'w') as f:
        f.write(user_installer_script)

    print("✓ Created dist/install_for_user.bat")

def copy_lang_folder():
    """Copy lang folder to dist directory for distribution"""
    print("Copying language files...")

    try:
        lang_source = Path("lang")
        dist_lang = Path("dist/lang")

        if not lang_source.exists():
            print("⚠ lang folder not found, skipping language files copy")
            return True

        # Create dist/lang directory if it doesn't exist
        dist_lang.mkdir(parents=True, exist_ok=True)

        # Copy all files from lang to dist/lang
        for file_path in lang_source.rglob("*"):
            if file_path.is_file():
                # Calculate relative path from lang source
                rel_path = file_path.relative_to(lang_source)
                dest_path = dist_lang / rel_path

                # Create parent directories if needed
                dest_path.parent.mkdir(parents=True, exist_ok=True)

                # Copy file
                shutil.copy2(file_path, dest_path)
                print(f"  Copied: {rel_path}")

        print("✓ Language files copied to dist/lang")
        return True

    except Exception as e:
        print(f"✗ Failed to copy language files: {e}")
        return False

def create_zip_package():
    """Create a ZIP package with all distribution files"""
    print("Creating ZIP package...")

    try:
        # Create ZIP file
        zip_path = "TemperatureMonitor.zip"
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            # Add all files from dist folder
            dist_path = Path("dist")
            if dist_path.exists():
                for file_path in dist_path.rglob("*"):
                    if file_path.is_file():
                        # Add file to ZIP with relative path
                        arcname = file_path.relative_to(dist_path)
                        zipf.write(file_path, arcname)
                        print(f"  Added: {arcname}")

            # Add README and Arduino code
            if Path("README.md").exists():
                zipf.write("README.md", "README.md")
                print("  Added: README.md")

            if Path("temperature_monitor.ino").exists():
                zipf.write("temperature_monitor.ino", "temperature_monitor.ino")
                print("  Added: temperature_monitor.ino")

            # Add language configuration utility
            if Path("language_config.py").exists():
                zipf.write("language_config.py", "language_config.py")
                print("  Added: language_config.py")

        print(f"✓ Created {zip_path}")
        return True

    except Exception as e:
        print(f"✗ Failed to create ZIP: {e}")
        return False

def main():
    """Main build process"""
    print("=== Temperature Monitor Build Script ===")
    print()

    # Check if we're in the right directory
    if not Path('temperature_monitor.py').exists():
        print("Error: temperature_monitor.py not found in current directory")
        print("Please run this script from the project root folder")
        return False

    # Install requirements
    if not install_requirements():
        print("Failed to install requirements")
        return False

    # Create icon
    create_icon()

    # Build executable
    if not build_executable():
        print("Failed to build executable")
        return False

    # Copy language files to dist
    if not copy_lang_folder():
        print("Warning: Language files copy failed, but build was successful")

    # Create installer and uninstaller
    create_installer()
    create_uninstaller()
    create_user_installer()

    # Create ZIP package
    if not create_zip_package():
        print("Warning: ZIP creation failed, but build was successful")

    print()
    print("=== Build Complete ===")
    print("Files created:")
    print("  - dist/TemperatureMonitor.exe (main executable)")
    print("  - dist/lang/ (language files)")
    print("  - dist/install.bat (smart installer - adapts to admin privileges)")
    print("  - dist/install_for_user.bat (user installer - no admin needed)")
    print("  - dist/uninstall.bat (uninstaller script)")
    print("  - TemperatureMonitor.zip (complete distribution package)")
    print()
    print("Distribution package includes:")
    print("  - TemperatureMonitor.exe")
    print("  - lang/ (language files for localization)")
    print("  - install.bat (smart installer)")
    print("  - install_for_user.bat (user installer)")
    print("  - uninstall.bat")
    print("  - README.md")
    print("  - temperature_monitor.ino (Arduino code)")
    print()
    print("Installation options:")
    print("  1. install.bat - Smart installer (adapts to privileges)")
    print("  2. install_for_user.bat - User installation (no admin needed)")
    print()
    print("To distribute:")
    print("  1. Share TemperatureMonitor.zip")
    print("  2. Recipients extract and run install.bat")
    print("     - Without admin: Installs to user directory")
    print("     - With admin: Installs to Program Files")

    return True

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)
