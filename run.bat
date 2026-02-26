@echo off
setlocal EnableDelayedExpansion

:: =====================================================================
::  DNDC Calibration Studio
::  Double-click this file to run the calibration tool.
::  First run: installs Python (if needed), creates venv, installs deps.
::  Subsequent runs: skips straight to launch.
:: =====================================================================

title DNDC Calibration Studio
color 0B

echo.
echo  ============================================================
echo    DNDC Calibration Studio
echo  ============================================================
echo.

:: ── Configuration ──
set "SCRIPT_NAME=caln.py"
set "VENV_DIR=%~dp0venv"
set "DEPS_MARKER=%~dp0.deps_installed"
set "PYTHON_VERSION=3.12.8"
set "PYTHON_INSTALLER=python-%PYTHON_VERSION%-amd64.exe"
set "PYTHON_URL=https://www.python.org/ftp/python/%PYTHON_VERSION%/%PYTHON_INSTALLER%"
set "REQUIRED_PACKAGES=numpy pandas scikit-learn scikit-optimize openpyxl portalocker"

:: ── Step 0: Check that the Python script exists ──
echo  [1/5]  Checking for %SCRIPT_NAME%...
if not exist "%~dp0%SCRIPT_NAME%" (
    echo.
    echo  ERROR: %SCRIPT_NAME% not found!
    echo  Please place this launcher in the same folder as the Python script.
    echo.
    pause
    exit /b 1
)
echo         Found.
echo.

:: ── Step 1: Check if Python is available ──
echo  [2/5]  Checking for Python installation...

:: Try multiple ways to find Python
set "PYTHON_CMD="

:: Check venv first (most reliable for repeat runs)
if exist "%VENV_DIR%\Scripts\python.exe" (
    set "PYTHON_CMD=%VENV_DIR%\Scripts\python.exe"
    echo         Found Python in virtual environment.
    goto :check_deps
)

:: Check system Python via 'py' launcher (standard Windows installer)
py --version >nul 2>&1
if %errorlevel% equ 0 (
    set "PYTHON_CMD=py -3"
    echo         Found Python via 'py' launcher.
    goto :have_python
)

:: Check system Python via 'python' command
python --version >nul 2>&1
if %errorlevel% equ 0 (
    :: Make sure it's Python 3, not a Windows Store stub
    for /f "tokens=2 delims= " %%v in ('python --version 2^>^&1') do set "PY_VER=%%v"
    echo !PY_VER! | findstr /b "3." >nul 2>&1
    if !errorlevel! equ 0 (
        set "PYTHON_CMD=python"
        echo         Found Python !PY_VER!.
        goto :have_python
    )
)

:: Check common install locations
for %%p in (
    "%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
    "%LOCALAPPDATA%\Programs\Python\Python311\python.exe"
    "%LOCALAPPDATA%\Programs\Python\Python310\python.exe"
    "C:\Python312\python.exe"
    "C:\Python311\python.exe"
    "C:\Python310\python.exe"
) do (
    if exist %%p (
        set "PYTHON_CMD=%%~p"
        echo         Found Python at %%~p
        goto :have_python
    )
)

:: ── Python not found — offer to install ──
echo.
echo  ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
echo    Python is not installed on this computer.
echo    Python 3.12 is required to run the calibration tool.
echo  ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
echo.
echo    The installer will be downloaded from python.org (~25 MB)
echo    and installed for the current user only (no admin needed).
echo.
set /p "INSTALL_CHOICE=  Install Python now? (Y/N): "
if /i not "!INSTALL_CHOICE!"=="Y" (
    echo.
    echo  To install manually, visit: https://www.python.org/downloads/
    echo  Then re-run this launcher.
    echo.
    pause
    exit /b 1
)

echo.
echo  [*]  Downloading Python %PYTHON_VERSION%...
echo       URL: %PYTHON_URL%
echo       This may take a minute depending on your connection.
echo.

:: Download using curl (built into Windows 10+) or PowerShell fallback
set "INSTALLER_PATH=%~dp0%PYTHON_INSTALLER%"

curl --version >nul 2>&1
if %errorlevel% equ 0 (
    curl -L -o "%INSTALLER_PATH%" "%PYTHON_URL%" --progress-bar
) else (
    echo       Using PowerShell to download...
    powershell -Command "& {[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; Invoke-WebRequest -Uri '%PYTHON_URL%' -OutFile '%INSTALLER_PATH%'}"
)

if not exist "%INSTALLER_PATH%" (
    echo.
    echo  ERROR: Download failed.
    echo  Please download Python manually from https://www.python.org/downloads/
    echo.
    pause
    exit /b 1
)

echo.
echo  [*]  Installing Python %PYTHON_VERSION%...
echo       Installing for current user (no admin required).
echo       Please wait — this takes about 1-2 minutes...
echo.

:: Silent install, per-user, add to PATH, include pip
"%INSTALLER_PATH%" /quiet InstallAllUsers=0 PrependPath=1 Include_pip=1 Include_launcher=1

if %errorlevel% neq 0 (
    echo.
    echo  ERROR: Python installation failed (error code: %errorlevel%).
    echo.
    echo  Possible causes:
    echo    - Antivirus blocked the installer
    echo    - Company policy prevents software installation
    echo    - Installer was corrupted during download
    echo.
    echo  Try installing Python manually from https://www.python.org/downloads/
    echo.
    pause
    exit /b 1
)

echo  [OK]  Python %PYTHON_VERSION% installed successfully.
echo.

:: Clean up installer
del "%INSTALLER_PATH%" 2>nul

:: Refresh PATH so we can find the new Python
set "PATH=%LOCALAPPDATA%\Programs\Python\Python312\;%LOCALAPPDATA%\Programs\Python\Python312\Scripts\;%PATH%"

:: Verify it works
py -3 --version >nul 2>&1
if %errorlevel% equ 0 (
    set "PYTHON_CMD=py -3"
) else (
    python --version >nul 2>&1
    if %errorlevel% equ 0 (
        set "PYTHON_CMD=python"
    ) else (
        echo  ERROR: Python was installed but cannot be found in PATH.
        echo  Please close this window, open a NEW terminal, and run this launcher again.
        echo.
        pause
        exit /b 1
    )
)

:have_python
echo.

:: ── Step 2: Create virtual environment ──
echo  [3/5]  Setting up virtual environment...

if exist "%VENV_DIR%\Scripts\python.exe" (
    echo         Virtual environment already exists.
) else (
    echo         Creating virtual environment in .\venv\ ...
    echo         This takes about 10-15 seconds...
    %PYTHON_CMD% -m venv "%VENV_DIR%"
    if !errorlevel! neq 0 (
        echo.
        echo  ERROR: Failed to create virtual environment.
        echo  Try deleting the 'venv' folder and running this launcher again.
        echo.
        pause
        exit /b 1
    )
    echo         Done.
    :: Clear deps marker so we reinstall into new venv
    del "%DEPS_MARKER%" 2>nul
)

set "PYTHON_CMD=%VENV_DIR%\Scripts\python.exe"
set "PIP_CMD=%VENV_DIR%\Scripts\pip.exe"
echo.

:check_deps
:: ── Step 3: Install dependencies ──
echo  [4/5]  Checking dependencies...

if exist "%DEPS_MARKER%" (
    echo         All dependencies already installed.
) else (
    echo         Installing required packages...
    echo         Packages: %REQUIRED_PACKAGES%
    echo.
    echo         This may take 2-5 minutes on first run.
    echo         You will see pip output below:
    echo  ------------------------------------------------------------

    "%PIP_CMD%" install --upgrade pip >nul 2>&1
    "%PIP_CMD%" install %REQUIRED_PACKAGES%

    if !errorlevel! neq 0 (
        echo  ------------------------------------------------------------
        echo.
        echo  ERROR: Some packages failed to install.
        echo  Check your internet connection and try again.
        echo  If the problem persists, run manually:
        echo    venv\Scripts\pip install %REQUIRED_PACKAGES%
        echo.
        pause
        exit /b 1
    )

    echo  ------------------------------------------------------------
    echo.
    echo         All packages installed successfully.

    :: Write marker file with timestamp
    echo Installed on %date% %time% > "%DEPS_MARKER%"
    echo Packages: %REQUIRED_PACKAGES% >> "%DEPS_MARKER%"
)
echo.

:: ── Step 4: Launch the application ──
echo  [5/5]  Launching DNDC Calibration Studio...
echo.
echo  ============================================================
echo    Starting application. This window will stay open for logs.
echo    Close this window to stop the application.
echo  ============================================================
echo.

"%PYTHON_CMD%" "%~dp0%SCRIPT_NAME%"

:: If we get here, the app has closed
echo.
echo  Application closed.
echo.
pause
