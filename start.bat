@echo off
setlocal
cd /d "%~dp0"

set "PYTHON_CMD="

py -3 --version >nul 2>nul
if not errorlevel 1 set "PYTHON_CMD=py -3"

if "%PYTHON_CMD%"=="" (
    python --version >nul 2>nul
    if not errorlevel 1 set "PYTHON_CMD=python"
)

if "%PYTHON_CMD%"=="" (
    set "BUNDLED_PY=%USERPROFILE%\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
    if exist "%BUNDLED_PY%" set "PYTHON_CMD=%BUNDLED_PY%"
)

if "%PYTHON_CMD%"=="" (
    echo Python was not found.
    echo Run install.bat after installing Python 3.10 or newer.
    echo Download: https://www.python.org/downloads/windows/
    echo.
    pause
    exit /b 1
)

%PYTHON_CMD% -c "import PIL, reportlab" >nul 2>nul
if errorlevel 1 (
    echo Required packages are missing.
    echo Please run install.bat first.
    echo.
    pause
    exit /b 1
)

%PYTHON_CMD% pdf_converter.py
exit /b %errorlevel%
