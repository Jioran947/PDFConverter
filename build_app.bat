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
    echo Python was not found. Build cannot continue.
    pause
    exit /b 1
)

%PYTHON_CMD% -m pip install -r requirements.txt pyinstaller
if errorlevel 1 (
    echo Failed to install build dependencies.
    pause
    exit /b 1
)

%PYTHON_CMD% -m PyInstaller --noconfirm --clean --onefile --windowed --name PDFConverter --manifest app.manifest --collect-data tkinterdnd2 --hidden-import=fitz --hidden-import=win32com --hidden-import=win32com.client --hidden-import=pythoncom --hidden-import=pywintypes pdf_converter.py
if errorlevel 1 (
    echo Build failed.
    pause
    exit /b 1
)

echo.
echo Build complete: dist\PDFConverter.exe
pause
