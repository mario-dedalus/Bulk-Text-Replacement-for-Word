@echo off
setlocal

REM =============================================================
REM  build_portable.bat  —  One-click portable build script
REM  Outputs: dist\WordTextReplacer\       (portable folder)
REM           dist\WordTextReplacer_portable.zip  (release zip)
REM
REM  Usage:
REM    build_portable.bat
REM    build_portable.bat --version 2.0.0
REM =============================================================

set "APP_NAME=WordTextReplacer"
set "VERSION="

REM ------ Parse optional --version argument ------
:parse_args
if "%~1"=="" goto :args_done
if /i "%~1"=="--version" (
    set "VERSION=%~2"
    shift
    shift
    goto :parse_args
)
shift
goto :parse_args
:args_done

REM ------ Find Python ------
echo [1/5] Checking virtual environment...
set "PYTHON_CMD="
where py >nul 2>&1 && set "PYTHON_CMD=py -3"
if not defined PYTHON_CMD (
    where python3 >nul 2>&1 && set "PYTHON_CMD=python3"
)
if not defined PYTHON_CMD (
    where python >nul 2>&1 && set "PYTHON_CMD=python"
)
if not defined PYTHON_CMD (
    echo ERROR: Python not found. Please install Python 3 and add it to PATH.
    echo        Download from https://www.python.org/downloads/
    echo        Make sure to check "Add Python to PATH" during installation.
    pause
    exit /b 1
)
echo      Found Python: %PYTHON_CMD%

REM ------ Set up venv if needed ------
if not exist ".venv\Scripts\python.exe" (
    echo      Creating .venv...
    %PYTHON_CMD% -m venv .venv
    if errorlevel 1 (
        echo ERROR: Failed to create virtual environment.
        pause
        exit /b 1
    )
)

REM ------ Install dependencies ------
echo [2/5] Installing dependencies...
.venv\Scripts\python.exe -m pip install --quiet -r requirements.txt
if errorlevel 1 (
    echo ERROR: pip install failed.
    pause
    exit /b 1
)

REM ------ Run tests ------
echo [3/5] Running tests...
.venv\Scripts\python.exe -m unittest discover -s tests -v
if errorlevel 1 (
    echo ERROR: Tests failed. Aborting build.
    pause
    exit /b 1
)

REM ------ Build portable (onedir) ------
echo [4/5] Building portable distribution...
.venv\Scripts\pyinstaller.exe --clean --noconfirm WordTextReplacer_portable.spec
if errorlevel 1 (
    echo ERROR: PyInstaller build failed.
    pause
    exit /b 1
)

REM ------ Package into ZIP ------
echo [5/5] Creating portable ZIP archive...

if defined VERSION (
    set "ZIP_NAME=%APP_NAME%_v%VERSION%_portable.zip"
) else (
    set "ZIP_NAME=%APP_NAME%_portable.zip"
)

REM Remove old zip if it exists
if exist "dist\%ZIP_NAME%" del "dist\%ZIP_NAME%"

REM Use Python zipfile (shared read, no file-lock issues unlike PowerShell)
.venv\Scripts\python.exe scripts\zip_portable.py "dist\%APP_NAME%" "dist\%ZIP_NAME%"
if errorlevel 1 (
    echo WARNING: ZIP creation failed. The portable folder is still available.
    echo          You can manually zip dist\%APP_NAME%\
)

echo.
echo ============================================================
echo  Build completed successfully!
echo.
echo  Portable folder : dist\%APP_NAME%\
echo    Entry point   : dist\%APP_NAME%\%APP_NAME%.exe
if exist "dist\%ZIP_NAME%" (
    echo  Release archive : dist\%ZIP_NAME%
)
echo.
echo  The entire folder is self-contained and portable.
echo  Users can run %APP_NAME%.exe directly — no installation needed.
echo ============================================================
echo.
pause
endlocal
