@echo off
setlocal

if not exist ".venv\Scripts\python.exe" (
    py -3.12 -m venv .venv
    if errorlevel 1 exit /b 1
)

.venv\Scripts\python.exe -m pip install -r requirements.txt
if errorlevel 1 exit /b 1

.venv\Scripts\python.exe -m unittest discover -s tests -v
if errorlevel 1 exit /b 1

.venv\Scripts\pyinstaller.exe --clean --noconfirm WordTextReplacer.spec
if errorlevel 1 exit /b 1

echo.
echo Build completed: dist\WordTextReplacer.exe
endlocal
