@echo off
setlocal
cd /d "%~dp0"
if not exist ".build-env\Scripts\python.exe" (
    where py >nul 2>&1
    if errorlevel 1 (
        python -m venv ".build-env"
    ) else (
        py -3 -m venv ".build-env"
    )
    if errorlevel 1 goto failed
)
".build-env\Scripts\python.exe" -m pip install "pyinstaller>=6,<7"
if errorlevel 1 goto failed
".build-env\Scripts\python.exe" -m PyInstaller --noconfirm --clean --onefile --windowed --name Senura_Bill_Generator --add-data "engine.ps1:." --add-data "template.xlsx:." Bill_Generator.pyw
if errorlevel 1 goto failed
echo Created "%CD%\dist\Senura_Bill_Generator.exe"
start "" explorer.exe "%CD%\dist"
pause
exit /b 0
:failed
echo Build failed. See the message above.
pause
exit /b 1
