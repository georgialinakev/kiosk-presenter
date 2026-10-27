@echo off
rem Builds MuseumKiosk-v<VERSION>.exe straight into this folder (next to Presentations).
rem VERSION is read from museum_kiosk.py. Uses the project's .venv if it exists,
rem otherwise the py launcher.
cd /d "%~dp0"

set "VER="
for /f tokens^=2^ delims^=^" %%v in ('findstr /b /c:"VERSION = " museum_kiosk.py') do set "VER=%%v"
if not defined VER (
    echo Could not read VERSION from museum_kiosk.py
    goto :fail
)
set "NAME=MuseumKiosk-v%VER%"

set "PY=py"
if exist ".venv\Scripts\python.exe" set "PY=.venv\Scripts\python.exe"

"%PY%" -m pip install -r requirements.txt || goto :fail
"%PY%" -m PyInstaller --noconfirm --onefile --noconsole --name "%NAME%" ^
    --distpath . --workpath build --specpath build museum_kiosk.py || goto :fail

echo.
echo Built: %~dp0%NAME%.exe
exit /b 0

:fail
echo.
echo Build failed.
exit /b 1