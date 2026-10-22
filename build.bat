@echo off
rem Builds MuseumKiosk.exe straight into this folder (next to Presentations).
rem Uses the project's .venv if it exists, otherwise the py launcher.
cd /d "%~dp0"

set "PY=py"
if exist ".venv\Scripts\python.exe" set "PY=.venv\Scripts\python.exe"

"%PY%" -m pip install -r requirements.txt || goto :fail
"%PY%" -m PyInstaller --noconfirm --onefile --noconsole --name MuseumKiosk ^
    --distpath . --workpath build --specpath build museum_kiosk.py || goto :fail

echo.
echo Built: %~dp0MuseumKiosk.exe
exit /b 0

:fail
echo.
echo Build failed.
exit /b 1
