@echo off
setlocal
cd /d "%~dp0" || exit /b 1
if not exist ".venv\Scripts\python.exe" (
    py -3.12 -m venv .venv || exit /b 1
)
".venv\Scripts\python.exe" -m pip install -r requirements-build.txt || exit /b 1
".venv\Scripts\python.exe" -m PyInstaller --noconfirm --clean GroceryTracker.spec || exit /b 1
echo.
echo Built "%CD%\dist\GroceryTracker.exe"
