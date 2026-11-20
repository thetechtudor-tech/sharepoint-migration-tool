@echo off
rem Builds SharePointMigration.exe from sp_migrate_ui.py and sp_migrate.py.
rem Run this on Windows, in the folder that holds both .py files.
rem Needs Python installed (python.org, with "Add python.exe to PATH" ticked).

cd /d "%~dp0"

if not exist sp_migrate_ui.py goto :missing
if not exist sp_migrate.py goto :missing

echo Installing build tools and packages...
python -m pip install --upgrade pyinstaller msal requests openpyxl paramiko || goto :failed

echo.
echo Building the program (this takes a minute or two)...
python -m PyInstaller --onefile --clean --noconfirm --name SharePointMigration sp_migrate_ui.py || goto :failed

echo.
echo Done. Your program is:  %~dp0dist\SharePointMigration.exe
pause
exit /b 0

:missing
echo Put this file in the same folder as sp_migrate_ui.py and sp_migrate.py, then run it again.
pause
exit /b 1

:failed
echo.
echo The build failed. The messages above say why.
pause
exit /b 1
