@echo off
REM ОгрОК_UA — збірка Windows .exe через PyInstaller
cd /d "%~dp0\.."
setlocal

if not exist .venv (
  python -m venv .venv
)

call .venv\Scripts\activate

python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python -m pip install pyinstaller

REM Clean old builds
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

REM Build (onedir) + bundle data folders рядом з exe
pyinstaller ^
  --noconsole ^
  --onedir ^
  --name OgrOK_UA ^
  --icon assets\icon.ico ^
  --add-data "templates;templates" ^
  --add-data "config;config" ^
  --add-data "assets;assets" ^
  --add-data "input;input" ^
  apps\parser_orders\src\gui\parser_gui.py

echo.
echo DONE. Реліз: dist\OgrOK_UA\OgrOK_UA.exe
pause
