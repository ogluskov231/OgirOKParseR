@echo off
setlocal

REM --- Create venv if missing
if not exist .venv (
  py -m venv .venv
)

call .venv\Scripts\activate.bat

python -m pip install --upgrade pip
python -m pip install -r requirements.txt

REM --- Run GUI
py apps\parser_orders\src\gui\parser_gui.py

endlocal
