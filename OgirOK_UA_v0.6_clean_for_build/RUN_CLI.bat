@echo off
setlocal

if not exist .venv (
  py -m venv .venv
)

call .venv\Scripts\activate.bat
python -m pip install --upgrade pip
python -m pip install -r requirements.txt

REM --- Example: process input\documents and export Excel + JSON
py run_parser_orders_cli.py --input "input\documents" --excel --json --txt

endlocal
