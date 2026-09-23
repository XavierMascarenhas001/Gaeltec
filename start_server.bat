@echo off
REM Starts the Gaeltec Tools web server. Leave this window open while people use the page.
cd /d "%~dp0"
python -m pip install -q -r requirements.txt
python app.py
pause
