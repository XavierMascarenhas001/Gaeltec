@echo off
title Gaeltec Tools server
cd /d "%~dp0"

REM ---------------------------------------------------------------
REM  Find Python: Anaconda first (it is often NOT on the PATH),
REM  then whatever "python" / "py" the PATH gives us.
REM ---------------------------------------------------------------
set "RUN="
for %%P in (
  "%LOCALAPPDATA%\anaconda3\python.exe"
  "%USERPROFILE%\anaconda3\python.exe"
  "C:\ProgramData\anaconda3\python.exe"
  "%LOCALAPPDATA%\miniconda3\python.exe"
  "%USERPROFILE%\miniconda3\python.exe"
) do (
  if not defined RUN if exist "%%~P" set "RUN="%%~P""
)
if not defined RUN python -c "import sys" >nul 2>nul && set "RUN=python"
if not defined RUN py -3 -c "import sys" >nul 2>nul && set "RUN=py -3"
if not defined RUN goto nopython

echo Using Python: %RUN%
echo Checking the required packages - the first run can take a few minutes...
%RUN% -m pip install -q --disable-pip-version-check -r requirements.txt
if errorlevel 1 echo Some packages could not be installed - see above. Trying to start anyway...

%RUN% app.py

echo.
echo  The server has stopped. If there is an error above, send a screenshot of it.
pause
exit /b 0

:nopython
echo.
echo  Could not find Python or Anaconda on this PC.
echo  Open an Anaconda Prompt, go to this folder and run:   python app.py
echo.
pause
exit /b 1
