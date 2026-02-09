@echo off
setlocal
cd /d "%~dp0"

echo Starting Word upload demo...
python word_upload_demo.py --open-browser

if errorlevel 1 (
  echo.
  echo Failed to start with "python". Trying "py -3"...
  py -3 word_upload_demo.py --open-browser
)

pause
