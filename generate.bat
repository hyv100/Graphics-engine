@echo off
setlocal
if "%~1"=="" (
  echo No calendar supplied. Generating the included sample calendar...
  python serene_engine.py --calendar content_calendar.csv --output output
) else (
  python serene_engine.py --calendar "%~1" --output output
)
echo.
echo Done. Open the output folder.
pause
