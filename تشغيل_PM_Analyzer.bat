@echo off
chcp 65001 > nul
cd /d "%~dp0"
set "PYTHONPATH=%~dp0src"
python -m pm_analyzer.gui
if errorlevel 1 (
  echo.
  echo تعذر تشغيل البرنامج. تأكد من تثبيت Python 3.11 أو أحدث وتفعيل Add Python to PATH.
  pause
)
