@echo off
cd /d "%~dp0"
set "MASK_PY=%USERPROFILE%\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
if exist "%MASK_PY%" (
  "%MASK_PY%" app.py
) else (
  python app.py
)
pause
