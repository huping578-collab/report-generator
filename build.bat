@echo off
setlocal
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
  echo [ERROR] Missing .venv. Create it and install requirements.txt first.
  exit /b 1
)

.venv\Scripts\python.exe -m PyInstaller --noconfirm --clean report_generator.spec
if errorlevel 1 exit /b %errorlevel%

.venv\Scripts\python.exe build-tools\stage_files.py
if errorlevel 1 exit /b %errorlevel%

echo.
echo Build complete: dist\报告生成工具\报告生成工具.exe
