@echo off
setlocal
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
  echo [ERROR] Missing .venv. Create it and install requirements.txt first.
  exit /b 1
)

.venv\Scripts\python.exe -m PyInstaller --noconfirm --clean report_generator.spec
if errorlevel 1 exit /b %errorlevel%

if not exist "dist\报告生成工具\templates" mkdir "dist\报告生成工具\templates"
copy /Y "templates\模板放置说明.txt" "dist\报告生成工具\templates\模板放置说明.txt" >nul
for %%F in ("templates\*.docx") do if exist "%%~fF" copy /Y "%%~fF" "dist\报告生成工具\templates\" >nul
copy /Y "使用说明.txt" "dist\报告生成工具\使用说明.txt" >nul

echo.
echo Build complete: dist\报告生成工具\报告生成工具.exe
