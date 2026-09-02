@echo off
setlocal EnableExtensions
set PYTHONIOENCODING=utf-8
cd /d "%~dp0"

set "PY=.venv\Scripts\python.exe"
if not exist "%PY%" (
  echo [ERROR] Missing .venv. Create it and install requirements.txt first.
  exit /b 1
)
if not exist "version.txt" (
  echo [ERROR] Missing version.txt.
  exit /b 1
)
set /p VERSION=<version.txt
"%PY%" -c "import re,sys; v=open('version.txt',encoding='utf-8').read().strip(); sys.exit(0 if re.fullmatch(r'[0-9]+\.[0-9]+\.[0-9]+',v) else 1)"
if errorlevel 1 (
  echo [ERROR] version.txt must contain X.Y.Z.
  exit /b 1
)

call build.bat
if errorlevel 1 exit /b %errorlevel%

if exist "dist\updater" rmdir /s /q "dist\updater"
if exist "build\updater" rmdir /s /q "build\updater"
"%PY%" -m PyInstaller --noconfirm --clean --onefile --windowed ^
  --name updater --icon "%CD%\assets\app.ico" ^
  --add-data "%CD%\version.txt;." ^
  --distpath "dist\updater" --workpath "build\updater" ^
  --specpath "build\updater" updater.py
if errorlevel 1 exit /b %errorlevel%

set "ISCC=%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe"
if not exist "%ISCC%" set "ISCC=%ProgramFiles%\Inno Setup 6\ISCC.exe"
if not exist "%ISCC%" set "ISCC=%LocalAppData%\Programs\Inno Setup 6\ISCC.exe"
if not exist "%ISCC%" set "ISCC="
if not exist "%ISCC%" (
  for /f "delims=" %%I in ('where ISCC.exe 2^>nul') do if not defined ISCC set "ISCC=%%I"
)
if not defined ISCC (
  echo [ERROR] Inno Setup 6 ISCC.exe not found.
  exit /b 1
)

"%ISCC%" /DAppVersion=%VERSION% installer.iss
if errorlevel 1 exit /b %errorlevel%

"%PY%" build-tools\verify_release.py
if errorlevel 1 exit /b %errorlevel%

echo.
echo Release build complete:
echo   dist\报告生成工具\报告生成工具.exe
echo   dist\updater\updater.exe
echo   dist\报告生成工具-Setup.exe
