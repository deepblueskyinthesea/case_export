@echo off
setlocal
rem Build a standalone exe folder with Nuitka.
rem Run on a Windows PC with internet access and uv (uv downloads Python 3.10 if needed).
rem Usage: build.bat [script name without .py]   (default: inspect_ui, so double-click works)
cd /d "%~dp0"
set "SCRIPT=%~1"
if "%SCRIPT%"=="" set "SCRIPT=inspect_ui"
if not exist "%SCRIPT%.py" (
  echo %SCRIPT%.py not found.
  goto :error
)
echo Building %SCRIPT%.py
where uv >nul 2>nul
if errorlevel 1 (
  echo ERROR: uv was not found. Install uv and run this again.
  goto :error
)
rem Call the build venv's python directly instead of "uv run" (keeps .venv and pyproject untouched).
set "VENV=.venv-build"
set "PY=%VENV%\Scripts\python.exe"

if not exist "%PY%" (
  echo [1/5] Creating %VENV% with Python 3.10 ...
  uv venv --python 3.10 "%VENV%" || goto :error
) else (
  echo [1/5] Using existing %VENV%
)

echo [2/5] Installing build tools ...
uv pip install --python "%PY%" Nuitka==4.0 pywinauto==0.6.9 pywin32==306 comtypes==1.4.16 || goto :error

rem pywinauto generates comtypes.gen.UIAutomationClient at first use; generate it now so it is bundled.
echo [3/5] Generating UI Automation wrappers ...
"%PY%" -c "import comtypes.client; comtypes.client.GetModule('UIAutomationCore.dll'); import pywinauto.uia_defines" || goto :error

echo [4/5] Building with Nuitka (the first run may take 10+ minutes) ...
rem Nuitka may fail with "getWindowsShortPathName" if the work folder does not exist yet.
if not exist "dist\%SCRIPT%.build" mkdir "dist\%SCRIPT%.build"
"%PY%" -m nuitka --standalone --windows-console-mode=force --assume-yes-for-downloads --remove-output --output-dir=dist --include-package=comtypes.gen --include-package=pywinauto "%SCRIPT%.py" || goto :error

echo [5/5] Zipping ...
powershell -NoProfile -Command "Compress-Archive -Path 'dist\%SCRIPT%.dist' -DestinationPath 'dist\%SCRIPT%.zip' -Force" || goto :error

echo.
echo Done: dist\%SCRIPT%.dist\%SCRIPT%.exe
echo       dist\%SCRIPT%.zip  (copy this folder or zip to the target PC)
pause
exit /b 0

:error
echo.
echo Build failed.
pause
exit /b 1
