@echo off
setlocal
rem Build a standalone exe folder with Nuitka.
rem Run on a Windows PC with internet access and python.org Python 3.10 (not the Store version).
rem Usage: build.bat [script name without .py]   (default: inspect_ui, so double-click works)
cd /d "%~dp0"
set "SCRIPT=%~1"
if "%SCRIPT%"=="" set "SCRIPT=inspect_ui"
if not exist "%SCRIPT%.py" (
  echo %SCRIPT%.py not found.
  goto :error
)
echo Building %SCRIPT%.py
set "VENV=.venv-build"
set "PY=%VENV%\Scripts\python.exe"

if not exist "%PY%" (
  echo [1/5] Creating %VENV% with Python 3.10 ...
  py -3.10 -m venv "%VENV%" || goto :error
) else (
  echo [1/5] Using existing %VENV%
)

echo [2/5] Installing build tools ...
"%PY%" -m pip install --disable-pip-version-check -q Nuitka==2.4.8 pywinauto==0.6.9 pywin32==306 comtypes==1.4.8 || goto :error

rem pywinauto generates comtypes.gen.UIAutomationClient at first use; generate it now so it is bundled.
echo [3/5] Generating UI Automation wrappers ...
"%PY%" -c "import comtypes.client; comtypes.client.GetModule('UIAutomationCore.dll'); import pywinauto.uia_defines" || goto :error

echo [4/5] Building with Nuitka (the first run may take 10+ minutes) ...
"%PY%" -m nuitka --standalone --assume-yes-for-downloads --remove-output --output-dir=dist --include-package=comtypes.gen --include-package=pywinauto "%SCRIPT%.py" || goto :error

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
