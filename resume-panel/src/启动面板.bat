@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"

set "PY="
if exist "C:\Python314\python.exe" set "PY=C:\Python314\python.exe"
if not defined PY if exist "%LOCALAPPDATA%\Programs\Python\Python313\python.exe" set "PY=%LOCALAPPDATA%\Programs\Python\Python313\python.exe"
if not defined PY if exist "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" set "PY=%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
if not defined PY ( where python >nul 2>nul && set "PY=python" )

if not defined PY (
  echo.
  echo   [ERROR] Python not found ^(need 3.10+^).
  echo   Install Python and run this file again.
  echo.
  pause
  exit /b 1
)

set "PYS="
for %%f in ("%~dp0*.py") do if not defined PYS set "PYS=%%~ff"
if not defined PYS (
  echo   [ERROR] panel script ^(*.py^) not found in this folder.
  pause
  exit /b 1
)

echo.
echo   ============================================================
echo    Enterprise Job Application Panel
echo    Address : http://127.0.0.1:8790
echo    Close this window (or press Ctrl+C) to stop the panel.
echo   ============================================================
echo.
"%PY%" -X utf8 "%PYS%"

echo.
echo   Panel stopped.
pause
