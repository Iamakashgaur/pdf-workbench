@echo off
setlocal

echo ============================================================
echo  PDF Workbench - Web App
echo ============================================================
echo.

set "ROOT=%~dp0"
set "VENV=%ROOT%.venv"

:: PY_CMD = base interpreter launcher ("py -3" or "python")
:: PY_EXE = full path to the venv interpreter, when it is usable.
:: :RunPy dispatches between them.
set "PY_CMD="
set "PY_EXE="

where py >nul 2>&1
if not errorlevel 1 set "PY_CMD=py -3"

if not defined PY_CMD (
    where python >nul 2>&1
    if not errorlevel 1 set "PY_CMD=python"
)

if not defined PY_CMD goto :NoPython

:: A bare "python" on Windows is often the Microsoft Store alias stub, which
:: resolves on PATH but fails to run. Confirm the interpreter actually works.
%PY_CMD% --version >nul 2>&1
if errorlevel 1 goto :NoPython

:: Prefer the venv only if it exists AND has the packages.
if not exist "%VENV%\Scripts\python.exe" goto :HaveInterpreter
"%VENV%\Scripts\python.exe" -c "import pandas, openpyxl, streamlit, reportlab, rich" >nul 2>&1
if errorlevel 1 goto :HaveInterpreter
set "PY_EXE=%VENV%\Scripts\python.exe"

:HaveInterpreter
:: Install Streamlit if missing BEFORE opening the browser.
call :RunPy -c "import streamlit" >nul 2>&1
if not errorlevel 1 goto :Launch

echo [INFO] Installing Streamlit...
call :RunPy -m pip install streamlit
if errorlevel 1 goto :StreamlitFailed

:Launch
echo [OK] Starting web converter at http://localhost:8501
echo      Press Ctrl+C in this window to stop the server.
echo.

:: Open browser after a short delay - AFTER streamlit is confirmed installed
start "" cmd /c "timeout /t 4 /nobreak >nul && start http://localhost:8501"

set "STREAMLIT_BROWSER_GATHER_USAGE_STATS=false"
call :RunPy -m streamlit run "%ROOT%app.py" --server.headless true --browser.gatherUsageStats false

pause
exit /b 0

:: ────────────────────────────────────────────────────────────
:: Error exits
:: ────────────────────────────────────────────────────────────

:NoPython
echo [ERROR] No working Python found. Run install.bat first.
echo         If "python" opens the Microsoft Store, disable the alias under
echo         Settings ^> Apps ^> Advanced app settings ^> App execution aliases.
pause
exit /b 1

:StreamlitFailed
echo [ERROR] Streamlit install failed. Check internet connection.
pause
exit /b 1

:: ────────────────────────────────────────────────────────────
:: Helper
:: ────────────────────────────────────────────────────────────

:RunPy
if defined PY_EXE (
    "%PY_EXE%" %*
) else (
    %PY_CMD% %*
)
exit /b %errorlevel%
