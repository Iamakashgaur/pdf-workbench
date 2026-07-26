@echo off
setlocal

echo ============================================================
echo  PDF Workbench - One Click Setup
echo ============================================================
echo.

set "ROOT=%~dp0"
set "VENV=%ROOT%.venv"

:: PY_CMD = base interpreter launcher ("py -3" or "python")
:: PY_EXE = full path to the venv interpreter, once one exists.
:: :RunPy dispatches between them, so nothing has to re-expand a variable
:: inside the same parenthesised block that assigned it.
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

for /f "tokens=*" %%i in ('%PY_CMD% --version 2^>^&1') do echo [OK] %%i

call :CheckDeps
if not errorlevel 1 (
    echo [OK] Existing Python environment already has required packages.
    goto :DepsReady
)

echo.
echo [1/5] Creating virtual environment with access to system packages...
if exist "%VENV%" rmdir /s /q "%VENV%" >nul 2>&1
%PY_CMD% -m venv --system-site-packages "%VENV%"
if errorlevel 1 goto :VenvFailed
if not exist "%VENV%\Scripts\python.exe" goto :VenvFailed

set "PY_EXE=%VENV%\Scripts\python.exe"
echo [OK] Virtual environment ready at %VENV%

call :CheckDeps
if not errorlevel 1 (
    echo [OK] Virtual environment already has required packages.
    goto :DepsReady
)

echo.
echo [2/5] Installing Python dependencies...
call :RunPy -m pip install -r "%ROOT%requirements.txt"
if errorlevel 1 goto :PipFailed
echo [OK] Python packages installed

:DepsReady
echo.
echo [3/5] Checking external system dependencies...
echo.
echo -- Tesseract OCR  (needed only for scanned PDFs) --
where tesseract >nul 2>&1
if errorlevel 1 (
    echo [WARN] Tesseract not found on PATH.
    echo        Digital PDFs still convert fine; scanned pages will be skipped.
    echo        Install from: https://github.com/UB-Mannheim/tesseract/wiki
) else (
    echo [OK] Tesseract found
)

echo.
echo -- Ghostscript  (optional - for the Camelot engine) --
where gswin64c >nul 2>&1
if not errorlevel 1 (
    echo [OK] Ghostscript 64-bit found
) else (
    where gswin32c >nul 2>&1
    if not errorlevel 1 (
        echo [OK] Ghostscript 32-bit found
    ) else (
        echo [INFO] Ghostscript not found - Camelot will be skipped.
        echo        Only needed if you install requirements-optional.txt.
        echo        https://www.ghostscript.com/releases/gsdnld.html
    )
)

echo.
echo -- Java  (optional - for the Tabula engine) --
where java >nul 2>&1
if errorlevel 1 (
    echo [INFO] Java not found - Tabula will be skipped.
    echo        Only needed if you install requirements-optional.txt.
    echo        https://www.java.com/en/download/
) else (
    echo [OK] Java found
)

echo.
echo -- LibreOffice  (optional - only for --to-pdf) --
where soffice >nul 2>&1
if not errorlevel 1 (
    echo [OK] LibreOffice found
) else (
    if exist "%ProgramFiles%\LibreOffice\program\soffice.exe" (
        echo [OK] LibreOffice found
    ) else (
        echo [INFO] LibreOffice not found - --to-pdf will be unavailable.
        echo        Only needed to convert documents INTO PDF.
        echo        https://www.libreoffice.org/download/download-libreoffice/
    )
)

echo.
echo [4/5] Verifying script...
call :RunPy "%ROOT%pdf_to_excel.py" --check-deps
if errorlevel 1 goto :ScriptFailed

echo.
echo ============================================================
echo  Installation complete.
echo  Run convert.bat to start converting PDFs.
echo ============================================================
echo.
pause
exit /b 0

:: ────────────────────────────────────────────────────────────
:: Error exits
:: ────────────────────────────────────────────────────────────

:NoPython
echo [ERROR] No working Python found. Install from https://python.org
echo         If "python" opens the Microsoft Store, disable the alias under
echo         Settings ^> Apps ^> Advanced app settings ^> App execution aliases.
pause
exit /b 1

:VenvFailed
echo [ERROR] Failed to create virtual environment at %VENV%
pause
exit /b 1

:PipFailed
echo [ERROR] Dependency installation failed. Check network access and requirements.txt.
echo [WARN]  You can still run the app if your base Python already has the packages.
pause
exit /b 1

:ScriptFailed
echo [ERROR] Script check failed.
pause
exit /b 1

:: ────────────────────────────────────────────────────────────
:: Helpers
:: ────────────────────────────────────────────────────────────

:CheckDeps
call :RunPy -c "import pandas, openpyxl, streamlit, reportlab, rich" >nul 2>&1
exit /b %errorlevel%

:RunPy
if defined PY_EXE (
    "%PY_EXE%" %*
) else (
    %PY_CMD% %*
)
exit /b %errorlevel%
