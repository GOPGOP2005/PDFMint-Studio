@echo off
setlocal EnableExtensions DisableDelayedExpansion
cd /d "%~dp0"
if errorlevel 1 (
    echo ERROR: Could not open the PDFMint Studio folder.
    pause
    exit /b 1
)

echo.
echo ================================================
echo   PDFMint Studio - Windows Starter
echo ================================================
echo.

set "VENV_PY=%~dp0.venv\Scripts\python.exe"
set "VENV_CFG=%~dp0.venv\pyvenv.cfg"
set "PY_VERSION="
set "USE_PY_LAUNCHER="

rem Prefer Python versions with broad prebuilt-wheel support on Windows.
where py >nul 2>&1
if errorlevel 1 goto TRY_PYTHON_PATH
for %%V in (3.11 3.12 3.13) do (
    py -%%V -c "import sys; assert '.'.join(map(str, sys.version_info[:2])) == '%%V'" >nul 2>&1
    if not errorlevel 1 (
        set "PY_VERSION=%%V"
        set "USE_PY_LAUNCHER=1"
        goto PYTHON_FOUND
    )
)

:TRY_PYTHON_PATH
where python >nul 2>&1
if errorlevel 1 goto PYTHON_FAIL
python -c "import sys; assert sys.version_info[:2] in ((3,11),(3,12),(3,13))" >nul 2>&1
if errorlevel 1 goto PYTHON_FAIL
set "PY_VERSION=PATH"

:PYTHON_FOUND
echo Compatible Python found: %PY_VERSION%

rem Rebuild an incomplete or unsupported virtual environment.
if exist "%VENV_PY%" if exist "%VENV_CFG%" (
    "%VENV_PY%" -c "import sys; assert sys.prefix != sys.base_prefix and sys.version_info[:2] in ((3,11),(3,12),(3,13))" >nul 2>&1
    if not errorlevel 1 goto VENV_READY
)

if exist ".venv\" (
    echo Existing .venv is incomplete or uses an unsupported Python version.
    echo Removing only the generated .venv environment...
    rmdir /s /q ".venv"
    if exist ".venv\" (
        echo ERROR: Cannot remove .venv. Close other Python/CMD windows and retry.
        pause
        exit /b 1
    )
)

echo Creating a clean Python environment...
if defined USE_PY_LAUNCHER (
    py -%PY_VERSION% -m venv ".venv"
) else (
    python -m venv ".venv"
)
if errorlevel 1 goto VENV_FAIL
if not exist "%VENV_PY%" goto VENV_FAIL
if not exist "%VENV_CFG%" goto VENV_FAIL
"%VENV_PY%" -c "import sys; assert sys.prefix != sys.base_prefix and sys.version_info[:2] in ((3,11),(3,12),(3,13))" >nul 2>&1
if errorlevel 1 goto VENV_FAIL

:VENV_READY
echo Using Python:
"%VENV_PY%" --version
echo.
echo Installing or verifying required packages...
"%VENV_PY%" -m pip install -r "requirements.txt"
if errorlevel 1 (
    echo.
    echo ERROR: Python dependencies could not be installed.
    echo Check that Python 3.11-3.13 is installed and internet is connected.
    pause
    exit /b 1
)

rem Pick a free local port so an old/stale server on 8000 cannot hijack the new UI.
set "PORT=8000"
:CHECK_PORT
netstat -ano | findstr /R /C:":%PORT% .*LISTENING" >nul 2>&1
if errorlevel 1 goto PORT_FOUND
set /a PORT+=1
if %PORT% GTR 8010 goto PORT_FAIL
goto CHECK_PORT

:PORT_FOUND
echo.
echo Starting FastAPI on http://127.0.0.1:%PORT%
echo A separate PDFMint FastAPI Server window will stay open while the site is running.
start "PDFMint FastAPI Server" cmd /k ""%VENV_PY%" -m uvicorn app:app --host 127.0.0.1 --port %PORT%"

rem Do not open the browser until the API is actually responding.
set "TRIES=0"
:WAIT_SERVER
set /a TRIES+=1
"%VENV_PY%" -c "import urllib.request,sys; r=urllib.request.urlopen('http://127.0.0.1:%PORT%/api/health',timeout=1); sys.exit(0 if r.status==200 else 1)" >nul 2>&1
if not errorlevel 1 goto SERVER_READY
if %TRIES% GEQ 30 goto SERVER_FAIL
>nul 2>&1 ping 127.0.0.1 -n 2
goto WAIT_SERVER

:SERVER_READY
echo FastAPI health check: OK
echo Opening PDFMint Studio...
start "" "http://127.0.0.1:%PORT%/"
echo.
echo PDFMint Studio is running at:
echo http://127.0.0.1:%PORT%
echo.
echo Keep the separate "PDFMint FastAPI Server" window open while using the website.
pause
exit /b 0

:SERVER_FAIL
echo.
echo ERROR: FastAPI did not answer its health check.
echo Look at the separate "PDFMint FastAPI Server" window for the exact Python error.
echo.
echo You can also run this command manually from this folder:
echo "%VENV_PY%" -m uvicorn app:app --host 127.0.0.1 --port %PORT%
pause
exit /b 1

:PORT_FAIL
echo.
echo ERROR: Ports 8000 through 8010 are already in use.
echo Close old PDFMint/Python server windows and run this starter again.
pause
exit /b 1

:VENV_FAIL
echo.
echo ERROR: Failed to create a working Python environment.
echo Close other Python programs and try again.
pause
exit /b 1

:PYTHON_FAIL
echo.
echo ERROR: A supported Python version was not found.
echo Install Python 3.11 64-bit, then run this file again.
echo To see installed versions, run: py -0p
pause
exit /b 1
