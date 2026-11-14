@echo off
setlocal
cd /d "%~dp0"

set "VENV_PYTHON=%~dp0.venv\Scripts\python.exe"

if not exist "%VENV_PYTHON%" (
    echo [ERROR] Virtual environment not found: "%VENV_PYTHON%"
    echo Please create .venv first. See README.md for setup instructions.
    pause
    exit /b 1
)

echo Starting PPTX Voice Studio Web UI...
echo Press Ctrl+C or close this window to stop the server.
echo.

"%VENV_PYTHON%" "%~dp0web_ui.py" %*

if errorlevel 1 (
    echo.
    echo [ERROR] Web UI exited with code %ERRORLEVEL%.
    pause
    exit /b %ERRORLEVEL%
)

endlocal