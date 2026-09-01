@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"

echo ============================================================
echo   GMP Automation System - Local CPU OCR Start
echo ============================================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python is not installed!
    echo Please install Python from https://www.python.org/downloads/
    echo Make sure to check "Add Python to PATH" during installation.
    pause
    exit /b 1
)

REM Create an isolated environment so model dependencies do not affect system Python.
if not exist ".venv\Scripts\python.exe" (
    echo [1/4] Creating Python virtual environment...
    python -m venv .venv
    if errorlevel 1 (
        echo [ERROR] Failed to create the virtual environment.
        pause
        exit /b 1
    )
)
set "PYTHON=%CD%\.venv\Scripts\python.exe"

REM Check if poppler is available
where pdftoppm >nul 2>&1
if %errorlevel% neq 0 (
    echo [WARNING] Poppler is not installed!
    echo PDF processing requires Poppler. Please install it:
    echo   1. Download from: https://github.com/oschwartz10612/poppler-windows/releases
    echo   2. Extract to C:\poppler
    echo   3. Add C:\poppler\Library\bin to your system PATH
    echo.
)

echo [2/4] Installing GMP application dependencies...
"%PYTHON%" -m pip install -r requirements.txt
if errorlevel 1 (
    echo [ERROR] Failed to install GMP application dependencies.
    pause
    exit /b 1
)

echo [3/4] Installing local DeepSeek-OCR CPU dependencies...
"%PYTHON%" -m pip install --index-url https://download.pytorch.org/whl/cpu torch==2.6.0 torchvision==0.21.0
if errorlevel 1 (
    echo [ERROR] Failed to install PyTorch CPU.
    pause
    exit /b 1
)
"%PYTHON%" -m pip install -r requirements-ocr.txt
if errorlevel 1 (
    echo [ERROR] Failed to install DeepSeek-OCR dependencies.
    pause
    exit /b 1
)

"%PYTHON%" -c "import urllib.request; urllib.request.urlopen('http://127.0.0.1:8000/health', timeout=2)" >nul 2>&1
if not errorlevel 1 goto ocr_ready

echo [4/4] Starting the local DeepSeek-OCR CPU model...
echo A separate window will download and load the model on first use.
start "GMP Offline OCR Model" /D "%CD%" "%PYTHON%" -m uvicorn deepseek_ocr.server:app --host 127.0.0.1 --port 8000

:wait_for_ocr
echo Waiting for the local OCR model. This may take a long time on CPU...
timeout /t 10 /nobreak >nul
"%PYTHON%" -c "import urllib.request; urllib.request.urlopen('http://127.0.0.1:8000/health', timeout=5)" >nul 2>&1
if errorlevel 1 goto wait_for_ocr

:ocr_ready
echo.
echo Local OCR model is ready.
echo Starting GMP Automation System with Waitress WSGI...
echo.
echo ============================================================
echo   Open your browser and go to: http://localhost:5002/offline
echo   Keep the GMP Offline OCR Model window open while using OCR.
echo   Press Ctrl+C to stop the GMP web server.
echo ============================================================
echo.

"%CD%\.venv\Scripts\waitress-serve.exe" --host=0.0.0.0 --port=5002 app:app
pause
