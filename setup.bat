@echo off
echo ============================================================
echo  Watch History Scraper - Setup
echo ============================================================
echo.

:: ── check Python ─────────────────────────────────────────────────────────────
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python not found. Please install Python 3.10+ from https://python.org
    pause
    exit /b 1
)

echo [1/3] Creating virtual environment...
python -m venv .venv
if errorlevel 1 (
    echo ERROR: Failed to create virtual environment.
    pause
    exit /b 1
)

echo [2/3] Installing Python dependencies...
call .venv\Scripts\activate.bat
pip install -r requirements.txt --quiet
if errorlevel 1 (
    echo ERROR: Failed to install dependencies.
    pause
    exit /b 1
)

echo [3/3] Installing Playwright browser (Chromium)...
python -m playwright install chromium
if errorlevel 1 (
    echo ERROR: Failed to install Playwright browser.
    pause
    exit /b 1
)

echo.
echo ============================================================
echo  Setup complete!
echo.
echo  Usage:
echo    run.bat              Run the full pipeline (Prime + Netflix + Process)
echo    run.bat Prime        Scrape Amazon Prime Video watch history
echo    run.bat Netflix      Scrape Netflix viewing history
echo    run.bat Process      Clean and consolidate into watch_history.xlsx
echo.
echo  Add --debug or --fresh to scraper steps, e.g.:
echo    run.bat Prime --debug
echo    run.bat Netflix --fresh
echo ============================================================
pause
