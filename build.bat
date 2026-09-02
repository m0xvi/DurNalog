@echo off
setlocal

echo ============================================
echo   Building the "TurNalog" application
echo ============================================
echo.

rem --- 1. Python check ---
where python >nul 2>nul
if errorlevel 1 (
    echo [ERROR] Python not found in PATH.
    echo Install Python from https://www.python.org/downloads/
    echo and enable "Add Python to PATH" during installation.
    pause
    exit /b 1
)
echo [1/5] Python found:
python --version
echo.

rem --- 2. Virtual environment ---
if not exist ".venv" (
    echo [2/5] Creating virtual environment .venv ...
    python -m venv .venv
    if errorlevel 1 (
        echo [ERROR] Failed to create virtual environment.
        pause
        exit /b 1
    )
) else (
    echo [2/5] Virtual environment .venv already exists.
)
call .venv\Scripts\activate.bat
echo.

rem --- 3. Dependencies ---
echo [3/5] Installing dependencies ...
python -m pip install --upgrade pip --quiet
pip install customtkinter pandas numpy openpyxl pyinstaller --quiet
if errorlevel 1 (
    echo [ERROR] Failed to install dependencies. Check internet connection.
    pause
    exit /b 1
)
echo Dependencies installed.
echo.

rem --- 4. customtkinter path (needed for --add-data) ---
echo [4/5] Locating customtkinter package ...
for /f "delims=" %%i in ('python -c "import customtkinter, os; print(os.path.dirname(customtkinter.__file__))"') do set CTK_PATH=%%i
if "%CTK_PATH%"=="" (
    echo [ERROR] customtkinter package not found.
    pause
    exit /b 1
)
echo customtkinter: %CTK_PATH%
echo.

rem --- 5. PyInstaller build ---
echo [5/5] Building exe (may take a few minutes) ...
pyinstaller --noconfirm --clean --onedir --windowed --name "TurNalog" --add-data "%CTK_PATH%;customtkinter" --collect-submodules openpyxl main.py

if errorlevel 1 (
    echo.
    echo [ERROR] Build failed. See messages above.
    pause
    exit /b 1
)

echo.
echo ============================================
echo   DONE!
echo   App: dist\TurNalog\TurNalog.exe
echo   Copy the whole dist\TurNalog folder
echo   to any Windows PC.
echo ============================================
pause
