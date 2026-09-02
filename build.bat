@echo off
REM ============================================================
REM  Word Merge Tool - one-click build (Windows only)
REM ============================================================
chcp 65001 >nul
cd /d "%~dp0"

echo [1/3] Installing dependencies...
pip install pywin32 pyinstaller
if errorlevel 1 (
    echo Failed to install dependencies. Please check your Python/pip.
    pause
    exit /b 1
)

echo.
echo [2/3] Building exe (with console, for debugging)...
pyinstaller --onefile --name WordMergeTool --icon app.ico word_merge_tool.py
if errorlevel 1 (
    echo Build failed. See the error above.
    pause
    exit /b 1
)

echo.
echo [3/3] Rebuilding without console window (final release)...
pyinstaller --noconsole --onefile --name WordMergeTool --icon app.ico word_merge_tool.py

echo.
echo ============================================================
echo  Done! Your exe is at:  dist\WordMergeTool.exe
echo  You can copy this single file to any Windows PC (needs Word).
echo ============================================================
pause
