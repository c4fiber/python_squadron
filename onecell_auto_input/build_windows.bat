@echo off
chcp 65001 > nul

echo [1/3] Installing dependencies...
pip install pyinstaller openpyxl beautifulsoup4 lxml --quiet
if errorlevel 1 (
    echo ERROR: pip install failed. Check Python/pip installation.
    pause
    exit /b 1
)

echo [2/3] Building with PyInstaller...
pyinstaller onecell_tool.spec --clean --noconfirm
if errorlevel 1 (
    echo ERROR: Build failed.
    pause
    exit /b 1
)

echo [3/3] Build complete!
echo.
echo   Output: dist\onecell_tool.exe
echo   Place settings.ini in the same folder as onecell_tool.exe
echo.
pause