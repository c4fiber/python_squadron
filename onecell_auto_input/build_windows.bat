@echo off
chcp 65001 > nul

echo [0/4] Closing existing process if running...
taskkill /f /im onecell_tool.exe > nul 2>&1

echo [1/4] Installing uv (if not present)...
where uv > nul 2>&1
if errorlevel 1 (
    pip install uv --quiet
)

echo [2/4] Syncing dependencies with uv...
uv sync
if errorlevel 1 (
    echo ERROR: uv sync failed.
    pause
    exit /b 1
)

echo [3/4] Building with PyInstaller...
uv run pyinstaller onecell_tool.spec --clean --noconfirm
if errorlevel 1 (
    echo ERROR: Build failed.
    pause
    exit /b 1
)

echo [4/4] Build complete!
echo.
echo   Output: dist\onecell_tool.exe
echo   Place settings.ini in the same folder as onecell_tool.exe
echo.
pause