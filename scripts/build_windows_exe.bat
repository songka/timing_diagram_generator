@echo off
setlocal enabledelayedexpansion

rem Build a single Windows exe with PyInstaller.
rem IMPORTANT: keep this file encoded as UTF-8 without BOM or ANSI to avoid
rem the "锘緻echo" error in cmd.exe.

set "SCRIPT_DIR=%~dp0"
pushd "%SCRIPT_DIR%.." || exit /b 1

if defined TIMING_DIAGRAM_PYTHON (
    set "PYTHON_EXE=%TIMING_DIAGRAM_PYTHON%"
) else (
    set "PYTHON_EXE=python"
)

"%PYTHON_EXE%" -c "import sys; print(sys.version)" >nul 2>nul
if errorlevel 1 (
    echo Python is not available. Install Python 3.10+ or set TIMING_DIAGRAM_PYTHON to python.exe.
    goto :error
)

if not exist .venv (
    "%PYTHON_EXE%" -m venv .venv
    if errorlevel 1 goto :error
)

if not exist .venv\Scripts\activate.bat (
    echo Virtual environment was not created correctly: .venv\Scripts\activate.bat is missing.
    echo Remove .venv and run this script again, or check that Python venv support is installed.
    goto :error
)

call .venv\Scripts\activate.bat
if errorlevel 1 goto :error

python -m pip install --upgrade pip
if errorlevel 1 goto :error
python -m pip install pyinstaller openpyxl matplotlib
if errorlevel 1 goto :error

if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist timing_diagram_generator.spec del /q timing_diagram_generator.spec

pyinstaller --noconfirm --clean --onefile --windowed ^
    --name "时序图自动生成工具" ^
    timing_diagram_generator105_v133.py
if errorlevel 1 goto :error

if exist "常用动作时间.json" (
    copy /y "常用动作时间.json" "dist\常用动作时间.json" >nul
) else (
    echo Warning: 常用动作时间.json was not found. The app will create a default file next to the exe on first run.
)

echo.
echo Build completed successfully.
echo EXE: dist\时序图自动生成工具.exe
echo Config: dist\常用动作时间.json
echo Keep 常用动作时间.json in the same folder as the exe so common action settings can be edited.
popd
exit /b 0

:error
echo.
echo Build failed. Please check the errors above.
popd
exit /b 1
