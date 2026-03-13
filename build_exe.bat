@echo off
setlocal
cd /d "%~dp0"

echo === Universal Thesis Formatter - Build EXE ===
echo.

REM Check pyinstaller
py -m PyInstaller --version >nul 2>&1
if errorlevel 1 (
    echo Installing PyInstaller...
    py -m pip install pyinstaller
)

REM Check pyyaml
py -c "import yaml" >nul 2>&1
if errorlevel 1 (
    echo Installing PyYAML...
    py -m pip install pyyaml
)

echo Building exe...
py -m PyInstaller ^
    --onefile ^
    --name thesis-format ^
    --add-data "defaults;defaults" ^
    --hidden-import pythoncom ^
    --hidden-import win32com ^
    --hidden-import win32com.client ^
    --hidden-import yaml ^
    --console ^
    --noconfirm ^
    thesis_format_cli.py

echo.
if exist "dist\thesis-format.exe" (
    echo SUCCESS: dist\thesis-format.exe
    echo.
    echo Usage:
    echo   thesis-format.exe --input "paper.txt"
    echo   thesis-format.exe --input "paper.docx" --output "out.docx"
    echo   thesis-format.exe --input "paper.docx" --config my_school.yaml
    echo   thesis-format.exe --dump-config ^> my_school.yaml
    echo.
    echo Copy pandoc.exe to dist\ for txt/md/tex support.
) else (
    echo BUILD FAILED - check output above.
)

endlocal
pause
