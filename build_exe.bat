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
    --hidden-import thesis_format_2024 ^
    --hidden-import thesis_config ^
    --hidden-import thesis_formatter ^
    --hidden-import thesis_formatter._common ^
    --hidden-import thesis_formatter._titles ^
    --hidden-import thesis_formatter.headings ^
    --hidden-import thesis_formatter.page ^
    --hidden-import thesis_formatter.headers ^
    --hidden-import thesis_formatter.toc ^
    --hidden-import thesis_formatter.cover ^
    --hidden-import thesis_formatter.structure ^
    --hidden-import thesis_formatter.references ^
    --hidden-import thesis_formatter.numbering ^
    --hidden-import thesis_formatter.formatter ^
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
