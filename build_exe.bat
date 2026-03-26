@echo off
setlocal
cd /d "%~dp0"

echo === Universal Thesis Formatter - Build EXE ===
echo.

set "PY_CMD="
if exist "%LocalAppData%\Programs\Python\Python310\python.exe" set "PY_CMD=%LocalAppData%\Programs\Python\Python310\python.exe"
if not defined PY_CMD if exist "%LocalAppData%\Programs\Python\Python311\python.exe" set "PY_CMD=%LocalAppData%\Programs\Python\Python311\python.exe"
if not defined PY_CMD if exist "%LocalAppData%\Programs\Python\Python312\python.exe" set "PY_CMD=%LocalAppData%\Programs\Python\Python312\python.exe"
if not defined PY_CMD for %%P in (python.exe py.exe) do (
    for /f "delims=" %%I in ('where %%P 2^>nul') do (
        if not defined PY_CMD set "PY_CMD=%%~fI"
    )
)

if not defined PY_CMD (
    echo ERROR: Python not found. Please install Python or add it to PATH.
    goto :end
)

echo Using Python: %PY_CMD%

echo Checking PyInstaller...
"%PY_CMD%" -m PyInstaller --version >nul 2>&1
if errorlevel 1 (
    echo Installing PyInstaller...
    "%PY_CMD%" -m pip install pyinstaller
)

echo Checking PyYAML...
"%PY_CMD%" -c "import yaml" >nul 2>&1
if errorlevel 1 (
    echo Installing PyYAML...
    "%PY_CMD%" -m pip install pyyaml
)

echo Building exe...
"%PY_CMD%" -m PyInstaller --noconfirm thesis-format.spec

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

:end
endlocal
pause
