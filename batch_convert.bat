@echo off
REM Batch PowerPoint to Markdown Converter - Windows Batch File
REM Usage: batch_convert.bat "input_folder" ["output_folder"]

setlocal enabledelayedexpansion

echo.
echo =====================================
echo   Batch PowerPoint to Markdown 
echo =====================================
echo.

REM Check if Python is available
python --version >nul 2>&1
if errorlevel 1 (
    echo Error: Python is not installed or not in PATH
    echo Please install Python and try again.
    pause
    exit /b 1
)

REM Check if virtual environment exists and activate it
if exist "venv\Scripts\activate.bat" (
    echo Activating virtual environment...
    call venv\Scripts\activate.bat
    echo.
)

REM Check if converter script exists
if not exist "pptx_to_md.py" (
    echo Error: pptx_to_md.py not found in current directory
    echo Please make sure you're running this from the correct folder.
    pause
    exit /b 1
)

REM Check if batch converter exists
if not exist "batch_convert.py" (
    echo Error: batch_convert.py not found in current directory
    echo Please make sure you're running this from the correct folder.
    pause
    exit /b 1
)

REM Check command line arguments
if "%~1"=="" (
    echo Usage: %0 "input_folder" ["output_folder"]
    echo.
    echo Examples:
    echo   %0 "C:\My Presentations"
    echo   %0 "C:\Lectures" "C:\Output"
    echo   %0 . .\output
    echo.
    pause
    exit /b 1
)

REM Run the batch converter
if "%~2"=="" (
    echo Running: python batch_convert.py "%~1"
    python batch_convert.py "%~1"
) else (
    echo Running: python batch_convert.py "%~1" "%~2"
    python batch_convert.py "%~1" "%~2"
)

set RESULT=%ERRORLEVEL%

echo.
if %RESULT%==0 (
    echo ✓ Batch conversion completed successfully!
) else (
    echo ✗ Batch conversion completed with errors.
)

echo.
pause
exit /b %RESULT%
