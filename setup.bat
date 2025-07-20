@echo off
REM Setup script for PowerPoint to Markdown Converter
REM This script creates a virtual environment and installs dependencies

echo Setting up PowerPoint to Markdown Converter...
echo.

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo Error: Python is not installed or not in PATH
    echo Please install Python 3.6+ and try again
    pause
    exit /b 1
)

echo Python found. Creating virtual environment...

REM Create virtual environment
python -m venv pptx_env
if errorlevel 1 (
    echo Error: Failed to create virtual environment
    pause
    exit /b 1
)

echo Virtual environment created successfully.
echo.

echo Activating virtual environment...
call pptx_env\Scripts\activate.bat

echo Installing dependencies...
pip install -r requirements.txt
if errorlevel 1 (
    echo Error: Failed to install dependencies
    pause
    exit /b 1
)

echo.
echo ✅ Setup completed successfully!
echo.
echo To use the converter:
echo 1. Activate the virtual environment: pptx_env\Scripts\activate.bat
echo 2. Run the script: python pptx_to_md.py your_presentation.pptx
echo.
echo Virtual environment is currently active. You can now test the script.
echo.
pause
