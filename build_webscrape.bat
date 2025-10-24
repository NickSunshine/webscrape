@echo off
where python >nul 2>&1
if errorlevel 1 (
    echo Python is not installed or not in PATH.
    echo Please install Python and ensure it is added to your system PATH.
    exit /b 1
)

pip show pyinstaller >nul 2>&1
if errorlevel 1 (
    echo PyInstaller not found. Installing...
    pip install pyinstaller
)

python -m PyInstaller --onefile webscrape.py
xcopy cfg dist\cfg /E /I /Y
copy readme.txt dist\readme.txt /Y

powershell -Command "Compress-Archive -Path dist\* -DestinationPath webscrape_1-1.zip -Force"

echo Build complete. cfg folder and readme.txt copied to dist. All dist contents zipped to webscrape.zip.