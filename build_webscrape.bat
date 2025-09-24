@echo off
python -m PyInstaller --onefile webscrape.py
xcopy cfg dist\cfg /E /I /Y
echo Build complete. cfg folder copied to dist.