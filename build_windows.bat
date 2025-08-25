@echo off
echo Building AMZ Excel Migration Tool for Windows...
echo.

echo Installing Python dependencies...
pip install -r requirements.txt

echo Creating executable with PyInstaller...
pyinstaller --onefile --windowed --name "AMZ_Excel_Migration_Tool" AMZ_excel_migrate.py

echo.
echo Build completed! Check the 'dist' folder for the executable.
pause