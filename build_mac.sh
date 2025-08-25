#!/bin/bash
echo "Building AMZ Excel Migration Tool for macOS..."
echo ""

echo "Installing Python dependencies..."
pip3 install -r requirements.txt

echo "Creating executable with PyInstaller..."
pyinstaller --onefile --windowed --name "AMZ_Excel_Migration_Tool" --icon=icon.icns AMZ_excel_migrate.py

echo ""
echo "Build completed! Check the 'dist' folder for the executable."
