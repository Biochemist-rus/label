@echo off
chcp 65001 >nul
setlocal

cd /d "%~dp0"

python -m pip install --upgrade pip
python -m pip install pyinstaller pillow openpyxl

python -m PyInstaller ^
  --noconfirm ^
  --clean ^
  --onefile ^
  --name print_datamatrix_zpl_graphics_from_excel_v3.5 ^
  --add-data "eac.png;." ^
  --add-data "fonts;fonts" ^
  print_datamatrix_zpl_graphics_from_excel_v3.5.py

pause
