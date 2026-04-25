@echo off
cd /d C:\labels

python -m pip install --upgrade pip
python -m pip install pyinstaller pillow openpyxl

python -m PyInstaller ^
  --noconfirm ^
  --clean ^
  --onefile ^
  --name print_datamatrix_zpl_graphics_from_excel_v3.1 ^
  --add-data "C:\labels\eac.png;." ^
  print_datamatrix_zpl_graphics_from_excel_v3.1.py

pause
