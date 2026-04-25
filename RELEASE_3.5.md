# Release 3.5

This release is the new baseline for the label-printing tool.

## Main changes

- Bilingual BAT menu: English and Russian.
- Startup language selection.
- Language value is saved in `menu_print_v3.1.ini`.
- Printer-specific label layout parameters moved from BAT into INI.
- INI contains profiles for:
  - Printer 1 / 60x30 mm labels;
  - Printer 2 / 60x40 mm labels.
- INI contains bilingual comments and a commented template for future label sizes.
- Removed hardcoded `C:\labels` dependency from menu and build scripts.
- Excel, EAC logo and font paths are now relative to the project folder.
- Fonts are expected in `fonts\`.
- Build script uses relative PyInstaller `--add-data` paths.
- Menu writes command output to `logs\`.

## Kept behavior from 3.1

- AI(21) is printed as the third text line.
- Empty barcode rows feed one blank label via `~PH`.
- Blank row handling works both in `--all` mode and when printing a direct `--row`.

## Recommended local folder structure

```text
label/
  menu_print_v3.1_en.bat
  menu_print_v3.1.ini
  build_exe_v3.1_download.bat
  print_datamatrix_zpl_graphics_from_excel_v3.1.py
  data.xlsx
  eac.png
  fonts/
    arial.ttf
  dist/
    print_datamatrix_zpl_graphics_from_excel_v3.1.exe
  logs/
```

## Release rule going forward

Starting with 3.5, every substantial feature addition should become a separate release.
Bug fixes and small parameter tuning may be patch-level changes.
