# Release 3.6

This release updates the active label-printing package from 3.5 to 3.6.

## Main change

For Printer 1 / 60x30 mm labels, the default EAC logo horizontal offset is changed to:

```ini
Printer1EacOffsetX=95
```

The value is set in `menu_print_v3.6.ini` and also in the default variables of `menu_print_v3.6.bat`, so saving the INI from the menu keeps the 3.6 default.

## Active release files

- `print_datamatrix_zpl_graphics_from_excel_v3.6.py`
- `build_exe_v3.6.bat`
- `menu_print_v3.6.bat`
- `menu_print_v3.6.ini`

## Expected local layout

```text
label/
  menu_print_v3.6.bat
  menu_print_v3.6.ini
  build_exe_v3.6.bat
  print_datamatrix_zpl_graphics_from_excel_v3.6.py
  data.xlsx
  eac.png
  fonts/
    arial.ttf
  dist/
    print_datamatrix_zpl_graphics_from_excel_v3.6.exe
  logs/
```

## Build

Run:

```bat
build_exe_v3.6.bat
```

The expected output is:

```text
dist\print_datamatrix_zpl_graphics_from_excel_v3.6.exe
```
