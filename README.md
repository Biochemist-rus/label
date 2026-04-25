# label

Chestniy Znak label-printing tool.

Current baseline release: **3.5**.

## What this tool does

- Prints Data Matrix labels from Excel.
- Adds GTIN, product name and AI(21) text lines.
- Adds EAC logo.
- Supports two printer profiles / label sizes:
  - 60x30 mm;
  - 60x40 mm.
- Supports blank separator rows: an empty barcode cell feeds one blank label via `~PH`.
- Provides a bilingual BAT menu: English / Russian.

## Expected local layout

```text
label/
  menu_print_v3.5.bat
  menu_print_v3.5.ini
  build_exe_v3.5.bat
  print_datamatrix_zpl_graphics_from_excel_v3.5.py
  data.xlsx
  eac.png
  fonts/
    arial.ttf
  dist/
    print_datamatrix_zpl_graphics_from_excel_v3.5.exe
  logs/
  legacy/
    v3.1/
      README.md
```

All runtime paths in the menu are relative to the folder containing the BAT file.

## Active release files

- `print_datamatrix_zpl_graphics_from_excel_v3.5.py`
- `build_exe_v3.5.bat`
- `menu_print_v3.5.bat`
- `menu_print_v3.5.ini`

## Legacy

Old release 3.1 files were removed from the active project root to avoid confusion. See `legacy/v3.1/README.md`.

## Release notes

See:

- `RELEASE_3.5.md`
- `CHANGELOG.md`

## Release policy

Starting from release 3.5, substantial feature additions should be shipped as separate releases.
