# label

Chestniy Znak label-printing tool.

Current baseline release: **3.6**.

## What this tool does

- Prints Data Matrix labels from Excel.
- Adds GTIN, product name and AI(21) text lines.
- Adds EAC logo.
- Supports two printer profiles / label sizes:
  - 60x30 mm;
  - 60x40 mm.
- Supports blank separator rows: an empty barcode cell feeds one blank label via `~PH`.
- Provides a BAT menu launcher.

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
  legacy/
    v3.1/
      README.md
```

All runtime paths in the menu are relative to the folder containing the BAT file.

## File roles

- `menu_print_v3.x.bat` is the main launcher for printing labels. It opens the interactive menu, lets the user select a printer and print mode, and runs the compiled EXE with the selected settings.
- `menu_print_v3.x.ini` stores precision print settings. It is mainly used to adjust layout parameters for label size, printer model, print position, font sizes, EAC logo position, and other printer-specific differences.
- `build_exe_v3.x.bat` is the binary build script. It is needed when the Python source code has changed and a new EXE must be compiled with PyInstaller.
- `print_datamatrix_zpl_graphics_from_excel_v3.x.py` is the Python source code used to generate ZPL and send it to the printer.
- `data.xlsx` is a template Excel file with label data.
- `eac.png` is the EAC logo image used on the label.
- `fonts/` contains fonts used for rendering label text.
- `dist/` contains the compiled EXE after build.
- `logs/` contains print/debug logs created by the menu launcher.

## Active release files

- `print_datamatrix_zpl_graphics_from_excel_v3.6.py`
- `build_exe_v3.6.bat`
- `menu_print_v3.6.bat`
- `menu_print_v3.6.ini`

## Release 3.6 note

For 60x30 mm labels, the default EAC logo horizontal offset is now:

```ini
Printer1EacOffsetX=95
```

## Legacy

Old release 3.1 files were removed from the active project root to avoid confusion. See `legacy/v3.1/README.md`.

## Release notes

See:

- `RELEASE_3.6.md`
- `RELEASE_3.5.md`
- `CHANGELOG.md`

## Release policy

Starting from release 3.5, substantial feature additions should be shipped as separate releases.
