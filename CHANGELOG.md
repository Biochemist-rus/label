# Changelog

## Release 3.6

Release 3.6 is the current stable baseline for the label-printing tool.

### Changed

- Bumped active release files from 3.5 to 3.6.
- Added `menu_print_v3.6.ini`.
- Added `menu_print_v3.6.bat`.
- Added `build_exe_v3.6.bat`.
- Added `print_datamatrix_zpl_graphics_from_excel_v3.6.py`.
- For Printer 1 / 60x30 mm labels, changed the default EAC logo horizontal offset:

```ini
Printer1EacOffsetX=95
```

### Expected project layout

```text
project-folder/
  build_exe_v3.6.bat
  menu_print_v3.6.bat
  menu_print_v3.6.ini
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

## Release 3.5

Release 3.5 was the previous stable baseline for the label-printing tool.

### Added

- Bilingual menu in the BAT launcher: English and Russian.
- Language selection on startup with the selected language stored in the INI file.
- Printer-specific layout profiles moved from the BAT file into the INI file.
- Separate INI profiles for 60x30 mm and 60x40 mm labels.
- Bilingual INI comments for printer profiles and future label-size templates.
- Relative project paths instead of hardcoded `C:\labels` paths.
- Relative Excel path: `data.xlsx`.
- Relative font path: `fonts\arial.ttf`.
- Relative EAC logo path: `eac.png`.
- Relative build script paths for PyInstaller resources.
- Logging of executed print commands and EXE output into the `logs` folder.
- Normalized release 3.5 file names.
- Legacy marker for release 3.1 files.

### Preserved from 3.1

- Data Matrix ZPL printing.
- EAC logo rendering.
- GTIN and product name rendering as image text.
- AI(21) extraction and printing.
- Blank Excel barcode rows feeding one empty label via `~PH`.
- Support for blank separator rows both in `--all` mode and for direct `--row` printing.

## Release policy

Starting from release 3.5, substantial feature additions should be released as separate versions.

Suggested versioning rule:

- Patch changes: bug fixes, small text changes, minor parameter tuning.
- Minor releases: new behavior, new menu flow, new printer profile logic, path/layout changes.
- Major releases: incompatible CLI/config changes or a major architecture rewrite.

Recommended workflow:

1. Make changes in source files.
2. Update `CHANGELOG.md`.
3. Update version labels in BAT/README/release notes when needed.
4. Build the EXE.
5. Test at least:
   - one normal label on each printer;
   - one blank separator row on each printer;
   - one `--no-send` debug run.
6. Publish the tested build as the next release.
