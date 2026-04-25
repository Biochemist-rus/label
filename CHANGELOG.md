# Changelog

## Release 3.5

Release 3.5 marks the current stable state after the 3.1 label-printing tool was extended and normalized.

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

### Preserved from 3.1

- Data Matrix ZPL printing.
- EAC logo rendering.
- GTIN and product name rendering as image text.
- AI(21) extraction and printing.
- Blank Excel barcode rows feeding one empty label via `~PH`.
- Support for blank separator rows both in `--all` mode and for direct `--row` printing.

### Expected project layout

```text
project-folder/
  build_exe_v3.1_download.bat
  menu_print_v3.1_en.bat
  menu_print_v3.1.ini
  print_datamatrix_zpl_graphics_from_excel_v3.1.py
  data.xlsx
  eac.png
  fonts/
    arial.ttf
  dist/
    print_datamatrix_zpl_graphics_from_excel_v3.1.exe
  logs/
```

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
